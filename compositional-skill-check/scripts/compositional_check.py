#!/usr/bin/env python3
"""
compositional_check.py — スキル組み合わせリスク検査の参照実装

発動・インストールしようとする複数スキルの能力(capability)を抽出し、
個別には安全でも組み合わせると危険になる「組み合わせリスク」を検出する。

- 能力抽出: コード兆候 + 文章兆候の regex、前後40字の否定ウィンドウで偽陽性を抑制
- Mode 1 (集合レベルゲート): 新規に開くパターンのみを判定し BLOCK/WARN/ALLOW を返す
- Mode 2 (ペアレベル証拠): 合成によってのみ生じるリスクを根拠付きで列挙

重要: 出力は危険な「能力面が到達可能になるか」を測るものであり、
エージェントが実際に危険な動作を「実行するか」を予測するものではない。
本チェックは per-skill スキャン・ランタイム隔離・整合性の補完であって代替ではない。

検出規則は再現率重視の出発点。ground truth ではないので運用に合わせて増補すること。

使い方:
    # ディレクトリ単位（各サブディレクトリ = 1スキル、SKILL.md と同梱コードを読む）
    python compositional_check.py path/to/skill_a path/to/skill_b ...
    # ベースライン(既存有効スキル)を指定して差分判定
    python compositional_check.py --baseline path/to/active1 -- path/to/new_skill
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from dataclasses import dataclass, field
from itertools import combinations
from pathlib import Path

# ──────────────────────────────────────────────────────────────────────────
# 能力モデル（8 capabilities）
# ──────────────────────────────────────────────────────────────────────────
CAPS = ["FR", "FW", "NO", "NI", "ER", "PS", "CA", "SI"]
CAP_NAMES = {
    "FR": "ファイル読み取り (file read)",
    "FW": "ファイル書き込み (file write)",
    "NO": "外向き通信 (network out)",
    "NI": "内向き通信 (network in)",
    "ER": "環境変数読み取り (env read)",
    "PS": "プロセス起動 (process spawn)",
    "CA": "認証情報アクセス (credential access)",
    "SI": "システム情報 (system info)",
}

# 各能力の検出規則。
#   "code"  : コード兆候(実装の存在を示唆)
#   "prose" : 文章兆候(散文での言及。偽陽性が出やすいので根拠で区別)
# net_out(NO) は load-bearing capability。最も厚く作り込む。
PATTERNS_CAP: dict[str, dict[str, list[str]]] = {
    "NO": {
        "code": [r"https?://", r"\brequests\.", r"\burllib", r"\bhttpx", r"\bcurl\b",
                 r"\bwget\b", r"fetch\s*\(", r"\baxios\b", r"socket\.connect",
                 r"http\.client", r"webhook", r"urlopen"],
        "prose": [r"アップロード", r"外部(に|へ)?送", r"送信", r"API(を|の)?(呼|叩|コール|call)",
                  r"exfiltrat", r"upload", r"call (an |the )?api", r"post (the |data)"],
    },
    "FR": {
        "code": [r"open\s*\([^)]*['\"][rR]['\"]", r"\.read\s*\(", r"read_text",
                 r"readFile", r"fs\.read", r"\bglob\b", r"os\.listdir", r"pathlib",
                 r"\bcat\s+[\w./]"],
        "prose": [r"読み(取|込)", r"ファイルを読", r"read (the )?file", r"load (the )?file"],
    },
    "FW": {
        "code": [r"open\s*\([^)]*['\"][waWA]['\"]", r"\.write\s*\(", r"write_text",
                 r"writeFile", r"fs\.write(File)?", r"shutil\.(copy|move)", r"os\.remove",
                 r"\btee\b"],
        "prose": [r"書き(込|出)", r"保存(する|します)?", r"ファイル(を)?作成",
                  r"write (to )?(a |the )?file", r"save (to )?(a |the )?file"],
    },
    "NI": {
        "code": [r"\.listen\s*\(", r"\.bind\s*\(", r"socketserver", r"http\.server",
                 r"Flask\s*\(", r"app\.route", r"\bserve\b", r"\bngrok\b", r"FastAPI"],
        "prose": [r"待ち受け", r"サーバ(ー)?を(立|起|建)", r"listen (on |for )", r"open (a )?port"],
    },
    "ER": {
        "code": [r"os\.environ", r"getenv", r"process\.env", r"\bdotenv\b",
                 r"\$\{?[A-Z_][A-Z0-9_]+\}?"],
        "prose": [r"環境変数", r"env(ironment)? var(iable)?"],
    },
    "PS": {
        "code": [r"subprocess", r"os\.system", r"\bPopen\b", r"child_process",
                 r"\bspawn\s*\(", r"\bexec\s*\(", r"\beval\s*\(", r"sh\s+-c", r"/bin/(ba)?sh"],
        "prose": [r"コマンド(を)?(実行|起動)", r"シェル(を)?(実行|起動|呼)", r"プロセスを起動",
                  r"run (a )?(shell )?command", r"spawn (a )?process"],
    },
    "CA": {
        "code": [r"api[_-]?key", r"\btoken\b", r"\bpassword\b", r"\bsecret\b",
                 r"credential", r"\boauth\b", r"\bbearer\b", r"private[_-]?key",
                 r"\.aws/credentials", r"access[_-]?key"],
        "prose": [r"認証情報", r"API\s*キー", r"トークン", r"パスワード", r"資格情報",
                  r"credential", r"api key"],
    },
    "SI": {
        "code": [r"platform\.", r"\buname\b", r"os\.name", r"\bpsutil\b",
                 r"socket\.gethostname", r"\bwhoami\b", r"\bifconfig\b", r"ip\s+addr"],
        "prose": [r"システム情報", r"ホスト名", r"OS(の)?情報", r"system info", r"hostname"],
    },
}

# 否定語（日英）。一致箇所の前後ウィンドウにあれば、その一致を無効化する。
NEGATORS = [
    "不要", "なし", "しない", "せず", "ない", "禁止", "持たない", "使わない", "せぬ",
    "no", "not", "without", "never", "n't", "none", "avoid", "disable", "do not",
]
NEG_WINDOW = 40  # 前後の文字数（出典研究に合わせる）


# ──────────────────────────────────────────────────────────────────────────
# 禁止パターン（10パターン）と深刻度
#   実証率は「フラグが本物だった監査上の目安」。BLOCK/WARN 調整と優先順位付けに使う。
# ──────────────────────────────────────────────────────────────────────────
@dataclass(frozen=True)
class Pattern:
    key: str
    name: str
    caps: frozenset[str]
    severity: str          # CRITICAL / HIGH / MEDIUM
    empirical_validity: float  # おおよその実証率(0-1)


FORBIDDEN: list[Pattern] = [
    Pattern("data_exfil",   "データ持ち出し",         frozenset({"FR", "NO"}),       "HIGH",     0.10),
    Pattern("cred_theft",   "認証情報窃取",           frozenset({"CA", "NO"}),       "CRITICAL", 0.25),
    Pattern("lethal_tri",   "Lethal Trifecta",        frozenset({"CA", "FR", "NO"}), "CRITICAL", 0.25),
    Pattern("priv_esc",     "権限昇格",               frozenset({"ER", "PS"}),       "HIGH",     0.45),
    Pattern("reverse_shell","リバースシェル",         frozenset({"NI", "PS"}),       "CRITICAL", 0.05),
    Pattern("backdoor",     "永続バックドア",         frozenset({"FW", "PS"}),       "HIGH",     0.25),
    Pattern("lateral",      "横展開",                 frozenset({"CA", "NO", "PS"}), "CRITICAL", 0.10),
    Pattern("env_exfil",    "環境変数持ち出し",       frozenset({"ER", "NO"}),       "HIGH",     0.25),
    Pattern("sys_recon",    "システム偵察持ち出し",   frozenset({"SI", "NO"}),       "MEDIUM",   0.15),
    Pattern("dropper",      "ドロッパー",             frozenset({"NO", "FW", "PS"}), "CRITICAL", 0.10),
]


# ──────────────────────────────────────────────────────────────────────────
# 能力抽出
# ──────────────────────────────────────────────────────────────────────────
@dataclass
class Evidence:
    cap: str
    kind: str       # "code" | "prose"
    file: str
    pattern: str    # 一致した regex
    snippet: str    # 前後を含む抜粋


def _negated(text: str, start: int, end: int) -> bool:
    """一致箇所の前後 NEG_WINDOW 字に否定語があれば True。"""
    lo = max(0, start - NEG_WINDOW)
    hi = min(len(text), end + NEG_WINDOW)
    window = text[lo:hi].lower()
    return any(neg.lower() in window for neg in NEGATORS)


def extract_capabilities(text: str, filename: str) -> tuple[set[str], list[Evidence]]:
    """1ファイルのテキストから能力集合と根拠を抽出。否定文脈の一致は捨てる。"""
    found: set[str] = set()
    evidence: list[Evidence] = []
    for cap, kinds in PATTERNS_CAP.items():
        for kind, regexes in kinds.items():
            for rgx in regexes:
                for m in re.finditer(rgx, text, flags=re.IGNORECASE):
                    if _negated(text, m.start(), m.end()):
                        continue  # 「APIキー不要」等は無効化
                    found.add(cap)
                    lo = max(0, m.start() - 30)
                    hi = min(len(text), m.end() + 30)
                    snippet = text[lo:hi].replace("\n", " ").strip()
                    evidence.append(Evidence(cap, kind, filename, rgx, snippet))
                    break  # 同じ regex の最初の有効一致のみ根拠化（ノイズ抑制）
    return found, evidence


@dataclass
class Skill:
    name: str
    caps: set[str] = field(default_factory=set)
    evidence: list[Evidence] = field(default_factory=list)

    def caps_str(self) -> str:
        return "{" + ", ".join(sorted(self.caps)) + "}"


CODE_EXTS = {".py", ".js", ".ts", ".tsx", ".jsx", ".sh", ".bash", ".rb", ".go", ".md"}


def load_skill(path: str) -> Skill:
    """スキルを読み込む。ディレクトリなら SKILL.md と同梱コードを走査、
    ファイルなら単体を走査する。"""
    p = Path(path)
    skill = Skill(name=p.name)
    files: list[Path] = []
    if p.is_dir():
        for root, _dirs, names in os.walk(p):
            for n in names:
                fp = Path(root) / n
                if fp.suffix.lower() in CODE_EXTS:
                    files.append(fp)
    elif p.is_file():
        files.append(p)
    else:
        raise FileNotFoundError(path)

    for fp in files:
        try:
            text = fp.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue
        caps, ev = extract_capabilities(text, str(fp.relative_to(p) if p.is_dir() else fp.name))
        skill.caps |= caps
        skill.evidence.extend(ev)
    return skill


# ──────────────────────────────────────────────────────────────────────────
# Mode 1: 集合レベルゲート（新規に開くパターンのみ判定）
# ──────────────────────────────────────────────────────────────────────────
@dataclass
class GateResult:
    verdict: str               # BLOCK / WARN / ALLOW
    new_violations: list[Pattern]
    union_caps: set[str]
    baseline_caps: set[str]


def mode1_gate(new_skills: list[Skill], baseline: list[Skill] | None = None) -> GateResult:
    base_caps: set[str] = set()
    for s in (baseline or []):
        base_caps |= s.caps
    union_caps = set(base_caps)
    for s in new_skills:
        union_caps |= s.caps

    # 和集合では立つが、ベースライン単独では立たなかったパターンのみ
    new_v = [f for f in FORBIDDEN
             if f.caps <= union_caps and not (f.caps <= base_caps)]

    if any(f.severity == "CRITICAL" for f in new_v):
        verdict = "BLOCK"
    elif new_v:
        verdict = "WARN"
    else:
        verdict = "ALLOW"
    return GateResult(verdict, new_v, union_caps, base_caps)


# ──────────────────────────────────────────────────────────────────────────
# Mode 2: ペアレベル証拠（合成によってのみ生じるリスクのみ）
# ──────────────────────────────────────────────────────────────────────────
@dataclass
class PairFinding:
    a: str
    b: str
    pattern: Pattern
    contributors: dict[str, str]  # cap -> どちらのスキルが提供したか


def mode2_evidence(skills: list[Skill]) -> list[PairFinding]:
    findings: list[PairFinding] = []
    for sa, sb in combinations(skills, 2):
        union = sa.caps | sb.caps
        for f in FORBIDDEN:
            if not (f.caps <= union):
                continue
            # 各スキル単独でも立つなら「合成によってのみ」ではない → 除外
            if f.caps <= sa.caps or f.caps <= sb.caps:
                continue
            contributors: dict[str, str] = {}
            for c in sorted(f.caps):
                who = []
                if c in sa.caps:
                    who.append(sa.name)
                if c in sb.caps:
                    who.append(sb.name)
                contributors[c] = " & ".join(who)
            findings.append(PairFinding(sa.name, sb.name, f, contributors))
    return findings


# ──────────────────────────────────────────────────────────────────────────
# レポート出力
# ──────────────────────────────────────────────────────────────────────────
def render(new_skills: list[Skill], baseline: list[Skill] | None) -> str:
    out: list[str] = []
    out.append("=" * 70)
    out.append("スキル組み合わせリスク検査")
    out.append("=" * 70)
    out.append("")
    if baseline:
        out.append("[ベースライン(既存有効スキル)]")
        for s in baseline:
            out.append(f"  - {s.name}: {s.caps_str()}")
        out.append("")
    out.append("[発動するスキル]")
    for s in new_skills:
        out.append(f"  - {s.name}: {s.caps_str()}")
    out.append("")

    # Mode 1
    g = mode1_gate(new_skills, baseline)
    out.append("-" * 70)
    out.append(f"Mode 1 判定: {g.verdict}")
    out.append("-" * 70)
    out.append(f"  発動後の総能力: " + "{" + ", ".join(sorted(g.union_caps)) + "}")
    if g.new_violations:
        out.append("  この発動で新たに開く違反パターン:")
        # 深刻度→実証率 の順で重要なものを上に
        sev_rank = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2}
        for f in sorted(g.new_violations,
                        key=lambda x: (sev_rank[x.severity], -x.empirical_validity)):
            caps = " ∧ ".join(sorted(f.caps))
            out.append(f"    [{f.severity:8}] {f.name} ({caps}) "
                       f"／実証率の目安 {int(f.empirical_validity*100)}%")
    else:
        out.append("  新たに開く違反パターンなし。")
    out.append("")

    # Mode 2
    findings = mode2_evidence(new_skills)
    out.append("-" * 70)
    out.append("Mode 2 ペアレベル証拠（合成によってのみ生じるリスク）")
    out.append("-" * 70)
    if not findings:
        out.append("  合成固有のペアリスクなし。")
    else:
        # ペアごとにまとめる
        by_pair: dict[tuple[str, str], list[PairFinding]] = {}
        for fnd in findings:
            by_pair.setdefault((fnd.a, fnd.b), []).append(fnd)
        for (a, b), fs in by_pair.items():
            out.append(f"  ◆ {a}  ×  {b}")
            for fnd in sorted(fs, key=lambda x: -x.pattern.empirical_validity):
                caps = " ∧ ".join(sorted(fnd.pattern.caps))
                out.append(f"      - {fnd.pattern.name} [{fnd.pattern.severity}] ({caps})")
                for c, who in fnd.contributors.items():
                    out.append(f"          {c} ({CAP_NAMES[c]}) ← {who}")
    out.append("")
    out.append("-" * 70)
    out.append("注意: 本結果は『危険な能力面が到達可能か』を示すものであり、")
    out.append("      エージェントが実際にそれを実行するかの予測ではない。")
    out.append("      フラグは人による確認を要するトリアージ信号として扱うこと。")
    out.append("-" * 70)
    return "\n".join(out)


def main() -> int:
    parser = argparse.ArgumentParser(description="スキル組み合わせリスク検査")
    parser.add_argument("--baseline", nargs="*", default=[],
                        help="既存有効スキルのパス（差分判定の基準）")
    parser.add_argument("skills", nargs="+", help="発動するスキルのパス")
    args = parser.parse_args()

    try:
        baseline = [load_skill(p) for p in args.baseline] if args.baseline else None
        new_skills = [load_skill(p) for p in args.skills]
    except FileNotFoundError as e:
        print(f"パスが見つかりません: {e}", file=sys.stderr)
        return 2

    print(render(new_skills, baseline))
    g = mode1_gate(new_skills, baseline)
    # 終了コード: BLOCK=1, WARN=0(注意のみ), ALLOW=0
    return 1 if g.verdict == "BLOCK" else 0


if __name__ == "__main__":
    raise SystemExit(main())
