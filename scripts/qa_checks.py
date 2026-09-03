# -*- coding: utf-8 -*-
"""jiaoan-skill 包质量门禁（CI/自动同步前运行）。

纯标准库实现，可在本地与 GitHub Actions 通用。
检查项：
  1. skill.json 为合法 JSON 且含 version
  2. SKILL.md 头部 `> 版本: vX.Y.Z` 与 skill.json version 一致
  3. CHANGELOG.md 最新版本行与 skill.json version 一致
  4. README.md 标题版本与 skill.json version 一致
  5. SKILL.md frontmatter（name/description）完整
  6. scripts/*.py 语法编译通过（py_compile，无需 python-docx）
  7. 必需资源存在：templates/ 含基底 docx、scripts/dump_full.py、verify_format.py、versions/CHANGELOG.md
  8. 敏感信息扫描（ghp_ token / GITHUB_TOKEN= / password= / Authorization）——文本文件仅扫描暂存变更内容
用法：
  python scripts/qa_checks.py            # 检查工作区
  python scripts/qa_checks.py --staged   # 仅检查 git 暂存区文本变更（同步前门禁）
退出码：0=通过，1=失败
"""
import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FAILED = []


def check(name, ok, detail=""):
    tag = "PASS" if ok else "FAIL"
    print(f"[{tag}] {name}" + (f" — {detail}" if detail else ""))
    if not ok:
        FAILED.append(name)


def load_skill_json():
    p = ROOT / "skill.json"
    if not p.exists():
        check("skill.json 存在", False, "缺失")
        return None
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        check("skill.json 合法 JSON 且含 version", isinstance(data.get("version"), str) and len(data["version"]) > 0, f"version={data.get('version')}")
        return data
    except Exception as e:
        check("skill.json 合法 JSON", False, str(e))
        return None


def head_version():
    m = re.search(r"^>\s*版本:\s*(v[\d.]+)", ROOT.joinpath("SKILL.md").read_text(encoding="utf-8"), re.M)
    return m.group(1) if m else None


def changelog_latest():
    txt = ROOT.joinpath("versions", "CHANGELOG.md").read_text(encoding="utf-8")
    for line in txt.splitlines():
        m = re.match(r"^\|\s*(v[\d.]+)\s*\|", line)
        if m:
            return m.group(1)
    return None


def readme_version():
    txt = ROOT.joinpath("README.md").read_text(encoding="utf-8")
    m = re.search(r"jiaoan-skill\s*[）)]?\s*v([\d.]+)", txt)
    return ("v" + m.group(1)) if m else None


def main():
    data = load_skill_json()
    ver = data.get("version", "?") if data else "?"
    want = "v" + ver

    sv = head_version()
    check("SKILL.md 头部版本一致", sv == want, f"SKILL.md={sv} skill.json={want}")

    cl = changelog_latest()
    check("CHANGELOG 最新版本一致", cl == want, f"CHANGELOG={cl} skill.json={want}")

    rv = readme_version()
    check("README 标题版本一致", rv == want, f"README={rv} skill.json={want}")

    # frontmatter
    sk = ROOT.joinpath("SKILL.md").read_text(encoding="utf-8")
    fm = re.match(r"^---\n(.*?)\n---", sk, re.S)
    fm_ok = bool(fm and re.search(r"^name:\s*\S+", fm.group(1), re.M) and re.search(r"^description:", fm.group(1), re.M))
    check("SKILL.md frontmatter 完整", fm_ok, "需要 name + description")

    # 脚本语法
    for py in sorted((ROOT / "scripts").glob("*.py")):
        r = subprocess.run([sys.executable, "-m", "py_compile", str(py)], capture_output=True)
        check(f"脚本语法 {py.name}", r.returncode == 0, "" if r.returncode == 0 else r.stderr.decode("utf-8", "ignore")[:200])

    # 必需资源
    need = [
        ROOT / "templates" / "教案模板基底_齐工程教2023-11号.docx",
        ROOT / "scripts" / "dump_full.py",
        ROOT / "scripts" / "verify_format.py",
        ROOT / "versions" / "CHANGELOG.md",
    ]
    for p in need:
        check(f"资源存在 {p.relative_to(ROOT)}", p.exists())

    # 敏感扫描（暂存文本变更）
    leak_pat = re.compile(r"(ghp_[A-Za-z0-9]{20,}|GITHUB_TOKEN\s*=|password\s*[:=]\s*\S+|Authorization\s*[:=]\s*[Tt]oken\s+\S+)")
    targets = []
    if "--staged" in sys.argv:
        r = subprocess.run(["git", "diff", "--cached", "--name-only", "--diff-filter=ACM"], capture_output=True, text=True)
        targets = [ROOT / n for n in r.stdout.splitlines() if n and Path(n).suffix in (".md", ".json", ".py", ".yml", ".yaml")]
    else:
        for pat in ("SKILL.md", "README.md", "skill.json"):
            targets.append(ROOT / pat)
        targets += list((ROOT / "scripts").glob("*.py"))
    leaks = []
    for p in targets:
        # 排除 QA 脚本自身：其中的检测模式为字面量而非真实凭证
        if p.name == "qa_checks.py":
            continue
        if p.exists():
            try:
                txt = p.read_text(encoding="utf-8")
            except Exception:
                continue
            if leak_pat.search(txt):
                leaks.append(str(p.relative_to(ROOT)))
    check("无敏感信息泄漏", not leaks, "、".join(leaks) if leaks else "已扫描文本文件")

    print("-" * 40)
    if FAILED:
        print(f"QA FAILED: {len(FAILED)} 项未通过")
        sys.exit(1)
    print(f"QA PASSED — jiaoan-skill v{ver} 质量门禁全绿")


if __name__ == "__main__":
    main()
