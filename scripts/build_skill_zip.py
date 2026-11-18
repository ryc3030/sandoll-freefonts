#!/usr/bin/env python3
"""설치용 스킬 ZIP(dist/free-korean-fonts-ppt.zip)을 만든다.

ZIP 안: free-korean-fonts-ppt/{SKILL.md, scripts/*.py, data/*.json, data/OFL-1.1.txt, README.md}
Claude·ChatGPT·Codex·Gemini CLI 등 Agent Skills(SKILL.md) 표준을 읽는 도구에 그대로 넣을 수 있다.
"""
import re, zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
NAME = "free-korean-fonts-ppt"
OUT = ROOT / "dist" / f"{NAME}.zip"

skill = (ROOT / "skill" / "SKILL.md").read_text("utf-8")
# 0단계: 함께 든 scripts/를 쓰도록 바꾼다(인터넷 없이도 분석·적용 가능)
start = skill.index("## 0. 준비")
end = skill.index("## 1. 덱 준비")
skill = skill[:start] + """## 0. 준비

이 스킬 폴더(이 SKILL.md가 있는 폴더, 아래 `$SKILL_DIR`)에 `scripts/`와 `data/`가 함께 들어 있다.

```bash
SKILL_DIR="<이 SKILL.md가 있는 폴더>"
pip install --break-system-packages -q fonttools lxml 2>/dev/null || pip install -q fonttools lxml
mkdir -p ~/ffskill && cp "$SKILL_DIR"/scripts/*.py ~/ffskill/ && cp -r "$SKILL_DIR"/data ~/data 2>/dev/null; true
```
- 스크립트는 저장소의 최신 데이터(매주 갱신)를 먼저 읽고, 인터넷이 막혀 있으면 함께 든 `data/`를 쓴다.
  `~/ffskill`로 복사해 쓸 때는 `--data "$SKILL_DIR/data"`를 붙이면 함께 든 데이터를 바로 쓴다.
- 폰트 파일은 인터넷으로 받는다. 인터넷이 막힌 환경(일부 챗봇 샌드박스)에서는 사용자가 폰트 파일을 올려 주면
  `fetch_fonts.py --local 파일.ttf`로 적용한다. 분석(`analyze_deck.py`)과 적용(`apply_fonts.py`)은 인터넷 없이도 된다.
- 최신 스크립트: https://github.com/ryc3030/sandoll-freefonts/tree/main/skill/scripts

""" + skill[end:]
OUT.parent.mkdir(exist_ok=True)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr(f"{NAME}/SKILL.md", skill)
    for f in sorted((ROOT / "skill" / "scripts").glob("*.py")):
        z.write(f, f"{NAME}/scripts/{f.name}")
    for f in ["fontfiles.json", "google-fonts.json", "glyph-coverage.json", "agency-fonts.json", "fonts.json", "OFL-1.1.txt"]:
        z.write(ROOT / "data" / f, f"{NAME}/data/{f}")
    z.writestr(f"{NAME}/README.md", (
        "# free-korean-fonts-ppt\n\n피피티를 분석해 어울리는 무료 한글 폰트를 고르고 적용하는 Agent Skill입니다.\n"
        "원본·최신판: https://github.com/ryc3030/sandoll-freefonts\n\n"
        "## 설치\n- Claude(claude.ai·데스크톱): 설정 → 기능(Capabilities) → 스킬에서 이 ZIP 업로드\n"
        "- Claude Code: `~/.claude/skills/`에 압축을 풀기\n"
        "- OpenAI Codex: 프로젝트의 `.agents/skills/`(또는 `.codex/skills/`)에 압축을 풀기\n"
        "- Gemini CLI: `gemini skills install <압축 푼 폴더>` 또는 `~/.gemini/skills/`에 압축을 풀기\n"
        "- ChatGPT(스킬 지원 요금제·데스크톱 앱): 스킬 메뉴에서 폴더/SKILL.md를 가져오기\n"
        "- Gemini 앱(Gem)·스킬이 없는 챗봇: SKILL.md 내용을 Gem·프로젝트 지침에 붙여 넣고, 이 ZIP을 대화에 올려 쓰기\n"
    ))
print(OUT, OUT.stat().st_size)
