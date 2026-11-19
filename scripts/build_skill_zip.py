#!/usr/bin/env python3
"""설치용 스킬 ZIP(dist/free-korean-fonts-ppt.zip)을 만든다.

ZIP 안: free-korean-fonts-ppt/{SKILL.md, scripts/*.py, data/*.json, data/OFL-1.1.txt, README.md, USER-GUIDE-ko.pdf}
사용설명서 PDF는 scripts/build_manual.py로 미리 만들어 docs/에 커밋해 둔 것을 넣는다.
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

이 스킬 폴더(이 SKILL.md가 있는 폴더, 아래 `$SKILL_DIR`)에 `scripts/`와 `data/`가 함께 들어 있다. **복사하지 말고 제자리에서 실행한다.**

```bash
SKILL_DIR="<이 SKILL.md가 있는 폴더>"
python3 -c "import fontTools, lxml, PIL" 2>/dev/null || pip install --break-system-packages -q fonttools lxml pillow 2>/dev/null || pip install -q fonttools lxml pillow
```
- 아래 명령의 스크립트 경로는 모두 `"$SKILL_DIR"/scripts/` 기준이다.
- 스크립트는 저장소의 최신 데이터(매주 갱신)를 먼저 읽고, 인터넷이 막혀 있으면 함께 든 `data/`를 쓴다.
- 폰트 파일은 인터넷으로 받는다. 인터넷이 막힌 환경(일부 챗봇 샌드박스, Codex 기본 샌드박스)에서는 사용자가 폰트 파일을 올려 주면
  `fetch_fonts.py --local 파일.ttf`로 적용한다. 분석(`analyze_deck.py`)과 적용(`apply_fonts.py`)은 인터넷 없이도 된다.
- PDF 내보내기(`export_pdf.py`)에는 LibreOffice(`soffice`)와 poppler(`pdffonts`, `pdftoppm`)가 필요하다. 없으면 바뀐 피피티만 넘기고
  PowerPoint에서 PDF로 저장하도록 안내한다.
- 최신 스크립트: https://github.com/ryc3030/sandoll-freefonts/tree/main/skill/scripts

""" + skill[end:]
skill = skill.replace("~/ffskill/", '"$SKILL_DIR"/scripts/')
OUT.parent.mkdir(exist_ok=True)


def put(z, arc, data):
    """고정 시각으로 기록해 내용이 같으면 ZIP도 똑같이 나오게 한다(매주 불필요한 커밋 방지)."""
    zi = zipfile.ZipInfo(arc, (2020, 1, 1, 0, 0, 0))
    zi.compress_type = zipfile.ZIP_STORED  # 압축 라이브러리 판에 따라 바이트가 달라지지 않게 무압축
    zi.external_attr = 0o644 << 16
    z.writestr(zi, data)


with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    put(z, f"{NAME}/SKILL.md", skill)
    for f in sorted((ROOT / "skill" / "scripts").glob("*.py")):
        put(z, f"{NAME}/scripts/{f.name}", f.read_bytes())
    for f in ["fontfiles.json", "google-fonts.json", "glyph-coverage.json", "agency-fonts.json", "design-tokens.json", "fonts.json", "OFL-1.1.txt"]:
        put(z, f"{NAME}/data/{f}", (ROOT / "data" / f).read_bytes())
    put(z, f"{NAME}/USER-GUIDE-ko.pdf", (ROOT / "docs" / "USER-GUIDE-ko.pdf").read_bytes())
    put(z, f"{NAME}/README.md", (
        "# free-korean-fonts-ppt\n\n피피티를 분석해 어울리는 무료 한글 폰트를 고르고 적용하는 Agent Skill입니다.\n"
        "원본·최신판: https://github.com/ryc3030/sandoll-freefonts\n\n"
        "**사용설명서와 전후 예시: 같은 폴더의 USER-GUIDE-ko.pdf** (설치·요청하는 법·결과물·폰트 고르는 기준·예시 3종·FAQ)\n\n"
        "## 설치\n- Claude(claude.ai·데스크톱): 설정 → 기능(Capabilities) → 스킬에서 이 ZIP 업로드\n"
        "- Claude Code: `~/.claude/skills/`에 압축을 풀기\n"
        "- OpenAI Codex: 프로젝트의 `.agents/skills/`(또는 `.codex/skills/`)에 압축을 풀기\n"
        "- Gemini CLI: `gemini skills install <압축 푼 폴더>` 또는 `~/.gemini/skills/`에 압축을 풀기\n"
        "- ChatGPT(스킬 지원 요금제·데스크톱 앱): 스킬 메뉴에서 폴더/SKILL.md를 가져오기\n"
        "- Gemini 앱(Gem)·스킬이 없는 챗봇: SKILL.md 내용을 Gem·프로젝트 지침에 붙여 넣고, 이 ZIP을 대화에 올려 쓰기\n\n"
        "## 쓰는 법\n피피티(.pptx)를 대화에 올리고 이렇게 말하면 됩니다.\n"
        "- 이 피피티에 어울리는 무료 폰트로 바꿔 줘\n"
        "- 폰트는 바꾸지 말고 어떤 폰트가 어울릴지 분석만 해 줘\n"
        "- 제목은 G마켓 산스 Bold, 본문은 프리텐다드로 적용해 줘\n"
        "- 공모전 제출용이야. 폰트 라이선스 내역서도 같이 만들어 줘\n"
        "- 받는 사람 컴퓨터에 폰트가 없어도 똑같이 보이게 해 줘\n\n"
        "결과: 바뀐 피피티 + fonts.zip(폰트와 라이선스) + 진단 요약.\n"
        "사용방법과 전후 예시: https://ryc3030.github.io/sandoll-freefonts/guide.html\n"
    ))
print(OUT, OUT.stat().st_size)
