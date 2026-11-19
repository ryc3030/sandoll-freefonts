# 무료폰트 색인

[산돌구름 무료폰트](https://www.sandollcloud.com/freefonts)와 [Google Fonts 한국어 폰트](https://fonts.google.com/?subset=korean)를 매주 자동으로 모아 보여 주는 정적 웹앱입니다. GitHub Pages에서 서버 없이 돌아갑니다.

- **검색·필터**: 이름·제작사·분류·분위기 검색, OFL / 상업용 / 임베딩 / 로고·CI / 신규 / 스테디 / 실시간 미리보기 / 찜
- **미리보기**: 공개 웹폰트가 있는 폰트는 입력한 문구로 바로 보여 주고, 나머지는 제작사 페이지로 안내(다른 사이트의 견본 이미지는 쓰지 않음)
- **상세**: 폰트 정보, 산돌구름 라이선스 요약표, 라이선스 원문 안내, 웹폰트 연결 코드 복사
- **다운로드**: 제작사 공식 배포처(없으면 산돌구름 상세 페이지)로 연결
- **자동 갱신**: GitHub Actions가 매주 월요일 목록을 다시 모으고, 바뀐 것이 있을 때만 커밋

**사용방법과 전후 예시: https://ryc3030.github.io/sandoll-freefonts/guide.html** · [사용설명서 PDF](docs/USER-GUIDE-ko.pdf)(설치 ZIP에도 들어 있음)

## 스킬 설치 (Claude · ChatGPT · Codex · Gemini)

**설치 ZIP: [free-korean-fonts-ppt.zip](https://github.com/ryc3030/sandoll-freefonts/raw/main/dist/free-korean-fonts-ppt.zip)** (매주 데이터와 함께 다시 만들어짐)

Agent Skills 표준(SKILL.md)으로 만들어 아래 도구에서 그대로 쓸 수 있습니다. ZIP에 스크립트와 데이터가 함께 들어 있어 인터넷이 막힌 환경에서도 분석과 적용은 되고, 폰트를 내려받을 때만 인터넷이 필요합니다.

| 도구 | 설치 | 확인 |
|---|---|---|
| Claude (claude.ai·데스크톱) | 설정 → 기능 → 스킬에서 ZIP 업로드 | 동작 확인 |
| Claude Code | `~/.claude/skills/`에 압축 풀기 | — |
| OpenAI Codex (CLI·앱) | 프로젝트 `.agents/skills/`에 압축 풀기 | 스킬 인식 확인(Codex CLI 0.170) |
| Gemini CLI | `gemini skills install <압축 푼 폴더>` | 설치·인식 확인(Gemini CLI 0.68) |
| ChatGPT (스킬 지원 요금제·데스크톱 앱) | 스킬 메뉴에서 가져오기 | 직접 시험 못 함 |
| Gemini 앱(Gem)·스킬 없는 챗봇 | SKILL.md를 Gem·프로젝트 지침에 붙여 넣고 ZIP을 대화에 올림 | 직접 시험 못 함 |

## 스킬: PPT 분석 → 무료 한글 폰트 자동 적용

`skill/` 폴더가 스킬 본체입니다. "이 피피티에 어울리는 무료 폰트 적용해 줘"라고 하면 AI가 덱의 분야·글의 문체·밀도·색감을 분석해 폰트 짝을 고릅니다.
이어서 OFL 폰트(산돌 OFL 23종, Google Fonts 38종)를 공개 배포처에서 받아 PPTX의 테마·마스터·슬라이드·표·차트 글꼴을 한 번에 바꾸고,
PDF로 내보내 확인한 뒤 **바뀐 피피티, PDF, 폰트 묶음(라이선스 포함)**을 함께 넘깁니다.

- `skill/SKILL.md` — 스킬 본문
- `skill/scripts/analyze_deck.py` — PPTX 진단(발행 주체·분야·문체·구조·시각·기호 등 30여 개 변수)과 폰트 짝 추천. 공공 덱의 기본값은 프리텐다드이고, 내용 성격(축제·전시·아동·학술 등)이 뚜렷하면 그에 맞는 서체를 고르며, 기관 전용 서체(서울·경기·부산·교육청 등)는 후보로 제시
- `scripts/glyph_coverage.py` → `data/glyph-coverage.json` — 폰트별 빠진 기호(□ ※ Ⅰ 등)와 한글 글자 수. 기호가 빠진 폰트는 추천에서 감점
- `skill/scripts/fetch_fonts.py` — `data/fontfiles.json`의 OFL 폰트를 받아 PowerPoint 글꼴 이름을 알려 줌
- `skill/scripts/apply_fonts.py` — PPTX 글꼴을 제목·본문 두 가지로 통일. 굵기가 이름에 든 원래 글꼴은 가까운 굵기로 옮기고, 빠진 기호·드문 한글을 경고하며, PowerPoint 글꼴 테마 파일(`--theme-xml`)도 만듦
- `skill/scripts/export_pdf.py` — 바뀐 피피티를 PDF로 내보냄. 받은 폰트만 보이는 빈 환경에서 변환해 폰트가 실제로 들어갔는지 확인하고, LibreOffice가 한글과 숫자 사이를 벌리는 설정을 끔(접근성 태그 포함)
- `skill/scripts/license_report.py` — 공모전·납품용 폰트 라이선스 내역서(폰트·굵기·라이선스·출처·파일 해시)
- `tests/run_tests.py` — 스킬 회귀 시험(성격이 다른 덱 11종, 분석 검사 2종, 적용 검사 2종). 워크플로가 매번 돌림
- `data/fontfiles.json` — 자동 다운로드 대상(OFL 23종)과 굵기별 파일 주소

```bash
python3 skill/scripts/analyze_deck.py in.pptx
python3 skill/scripts/fetch_fonts.py --list
python3 skill/scripts/fetch_fonts.py 프리텐다드 --weights Regular,Bold,ExtraBold --out fonts
python3 skill/scripts/apply_fonts.py in.pptx out.pptx --manifest fonts/manifest.json --title 프리텐다드:ExtraBold --body 프리텐다드:Regular
python3 skill/scripts/export_pdf.py out.pptx --fonts fonts --manifest fonts/manifest.json --png after
```

### 참고한 프로젝트

- [achimnol/pptx-tool](https://github.com/achimnol/pptx-tool)(MIT): 굵기 이름이 든 글꼴('Noto Sans KR Black')을 새 가족의 같은 굵기로 옮기는 방식, 테마와 글자의 cs(복합 문자) 글꼴을 함께 맞추는 방식, Office 글꼴 테마 파일 만들기를 참고했습니다.
- [liyao1520/pptx-embed-fonts](https://github.com/liyao1520/pptx-embed-fonts)(MIT): 폰트를 PPTX 안에 직접 넣는 방식(EOT 형식 `.fntdata`)을 검토했습니다. 이 환경의 LibreOffice가 넣은 폰트를 읽지 못해 결과를 확인할 수 없었고, 한글 폰트는 굵기 하나가 수 MB~14MB라 파일이 크게 불어나 넣지 않았습니다. 받는 사람에게는 PDF를 함께 보내고, 편집할 사람에게는 PowerPoint의 '파일의 글꼴 포함' 저장을 안내합니다.

## 폰트 파일을 담지 않는 이유

무료폰트라도 상당수는 라이선스에서 **임의 재배포를 금지**합니다(OFL이 아닌 폰트, 일부 기업·지자체 전용 폰트 등). 그래서 이 저장소는 폰트 파일을 보관하지 않고 목록, 라이선스 정보, 공식 배포처 링크만 다룹니다. 실시간 미리보기도 제작사·Google Fonts·네이버 한글한글아름답게·jsDelivr처럼 이미 공개된 웹폰트 배포처의 CSS를 불러옵니다.

## 구성

```
index.html, style.css, app.js     화면
data/fonts.json                   수집 결과 (자동 생성)
data/webfonts.json                실시간 미리보기용 웹폰트 연결표 (직접 관리)
data/overrides.json               산돌구름 원본 정보 보정표(끊긴 배포처 주소 등)
data/fontfiles.json               PPT 스킬용 산돌 OFL 폰트 파일 주소·성격 태그 (직접 관리)
data/google-fonts.json            Google Fonts 한국어 폰트 (자동 생성)
scripts/google_fonts.py           Google Fonts 수집기
guide.html, guide.js               사용방법 페이지(웹 사용법·스킬 설치·요청 예시·전후 비교)
docs/examples/make_examples.py    전후 비교 예시 다시 만들기(스킬을 실제로 돌려 렌더링)
scripts/build_manual.py           사용설명서 PDF(docs/USER-GUIDE-ko.pdf) 만들기 → 스킬 ZIP에 포함
scripts/build_skill_zip.py        설치용 스킬 ZIP 만들기 → dist/free-korean-fonts-ppt.zip
skill/                            스킬 본체 (PPT 분석·폰트 적용·PDF 내보내기)
scripts/scrape.py                 수집기 (Python 3 + beautifulsoup4). 배포처 주소의 추적 꼬리표(?ref=, utm_)는 지움
.github/workflows/update.yml      매주 자동 갱신
```

## 배포 방법

1. 이 폴더를 새 GitHub 저장소에 올립니다.
2. 저장소 **Settings → Pages**에서 Source를 `Deploy from a branch`, 브랜치 `main` / 폴더 `/ (root)`로 지정합니다.
3. **Actions → 무료폰트 목록 갱신 → Run workflow**로 한 번 돌려 동작을 확인합니다.

## 직접 돌려 보기

```bash
pip install beautifulsoup4
python scripts/scrape.py          # data/fonts.json 갱신
python -m http.server 8000        # http://localhost:8000
```

## 실시간 미리보기 추가하기

`data/webfonts.json`에 산돌구름 폰트 ID(상세 주소 `/free-font/<ID>/...`의 숫자)를 키로, 공개 웹폰트 CSS 주소와 `font-family` 이름을 넣습니다. 넣기 전에 해당 폰트 라이선스가 웹 사용(임베딩)을 허용하는지 확인하세요.

```json
"16951": { "css": "https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.min.css", "family": "Pretendard Variable" }
```

## 알림

- 라이선스 정보는 산돌구름 표기를 옮긴 것이며 바뀔 수 있습니다. 쓰기 전에 배포처의 최신 원문을 확인하세요.
- 각 폰트의 저작권은 제작사에 있습니다. 이 프로젝트는 산돌과 관계없습니다.
- 산돌구름 페이지 구조가 바뀌면 수집기가 멈출 수 있습니다(목록이 10종 미만이면 기존 데이터를 지키려고 실패 처리합니다).
