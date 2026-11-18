# 무료폰트 색인

[산돌구름 무료폰트](https://www.sandollcloud.com/freefonts)에 올라온 폰트를 매주 자동으로 모아 보여 주는 정적 웹앱입니다. GitHub Pages에서 그대로 돌아갑니다.

- **검색·필터**: 이름·제작사·분류·분위기 검색, OFL / 상업용 / 임베딩 / 로고·CI / 신규 / 스테디 / 실시간 미리보기 / 찜
- **미리보기**: 공개 웹폰트가 있는 폰트는 입력한 문구로 실시간 미리보기, 나머지는 견본 이미지
- **상세**: 폰트 정보, 산돌구름 라이선스 요약표, 라이선스 원문 안내, 웹폰트 연결 코드 복사
- **다운로드**: 제작사 공식 배포처(없으면 산돌구름 상세 페이지)로 연결
- **자동 갱신**: GitHub Actions가 매주 월요일 목록을 다시 수집하고, 바뀐 것이 있을 때만 커밋

## Claude 스킬: 무료 한글 폰트 → PPT 자동 적용

`skill/` 폴더는 Claude 스킬입니다. "이 피피티에 어울리는 무료 폰트 적용해 줘"라고 하면 Claude가
목적에 맞는 폰트 짝을 고르고, OFL 폰트를 공개 배포처에서 받아, PPTX의 테마·마스터·슬라이드·표·차트 글꼴을 한 번에 바꾼 뒤
렌더링으로 확인하고, 덱과 폰트 묶음(라이선스 포함)을 함께 넘깁니다.

- `skill/SKILL.md` — 스킬 본문
- `skill/scripts/fetch_fonts.py` — `data/fontfiles.json`의 OFL 폰트를 받아 PowerPoint 글꼴 이름을 알려 줌
- `skill/scripts/apply_fonts.py` — PPTX 글꼴을 제목/본문 두 가지로 통일
- `data/fontfiles.json` — 자동 다운로드 대상(OFL 23종)과 굵기별 파일 주소

```bash
python3 skill/scripts/fetch_fonts.py --list
python3 skill/scripts/fetch_fonts.py 프리텐다드 --weights Regular,ExtraBold --out fonts
python3 skill/scripts/apply_fonts.py in.pptx out.pptx --title "Pretendard ExtraBold" --body "Pretendard"
```

## 폰트 파일을 담지 않는 이유

무료폰트라도 상당수는 라이선스에서 **임의 재배포를 금지**합니다(OFL이 아닌 폰트, 일부 기업·지자체 전용 폰트 등). 그래서 이 저장소는 폰트 파일을 보관하지 않고, 목록·라이선스 정보·공식 배포처 링크만 다룹니다. 실시간 미리보기도 제작사·Google Fonts·네이버 한글한글아름답게·jsDelivr처럼 이미 공개된 웹폰트 배포처의 CSS를 그대로 불러옵니다.

## 구성

```
index.html, style.css, app.js     화면
data/fonts.json                   수집 결과 (자동 생성)
data/webfonts.json                실시간 미리보기용 웹폰트 연결표 (직접 관리)
data/fontfiles.json               PPT 스킬용 OFL 폰트 파일 주소 (직접 관리)
skill/                            Claude 스킬 (PPT 자동 적용)
scripts/scrape.py                 수집기 (Python 3 + beautifulsoup4)
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
