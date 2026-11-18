---
name: free-korean-fonts-ppt
description: 피피티(.pptx)에 무료 한글 폰트를 골라 자동으로 내려받아 제목·본문에 적용할 때 사용. 무료폰트 색인(ryc3030/sandoll-freefonts)의 OFL 폰트만 받는다.
---

# 무료 한글 폰트 → PPT 자동 적용

무료폰트 색인(https://ryc3030.github.io/sandoll-freefonts/)의 데이터를 써서
① 발표 목적에 맞는 폰트 짝을 고르고 ② 폰트 파일을 공식·공개 배포처에서 내려받고
③ PPTX 전체(테마·마스터·레이아웃·슬라이드·표·차트)에 적용한 뒤 ④ 렌더링으로 확인하고 ⑤ 덱과 폰트 묶음을 함께 넘긴다.

## 지켜야 할 것

- **자동으로 받는 폰트는 `data/fontfiles.json`에 있는 것뿐이다.** 모두 산돌구름 표기상 OFL(수정·재배포 허용)이다.
  그 밖의 폰트(이서윤체, 여기어때 잘난체, 에스코어 드림 등)는 재배포가 막혀 있을 수 있으므로 내려받지 않는다.
  사용자가 그런 폰트를 원하면 `fetch_fonts.py`가 출력하는 공식 배포처 링크를 주고, 가장 비슷한 OFL 폰트를 대안으로 제안한다.
- 산돌구름 사이트에서 폰트 파일을 직접 긁어 오지 않는다.
- 한 덱에 글꼴 패밀리는 두 개까지(제목 1 + 본문 1). 같은 패밀리의 굵기 차이로 위계를 만드는 것을 먼저 고려한다.

## 0. 준비

```bash
mkdir -p ~/ffskill && cd ~/ffskill
for f in fetch_fonts.py apply_fonts.py; do
  curl -fsSLO https://raw.githubusercontent.com/ryc3030/sandoll-freefonts/main/skill/scripts/$f
done
pip install --break-system-packages -q fonttools lxml 2>/dev/null || pip install -q fonttools lxml
python3 ~/ffskill/fetch_fonts.py --list      # 받을 수 있는 폰트, 굵기, 성격
```

## 1. 덱 준비

- 사용자가 PPTX를 올렸으면 작업 폴더로 **복사**해서 쓴다(원본은 건드리지 않음).
- 새 덱을 만들어야 하면 먼저 pptx 스킬로 덱을 만든다. 이때 글꼴은 아무거나 둬도 된다. 3단계에서 통일한다.
- 사용자가 PPTX 파일이 아니라 Slides 아티팩트를 원하는 경우에는 이 스킬 대신 `data/webfonts.json`의 웹폰트 CSS를 쓴다.

## 2. 폰트 짝 고르기

사용자가 폰트를 지정했으면 그것을 쓴다. 아니면 덱의 목적·청중으로 아래에서 고르고, 고른 이유를 한 줄로 밝힌다.
목적이 정말 불분명할 때만 한 번 묻는다.

| 목적 | 제목 | 본문 |
|---|---|---|
| 공공기관·정책 보고 | 나눔스퀘어 네오 Heavy | 나눔스퀘어 네오 Regular |
| 기업 보고·제안서·IR | 프리텐다드 ExtraBold | 프리텐다드 Regular |
| 스타트업·제품 소개 | Wanted Sans Black | Wanted Sans Regular |
| 표·수치가 많은 자료 | SUIT ExtraBold | SUIT Regular |
| 강의·공모전 발표 | 페이퍼로지 ExtraBold | 페이퍼로지 Regular (또는 프리젠테이션) |
| 학술·논문 발표 | 본명조 Bold | 프리텐다드 Regular 또는 Noto Sans KR |
| 인문·문화·예술 | 마루 부리 Bold | 마루 부리 Regular (또는 고운바탕) |
| 홍보·행사·마케팅 | G마켓 산스 Bold 또는 검은고딕 | 프리텐다드 Regular |
| 아동·교육·친근한 안내 | 배민 주아체 | 나눔스퀘어라운드 Regular |
| 레트로·게임 콘셉트 | 갈무리9 | 프리텐다드 Regular |

- 본문은 Regular 이상. Light는 프로젝터에서 흐려지므로 본문에 쓰지 않는다.
- Instrument Serif처럼 `latinOnly` 폰트는 한글이 없으므로 `--title-ea`로 한글 글꼴을 따로 준다.
- 검은고딕·도현·주아·베이글 팻·갈무리9는 제목 전용. 본문에 쓰지 않는다.

## 3. 내려받기와 적용

```bash
cd <작업폴더>
python3 ~/ffskill/fetch_fonts.py 프리텐다드 --weights Regular,Bold,ExtraBold --out fonts --install
```

출력의 `PowerPoint 이름`이 실제로 넣을 글꼴 이름이다. Regular·Bold는 같은 이름(예: `Pretendard`)을 쓰고
그 밖의 굵기는 이름이 따로 있다(예: `Pretendard ExtraBold`). `fonts/manifest.json`에도 같은 정보가 있다.

```bash
python3 ~/ffskill/apply_fonts.py input.pptx output.pptx --title "Pretendard ExtraBold" --body "Pretendard"
# 제목이 영문 전용 폰트일 때
python3 ~/ffskill/apply_fonts.py input.pptx output.pptx --title "Instrument Serif" --title-ea "MaruBuri" --body "MaruBuri"
# 제목 굵기의 PowerPoint 이름이 Regular와 같고 '(굵게)'로 나왔을 때(예: 본명조 Bold → "Noto Serif KR"): --title-bold
# 코드 글꼴 등 지킬 글꼴이 있을 때: --keep "Consolas,D2Coding"
# Google Slides로 옮기거나 테마를 무시하는 프로그램용: --mode force
```

`apply_fonts.py`는 테마의 제목/본문 글꼴(라틴·한글·Hang)을 바꾸고, 슬라이드에 직접 박힌 글꼴을 테마 참조로 돌려
모든 글자가 새 글꼴을 따르게 한다. 차트는 글꼴 이름을 직접 쓴다. 기호 글꼴(Wingdings 등)과 글머리표 글꼴은 그대로 둔다.

## 4. 확인

```bash
soffice --headless --convert-to pdf output.pptx >/dev/null 2>&1
pdffonts output.pdf            # 고른 글꼴이 실제로 쓰였는지
pdftoppm -r 60 -png output.pdf page   # 이미지로 떨어뜨려 Read로 직접 본다
```

- 모든 슬라이드 이미지를 직접 보고, 글꼴이 바뀌면서 **넘치거나 잘린 글자**가 없는지 확인한다.
  한글 폰트마다 글자 폭이 달라 줄바꿈이 바뀐다. 넘치면 그 상자의 글자 크기를 줄이거나 상자를 넓히고 다시 확인한다.
- LibreOffice는 한글과 숫자 사이에 띄어쓰기를 조금 더 넣어 보여 줄 수 있다. PowerPoint에서는 생기지 않으므로 고치지 않는다.

## 5. 넘기기

- `output.pptx`와 함께 `fonts.zip`(폰트 파일 + `LICENSE-*.txt`)을 만들어 보낸다.
  ```bash
  cd fonts && zip -q ../fonts.zip *.otf *.ttf LICENSE-*.txt && cd ..
  ```
- 답에 한두 줄로 알린다: 고른 폰트와 이유, 그리고
  **"PowerPoint에서 제대로 보이려면 fonts.zip의 폰트를 설치하세요. 다른 사람에게 보낼 때는 설치 후 파일 → 옵션 → 저장 → '파일의 글꼴 포함'을 켜고 저장하면 됩니다."**
- 받을 수 없었던 폰트가 있으면 공식 배포처 링크를 함께 적는다.

## 참고

- 폰트 목록·라이선스: https://raw.githubusercontent.com/ryc3030/sandoll-freefonts/main/data/fonts.json (매주 자동 갱신)
- 파일 경로: https://raw.githubusercontent.com/ryc3030/sandoll-freefonts/main/data/fontfiles.json
- 새 OFL 폰트를 자동 받기 대상에 넣으려면 `fontfiles.json`에 공개 배포처의 TTF/OTF 주소를 추가한다.
