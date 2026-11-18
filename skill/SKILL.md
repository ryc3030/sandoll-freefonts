---
name: free-korean-fonts-ppt
description: 피피티(.pptx)를 분석해 내용에 어울리는 무료 한글 폰트(산돌구름 OFL·Google Fonts·기관 전용 서체)를 고르고, 내려받아 제목·본문에 적용할 때 사용.
---

# 무료 한글 폰트 → PPT 분석·자동 적용

무료폰트 색인(https://ryc3030.github.io/sandoll-freefonts/)의 데이터로
① 덱을 분석하고 ② 폰트 짝을 정하고 ③ 재배포가 허용된 폰트를 내려받아 ④ PPTX 전체에 적용한 뒤
⑤ 렌더링으로 확인하고 ⑥ 덱과 폰트 묶음(라이선스 포함)을 함께 넘긴다.

## 지켜야 할 것

- **내용이 먼저다.** 공공기관 덱(문서 속성·바닥글의 기관명, 공공 보고서 기호 □ ○ ※ Ⅰ., 기관명·직함)의 기본값은 프리텐다드지만,
  내용 성격이 뚜렷하면(축제·행사 홍보, 전시·문화, 아동, 학술 연구, 강의, 게임) 그 성격에 맞는 서체를 쓴다.
  예: 도청 주최 축제 홍보 → G마켓 산스+프리텐다드, 교육청 어린이 안전교육 → 주아+나눔스퀘어라운드, 국책연구원 연구 발표 → 본명조.
  보고·정책·IT 사업 보고처럼 성격이 '공공 업무' 그 자체면 프리텐다드.
- **기관 전용 서체가 있으면 먼저 제안한다.** 서울(서울남산체·서울한강체·서울알림체), 경기(경기천년체), 부산(부산단디체),
  교육청·학교(학교안심체), 보훈(62570체) 등. 분석 결과의 '기관 전용 서체 후보'를 보고, 사용자에게 공식 배포처 링크를 주며
  그 서체를 쓸지 묻는다. 라이선스가 기관마다 달라 자동으로 받지 않는다. 사용자가 파일을 주면 `fetch_fonts.py --local`로 적용한다.
- 사용자가 폰트를 지정하면 그것을 따른다.
- **자동으로 받는 폰트는 두 출처뿐이다**: 산돌구름 OFL 폰트(`fontfiles.json`)와 Google Fonts 한국어 폰트(`google-fonts.json`).
  그 밖의 폰트(이서윤체, 여기어때 잘난체, 에스코어 드림, 눈누에만 있는 폰트 등)는 재배포가 막혀 있을 수 있으므로 받지 않는다.
  원하면 `fetch_fonts.py`가 출력하는 공식 배포처 링크를 주고, 비슷한 OFL 폰트를 대안으로 제안한다.
- 산돌구름·눈누 사이트에서 폰트 파일을 직접 긁어 오지 않는다(눈누는 약관상 자동 수집 금지).
- 한 덱에 글꼴 패밀리는 두 개까지(제목 1 + 본문 1). 같은 가족의 굵기 차이로 위계를 만드는 것을 먼저 고려한다.

## 0. 준비

```bash
mkdir -p ~/ffskill && cd ~/ffskill
V=3cc0c746621f713f29e319197897c977b5668314   # 검증된 판(커밋)으로 고정. 최신을 쓰려면 V=main
for f in analyze_deck.py fetch_fonts.py apply_fonts.py; do
  curl -fsSLO https://raw.githubusercontent.com/ryc3030/sandoll-freefonts/$V/skill/scripts/$f
done
pip install --break-system-packages -q fonttools lxml 2>/dev/null || pip install -q fonttools lxml
```
데이터(`fontfiles.json` 등)는 스크립트가 저장소 `main`에서 읽는다(매주 갱신). 네트워크가 막혔으면 저장소의 `data/` 폴더를 받아 `--data <폴더>`로 준다.

## 1. 덱 준비

- 사용자가 PPTX를 올렸으면 작업 폴더로 **복사**해서 쓴다(원본은 건드리지 않음).
- 새 덱을 만들어야 하면 먼저 pptx 스킬로 덱을 만든다. 글꼴은 아무거나 둬도 된다. 4단계에서 통일한다.
- PPTX가 아니라 Slides 아티팩트를 원하면 이 스킬 대신 `data/webfonts.json`·`google-fonts.json`의 웹폰트 CSS를 쓴다.

## 2. 덱 분석

```bash
python3 ~/ffskill/analyze_deck.py input.pptx --json analysis.json
soffice --headless --convert-to pdf input.pptx >/dev/null 2>&1 && pdftoppm -r 40 -png input.pdf before
```

`analyze_deck.py`가 보는 변수: 발행 주체(문서 속성·바닥글·기관명), 분야 어휘(한·영), 공공 보고서 기호, 학술 인용·통계, 금액·비율,
글의 문체(개조식·합니다체·한다체·해요체·반말), 청중 호칭, 노트·차트·SmartArt 글, 슬라이드당 글자 수(발표형/읽는 보고서형),
제목이 자리표시자인지 텍스트 상자인지, 글자 크기, 굵은 글자 비율, 표의 숫자 비율, 이미지 면적, 배경 밝기(clrMap 반영), 채운 색,
지금 글꼴, 덱에 쓰인 특수 기호·드문 한글. 결과로 서체 성격 태그, 굵기·크기 조언, 폰트 짝 후보(빠진 기호 경고 ⚠ 포함)를 낸다.

1순위 방식이 '공공 발행 + 내용 우선(분야)'이면 공공기관 덱이지만 내용 성격을 따른 것이다. 2순위에 '공공 기본값'(프리텐다드)이 함께 나온다.
`before-*.png`를 직접 보고 판단을 보정한다. 예: 공공 보고서인데 표지만 축제 사진이면 본문 성격(보고)을 따라 프리텐다드.
**'판단 불확실'이면** 다음 순서로 정한다.
1. 공공 단서가 있고 내용 성격이 따로 드러나지 않으면 → 프리텐다드.
2. 이미지·노트·제목에서 목적을 읽고(예: 사진 위주 전시 소개 → 문화), 어떤 가정을 했는지 답에 한 줄로 밝힌다.
3. 그래도 모르겠으면 1순위(프리텐다드)로 적용하고, 사용자가 대화 중이면 대안 2개를 함께 제시한다.

## 3. 폰트 짝 정하기

사용자가 폰트를 지정했으면 그것을 쓴다. 아니면 1순위 후보를 기본으로 하되, 이미지를 보고 바꿀 이유가 있으면 바꾼다.
답에는 진단을 두세 줄로 밝힌다: "공공·정책 보고, 개조식, 표 많음 → 프리텐다드 ExtraBold/Regular. 대안: 나눔스퀘어 네오."

판단 기준
- 본문은 Regular 이상의 고딕·명조. 손글씨·디스플레이·픽셀 폰트는 본문에 쓰지 않는다. Light는 프로젝터에서 흐려진다.
- 제목만 개성 있는 폰트(검은고딕·도현·주아·동글·베이글 팻·갈무리9 등)를 쓰고 본문은 안정적인 고딕으로 받친다.
- 명조 제목은 Bold, 고딕 제목은 발표형이면 ExtraBold~Black, 읽는 보고서형이면 Bold.
- **⚠ 빠진 기호가 있는 후보는 피한다.** 예: SUIT·Wanted Sans·나눔고딕에는 Ⅰ Ⅱ가, 함렛·주아·검은고딕에는 □ ※가 없다.
  꼭 써야 하면 그 기호만 다른 글꼴로 보인다는 점을 사용자에게 알린다. 프리텐다드·본명조·Noto Sans KR·고딕 A1은 거의 다 있다.
- 영문 전용 폰트(Instrument Serif)는 영문 비중이 높을 때만, `--title-ea`로 한글 글꼴을 따로 준다.

참고 짝(분석 결과가 애매할 때)

| 목적 | 제목 | 본문 |
|---|---|---|
| 공공 업무(보고·정책·사업 보고·안내) | 프리텐다드 ExtraBold (기관 전용 서체가 있으면 그것) | 프리텐다드 Regular |
| 공공 주최 축제·행사 홍보 | G마켓 산스 Bold | 프리텐다드 Regular |
| 기업 보고·제안서·IR | 프리텐다드 ExtraBold | 프리텐다드 Regular |
| 스타트업·IT | Wanted Sans Black / IBM Plex Sans KR Bold | 같은 가족 Regular |
| 표·수치가 많은 자료 | 프리텐다드 Bold (Ⅰ·— 없는 덱이면 SUIT) | 같은 가족 Regular |
| 강의·공모전 발표 | 페이퍼로지 ExtraBold | 페이퍼로지 Regular |
| 학술·논문 발표 | 본명조 Bold | 본명조 Regular(글 많으면 프리텐다드 Regular) |
| 인문·문화·예술 | 마루 부리 Bold / 고운바탕 Bold | 같은 가족 Regular |
| 홍보·행사(민간) | G마켓 산스 Bold | 프리텐다드 Regular |
| 아동·가족 | 배민 주아체 / 동글 Bold | 나눔스퀘어라운드 Regular |
| 레트로·게임 | 갈무리9 | 프리텐다드 Regular |

## 4. 내려받기와 적용

제목·본문 가족을 **한 명령으로**, 본문의 굵은 글자용 Bold까지 받는다.

```bash
cd <작업폴더>
python3 ~/ffskill/fetch_fonts.py 프리텐다드 --weights Regular,Bold,ExtraBold --out fonts --install
python3 ~/ffskill/fetch_fonts.py 마루부리 프리텐다드 --weights Regular,Bold --out fonts --install   # 가족이 둘일 때
python3 ~/ffskill/apply_fonts.py input.pptx output.pptx --manifest fonts/manifest.json \
        --title 프리텐다드:ExtraBold --body 프리텐다드:Regular
```

- 기관 전용 서체 등 사용자가 준 파일: `fetch_fonts.py --local 파일1.ttf 파일2.ttf --out fonts --install` → 출력된 `이름:굵기`로 적용.
  이 파일은 재배포 조건이 기관마다 다르므로 6단계 fonts.zip에 넣지 않는다.
- 이름은 한글 이름·영문 family·산돌구름 ID 모두 된다(`함렛`, `"Gothic A1"`, `16951`). `--list`로 전체 목록.
- 같은 `--out`에 나눠 받아도 manifest가 합쳐진다. 받지 못한 것이 있으면 목록을 출력하고 종료 코드 1로 끝난다.
- **`이름:굵기`는 반드시 `--manifest`와 함께** 쓴다. 실제 PowerPoint 글꼴 이름(예: 'Pretendard ExtraBold'), 굵게 처리,
  빠진 기호 경고를 manifest로 맞춘다.
- `apply_fonts.py`가 하는 일: 테마 제목/본문 글꼴(라틴·한글·Hang) 교체, 슬라이드·레이아웃·마스터·노트·차트·SmartArt·기본 텍스트 스타일의
  직접 지정 글꼴 정리, **텍스트 상자로 만든 제목**(슬라이드에서 가장 큰 28pt 이상 상자)에도 제목 글꼴 적용,
  굵기 이름이 붙은 제목 글꼴의 가짜 굵게 제거, 본문 굵은 글자를 별도 Bold 글꼴로 교체(페이퍼로지·나눔스퀘어 계열 등).
  기호·아이콘·코드 글꼴(Wingdings, Font Awesome, Consolas, D2Coding 등)은 그대로 둔다. 더 지킬 글꼴은 `--keep "이름,이름"`.
- Google Slides로 옮기거나 테마를 무시하는 프로그램용이면 `--mode force`.
- 출력의 `!` 줄(빠진 기호, 드문 한글, Bold 미수신)은 반드시 읽고 처리한다.

## 5. 확인

이미 설치된 폰트가 섞여 거짓 통과가 나지 않도록 **빈 HOME에서** 렌더링한다.

```bash
export HOME=$(mktemp -d) && mkdir -p $HOME/.fonts && cp fonts/*.ttf fonts/*.otf $HOME/.fonts/ 2>/dev/null; fc-cache -f
soffice --headless --convert-to pdf output.pptx >/dev/null 2>&1
pdffonts output.pdf                    # 고른 글꼴이 쓰였는지. LiberationSans는 글머리표(•)뿐이면 정상
pdftoppm -r 60 -png output.pdf after   # 모든 장을 Read로 직접 본다
```

- 모든 슬라이드 이미지에서 **넘치거나 잘린 글자, 어색하게 두 줄이 된 제목, 제목인데 본문 글꼴로 보이는 상자**를 찾는다.
- 넘치면 python-pptx로 그 상자의 글자 크기를 줄이거나 자동 맞춤을 켠다:
  ```python
  from pptx import Presentation; from pptx.util import Pt
  p = Presentation("output.pptx"); sh = p.slides[0].shapes[0]
  for para in sh.text_frame.paragraphs:
      for r in para.runs: r.font.size = Pt(32)
  # 또는 sh.text_frame.fit_text(max_size=32)  /  bodyPr에 <a:normAutofit/>
  p.save("output.pptx")
  ```
- LibreOffice는 한글과 숫자 사이를 조금 더 띄워 보여 줄 수 있다. PowerPoint에서는 생기지 않으므로 고치지 않는다.

## 6. 넘기기

- `output.pptx`와 함께 **실제로 쓴 굵기만** 담은 `fonts.zip`(폰트 + `LICENSE-*.txt`)을 보낸다. LICENSE에는 저작권 표시와 OFL 1.1 원문이 들어 있어 재배포 조건을 지킨다.
  ```bash
  cd fonts && zip -q ../fonts.zip Pretendard-Regular.ttf Pretendard-Bold.ttf Pretendard-ExtraBold.ttf LICENSE-*.txt && cd ..
  ```
- 답에는 진단 두세 줄, 고른 폰트와 대안, 그리고 manifest의 `outline`·`embeddable`에 따라 다음 중 맞는 안내를 쓴다.
  - 모두 TrueType이고 `편집 가능`: "PowerPoint에서 제대로 보이려면 fonts.zip의 폰트를 설치하세요. 다른 사람에게 보낼 때는 파일 → 옵션 → 저장 → '파일의 글꼴 포함'을 켜고 저장하면 됩니다."
  - `CFF`가 있으면: "이 폰트는 PowerPoint 파일에 포함되지 않습니다. 받는 사람도 fonts.zip의 폰트를 설치해야 합니다."
  - `보기·인쇄만`(SUIT·나눔스퀘어·G마켓 산스 등): "파일에 포함해 보내면 받는 사람은 보기·인쇄만 되고 편집하려면 폰트를 설치해야 합니다."
- 받을 수 없었던 폰트와 기관 전용 서체는 공식 배포처 링크를 함께 적는다.

## 참고

- 데이터(매주 자동 갱신): https://raw.githubusercontent.com/ryc3030/sandoll-freefonts/main/data/
  `fonts.json`(산돌구름 전체), `fontfiles.json`(산돌 OFL 파일 주소·성격 태그), `google-fonts.json`(Google 한국어·성격 태그),
  `glyph-coverage.json`(폰트별 빠진 기호·한글 글자 수), `agency-fonts.json`(기관 전용 서체·공식 배포처), `OFL-1.1.txt`
- 새 OFL 폰트를 자동 받기 대상에 넣으려면 `fontfiles.json`에 공개 배포처의 TTF 주소와 `tags`를 추가하고 `scripts/glyph_coverage.py`를 돌린다.
