---
name: free-korean-fonts-ppt
description: 피피티(.pptx)를 분석해 어울리는 무료 한글 폰트(산돌구름 OFL·Google Fonts)를 고르고, 자동으로 내려받아 제목·본문에 적용할 때 사용.
---

# 무료 한글 폰트 → PPT 분석·자동 적용

무료폰트 색인(https://ryc3030.github.io/sandoll-freefonts/)의 데이터로
① 덱을 분석해 분야·글의 문체·밀도·색감을 진단하고 ② 어울리는 서체 성격과 폰트 짝을 고르고
③ 폰트 파일을 공개 배포처에서 내려받아 ④ PPTX 전체(테마·마스터·레이아웃·슬라이드·표·차트)에 적용한 뒤
⑤ 렌더링으로 확인하고 ⑥ 덱과 폰트 묶음을 함께 넘긴다.

## 지켜야 할 것

- **자동으로 받는 폰트는 두 출처뿐이다.**
  - 산돌구름 무료폰트 중 OFL 폰트(`data/fontfiles.json`, 제작사·네이버·jsDelivr 배포처)
  - Google Fonts 한국어 폰트(`data/google-fonts.json`, 전부 OFL)
- 그 밖의 폰트(이서윤체, 여기어때 잘난체, 에스코어 드림, 눈누에만 있는 폰트 등)는 재배포가 막혀 있을 수 있으므로 내려받지 않는다.
  사용자가 그런 폰트를 원하면 공식 배포처 링크를 주고, 가장 비슷한 OFL 폰트를 대안으로 제안한다.
- 산돌구름·눈누 사이트에서 폰트 파일을 직접 긁어 오지 않는다(눈누는 약관상 자동 수집 금지).
- 한 덱에 글꼴 패밀리는 두 개까지(제목 1 + 본문 1). 같은 패밀리의 굵기 차이로 위계를 만드는 것을 먼저 고려한다.

## 0. 준비

```bash
mkdir -p ~/ffskill && cd ~/ffskill
for f in analyze_deck.py fetch_fonts.py apply_fonts.py; do
  curl -fsSLO https://raw.githubusercontent.com/ryc3030/sandoll-freefonts/main/skill/scripts/$f
done
pip install --break-system-packages -q fonttools lxml 2>/dev/null || pip install -q fonttools lxml
```

## 1. 덱 준비

- 사용자가 PPTX를 올렸으면 작업 폴더로 **복사**해서 쓴다(원본은 건드리지 않음).
- 새 덱을 만들어야 하면 먼저 pptx 스킬로 덱을 만든다. 글꼴은 아무거나 둬도 된다. 4단계에서 통일한다.
- PPTX 파일이 아니라 Slides 아티팩트를 원하면 이 스킬 대신 `data/webfonts.json`·`google-fonts.json`의 웹폰트 CSS를 쓴다.

## 2. 덱 분석

```bash
python3 ~/ffskill/analyze_deck.py input.pptx --json analysis.json
soffice --headless --convert-to pdf input.pptx >/dev/null 2>&1 && pdftoppm -r 40 -png input.pdf before
```

`analyze_deck.py`가 내는 것:
- **분야**: 공공·정책 / 기업·비즈니스 / 학술·연구 / 교육·강의 / IT·스타트업 / 홍보·행사 / 문화·예술·인문 / 아동·가족 / 게임·레트로
- **글의 문체**: 개조식(~함·명사형), 합니다체, 해요체, 반말·구어, 느낌표·이모지 수
- **밀도**: 슬라이드당 글자 수, 표·차트·이미지 수 → 빽빽하면 본문은 가독성 최우선
- **색감**: 차분한 무채색 / 파스텔 / 선명 / 어두움
- **어울리는 서체 성격 태그**(formal·neutral·friendly·playful·cute·handwritten·retro·elegant·traditional·academic·tech·impact·calm)와 **폰트 짝 후보 3~4개**

점수는 출발점이다. `before-*.png` 몇 장을 직접 보고 판단을 보정한다.
예: 어휘는 '공공'인데 디자인이 밝은 일러스트 위주라면 친근한 쪽으로, 학술 발표인데 표가 많으면 본문은 고딕으로.

## 3. 폰트 짝 정하기

사용자가 폰트를 지정했으면 그것을 쓴다. 아니면 분석 결과를 바탕으로 정하고, 답에 진단을 두세 줄로 밝힌다:
"공공·정책 발표, 개조식, 표가 많음 → 격식 있고 중립적인 고딕. 프리텐다드 ExtraBold/Regular로 적용. 대안: 나눔스퀘어 네오, 본명조+프리텐다드."
목적이 정말 불분명하고 사용자가 대화 중일 때만 후보를 보여 주고 고르게 한다.

판단 기준
- 본문은 Regular 이상의 고딕·명조. 손글씨·디스플레이·픽셀 폰트는 본문에 쓰지 않는다. Light는 프로젝터에서 흐려진다.
- 제목만 개성 있는 폰트(검은고딕·도현·주아·동글·베이글 팻·갈무리9 등)를 쓰고 본문은 안정적인 고딕으로 받친다.
- 명조 제목은 Bold, 고딕 제목은 ExtraBold~Black이 무난하다.
- 영문 전용 폰트(Instrument Serif 등)는 영문 비중이 높을 때만, `--title-ea`로 한글 글꼴을 따로 준다.

참고 짝(분석 결과가 애매할 때)

| 목적 | 제목 | 본문 |
|---|---|---|
| 공공기관·정책 보고 | 나눔스퀘어 네오 Heavy / 프리텐다드 ExtraBold | 같은 가족 Regular |
| 기업 보고·제안서·IR | 프리텐다드 ExtraBold | 프리텐다드 Regular |
| 스타트업·IT | Wanted Sans Black / IBM Plex Sans KR Bold | 같은 가족 Regular |
| 표·수치가 많은 자료 | SUIT ExtraBold | SUIT Regular |
| 강의·공모전 발표 | 페이퍼로지 ExtraBold | 페이퍼로지 Regular |
| 학술·논문 발표 | 본명조 Bold / 함렛 Bold | 본명조 Regular 또는 Noto Sans KR |
| 인문·문화·예술 | 마루 부리 Bold / 고운바탕 Bold | 같은 가족 Regular |
| 홍보·행사 | G마켓 산스 Bold / 검은고딕 | 프리텐다드 Regular |
| 아동·가족·친근한 안내 | 배민 주아체 / 동글 Bold | 나눔스퀘어라운드 Regular |
| 레트로·게임 | 갈무리9 / 오르빗 | 프리텐다드 Regular |

## 4. 내려받기와 적용

```bash
cd <작업폴더>
python3 ~/ffskill/fetch_fonts.py --list                     # 받을 수 있는 폰트 전체(산돌 OFL + Google)
python3 ~/ffskill/fetch_fonts.py 프리텐다드 --weights Regular,ExtraBold --out fonts --install
python3 ~/ffskill/apply_fonts.py input.pptx output.pptx --manifest fonts/manifest.json \
        --title 프리텐다드:ExtraBold --body 프리텐다드:Regular
```

- 이름은 한글 이름·영문 family·산돌구름 ID 모두 된다(예: `함렛`, `"Gothic A1"`, `16951`).
- `--manifest`를 주고 `이름:굵기`로 지정하면 PowerPoint에 넣을 실제 글꼴 이름과 굵게 여부를 알아서 맞춘다.
  없는 굵기를 고르면 받은 굵기 목록을 보여 주며 멈추므로, 그 굵기를 받거나 목록에서 고른다.
- 그 밖의 옵션: 제목만 영문 폰트 `--title-ea 마루부리:Regular`, 지킬 글꼴 `--keep "Consolas,D2Coding"`,
  Google Slides로 옮길 때 `--mode force`.

`apply_fonts.py`는 테마의 제목/본문 글꼴(라틴·한글·Hang)을 바꾸고, 슬라이드에 직접 박힌 글꼴을 테마 참조로 돌려
모든 글자가 새 글꼴을 따르게 한다. 차트는 글꼴 이름을 직접 쓴다. 기호 글꼴과 글머리표 글꼴은 그대로 둔다.

## 5. 확인

```bash
soffice --headless --convert-to pdf output.pptx >/dev/null 2>&1
pdffonts output.pdf                    # 고른 글꼴이 실제로 쓰였는지
pdftoppm -r 60 -png output.pdf after   # 이미지로 떨어뜨려 Read로 직접 본다
```

- 모든 슬라이드 이미지를 보고 글꼴이 바뀌면서 **넘치거나 잘린 글자, 어색하게 두 줄이 된 제목**이 없는지 확인한다.
  넘치면 그 상자의 글자 크기를 줄이거나 상자를 넓히고 다시 확인한다.
- LibreOffice는 한글과 숫자 사이를 조금 더 띄워 보여 줄 수 있다. PowerPoint에서는 생기지 않으므로 고치지 않는다.

## 6. 넘기기

- `output.pptx`와 함께 `fonts.zip`(폰트 파일 + `LICENSE-*.txt`)을 보낸다.
  ```bash
  cd fonts && zip -q ../fonts.zip *.otf *.ttf LICENSE-*.txt && cd ..
  ```
- 답에는 진단 두세 줄, 고른 폰트와 대안, 그리고 다음 안내를 적는다:
  **"PowerPoint에서 제대로 보이려면 fonts.zip의 폰트를 설치하세요. 다른 사람에게 보낼 때는 설치 후 파일 → 옵션 → 저장 → '파일의 글꼴 포함'을 켜고 저장하면 됩니다."**
- 받을 수 없었던 폰트가 있으면 공식 배포처 링크를 함께 적는다.

## 참고

- 데이터(매주 자동 갱신): https://raw.githubusercontent.com/ryc3030/sandoll-freefonts/main/data/
  `fonts.json`(산돌구름 전체), `fontfiles.json`(산돌 OFL 파일 주소·성격 태그), `google-fonts.json`(Google 한국어·성격 태그)
- 새 OFL 폰트를 자동 받기 대상에 넣으려면 `fontfiles.json`에 공개 배포처의 TTF/OTF 주소와 `tags`를 추가한다.
