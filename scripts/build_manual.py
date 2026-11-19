#!/usr/bin/env python3
"""스킬 사용설명서 PDF(docs/USER-GUIDE-ko.pdf)를 만든다. 스킬 ZIP에 함께 들어간다.

  python3 scripts/build_manual.py      # Playwright(Chromium)·인터넷(프리텐다드 받기) 필요
예시 그림과 진단은 docs/examples/(make_examples.py 결과)를 쓴다. 예시를 바꿨으면 먼저 make_examples.py를 돌린다.
PDF는 저장소에 커밋해 두므로 자동 갱신(CI)에서는 다시 만들지 않는다.
"""
import html, json, subprocess, sys, tempfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
EX = ROOT / "docs/examples"
OUT = ROOT / "docs/USER-GUIDE-ko.pdf"
SITE = "https://ryc3030.github.io/sandoll-freefonts/"
ZIPURL = "https://github.com/ryc3030/sandoll-freefonts/raw/main/dist/free-korean-fonts-ppt.zip"
e = html.escape

ASKS = [
    ("기본", "이 피피티에 어울리는 무료 폰트로 바꿔 줘", "분석부터 적용, PDF 내보내기까지 한 번에"),
    ("분석만", "폰트는 바꾸지 말고, 어떤 폰트가 어울릴지 분석만 해 줘", "진단과 추천 짝 3개, 이유"),
    ("그래프·디자인", "그래프랑 디자인도 같이 손봐 줘. 글자가 넘치거나 흐린 곳도 고쳐 줘", "폰트 적용 + 그래프·디자인 보정, 흐린·넘치는 글자 보정"),
    ("점검만", "폰트는 그대로 두고 그래프·디자인에 문제가 있는지만 알려 줘", "보정 없이 슬라이드별 지적 목록"),
    ("폰트 지정", "제목은 G마켓 산스 Bold, 본문은 프리텐다드로 적용해 줘", "지정한 폰트로 적용(기호 누락 경고는 알려 줌)"),
    ("공모전·납품", "공모전 제출용이야. 폰트 라이선스 내역서도 같이 만들어 줘", "적용 + 폰트·굵기·라이선스·출처·파일 해시 내역서"),
    ("기관 서체", "서울시 행사 자료야. 서울남산체 파일을 첨부할게, 이걸로 적용해 줘", "첨부한 파일로 적용(재배포 묶음에는 넣지 않음)"),
    ("보내기", "다른 사람이 이 피피티를 고쳐야 해. 폰트까지 챙겨서 보낼 수 있게 해 줘", "폰트를 넣어 저장할 수 있는지와 방법 안내"),
    ("새 슬라이드", "앞으로 새 슬라이드를 만들 때도 이 글꼴이 나오게 글꼴 테마 파일도 만들어 줘", "PowerPoint 글꼴 테마 파일(.xml)과 넣을 위치"),
    ("Google Slides", "Google Slides로 옮길 거야. 거기서도 폰트가 유지되게 해 줘", "글자마다 글꼴을 직접 지정하는 방식으로 적용"),
    ("새로 만들기", "이 보고서로 10장짜리 발표 자료를 만들고 어울리는 무료 폰트로 맞춰 줘", "덱을 만든 뒤 같은 절차로 폰트 적용"),
]
PAIRS = [
    ("공공 업무(보고·정책·사업 보고·안내)", "프리텐다드 ExtraBold", "프리텐다드 Regular"),
    ("공공 주최 축제·행사 홍보", "G마켓 산스 Bold", "프리텐다드 Regular"),
    ("추모·기념식·의전", "본명조 Bold", "본명조 Regular"),
    ("기업 보고·제안서·IR", "프리텐다드 ExtraBold", "프리텐다드 Regular"),
    ("스타트업·IT", "Wanted Sans Black", "Wanted Sans Regular"),
    ("강의·공모전 발표", "페이퍼로지 ExtraBold", "페이퍼로지 Regular"),
    ("학술·논문 발표", "본명조 Bold", "본명조 Regular"),
    ("인문·문화·예술", "마루 부리 Bold", "마루 부리 Regular"),
    ("아동·가족", "배민 주아체", "나눔스퀘어라운드 Regular"),
    ("레트로·게임", "갈무리9", "프리텐다드 Regular"),
]
FAQ = [
    ("공공기관 자료는 무조건 프리텐다드인가요?",
     "보고·정책처럼 공공 업무가 내용이면 프리텐다드가 기본입니다. 추모·기념식은 명조, 축제 홍보는 G마켓 산스, 어린이 교육은 둥근 서체처럼 내용 성격이 뚜렷하면 그에 맞춥니다. 이 설명서의 예시 2·3이 그 경우입니다."),
    ("어떤 폰트를 자동으로 받나요?",
     "재배포가 허용된 산돌구름 OFL 폰트와 Google Fonts 한국어 폰트만 받습니다. 그 밖의 폰트는 공식 배포처 링크를 알려 주고, 비슷한 OFL 폰트를 대안으로 제안합니다."),
    ("기관 전용 서체(서울남산체·학교안심체 등)는요?",
     "자료를 만든 기관을 알아보면 그 기관의 전용 서체와 공식 배포처를 먼저 알려 줍니다. 라이선스가 기관마다 달라 자동으로 받지 않으니, 직접 받아 첨부하면 그 파일로 적용합니다."),
    ("받는 사람 컴퓨터에 폰트가 없으면요?",
     "피피티는 다른 글꼴로 바뀌어 보입니다. 보기만 할 사람에게는 함께 받은 PDF를 보내세요. 고칠 사람에게는 fonts.zip의 폰트를 설치하게 하거나, PowerPoint에서 파일 → 옵션 → 저장 → '파일의 글꼴 포함'을 켜고 저장해 보내세요. SUIT·나눔스퀘어·G마켓 산스 등은 넣어 보내도 받는 사람이 보기·인쇄만 할 수 있습니다. 스킬이 폰트마다 알려 줍니다."),
    ("그래프와 디자인은 무엇을 고치나요? 제 디자인이 바뀌지는 않나요?",
     "거의 언제나 잘못인 것만 자동으로 고칩니다. 3D 그래프는 2D로, 0이 아닌 값에서 시작하는 막대 축은 0부터, 작은 그래프 글자는 12pt 이상으로 바꾸고, 늘어난 사진과 조금씩 어긋난 제목 위치를 바로잡습니다. 색·배치처럼 취향이 갈리는 것은 바꾸지 않습니다. 흐린 글자, 넘치는 글자, 그래프 색 통일은 알려 준 뒤 원할 때만 고치며, 고친 내역은 답에 모두 적습니다."),
    ("PDF는 PowerPoint에서 저장한 것과 같나요?",
     "거의 같지만 LibreOffice로 만든 PDF라 줄바꿈이 조금 다를 수 있습니다. 제출용 PDF가 PowerPoint 화면과 똑같아야 하면 fonts.zip의 폰트를 설치한 뒤 PowerPoint에서 파일 → 다른 이름으로 저장 → PDF로 다시 만드세요."),
    ("목록에 없는 폰트를 쓰고 싶어요.",
     "공식 배포처에서 직접 받은 폰트 파일을 첨부하면 그 파일로 적용합니다. 재배포 조건을 알 수 없는 파일이라 fonts.zip에는 넣지 않습니다."),
    ("원본 피피티가 바뀌나요?",
     "아니요. 원본은 복사해서 작업하고, 바뀐 파일을 새로 만들어 돌려줍니다."),
    ("인터넷이 막힌 곳에서도 되나요?",
     "데이터가 ZIP에 들어 있어 분석과 적용은 됩니다. 인터넷은 폰트를 새로 내려받을 때만 필요하고, 폰트 파일을 직접 첨부하면 인터넷 없이도 적용할 수 있습니다."),
]


def reg(x):
    n = x[0][0] if x else ""
    return "" if "보류" in n else n


def why(x):
    if "내용 우선" in x["way"]:
        return f"만든 곳은 공공기관이지만 내용이 '{x['genre']}' 성격이 뚜렷해 그에 맞는 서체를 골랐습니다."
    if "공공" in x["way"]:
        return "공공기관의 정책 보고 자료라 공공 기본값인 프리텐다드를 골랐습니다. □ ○ ※ Ⅰ 기호를 모두 갖춘 폰트입니다."
    if "기업" in x["genre"] + x["way"]:
        return "회사의 실적 보고 자료라 숫자와 표가 잘 읽히는 프리텐다드를 골랐습니다."
    return ""


def example_page(i, x):
    rows = "".join(
        f'<div class="pair"><figure><img src="{(EX / f"{x["key"]}-before-{n}.jpg").as_uri()}"><figcaption>원본 {n}쪽 · 맑은 고딕</figcaption></figure>'
        f'<figure><img src="{(EX / f"{x["key"]}-after-{n}.jpg").as_uri()}"><figcaption>적용 후 {n}쪽</figcaption></figure></div>'
        for n in (range(2, x["pages"] + 1) if x["pages"] > 3 else range(1, x["pages"] + 1)))
    ag = "".join(f"<li>{e(a['agency'])}: {e(', '.join(a['fonts']))}<br><span class='url'>{e(a['url'])}</span></li>" for a in x["agency"])
    r = reg(x["register"])
    compact = " compact" if len(x.get("polish", [])) > 3 else ""
    return f"""<section class="page ex{compact}">
  <p class="eyeline">6 · 예시 {i}</p>
  <h2>{e(x['title'])}</h2>
  <p class="who">만든 곳: {e(x['author'])} · 요청: “{'어울리는 무료 폰트로 바꾸고, 그래프랑 디자인도 같이 손봐 줘' if len(x.get('polish', [])) > 3 else '이 피피티에 어울리는 무료 폰트로 바꿔 줘'}”</p>
  {rows}
  <div class="diag">
    <table class="kv">
      <tr><th>진단</th><td>{e(x['genre'])}{(' · 문체 ' + e(r)) if r else ''}</td></tr>
      <tr><th>제목</th><td><b>{e(x['titleFont'])}</b></td></tr>
      <tr><th>본문</th><td><b>{e(x['bodyFont'])}</b></td></tr>
      <tr><th>대안</th><td>{'<br>'.join(map(e, x['alt']))}</td></tr>
    </table>
    <div class="why"><p>{e(why(x))}</p>{('<p class="agh">기관 전용 서체 제안</p><ul>' + ag + '</ul><p class="note">라이선스가 기관마다 달라 자동으로 받지 않습니다. 파일을 첨부하면 그 파일로 적용합니다.</p>') if ag else ''}</div>
  </div>
  {('<p class="agh">그래프·디자인 보정</p><ul class="pol">' + ''.join(f"<li>{e(m)}</li>" for m in x.get('polish', [])) + '</ul>') if x.get('polish') else ''}
</section>"""


def build_html(fontdir):
    ex = json.loads((EX / "examples.json").read_text("utf-8"))
    ff = "".join(f"@font-face{{font-family:P;font-weight:{w};src:url('{(fontdir / f'Pretendard-{n}.ttf').as_uri()}')}}"
                 for n, w in [("Regular", 400), ("SemiBold", 600), ("Bold", 700), ("ExtraBold", 800)])
    asks = "".join(f"<tr><td class='tag'>{e(a)}</td><td>“{e(b)}”</td><td>{e(c)}</td></tr>" for a, b, c in ASKS)
    pairs = "".join(f"<tr><td>{e(a)}</td><td>{e(b)}</td><td>{e(c)}</td></tr>" for a, b, c in PAIRS)
    faq = "".join(f"<dt>{e(q)}</dt><dd>{e(a)}</dd>" for q, a in FAQ)
    exs = "".join(example_page(i + 1, x) for i, x in enumerate(ex))
    today = date.today()
    return f"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8"><style>
{ff}
@page {{ size: A4; margin: 18mm 17mm 16mm; }}
* {{ box-sizing: border-box; }}
body {{ font-family: P, sans-serif; color: #1b1a18; font-size: 10pt; line-height: 1.6; margin: 0; word-break: keep-all; }}
.page {{ break-after: page; }} .page:last-child {{ break-after: auto; }}
h1 {{ font-size: 34pt; line-height: 1.15; letter-spacing: -0.03em; margin: 0; font-weight: 800; }}
h2 {{ font-size: 18pt; letter-spacing: -0.02em; margin: 0 0 4mm; font-weight: 800; }}
h3 {{ font-size: 11.5pt; margin: 7mm 0 2.5mm; font-weight: 700; }}
p {{ margin: 0 0 2.5mm; }}
.eyeline {{ font-size: 9pt; font-weight: 700; color: #b5371f; margin-bottom: 1.5mm; }}
.cover {{ padding-top: 38mm; }}
.cover .sub {{ font-size: 13pt; color: #4a4741; margin-top: 6mm; max-width: 140mm; }}
.cover .meta {{ margin-top: 40mm; font-size: 9.5pt; color: #4a4741; border-top: 1px solid #d6d1c6; padding-top: 4mm; }}
.cover .meta b {{ display: inline-block; width: 22mm; color: #1b1a18; }}
.toc {{ margin-top: 12mm; font-size: 10.5pt; }} .toc li {{ margin-bottom: 1.2mm; }}
table {{ width: 100%; border-collapse: collapse; font-size: 9.3pt; }}
th, td {{ text-align: left; vertical-align: top; padding: 1.5mm 2.2mm; border-bottom: 0.6pt solid #d6d1c6; }}
thead th {{ font-weight: 600; color: #6b665d; border-bottom: 0.9pt solid #1b1a18; }}
td.tag {{ font-weight: 700; color: #b5371f; white-space: nowrap; }}
code {{ font-family: P, monospace; background: #f1eee7; padding: 0 1.2mm; border-radius: 1mm; font-size: 9pt; }}
.steps {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 2.5mm; margin: 0; padding: 0; list-style: none; counter-reset: s; }}
.steps li {{ counter-increment: s; border: 0.6pt solid #d6d1c6; border-radius: 2mm; padding: 2.5mm 3mm; font-size: 8.8pt; line-height: 1.5; }}
.steps li b {{ display: block; font-size: 10.3pt; margin-bottom: 1mm; }} .steps li b::before {{ content: counter(s) ". "; color: #b5371f; }}
.box {{ background: #f4f1ea; border-radius: 2mm; padding: 3.5mm 4.5mm; margin: 4mm 0; font-size: 9.3pt; }}
.box p:last-child {{ margin: 0; }}
.out li {{ margin-bottom: 1.5mm; }}
.ex .who {{ color: #4a4741; font-size: 9.3pt; margin-bottom: 4mm; }}
.pair {{ display: grid; grid-template-columns: 1fr 1fr; gap: 3.5mm; margin-bottom: 3mm; }}
figure {{ margin: 0; }} figure img {{ width: 100%; display: block; border: 0.6pt solid #cfc9bc; border-radius: 1.2mm; }}
figcaption {{ font-size: 8pt; color: #6b665d; margin-top: 1mm; }}
.diag {{ display: grid; grid-template-columns: 1fr 1fr; gap: 6mm; margin-top: 4mm; border-top: 0.9pt solid #1b1a18; padding-top: 3mm; }}
.kv th {{ width: 13mm; color: #6b665d; font-weight: 600; }} .kv td, .kv th {{ padding: 1.4mm 1mm; }}
.why {{ font-size: 9.3pt; }} .agh {{ font-weight: 700; margin-top: 3mm; }} .why ul {{ margin: 0 0 2mm; padding-left: 4.5mm; }}
.url {{ font-size: 8pt; color: #6b665d; word-break: break-all; }}
.pol {{ font-size: 8.2pt; line-height: 1.4; column-count: 2; column-gap: 6mm; margin: 0; padding-left: 4mm; }} .pol li {{ margin-bottom: 0.8mm; break-inside: avoid; }}
.compact .pair {{ width: 70%; margin: 0 auto 2mm; }} .compact .diag {{ margin-top: 2mm; }}
.note {{ font-size: 8.3pt; color: #6b665d; }}
dl.faq dt {{ font-weight: 700; margin-top: 4mm; }} dl.faq dd {{ margin: 1mm 0 0; color: #3a3833; }}
</style></head><body>

<section class="page cover">
  <p class="eyeline">free-korean-fonts-ppt 스킬</p>
  <h1>무료 한글 폰트<br>피피티 자동 적용<br>사용설명서</h1>
  <p class="sub">피피티를 올리면 내용에 어울리는 무료 한글 폰트를 적용하고 그래프와 디자인을 바로잡아, 바뀐 피피티와 PDF를 돌려줍니다.</p>
  <ol class="toc"><li>이 스킬이 하는 일</li><li>설치</li><li>요청하는 법</li><li>받는 결과물</li><li>폰트를 고르는 기준</li><li>사용 예시 4종 (전후 비교, 그래프·디자인 보정 포함)</li><li>자주 묻는 질문</li></ol>
  <div class="meta">
    <p><b>만든 날</b>{today.year}. {today.month}. {today.day}.</p>
    <p><b>사용방법</b>{SITE}guide.html</p>
    <p><b>폰트 색인</b>{SITE}</p>
    <p><b>최신 ZIP</b>{ZIPURL}</p>
  </div>
</section>

<section class="page">
  <p class="eyeline">1</p><h2>이 스킬이 하는 일</h2>
  <p>피피티(.pptx)를 대화에 올리고 “어울리는 무료 폰트로 바꿔 줘”라고 하면 AI가 아래 일곱 단계를 차례로 진행합니다.</p>
  <ol class="steps">
    <li><b>분석</b>만든 곳(기관명·문서 속성), 내용(정책·추모·아동·학술 등), 글의 문체, 슬라이드당 글자 양, 배경색, 쓰인 기호까지 30여 가지를 봅니다.</li>
    <li><b>고르기</b>제목 1개와 본문 1개, 두 가족 이내로 짝을 정합니다. □ ※ Ⅰ 같은 기호가 빠진 폰트는 피합니다.</li>
    <li><b>내려받기</b>재배포가 허용된 폰트(산돌구름 OFL·Google Fonts)만 공식 배포처에서 받습니다.</li>
    <li><b>적용</b>테마·마스터·슬라이드·표·차트·노트의 글꼴을 한 번에 바꿉니다. 텍스트 상자로 만든 제목에도 제목 글꼴을 입히고, 원래 굵기 차이도 살립니다.</li>
    <li><b>보정</b>3D 그래프, 0에서 시작하지 않는 막대 축, 작은 그래프 글자, 늘어난 사진, 조금씩 어긋난 제목 위치를 고칩니다. 흐린 글자와 넘치는 글자는 알려 주고, 원하면 고칩니다.</li>
    <li><b>PDF·확인</b>받은 폰트만 있는 빈 환경에서 PDF로 내보내고, 폰트가 실제로 들어갔는지와 넘치거나 잘린 글자가 없는지 봅니다.</li>
    <li><b>넘기기</b>바뀐 피피티, PDF, 폰트 묶음(라이선스 포함)과 진단 요약을 돌려줍니다.</li>
  </ol>
  <div class="box"><p><b>원칙</b> — 내용이 먼저입니다. 공공기관 자료의 기본값은 프리텐다드이고, 추모·기념식, 축제 홍보, 아동 교육, 학술 발표처럼 성격이 뚜렷하면 그에 맞는 서체를 씁니다. 사용자가 폰트를 지정하면 그것을 따릅니다.</p></div>

  <p class="eyeline" style="margin-top:9mm">2</p><h2>설치</h2>
  <p>이 ZIP 파일 하나로 설치합니다. 폰트를 실제로 바꾸려면 코드를 실행할 수 있는 대화(파일 업로드·코드 실행이 켜진 대화)여야 합니다.</p>
  <table><thead><tr><th style="width:42mm">쓰는 도구</th><th>설치 방법</th></tr></thead><tbody>
    <tr><td><b>Claude</b> (웹·데스크톱 앱)</td><td>설정 → 기능 → 스킬에서 ZIP을 그대로 업로드</td></tr>
    <tr><td><b>Claude Code</b></td><td><code>~/.claude/skills/</code> 폴더에 압축 풀기</td></tr>
    <tr><td><b>OpenAI Codex</b></td><td>프로젝트의 <code>.agents/skills/</code> 폴더에 압축 풀기</td></tr>
    <tr><td><b>Gemini CLI</b></td><td>압축을 푼 뒤 <code>gemini skills install 폴더경로</code></td></tr>
    <tr><td><b>ChatGPT</b> (스킬 지원 요금제)</td><td>스킬 메뉴에서 ZIP 가져오기</td></tr>
    <tr><td><b>Gemini 앱·그 밖의 챗봇</b></td><td>ZIP 안 <code>SKILL.md</code> 내용을 Gem·프로젝트 지침에 붙여 넣고, ZIP을 대화에 함께 올림</td></tr>
  </tbody></table>
</section>

<section class="page">
  <p class="eyeline">3</p><h2>요청하는 법</h2>
  <p>피피티를 올리고 상황에 맞는 문장을 그대로 쓰거나 고쳐 쓰면 됩니다.</p>
  <table><thead><tr><th style="width:22mm">상황</th><th style="width:78mm">이렇게 말하기</th><th>결과</th></tr></thead><tbody>{asks}</tbody></table>

  <p class="eyeline" style="margin-top:9mm">4</p><h2>받는 결과물</h2>
  <ul class="out">
    <li><b>바뀐 피피티</b> — 원본은 그대로 두고 새 파일로 돌려줍니다.</li>
    <li><b>PDF</b> — 바뀐 피피티를 PDF로 내보낸 파일입니다. 폰트가 들어 있어 폰트가 없는 컴퓨터에서도 같은 모양으로 열리고, 접근성 태그도 붙어 있습니다.</li>
    <li><b>fonts.zip</b> — 실제로 쓴 폰트 파일과 라이선스(저작권 표시·OFL 원문). 피피티를 편집할 컴퓨터에 설치하세요.</li>
    <li><b>진단 요약</b> — “공공·정책 보고, 개조식 → 프리텐다드 ExtraBold/Regular. 대안: 고딕 A1”처럼 두세 줄로 적습니다.</li>
    <li><b>보정 내역</b> — 그래프·디자인에서 고친 것과 알려 주기만 한 것을 슬라이드별로 적습니다.</li>
    <li><b>글꼴 테마 파일</b>(요청 시) — PowerPoint ‘디자인 → 변형 → 글꼴’에 등록해 새 슬라이드에도 같은 글꼴을 쓰게 합니다.</li>
    <li><b>폰트 라이선스 내역서</b>(요청 시) — 폰트·굵기·라이선스·출처·파일 해시를 담은 공모전 제출·기관 납품용 문서입니다.</li>
  </ul>
</section>

<section class="page">
  <p class="eyeline">5</p><h2>폰트를 고르는 기준</h2>
  <p>분석 결과에 따라 아래 짝을 기본으로 삼고, 슬라이드 이미지를 보고 바꿀 이유가 있으면 바꿉니다.</p>
  <table><thead><tr><th>목적</th><th>제목</th><th>본문</th></tr></thead><tbody>{pairs}</tbody></table>
  <h3>함께 지키는 규칙</h3>
  <ul class="out">
    <li>한 덱에 글꼴 가족은 두 개까지. 같은 가족의 굵기 차이로 위계를 먼저 만듭니다.</li>
    <li>본문은 Regular 이상의 고딕·명조. 손글씨·디스플레이·픽셀 폰트는 제목에만 씁니다. Light는 프로젝터에서 흐려집니다.</li>
    <li>개조식 보고서는 □ ○ ※ Ⅰ ①을 모두 갖춘 폰트를 고릅니다. 빠진 기호가 있으면 미리 경고합니다.</li>
    <li>기관 전용 서체(서울·경기·부산·교육청·보훈 등)가 있으면 먼저 제안합니다.</li>
  </ul>
</section>

{exs}

<section class="page">
  <p class="eyeline">7</p><h2>자주 묻는 질문</h2>
  <dl class="faq">{faq}</dl>
  <div class="box" style="margin-top:8mm"><p>예시 그림은 스킬을 실제로 돌려 만든 PDF를 쪽마다 옮긴 것입니다. 원본 그림의 ‘2027 년’처럼 벌어진 간격은 원본 글꼴(맑은 고딕)이 없는 환경에서 대체 글꼴이 섞인 탓입니다.</p>
  <p>폰트 라이선스는 바뀔 수 있습니다. 쓰기 전에 각 폰트의 원문 라이선스를 확인하세요. 각 폰트의 저작권은 해당 제작사에 있습니다.</p></div>
</section>
</body></html>"""


def main():
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        subprocess.run([sys.executable, ROOT / "skill/scripts/fetch_fonts.py", "프리텐다드", "--weights", "Regular,SemiBold,Bold,ExtraBold",
                        "--out", td / "f", "--data", ROOT / "data"], check=True, capture_output=True)
        page = td / "manual.html"
        page.write_text(build_html(td / "f"), "utf-8")
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch(); pg = b.new_page()
            pg.goto(page.as_uri()); pg.wait_for_load_state("networkidle")
            pg.pdf(path=str(OUT), format="A4", print_background=True, prefer_css_page_size=True,
                   display_header_footer=True, header_template="<span></span>",
                   footer_template='<div style="width:100%;font-size:7pt;color:#8a857b;text-align:center"><span class="pageNumber"></span> / <span class="totalPages"></span></div>')
            b.close()
    print(OUT, OUT.stat().st_size)


if __name__ == "__main__":
    main()
