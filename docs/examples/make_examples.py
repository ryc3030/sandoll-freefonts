#!/usr/bin/env python3
"""사용방법 페이지(guide.html)의 전후 비교 예시를 다시 만든다.

  python3 docs/examples/make_examples.py      # LibreOffice·poppler·인터넷 필요
예시 덱 4개를 만들고 → 스킬 스크립트로 분석·폰트 받기·적용·그래프 보정·디자인 다듬기 → 원본과 결과를 같은 조건(빈 HOME)에서 렌더링한다.
결과: docs/examples/<이름>-before-N.jpg, -after-N.jpg, examples.json(진단 요약)
렌더링은 스킬의 export_pdf.py를 그대로 쓴다(사용자가 받는 PDF와 같은 방식).
"""
import json, os, shutil, subprocess, sys, tempfile
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent
S = ROOT / "skill/scripts"
OLD = "맑은 고딕"

NOTE = "※ 예시용 가상 수치"
DECKS = [
    dict(key="public", title="공공 정책 보고", author="행정안전부 지방재정정책관", level="standard",
         ask="어울리는 무료 폰트로 바꾸고 디자인도 다듬어 줘",
         slides=[dict(t="2027년 지방재정 운용 방향", sub="지방재정정책관 · 2026. 11."),
                 dict(t="Ⅰ. 추진 배경", b=["□ 지방세수 둔화로 세출 구조조정 필요", "  ○ 의무지출 비중 68%로 3년 새 5%p 증가",
                                          "  ○ 자율 재원 비중은 같은 기간 4%p 감소", "□ 성과평가 결과를 다음 연도 예산에 반영",
                                          "※ 관계 부처 협의 후 12월 국무회의 보고"]),
                 dict(t="Ⅱ. 2027년 재원별 편성 현황(조 원)", table=[["구분", "2026년", "2027년", "증감"],
                                                                 ["지방세", "112.4", "115.1", "+2.7"], ["지방교부세", "66.8", "68.0", "+1.2"],
                                                                 ["국고보조금", "79.5", "81.3", "+1.8"], ["합계", "258.7", "264.4", "+5.7"]], note=NOTE),
                 dict(t="Ⅲ. 연도별 지방세 수입(조 원)", chart=("bar", ["2022", "2023", "2024", "2025", "2026"], {"지방세 수입": (98.4, 104.1, 109.0, 112.4, 115.1)}), note=NOTE,
                      alt="2022년 98.4조 원에서 2026년 115.1조 원으로 해마다 늘어난 지방세 수입 막대 그래프")]),
    dict(key="memorial", title="현충일 추념식", author="서울지방보훈청", level="standard",
         ask="어울리는 무료 폰트로 바꾸고 디자인도 다듬어 줘",
         slides=[dict(t="제72회 현충일 추념식", sub="나라를 위해 희생하신 분들을 기억합니다"),
                 dict(t="행사 개요", table=[["구분", "내용"], ["일시", "6월 6일(토) 10:00"], ["장소", "국립서울현충원 현충문 앞"],
                                          ["주관", "서울지방보훈청"], ["참석", "보훈 가족, 시민 등 1,000여 명"]], note="※ 예시용 가상 일정"),
                 dict(t="식순", b=["개식", "국민의례", "헌화 및 분향", "추념사", "추모 공연", "폐식"])]),
    dict(key="kids", title="어린이 교통안전 교육", author="서울특별시교육청 학생안전과", level="bold",
         ask="아이들이 보기 좋게 폰트와 디자인을 확 바꿔 줘",
         slides=[dict(t="어린이 교통안전, 이렇게 지켜요", sub="서울특별시교육청 · 초등 저학년용"),
                 dict(t="등굣길 사고는 언제 일어날까요?", chart=("pie", ["길을 건널 때", "차 사이에서 뛰어나올 때", "놀다가", "그 밖에"],
                                                        {"비율": (46, 22, 18, 14)}), note=NOTE,
                      alt="등굣길 사고 원인 원그래프: 길을 건널 때 46%, 차 사이에서 뛰어나올 때 22%, 놀다가 18%, 그 밖에 14%"),
                 dict(t="우리 반 안전 약속", b=["횡단보도에서는 손을 들고 건너요", "초록불이 켜져도 좌우를 살펴요",
                                             "등하굣길에는 보호자와 함께해요", "놀이는 차도에서 멀리 떨어져서 해요"])]),
    dict(key="academic", title="학술대회 발표", author="한국행정학회 추계학술대회 발표자", level="standard",
         ask="학회 발표 자료야. 어울리는 폰트로 바꾸고 디자인도 다듬어 줘",
         slides=[dict(t="주민참여예산제의 제도화와 참여의 질", sub="광역자치단체 17곳 패널 분석(2018~2025)"),
                 dict(t="연구 질문과 가설", b=["연구 질문: 제도화 수준이 높을수록 참여의 질도 높아지는가?",
                                          "가설 1. 위원회 상설화는 제안 채택률을 높인다", "가설 2. 예산 한도 확대는 제안 수를 늘린다",
                                          "선행연구(Wampler, 2007; 곽채기, 2005)는 제도 설계의 효과를 주로 사례로 확인"]),
                 dict(t="제안 채택률 추이(%)", chart=("line", ["2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025"],
                                                {"상설 위원회": (31, 34, 36, 38, 41, 43, 44, 47), "비상설 위원회": (28, 29, 29, 30, 31, 31, 32, 33)}),
                      note="※ 예시용 가상 수치", alt="2018~2025년 제안 채택률 선 그래프: 상설 위원회는 31%에서 47%로, 비상설 위원회는 28%에서 33%로 늘었다"),
                 dict(t="고정효과 모형 추정 결과", table=[["변수", "계수", "표준오차", "p값"], ["위원회 상설화", "0.142", "0.041", "0.001"],
                                                    ["예산 한도(로그)", "0.087", "0.052", "0.096"], ["재정자립도", "0.013", "0.009", "0.152"],
                                                    ["관측치", "136", "", ""]], note="※ 예시용 가상 추정치")]),
    dict(key="startup", title="스타트업 투자 설명", author="주식회사 루프랩 대표", level="standard",
         ask="투자 설명회 자료야. 폰트랑 디자인 깔끔하게 정리해 줘",
         slides=[dict(t="루프랩 시리즈 A 투자 제안", sub="동네 재사용 용기 순환 플랫폼 · 2026. 11."),
                 dict(t="문제와 해결", b=["배달 용기 하루 830만 개가 한 번 쓰고 버려진다", "루프랩은 동네 가게에 재사용 용기를 빌려주고 회수·세척까지 맡는다",
                                       "가게는 용기값을 30% 아끼고, 이용자는 반납만 하면 된다"]),
                 dict(t="월 이용 건수(만 건)", chart=("bar", ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월"],
                                               {"이용 건수": (1.2, 1.6, 2.3, 3.1, 3.9, 5.2, 6.4, 7.8, 9.5, 11.6)}),
                      note="※ 예시용 가상 수치", alt="2026년 1월 1.2만 건에서 10월 11.6만 건으로 꾸준히 늘어난 월 이용 건수 막대 그래프"),
                 dict(t="투자금 사용 계획", chart=("pie", ["세척 센터 2곳", "영업·제휴", "개발", "운영 자금"], {"비중": (45, 25, 20, 10)}),
                      note="※ 예시용 가상 계획", alt="투자금 사용 계획 원그래프: 세척 센터 2곳 45%, 영업·제휴 25%, 개발 20%, 운영 자금 10%")]),
    dict(key="exhibit", title="전시 소개", author="시립미술관 학예팀", level="bold",
         ask="전시 소개 자료야. 분위기에 맞게 폰트와 디자인을 확 바꿔 줘",
         slides=[dict(t="빛의 기록: 한국 근대 사진 100년", sub="시립미술관 기획전시 · 2027. 3. 2.~6. 28."),
                 dict(t="전시 구성", b=["1부 낯선 기계, 사진관의 탄생", "2부 거리로 나간 카메라", "3부 기록과 증언의 사진",
                                     "4부 일상의 얼굴들"]),
                 dict(t="관람 안내", table=[["구분", "내용"], ["관람 시간", "화~일 10:00~19:00(월요일 휴관)"], ["관람료", "무료"],
                                         ["도슨트", "매일 11:00, 15:00"], ["장소", "본관 2층 전시실"]], note="※ 예시용 가상 일정")]),
    dict(key="festival", title="축제 홍보", author="강원특별자치도 관광마케팅과", level="bold",
         ask="도청 주최 축제 홍보 자료야. 눈에 띄게 폰트와 디자인을 확 바꿔 줘",
         slides=[dict(t="2027 강원 겨울 빛 축제", sub="1. 15.(금)~2. 14.(일) · 춘천 의암호 일원"),
                 dict(t="이렇게 즐겨요", b=["호수 위 빛 터널과 미디어 아트", "얼음 위 썰매·스케이트장", "지역 먹거리 장터와 공연",
                                         "매일 저녁 7시 불꽃놀이"]),
                 dict(t="지난해 방문객(만 명)", chart=("bar", ["2023", "2024", "2025", "2026"], {"방문객": (38, 45, 52, 61)}),
                      note="※ 예시용 가상 수치", alt="2023년 38만 명에서 2026년 61만 명으로 늘어난 축제 방문객 막대 그래프")]),
    dict(key="lecture", title="대학 강의", author="OO대학교 행정학과", level="standard",
         ask="수업 자료야. 학생들이 보기 편하게 폰트랑 디자인 다듬어 줘",
         slides=[dict(t="행정학원론 7주차: 정책 과정과 정책 도구", sub="OO대학교 행정학과 · 2027학년도 1학기"),
                 dict(t="오늘 배울 것", b=["정책 과정의 다섯 단계", "정책 도구의 네 가지 유형(규제·유인·정보·직접 공급)",
                                         "사례: 일회용 컵 보증금제는 왜 흔들렸나", "토론: 우리 학교에 맞는 정책 도구는?"]),
                 dict(t="정책 도구별 장단점", table=[["도구", "장점", "단점"], ["규제", "효과가 빠르고 분명함", "반발과 집행 비용"],
                                               ["유인", "자발적 참여", "예산이 듦"], ["정보", "비용이 적음", "효과가 더딤"],
                                               ["직접 공급", "품질 관리", "정부 규모가 커짐"]])]),
    dict(key="report", title="분기 실적 보고(그래프·디자인 보정)", author="에이크미 경영기획팀", messy=True, level="bold", image_fit="crop",
         ask="폰트 바꾸고, 그래프랑 디자인도 확 바꿔 줘", slides=[None] * 4,
         alt={2: "1분기 92억 원에서 4분기 전망 101억 원으로 늘어나는 분기별 매출 막대 그래프",
              3: "점검을 마친 공장 설비와 지붕 위로 뜬 해를 그린 그림",
              4: "제품군 A~H의 매출 비중 원그래프: A 31%, B 22%, C 14%, D 11%, 나머지 네 제품군 22%"}),
]


def _font(shape):
    """원본 덱처럼 모든 글자에 '맑은 고딕'을 직접 지정한다(라틴·한글 모두)."""
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    for para in shape.text_frame.paragraphs:
        for r in para.runs:
            r.font.name = OLD
            rpr = r._r.get_or_add_rPr()
            rpr.append(rpr.makeelement(A + "ea", {"typeface": OLD}))


def build(d, path):
    """흔한 기본 서식 그대로인 원본 덱: 기본 Office 테마, 가운데 제목, 기본 표·그래프 스타일, 맑은 고딕."""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    p = Presentation(); p.slide_width, p.slide_height = Inches(13.333), Inches(7.5)
    p.core_properties.author = d["author"]; p.core_properties.title = d["slides"][0]["t"]
    for i, sp in enumerate(d["slides"]):
        if i == 0:
            s = p.slides.add_slide(p.slide_layouts[0])
            s.shapes.title.text = sp["t"]; s.placeholders[1].text = sp["sub"]
            for shp, (y, h) in ((s.shapes.title, (2.4, 1.5)), (s.placeholders[1], (4.0, 0.9))):
                shp.left, shp.top, shp.width, shp.height = Inches(1), Inches(y), Inches(11.33), Inches(h)
            _font(s.shapes.title); _font(s.placeholders[1])
            continue
        layout = p.slide_layouts[1] if "b" in sp else p.slide_layouts[5]
        s = p.slides.add_slide(layout)
        s.shapes.title.text = sp["t"]
        s.shapes.title.left, s.shapes.title.top, s.shapes.title.width, s.shapes.title.height = Inches(0.6), Inches(0.45), Inches(12.1), Inches(1.0)
        _font(s.shapes.title)
        if "b" in sp:
            body = s.placeholders[1]
            body.left, body.top, body.width, body.height = Inches(0.6), Inches(1.7), Inches(12.1), Inches(5.2)
            tf = body.text_frame
            for j, l in enumerate(sp["b"]):
                para = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                if l.startswith("  "):
                    para.level = 1
                para.text = l.strip()
            _font(body)
        if "table" in sp:
            rows, cols = len(sp["table"]), len(sp["table"][0])
            gf = s.shapes.add_table(rows, cols, Inches(0.9), Inches(1.8), Inches(11.5), Inches(0.6) * rows)
            for r, row in enumerate(sp["table"]):
                for c, v in enumerate(row):
                    cell = gf.table.cell(r, c); cell.text = v
                    for para in cell.text_frame.paragraphs:
                        for run in para.runs:
                            run.font.size = Pt(18); run.font.name = OLD
        if "chart" in sp:
            kind, cats, series = sp["chart"]
            cd = CategoryChartData(); cd.categories = cats
            for name, vals in series.items():
                cd.add_series(name, vals)
            ct = {"pie": XL_CHART_TYPE.PIE, "line": XL_CHART_TYPE.LINE_MARKERS, "hbar": XL_CHART_TYPE.BAR_CLUSTERED}.get(kind, XL_CHART_TYPE.COLUMN_CLUSTERED)
            gf = s.shapes.add_chart(ct, Inches(0.9), Inches(1.6), Inches(11.5), Inches(5.1), cd)
            gf.chart.has_legend = True
        if sp.get("note"):
            tb = s.shapes.add_textbox(Inches(0.9), Inches(6.85), Inches(6), Inches(0.4))
            tb.text_frame.text = sp["note"]
            for r in tb.text_frame.paragraphs[0].runs:
                r.font.size = Pt(12)
            _font(tb)
    p.save(path)


def render(pptx, prefix, fonts_dir=None):
    """스킬의 export_pdf.py로 PDF를 만들고(빈 HOME, 한글-숫자 간격 보정) 쪽마다 JPG로 뽑는다."""
    pdf = Path(pptx).with_suffix(".pdf")
    cmd = [sys.executable, S / "export_pdf.py", pptx, "--out", pdf]
    if fonts_dir:
        cmd += ["--fonts", fonts_dir]
    subprocess.run(cmd, check=True, capture_output=True)
    subprocess.run(["pdftoppm", "-r", "72", "-jpeg", "-jpegopt", "quality=82", str(pdf), str(OUT / prefix)], check=True)
    for f in OUT.glob(prefix + "-*.jpg"):  # pdftoppm이 붙이는 0 채움(-1/-01) 정리
        f.rename(OUT / f"{prefix}-{int(f.stem.rsplit('-', 1)[1])}.jpg")


def main():
    summary = []
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        for d in sorted(DECKS, key=lambda d: d.get("messy", False)):   # 보정 예시는 맨 뒤
            src, dst, fonts = td / f"{d['key']}.pptx", td / f"{d['key']}-out.pptx", td / f"fonts-{d['key']}"
            if d.get("messy"):   # 그래프·디자인 실수를 일부러 넣은 덱(tests/messy_deck.py)
                sys.path.insert(0, str(ROOT / "tests"))
                import messy_deck
                messy_deck.build(src, messy_deck.sample_image())
            else:
                build(d, src)
            aj = td / f"{d['key']}.json"
            subprocess.run([sys.executable, S / "analyze_deck.py", src, "--json", aj, "--data", ROOT / "data"], check=True, capture_output=True)
            a = json.loads(aj.read_text())
            pair = a["recommendation"]["pairs"][0]
            (tn, tw), (bn, bw) = pair["제목"].rsplit(" ", 1), pair["본문"].rsplit(" ", 1)
            fams = list(dict.fromkeys([tn, bn]))
            weights = sorted({tw, bw, "Bold"})
            subprocess.run([sys.executable, S / "fetch_fonts.py", *fams, "--weights", ",".join(weights), "--out", fonts,
                            "--data", ROOT / "data"], check=True, capture_output=True)
            mid = td / f"{d['key']}-fonted.pptx"
            subprocess.run([sys.executable, S / "apply_fonts.py", src, mid, "--manifest", fonts / "manifest.json",
                            "--title", f"{tn}:{tw}", "--body", f"{bn}:{bw}"], check=True, capture_output=True)
            pj, sj = td / f"{d['key']}-polish.json", td / f"{d['key']}-style.json"
            fixed = td / f"{d['key']}-fixed.pptx"
            # 대체 텍스트: 스킬을 쓰는 AI가 그림을 보고 쓰는 설명을 예시에서는 미리 적어 둔다
            mj = td / f"{d['key']}-media.json"
            subprocess.run([sys.executable, S / "polish_deck.py", mid, "--json", mj], check=True, capture_output=True)
            alts = d.get("alt") or {i + 1: sp.get("alt") for i, sp in enumerate(d["slides"]) if sp and sp.get("alt")}
            altj = td / f"{d['key']}-alt.json"
            altj.write_text(json.dumps([{"slide": m["slide"], "name": m["name"], "alt": alts[m["slide"]]}
                                        for m in json.loads(mj.read_text())["media"] if alts.get(m["slide"])], ensure_ascii=False), "utf-8")
            subprocess.run([sys.executable, S / "polish_deck.py", mid, "--fix", fixed, "--manifest", fonts / "manifest.json",
                            "--contrast", "--alt", altj, "--json", pj, "--image-fit", d.get("image_fit", "contain")], check=True, capture_output=True)
            subprocess.run([sys.executable, S / "style_deck.py", fixed, "--out", dst, "--level", d["level"], "--json", sj,
                            "--manifest", fonts / "manifest.json", "--data", ROOT / "data"], check=True, capture_output=True)
            pol = json.loads(pj.read_text())
            polish = [f"{k}장: {m}" for k, ms in sorted(pol["slides"].items(), key=lambda t: int(t[0])) for m in ms] + pol["deck"]
            style = json.loads(sj.read_text())["notes"]
            render(src, f"{d['key']}-before"); render(dst, f"{d['key']}-after", fonts)
            diag = a["analysis"]
            summary.append(dict(key=d["key"], title=d["title"], author=d["author"], pages=len(d["slides"]), polish=polish, style=style, ask=d["ask"], level=d["level"],
                                genre=diag["genres"][0][0] if diag["genres"] else "", way=pair["방식"],
                                titleFont=pair["제목"], bodyFont=pair["본문"], 
                                agency=[dict(agency=x["agency"], fonts=x["fonts"], url=x["url"]) for x in a["recommendation"]["agency_fonts"]],
                                tags=diag.get("tags", [])[:6], register=diag.get("register", ""),
                                alt=[f"{x['제목']} / {x['본문']}" for x in a["recommendation"]["pairs"][1:3]]))
            print(d["key"], pair["방식"], pair["제목"], "/", pair["본문"])
    (OUT / "examples.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", "utf-8")


if __name__ == "__main__":
    main()
