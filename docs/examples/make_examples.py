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
                 dict(t="Ⅲ. 연도별 지방세 수입(조 원)", chart=("bar", ["2022", "2023", "2024", "2025", "2026"], {"지방세 수입": (98.4, 104.1, 109.0, 112.4, 115.1)}), note=NOTE)]),
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
                                                        {"비율": (46, 22, 18, 14)}), note=NOTE),
                 dict(t="우리 반 안전 약속", b=["횡단보도에서는 손을 들고 건너요", "초록불이 켜져도 좌우를 살펴요",
                                             "등하굣길에는 보호자와 함께해요", "놀이는 차도에서 멀리 떨어져서 해요"])]),
    dict(key="report", title="분기 실적 보고(그래프·디자인 보정)", author="에이크미 경영기획팀", messy=True, level="bold",
         ask="폰트 바꾸고, 그래프랑 디자인도 확 바꿔 줘", slides=[None] * 4),
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
            ct = XL_CHART_TYPE.PIE if kind == "pie" else XL_CHART_TYPE.COLUMN_CLUSTERED
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
            subprocess.run([sys.executable, S / "polish_deck.py", mid, "--fix", fixed, "--manifest", fonts / "manifest.json",
                            "--contrast", "--json", pj], check=True, capture_output=True)
            subprocess.run([sys.executable, S / "style_deck.py", fixed, "--out", dst, "--level", d["level"], "--json", sj,
                            "--data", ROOT / "data"], check=True, capture_output=True)
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
