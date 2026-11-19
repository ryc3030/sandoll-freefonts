#!/usr/bin/env python3
"""사용방법 페이지(guide.html)의 전후 비교 예시를 다시 만든다.

  python3 docs/examples/make_examples.py      # LibreOffice·poppler·인터넷 필요
예시 덱 3개를 만들고 → 스킬 스크립트로 분석·폰트 받기·적용 → 원본과 결과를 같은 조건(빈 HOME)에서 렌더링한다.
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

DECKS = [
    dict(key="public", title="공공 정책 보고", author="행정안전부 지방재정정책관", accent="1F4E79",
         slides=[("2027년 지방재정 운용 방향", ["지방재정정책관 · 2026. 11."]),
                 ("Ⅰ. 추진 배경", ["□ 지방세수 둔화로 세출 구조조정 필요", "  ○ 의무지출 비중 68%로 3년 새 5%p 증가",
                                 "  ○ 자율 재원 비중은 같은 기간 4%p 감소", "□ 성과평가 결과를 다음 연도 예산에 반영",
                                 "※ 관계 부처 협의 후 12월 국무회의 보고"])]),
    dict(key="memorial", title="현충일 추념식", author="서울지방보훈청", accent="2B2B2B",
         slides=[("제72회 현충일 추념식", ["나라를 위해 희생하신 분들을 기억합니다"]),
                 ("식순", ["개식", "국민의례", "헌화 및 분향", "추념사", "추모 공연", "폐식"])]),
    dict(key="kids", title="어린이 교통안전 교육", author="서울특별시교육청 학생안전과", accent="E8743B",
         slides=[("어린이 교통안전, 이렇게 지켜요", ["서울특별시교육청 · 초등 저학년용"]),
                 ("우리 반 안전 약속", ["횡단보도에서는 손을 들고 건너요", "초록불이 켜져도 좌우를 살펴요",
                                     "등하굣길에는 보호자와 함께해요", "놀이는 차도에서 멀리 떨어져서 해요"])]),
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
    p = Presentation(); p.slide_width, p.slide_height = Inches(13.333), Inches(7.5)
    p.core_properties.author = d["author"]; p.core_properties.title = d["slides"][0][0]
    acc = RGBColor.from_string(d["accent"]); grey = RGBColor(0x6B, 0x6B, 0x6B)
    for i, (t, lines) in enumerate(d["slides"]):
        s = p.slides.add_slide(p.slide_layouts[0 if i == 0 else 1])
        title, body = s.shapes.title, s.placeholders[1]
        if i == 0:  # 표지: 왼쪽 색 띠 + 왼쪽 정렬 큰 제목
            band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.35), p.slide_height)
            band.fill.solid(); band.fill.fore_color.rgb = acc; band.line.fill.background()
            title.left, title.top, title.width, title.height = Inches(1.1), Inches(2.3), Inches(11.2), Inches(1.6)
            body.left, body.top, body.width, body.height = Inches(1.1), Inches(4.0), Inches(11.2), Inches(1.0)
        else:      # 본문: 위쪽 색 줄 + 제목 + 본문 + 바닥글
            bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, p.slide_width, Inches(0.18))
            bar.fill.solid(); bar.fill.fore_color.rgb = acc; bar.line.fill.background()
            title.left, title.top, title.width, title.height = Inches(0.9), Inches(0.55), Inches(11.5), Inches(1.1)
            body.left, body.top, body.width, body.height = Inches(0.9), Inches(1.9), Inches(11.5), Inches(4.8)
            foot = s.shapes.add_textbox(Inches(0.9), Inches(6.85), Inches(11.5), Inches(0.4))
            foot.text_frame.word_wrap = True
            foot.text_frame.text = f"{d['author']}  |  {i + 1}"
            foot.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
            foot.text_frame.paragraphs[0].runs[0].font.color.rgb = grey
            foot.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
            _font(foot)
        title.text = t
        tf = body.text_frame
        for j, l in enumerate(lines):
            para = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            if l.startswith("  "): para.level = 1
            para.text = l.strip()
        for para in title.text_frame.paragraphs:
            para.alignment = PP_ALIGN.LEFT
            for r in para.runs:
                r.font.size = Pt(54 if i == 0 else 36); r.font.color.rgb = acc if i == 0 else RGBColor(0x1A, 0x1A, 0x1A)
        for para in tf.paragraphs:
            if i == 0:
                para.alignment = PP_ALIGN.LEFT
            for r in para.runs:
                r.font.size = Pt(24 if i == 0 else (22 if para.level else 26))
                if i == 0: r.font.color.rgb = grey
        if i == 0: body.text_frame.paragraphs[0]._p.get_or_add_pPr().set("indent", "0")
        _font(title); _font(body)
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
        for d in DECKS:
            src, dst, fonts = td / f"{d['key']}.pptx", td / f"{d['key']}-out.pptx", td / f"fonts-{d['key']}"
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
            subprocess.run([sys.executable, S / "apply_fonts.py", src, dst, "--manifest", fonts / "manifest.json",
                            "--title", f"{tn}:{tw}", "--body", f"{bn}:{bw}"], check=True, capture_output=True)
            render(src, f"{d['key']}-before"); render(dst, f"{d['key']}-after", fonts)
            diag = a["analysis"]
            summary.append(dict(key=d["key"], title=d["title"], author=d["author"], pages=len(d["slides"]),
                                genre=diag["genres"][0][0] if diag["genres"] else "", way=pair["방식"],
                                titleFont=pair["제목"], bodyFont=pair["본문"], 
                                agency=[dict(agency=x["agency"], fonts=x["fonts"], url=x["url"]) for x in a["recommendation"]["agency_fonts"]],
                                tags=diag.get("tags", [])[:6], register=diag.get("register", ""),
                                alt=[f"{x['제목']} / {x['본문']}" for x in a["recommendation"]["pairs"][1:3]]))
            print(d["key"], pair["방식"], pair["제목"], "/", pair["본문"])
    (OUT / "examples.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", "utf-8")


if __name__ == "__main__":
    main()
