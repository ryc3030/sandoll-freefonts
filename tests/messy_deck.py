"""그래프·디자인 보정 시험용 덱: 흔한 실수를 일부러 넣는다(polish_deck.py 시험과 사용 예시에 씀)."""
import io
from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
OLD = "맑은 고딕"


def _font(shape, size=None, color=None):
    for para in shape.text_frame.paragraphs:
        for r in para.runs:
            r.font.name = OLD
            if size:
                r.font.size = Pt(size)
            if color:
                r.font.color.rgb = RGBColor.from_string(color)
            rpr = r._r.get_or_add_rPr()
            rpr.append(rpr.makeelement(f"{{{A}}}ea", {"typeface": OLD}))


def _to_3d(chart, kind):
    """python-pptx는 3D 그래프를 만들지 못해 2D 그래프 XML을 3D로 바꾼다."""
    cs = chart._chartSpace
    plot = cs.find(f".//{{{C}}}plotArea")
    ch = plot.getparent()
    el = plot.find(f"{{{C}}}{kind}Chart")
    el.tag = f"{{{C}}}{kind}3DChart"
    view = etree.SubElement(ch, f"{{{C}}}view3D")
    for t, v in (("rotX", "15"), ("rotY", "20"), ("rAngAx", "1")):
        etree.SubElement(view, f"{{{C}}}{t}").set("val", v)
    ch.remove(view); ch.insert(list(ch).index(plot), view)
    if kind == "bar":
        g = el.find(f"{{{C}}}grouping")
        ax = el.findall(f"{{{C}}}axId")
        gd = etree.Element(f"{{{C}}}gapDepth"); gd.set("val", "150")
        shp = etree.Element(f"{{{C}}}shape"); shp.set("val", "box")
        for x in el.findall(f"{{{C}}}overlap"):
            el.remove(x)
        ax[0].addprevious(gd); ax[0].addprevious(shp)
        new = etree.Element(f"{{{C}}}axId"); new.set("val", "99999"); ax[-1].addnext(new)
        ser_ax = etree.fromstring(f'<c:serAx xmlns:c="{C}"><c:axId val="99999"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
                                  f'<c:delete val="1"/><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="{ax[1].get("val")}"/>'
                                  f'<c:crosses val="autoZero"/></c:serAx>')
        plot.findall(f"{{{C}}}valAx")[-1].addnext(ser_ax)


def build(path, image_bytes=None):
    p = Presentation(); p.slide_width, p.slide_height = Inches(13.333), Inches(7.5)
    p.core_properties.author = "에이크미 경영기획팀"
    # 1 표지
    s = p.slides.add_slide(p.slide_layouts[0])
    s.shapes.title.text = "2026년 3분기 사업 실적"
    s.placeholders[1].text = "경영기획팀 · 2026. 10."
    _font(s.shapes.title, 44); _font(s.placeholders[1], 20)
    # 2 3D 막대 그래프: 축이 90에서 시작, 9pt 글자, 보조 눈금선, 계열 하나인데 범례
    s = p.slides.add_slide(p.slide_layouts[5])
    s.shapes.title.text = "분기별 매출(억 원)"; _font(s.shapes.title, 32)
    s.shapes.title.left, s.shapes.title.top = Inches(0.6), Inches(0.4)
    cd = CategoryChartData(); cd.categories = ["1분기", "2분기", "3분기", "4분기(전망)"]; cd.add_series("매출", (92, 95, 98, 101))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.6), Inches(11), Inches(5.4), cd)
    ch = gf.chart
    ch.has_legend = True
    va = ch.value_axis; va.minimum_scale = 90; va.has_minor_gridlines = True
    ch.font.size = Pt(9); ch.font.name = OLD
    _to_3d(ch, "bar")
    # 3 사진 늘어남, 넘치는 글, 낮은 대비, 작은 글자, 제목 위치 어긋남
    s = p.slides.add_slide(p.slide_layouts[5])
    s.shapes.title.text = "현장 점검 결과"; _font(s.shapes.title, 32)
    s.shapes.title.left, s.shapes.title.top = Inches(0.72), Inches(0.46)
    if image_bytes:
        s.shapes.add_picture(io.BytesIO(image_bytes), Inches(0.8), Inches(1.7), Inches(6.2), Inches(2.6))  # 원본 4:3을 2.4:1로 늘림
    tb = s.shapes.add_textbox(Inches(7.4), Inches(1.7), Inches(5.2), Inches(2.2))
    tb.text_frame.word_wrap = True
    from pptx.enum.text import MSO_AUTO_SIZE
    tb.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    tb.text_frame.text = ("3개 공장 모두 안전 점검을 통과했고, 2공장은 설비 교체 뒤 불량률이 1.8%에서 0.9%로 낮아졌습니다. "
                          "다만 1공장 포장 라인은 인력 부족으로 야간 가동률이 72%에 머물러 4분기에 인력 충원과 자동화 설비 도입을 함께 검토합니다.")
    _font(tb, 20)
    note = s.shapes.add_textbox(Inches(7.4), Inches(4.3), Inches(5.2), Inches(0.6))
    note.text_frame.text = "※ 점검 기간: 9월 1일~9월 20일, 외부 감리 동행"; _font(note, 11, "C8C8C8")
    # 4 원그래프 조각 8개(3D), 왼쪽 끝 어긋남
    s = p.slides.add_slide(p.slide_layouts[5])
    s.shapes.title.text = "제품군별 매출 비중"; _font(s.shapes.title, 32)
    cd = CategoryChartData(); cd.categories = ["A", "B", "C", "D", "E", "F", "G", "H"]; cd.add_series("비중", (31, 22, 14, 11, 8, 6, 5, 3))
    gf = s.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(0.6), Inches(1.6), Inches(6), Inches(5.4), cd)
    gf.chart.has_legend = True; gf.chart.font.size = Pt(10)
    _to_3d(gf.chart, "pie")
    t1 = s.shapes.add_textbox(Inches(7.2), Inches(2.0), Inches(5.4), Inches(0.8)); t1.text_frame.text = "A·B 두 제품군이 매출의 53%"
    t2 = s.shapes.add_textbox(Inches(7.28), Inches(3.0), Inches(5.4), Inches(0.8)); t2.text_frame.text = "하위 4개 제품군은 합쳐도 22%"
    t3 = s.shapes.add_textbox(Inches(7.2), Inches(4.0), Inches(5.4), Inches(0.8)); t3.text_frame.text = "4분기에 제품군 정리 검토"
    for t in (t1, t2, t3):
        _font(t, 20)
    p.save(path)


def sample_image():
    """4:3 견본 사진(그라데이션과 도형)."""
    from PIL import Image, ImageDraw
    im = Image.new("RGB", (800, 600), (40, 90, 140))
    d = ImageDraw.Draw(im)
    for y in range(600):
        d.line([(0, y), (800, y)], fill=(40 + y // 8, 90 + y // 10, 140 + y // 12))
    d.ellipse([300, 200, 500, 400], fill=(240, 200, 80))
    d.rectangle([80, 380, 720, 560], fill=(60, 60, 70))
    for x in range(100, 700, 90):
        d.rectangle([x, 320, x + 50, 380], fill=(200, 200, 210))
    b = io.BytesIO(); im.save(b, "PNG"); return b.getvalue()
