#!/usr/bin/env python3
"""피피티를 분야에 맞는 디자인 프리셋으로 다듬는다(색 체계·글자 단계·그래프·표).

  python style_deck.py output.pptx --out styled.pptx                       # 분야를 분석해 프리셋 자동 선택(표준 강도)
  python style_deck.py output.pptx --out styled.pptx --preset public       # 프리셋 지정
  python style_deck.py output.pptx --out styled.pptx --level bold          # 직접 지정한 도형 색까지 프리셋 색으로
  python style_deck.py output.pptx --out styled.pptx --chart-palette colorblind   # 그래프만 색각 이상 친화 팔레트
  python style_deck.py --list                                              # 프리셋 목록

프리셋(data/design-tokens.json)은 공개 디자인 자원에서 값을 가져왔다.
  공공·추모: KRDS 범정부 UI/UX 디자인 시스템 토큰(행정안전부, 공공누리 제1유형)
  기업 보고: Urban Institute 그래프 스타일(vega-themes, BSD-3-Clause)
  IT·문화·아동·홍보·교육·게임: Open Color(MIT)   학술: ColorBrewer Dark2(Apache-2.0)   색각 이상 친화: Okabe-Ito

강도
  light     그래프(Urban Institute 방식: 테두리·값 축선·눈금 표시 없이 옅은 눈금선, 막대 간격 좁게, 막대 하나 계열은 값 표시)와
            표(제목 줄 진한 바탕, 옅은 줄무늬, 세로 선 없이 가로 선만, 숫자 열 오른쪽 정렬), 본문 행간 1.2·문단 간격
  standard  light + 테마 색 체계(기본 Office 테마일 때만), 내용 슬라이드 제목의 크기·색·왼쪽 정렬·자간(-2%), 글머리표 색,
            꾸밈 없는 기본 서식 덱이면 장식(표지 왼쪽 띠, 내용 슬라이드 위쪽 가는 띠, 쪽 번호)
  bold      standard + 브랜드 테마여도 색 체계 교체·장식, 도형에 직접 칠한 색을 제목색·강조색 중 가까운 쪽으로

색 원칙: 색은 세 가지까지(제목색·강조색 하나·무채색). 그래프 계열이 넷 이상이면 강조색의 옅은 단계를 쓴다.
--manifest를 주면 두 줄 제목의 줄 길이를 고르게 나눈다(adobe/balance-text 방식). 장식 도형은 '장식용'으로 표시해 화면 낭독기가 건너뛴다.

브랜드 색을 지키려면 --level light, 위치·정렬을 건드리지 않으려면 --keep-align.
적용 뒤에는 polish_deck.py(실수 바로잡기)와 export_pdf.py(그림 확인)를 이어서 돌린다.
"""
import argparse, colorsys, json, os, re, sys
from collections import Counter
from pathlib import Path
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
import polish_deck as pdk  # noqa: E402
from polish_deck import Deck, q, ln, NS, A, P, C, shapes, xfrm_of, ph_type, title_shape, TITLE_PH  # noqa: E402

OFFICE_ACCENT1 = {"4472C4", "4F81BD", "156082", "5B9BD5", "0F9ED5", "4E67C8", "0070C0"}
PPR_ORDER = ["lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont",
             "buNone", "buAutoNum", "buChar", "buBlip", "tabLst", "defRPr", "extLst"]
RPR_ORDER = ["ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "effectLst", "effectDag", "highlight",
             "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]
SPPR_ORDER = ["xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "ln",
              "effectLst", "effectDag", "scene3d", "sp3d", "extLst"]
SER_ORDER = ["idx", "order", "tx", "spPr", "invertIfNegative", "pictureOptions", "marker", "explosion", "dPt", "dLbls", "trendline",
             "errBars", "cat", "val", "xVal", "yVal", "smooth", "shape", "bubbleSize", "bubble3D", "extLst"]
AX_ORDER = ["axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt", "majorTickMark",
            "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "crossBetween", "auto", "lblAlgn",
            "lblOffset", "majorUnit", "minorUnit", "noMultiLvlLbl", "dispUnits", "extLst"]
DLBLS_ORDER = ["dLbl", "delete", "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey", "showVal", "showCatName", "showSerName",
               "showPercent", "showBubbleSize", "separator", "showLeaderLines", "extLst"]
CHART_ORDER = ["title", "autoTitleDeleted", "pivotFmts", "view3D", "floor", "sideWall", "backWall", "plotArea", "legend",
               "plotVisOnly", "dispBlanksAs", "extLst"]
BAR_ORDER = ["barDir", "grouping", "varyColors", "ser", "dLbls", "gapWidth", "overlap", "serLines", "axId", "extLst"]
PIE_ORDER = ["varyColors", "ser", "dLbls", "firstSliceAng", "holeSize", "extLst"]
LEGEND_ORDER = ["legendPos", "legendEntry", "layout", "overlay", "spPr", "txPr", "extLst"]
TCPR_ORDER = ["lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D", "noFill", "solidFill", "gradFill", "blipFill",
              "pattFill", "grpFill", "headers", "extLst"]
NUM = re.compile(r"^[\s₩$€(]*[-+−△▲▼]?\s*[\d][\d,]*(\.\d+)?\s*(%p|%|원|억|만|천|조|건|명|개|배|점|년|회|곳|t|kg|km)?[\s)]*$")


def put(parent, tag, order, ns=A, attrs=None, replace=False):
    """order 순서를 지켜 자식 요소를 찾거나 만든다."""
    el = parent.find(f"{{{ns}}}{tag}")
    if el is not None and replace:
        parent.remove(el); el = None
    if el is None:
        el = etree.Element(f"{{{ns}}}{tag}")
        idx = order.index(tag) if tag in order else len(order)
        pos = len(parent)
        for i, ch in enumerate(parent):
            n = ln(ch)
            if n in order and order.index(n) > idx:
                pos = i; break
        parent.insert(pos, el)
    for k, v in (attrs or {}).items():
        el.set(k, str(v))
    return el


def set_fill(parent, order, color=None, scheme=None, ns_tag="solidFill"):
    """parent(spPr·rPr·tcPr 등)의 채우기를 하나의 단색으로."""
    for t in ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"):
        for x in parent.findall(f"a:{t}", NS):
            parent.remove(x)
    if color is None and scheme is None:
        put(parent, "noFill", order)
        return
    f = put(parent, "solidFill", order)
    c = etree.SubElement(f, q(A, "schemeClr" if scheme else "srgbClr"))
    c.set("val", scheme or color)


def set_line(sppr, order, color=None, w=None, scheme=None):
    l_ = put(sppr, "ln", order)
    for x in list(l_):
        if ln(x) in ("noFill", "solidFill", "gradFill", "pattFill", "prstDash"):
            l_.remove(x)
    if color is None and scheme is None:
        l_.insert(0, etree.Element(q(A, "noFill")))
    else:
        f = etree.Element(q(A, "solidFill")); c = etree.SubElement(f, q(A, "schemeClr" if scheme else "srgbClr")); c.set("val", scheme or color)
        l_.insert(0, f)
        if w:
            l_.set("w", str(w))
    return l_


def txpr(parent, order, size=None, color=None, bold=None, ns=C):
    t = parent.find(f"{{{ns}}}txPr")
    if t is None:
        t = put(parent, "txPr", order, ns=ns)
        etree.SubElement(t, q(A, "bodyPr")); etree.SubElement(t, q(A, "lstStyle"))
        p = etree.SubElement(t, q(A, "p")); ppr = etree.SubElement(p, q(A, "pPr")); etree.SubElement(ppr, q(A, "defRPr"))
        etree.SubElement(p, q(A, "endParaRPr")).set("lang", "ko-KR")
    for d in t.iter(q(A, "defRPr")):
        if size:
            d.set("sz", str(size))
        if bold is not None:
            d.set("b", "1" if bold else "0")
        if color:
            set_fill(d, RPR_ORDER, color=color)
    return t


def hexs(h):
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def nearest(hexv, palette):
    h, l_, s = colorsys.rgb_to_hls(*hexs(hexv))
    best, bd = palette[0], 9
    for c in palette:
        h2, l2, s2 = colorsys.rgb_to_hls(*hexs(c))
        d = min(abs(h - h2), 1 - abs(h - h2)) * 2 + abs(l_ - l2) * 0.5 + abs(s - s2) * 0.3
        if d < bd:
            best, bd = c, d
    return best


# ── 프리셋 고르기 ────────────────────────────────────────────────────────
def load_tokens(data_dir):
    import analyze_deck as ad
    return ad.load("design-tokens.json", data_dir)


def auto_preset(path, tokens, data_dir):
    import analyze_deck as ad
    d = ad.diagnose(ad.extract(path))
    top = d["genres"][0][0] if d["genres"] else None
    if d["publisher_public"] and (top is None or top not in ad.CONTENT_OVERRIDE):
        return "public", "공공 발행"
    if top in tokens["genre"]:
        return tokens["genre"][top], top
    return ("public", "공공 발행") if d["publisher_public"] else ("corporate", "분야 불확실 → 기업 보고")


# ── 테마 ────────────────────────────────────────────────────────────────
def style_theme(deck, preset, force, notes, dark=False):
    done = 0
    for n in sorted(x for x in deck.names if re.match(r"ppt/theme/theme\d+\.xml$", x)):
        t = deck.xml(n)
        cs = t.find(".//a:clrScheme", NS)
        if cs is None:
            continue
        a1 = cs.find("a:accent1/a:srgbClr", NS)
        default = (a1 is not None and a1.get("val", "").upper() in OFFICE_ACCENT1) or "office" in (cs.get("name") or "").lower()
        if not default and not force:
            notes.append(f"테마 색 체계가 기본 Office 테마가 아니라(브랜드 템플릿으로 보임) 그대로 둠 — 바꾸려면 --level bold")
            continue
        vals = {} if dark else {"dk1": preset["text"], "lt1": "FFFFFF", "dk2": preset["title"], "lt2": preset["light"]}
        vals.update({"hlink": preset["hlink"], "folHlink": preset["muted"]})
        for i, c in enumerate(preset["accents"][:6], 1):
            vals[f"accent{i}"] = c
        for k, v in vals.items():
            el = cs.find(f"a:{k}", NS)
            if el is None:
                continue
            for x in list(el):
                el.remove(x)
            etree.SubElement(el, q(A, "srgbClr")).set("val", v)
        cs.set("name", f"무료폰트 색인 · {preset['name']}")
        deck.touch(n); done += 1
    if done:
        notes.append(f"테마 색 체계를 '{preset['name']}'로 바꿈(제목 #{preset['title']}, 강조 #{preset['accents'][0]})")
    return done


# ── 글자 ────────────────────────────────────────────────────────────────
def style_text(deck, slides, preset, tokens, level, keep_align, notes, dark=False):
    W, H = deck.size()
    scale = tokens["type"]["16:9" if W / H > 1.5 else "4:3"]
    rules = tokens["rules"]
    st = Counter()
    dark_bg = dark
    for i, s in enumerate(slides, 1):
        root = deck.xml(s)
        t = title_shape(root)
        cover = i == 1 or (t is not None and ph_type(t) == "ctrTitle")
        if t is not None and level != "light":
            size = None if cover else scale["title"] * 100
            for p in t.iter(q(A, "p")):
                if not keep_align and not cover:
                    put(p, "pPr", ["pPr", "r", "br", "fld", "endParaRPr"]).set("algn", "l")
                for r in p.iter(q(A, "r")):
                    rp = r.find("a:rPr", NS)
                    if rp is None:
                        rp = etree.Element(q(A, "rPr")); r.insert(0, rp)
                    if size:
                        rp.set("sz", str(size))
                    pt = int(rp.get("sz") or (size or scale["cover"] * 100)) / 100
                    rp.set("spc", str(int(round(pt * 100 * rules["titleTracking"]))))
                    sf = rp.find("a:solidFill", NS)
                    dark = sf is None or (sf.find("a:srgbClr", NS) is not None and pdk.rel_lum(sf.find("a:srgbClr", NS).get("val")) < 0.05) \
                        or (sf.find("a:schemeClr", NS) is not None and sf.find("a:schemeClr", NS).get("val") in ("tx1", "dk1"))
                    if not dark_bg and (dark or level == "bold"):
                        set_fill(rp, RPR_ORDER, scheme="tx2")
            st["titles"] += 0 if cover else 1
            deck.touch(s)
        # 본문 행간·문단 간격·글머리표 색
        for sh in shapes(root):
            if ln(sh) != "sp" or sh is t or sh.find("p:txBody", NS) is None:
                continue
            paras = [p for p in sh.findall("p:txBody/a:p", NS) if "".join(x.text or "" for x in p.iter(q(A, "t"))).strip()]
            typ = ph_type(sh)
            if typ in ("ctrTitle", "subTitle", "dt", "ftr", "sldNum") or (typ is None and len(paras) < 2):
                continue
            for k, p in enumerate(paras):
                ppr = put(p, "pPr", ["pPr", "r", "br", "fld", "endParaRPr"])
                if ppr.find("a:lnSpc", NS) is None:
                    put(put(ppr, "lnSpc", PPR_ORDER), "spcPct", ["spcPct"], attrs={"val": int(rules["lineSpacing"] * 100000)})
                lvl = int(ppr.get("lvl", "0"))
                if k and lvl == 0 and ppr.find("a:spcBef", NS) is None:
                    put(put(ppr, "spcBef", PPR_ORDER), "spcPts", ["spcPts"], attrs={"val": rules["paraSpaceBefore"] * 100})
                if level != "light" and typ in ("body", "obj") and ppr.find("a:buNone", NS) is None and ppr.find("a:buClr", NS) is None:
                    c = put(ppr, "buClr", PPR_ORDER); etree.SubElement(c, q(A, "schemeClr")).set("val", "accent1" if lvl == 0 else "tx2")
                st["paras"] += 1
            deck.touch(s)
    if st["titles"]:
        notes.append(f"내용 슬라이드 제목 {st['titles']}개: {scale['title']}pt로 맞추고 제목 색·자간 -2%" + ("" if keep_align else "·왼쪽 정렬"))
    if st["paras"]:
        notes.append(f"본문 문단 {st['paras']}개: 행간 {rules['lineSpacing']}·문단 앞 간격 {rules['paraSpaceBefore']}pt" + ("·글머리표 강조색" if level != "light" else ""))


# ── 그래프 ──────────────────────────────────────────────────────────────
def style_chart(root, preset, accents, rules, notes_c):
    chart = root.find("c:chart", NS)
    plot = chart.find("c:plotArea", NS) if chart is not None else None
    if plot is None:
        return False
    cr = rules["chart"]
    grid, muted, line, text = preset["grid"], preset["muted"], preset["line"], preset["text"]
    # 바탕·테두리 없음, 둥근 모서리 없음
    rc = root.find("c:roundedCorners", NS)
    if rc is None:
        rc = etree.Element(q(C, "roundedCorners"))
        anchor = root.find("c:lang", NS) if root.find("c:lang", NS) is not None else root.find("c:date1904", NS)
        (anchor.addnext(rc) if anchor is not None else root.insert(0, rc))
    rc.set("val", "0")
    sp = root.find("c:spPr", NS)
    if sp is None:
        sp = etree.Element(q(C, "spPr")); chart.addnext(sp)
    set_fill(sp, SPPR_ORDER); set_line(sp, SPPR_ORDER)
    psp = put(plot, "spPr", ["layout", "spPr", "extLst"], ns=C)
    set_fill(psp, SPPR_ORDER); set_line(psp, SPPR_ORDER)
    # 자동 제목(계열 이름) 지우기: 슬라이드 제목이 이미 설명한다
    if chart.find("c:title", NS) is None:
        put(chart, "autoTitleDeleted", CHART_ORDER, ns=C, attrs={"val": "1"})
    txpr(root, ["chart", "spPr", "txPr", "externalData", "printSettings", "userShapes", "extLst"], size=1200, color=muted)
    types = [ct for ct in plot if ln(ct).endswith("Chart")]
    all_ser = [s for ct in types for s in ct.findall("c:ser", NS)]
    k = 0
    for ct in types:
        kind = ln(ct)
        sers = ct.findall("c:ser", NS)
        if kind in ("pieChart", "doughnutChart"):
            put(ct, "varyColors", PIE_ORDER, ns=C, attrs={"val": "1"})
            for s in sers:
                pts = s.findall("c:val//c:pt", NS)
                for d in s.findall("c:dPt", NS):
                    s.remove(d)
                for j in range(len(pts)):
                    d = etree.Element(q(C, "dPt"))
                    etree.SubElement(d, q(C, "idx")).set("val", str(j))
                    etree.SubElement(d, q(C, "bubble3D")).set("val", "0")
                    dsp = etree.SubElement(d, q(C, "spPr"))
                    set_fill(dsp, SPPR_ORDER, color=accents[j % len(accents)])
                    set_line(dsp, SPPR_ORDER, color="FFFFFF", w=19050)
                    anchor = next((x for x in (s.find("c:dLbls", NS), s.find("c:cat", NS), s.find("c:val", NS)) if x is not None), None)
                    (anchor.addprevious(d) if anchor is not None else s.append(d))
                if len(pts) <= cr["pieMaxSlices"]:
                    dl = put(s, "dLbls", SER_ORDER, ns=C, replace=True)
                    txpr(dl, DLBLS_ORDER, size=1300, color=text, bold=True)
                    for t_, v in (("dLblPos", "outEnd"), ("showLegendKey", "0"), ("showVal", "0"), ("showCatName", "1"),
                                  ("showSerName", "0"), ("showPercent", "1"), ("showBubbleSize", "0"), ("showLeaderLines", "1")):
                        put(dl, t_, DLBLS_ORDER, ns=C, attrs={"val": v})
                    put(dl, "separator", DLBLS_ORDER, ns=C).text = "\n"
                    leg = chart.find("c:legend", NS)
                    if leg is not None:
                        chart.remove(leg)
                    notes_c.append("원그래프: 조각마다 이름·비율 표시, 범례 삭제")
            continue
        for s in sers:
            ssp = put(s, "spPr", SER_ORDER, ns=C)
            col = accents[k % len(accents)]
            if kind in ("lineChart", "scatterChart", "radarChart"):
                set_fill(ssp, SPPR_ORDER)
                l_ = set_line(ssp, SPPR_ORDER, color=col, w=int(cr["lineWidthPt"] * 12700)); l_.set("cap", "rnd")
                npts = len(s.findall("c:val//c:pt", NS))
                m = put(s, "marker", SER_ORDER, ns=C, replace=True)
                if npts <= 8:
                    etree.SubElement(m, q(C, "symbol")).set("val", "circle"); etree.SubElement(m, q(C, "size")).set("val", "7")
                    msp = etree.SubElement(m, q(C, "spPr")); set_fill(msp, SPPR_ORDER, color=col); set_line(msp, SPPR_ORDER, color="FFFFFF", w=12700)
                else:
                    etree.SubElement(m, q(C, "symbol")).set("val", "none")
            else:
                set_fill(ssp, SPPR_ORDER, color=col); set_line(ssp, SPPR_ORDER)
            for d in s.findall("c:dPt", NS):
                dsp = d.find("c:spPr", NS)
                if dsp is not None and kind not in ("lineChart",):
                    set_fill(dsp, SPPR_ORDER, color=col)
            k += 1
        if kind == "barChart":
            put(ct, "gapWidth", BAR_ORDER, ns=C, attrs={"val": cr["gapWidth"]})
            g = ct.find("c:grouping", NS)
            if g is not None and g.get("val") in ("stacked", "percentStacked"):
                put(ct, "overlap", BAR_ORDER, ns=C, attrs={"val": "100"})
        # 계열 하나·항목 적은 막대·선: 값을 직접 표시하고 값 축과 눈금선을 지운다(Urban·Datawrapper 방식)
        vals = [float(v.text) for s in sers for v in s.iterfind("c:val//c:v", NS) if (v.text or "").replace(".", "", 1).replace("-", "", 1).isdigit()]
        npts = max((len(s.findall("c:val//c:pt", NS)) for s in sers), default=0)
        if kind in ("barChart", "lineChart") and len(all_ser) == 1 and 0 < npts <= cr["dataLabelsMaxPoints"]:
            s = sers[0]
            dl = put(s, "dLbls", SER_ORDER, ns=C, replace=True)
            txpr(dl, DLBLS_ORDER, size=1300, color=text, bold=True)
            pos = "outEnd" if kind == "barChart" and (not vals or min(vals) >= 0) else ("t" if kind == "lineChart" else None)
            if any(v != int(v) for v in vals):   # 소수가 있으면 값 표시도 같은 자릿수로(109 → 109.0)
                put(dl, "numFmt", DLBLS_ORDER, ns=C, attrs={"formatCode": "#,##0.0", "sourceLinked": "0"})
            for t_, v in (("dLblPos", pos), ("showLegendKey", "0"), ("showVal", "1"), ("showCatName", "0"),
                          ("showSerName", "0"), ("showPercent", "0"), ("showBubbleSize", "0")):
                if v is not None:
                    put(dl, t_, DLBLS_ORDER, ns=C, attrs={"val": v})
            ids = {a.get("val") for a in ct.findall("c:axId", NS)}
            for ax in plot.findall("c:valAx", NS):
                if ax.find("c:axId", NS).get("val") in ids:
                    put(ax, "delete", AX_ORDER, ns=C, attrs={"val": "1"})
                    for gl in ax.findall("c:majorGridlines", NS):
                        ax.remove(gl)
            notes_c.append("계열 하나인 그래프: 막대·점 위에 값을 직접 표시하고 값 축·눈금선 삭제")
    # 축: 값 축선 없음·옅은 눈금선, 항목 축은 가는 회색 선, 눈금 표시 없음
    for ax in plot.findall("c:valAx", NS):
        if ax.find("c:delete", NS) is not None and ax.find("c:delete", NS).get("val") == "1":
            continue
        gl = put(ax, "majorGridlines", AX_ORDER, ns=C)
        set_line(put(gl, "spPr", ["spPr"], ns=C), SPPR_ORDER, color=grid, w=int(cr["gridWidthPt"] * 12700))
        set_line(put(ax, "spPr", AX_ORDER, ns=C), SPPR_ORDER)
        put(ax, "majorTickMark", AX_ORDER, ns=C, attrs={"val": "none"}); put(ax, "minorTickMark", AX_ORDER, ns=C, attrs={"val": "none"})
        for x in ax.findall("c:minorGridlines", NS):
            ax.remove(x)
    for ax in plot.findall("c:catAx", NS) + plot.findall("c:dateAx", NS):
        set_line(put(ax, "spPr", AX_ORDER, ns=C), SPPR_ORDER, color=line, w=9525)
        put(ax, "majorTickMark", AX_ORDER, ns=C, attrs={"val": "none"}); put(ax, "minorTickMark", AX_ORDER, ns=C, attrs={"val": "none"})
        for x in ax.findall("c:majorGridlines", NS) + ax.findall("c:minorGridlines", NS):
            ax.remove(x)
        txpr(ax, AX_ORDER, size=1200, color=muted)
    leg = chart.find("c:legend", NS)
    if leg is not None and len(all_ser) > 1:
        put(leg, "legendPos", LEGEND_ORDER, ns=C, attrs={"val": "t"}); put(leg, "overlay", LEGEND_ORDER, ns=C, attrs={"val": "0"})
        txpr(leg, LEGEND_ORDER, size=1200, color=text)
    return True


# ── 표 ──────────────────────────────────────────────────────────────────
def style_tables(root, preset):
    n = 0
    for tbl in root.iter(q(A, "tbl")):
        rows = tbl.findall("a:tr", NS)
        if not rows:
            continue
        tp = tbl.find("a:tblPr", NS)
        if tp is not None:
            tp.set("firstRow", "1"); tp.set("bandRow", "1")
        ncol = max(len(r.findall("a:tc", NS)) for r in rows)
        numeric = []
        for c in range(ncol):
            cells = [r.findall("a:tc", NS)[c] for r in rows[1:] if len(r.findall("a:tc", NS)) > c]
            texts = ["".join(cell.itertext()).strip() for cell in cells]
            texts = [t for t in texts if t]
            numeric.append(bool(texts) and sum(bool(NUM.match(t)) for t in texts) / len(texts) >= 0.7)
        for ri, r in enumerate(rows):
            for ci, tc in enumerate(r.findall("a:tc", NS)):
                tcpr = tc.find("a:tcPr", NS)
                if tcpr is None:
                    tcpr = etree.SubElement(tc, q(A, "tcPr"))
                tcpr.set("marL", "91440"); tcpr.set("marR", "91440"); tcpr.set("marT", "54864"); tcpr.set("marB", "54864"); tcpr.set("anchor", "ctr")
                for side in ("lnL", "lnR", "lnT", "lnB"):
                    el = put(tcpr, side, TCPR_ORDER, replace=True)
                    if side in ("lnL", "lnR") or (side == "lnT" and ri > 0):
                        etree.SubElement(el, q(A, "noFill"))
                    else:
                        el.set("w", "19050" if (ri == 0 and side == "lnB") else "6350")
                        f = etree.SubElement(el, q(A, "solidFill"))
                        etree.SubElement(f, q(A, "srgbClr")).set("val", preset["title"] if (ri == 0 and side == "lnB") else preset["line"])
                if ri == 0:
                    set_fill(tcpr, TCPR_ORDER, color=preset["title"])
                elif ri % 2 == 0:
                    set_fill(tcpr, TCPR_ORDER, color=preset["light"])
                else:
                    set_fill(tcpr, TCPR_ORDER, color="FFFFFF")
                for p in tc.iter(q(A, "p")):
                    if numeric[ci] and ci > 0:
                        put(p, "pPr", ["pPr", "r", "br", "fld", "endParaRPr"]).set("algn", "r")
                    for r_ in p.iter(q(A, "r")):
                        rp = r_.find("a:rPr", NS)
                        if rp is None:
                            rp = etree.Element(q(A, "rPr")); r_.insert(0, rp)
                        if ri == 0:
                            rp.set("b", "1"); set_fill(rp, RPR_ORDER, color="FFFFFF")
                        else:
                            set_fill(rp, RPR_ORDER, color=preset["text"])
        n += 1
    return n


# ── 직접 지정한 도형 색(bold) ───────────────────────────────────────────
def recolor_shapes(root, preset):
    pal = [preset["accent"], preset["title"]]   # 색은 세 가지까지: 제목색·강조색 + 무채색
    n = 0
    for sh in root.iter(q(P, "sp")):
        f = sh.find("p:spPr/a:solidFill/a:srgbClr", NS)
        if f is None:
            continue
        h, l_, s = colorsys.rgb_to_hls(*hexs(f.get("val")))
        if s < 0.18 or l_ > 0.93 or l_ < 0.07:
            continue   # 무채색·아주 밝은·아주 어두운 색은 그대로
        new = nearest(f.get("val"), pal)
        if new.upper() != f.get("val").upper():
            f.set("val", new); n += 1
    return n


DECO = "ffs-deco"
DECORATIVE = ('<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">'
              '<adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>')


def add_rect(root, name, x, y, cx, cy, scheme, front=False):
    tree = root.find("p:cSld/p:spTree", NS)
    ids = [int(e.get("id")) for e in root.iter(q(P, "cNvPr")) if (e.get("id") or "").isdigit()]
    sp = etree.fromstring(
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{max(ids + [1]) + 1}" name="{name}">{DECORATIVE}</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{int(x)}" y="{int(y)}"/><a:ext cx="{int(cx)}" cy="{int(cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:schemeClr val="{scheme}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ko-KR"/></a:p></p:txBody></p:sp>')
    first = next((i for i, ch in enumerate(tree) if ln(ch) in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp")), len(tree))
    tree.insert(len(tree) if front else first, sp)
    return sp


def add_slide_number(root, W, H, muted):
    tree = root.find("p:cSld/p:spTree", NS)
    ids = [int(e.get("id")) for e in root.iter(q(P, "cNvPr")) if (e.get("id") or "").isdigit()]
    IN = 914400
    sp = etree.fromstring(
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{max(ids + [1]) + 1}" name="{DECO}-num"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{int(W - 1.4 * IN)}" y="{int(H - 0.5 * IN)}"/><a:ext cx="{int(1.0 * IN)}" cy="{int(0.32 * IN)}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody><a:bodyPr wrap="none" lIns="0" rIns="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="r"/>'
        f'<a:fld id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum"><a:rPr lang="ko-KR" sz="1100"><a:solidFill><a:srgbClr val="{muted}"/></a:solidFill></a:rPr><a:t>‹#›</a:t></a:fld>'
        f'</a:p></p:txBody></p:sp>')
    tree.append(sp)


def decorate(deck, slides, preset, keep_align=False):
    """bold: 내용 슬라이드 위쪽 가는 띠(제목 색)와 쪽 번호, 표지 왼쪽 띠. 이미 넣은 장식(ffs-deco)은 다시 넣지 않는다."""
    W, H = deck.size()
    IN = 914400
    n = 0
    for i, s in enumerate(slides, 1):
        root = deck.xml(s)
        if any((e.get("name") or "").startswith(DECO) for e in root.iter(q(P, "cNvPr"))):
            continue
        bgimg = root.find("p:cSld/p:bg//a:blipFill", NS) is not None
        t = title_shape(root)
        cover = i == 1 or (t is not None and ph_type(t) == "ctrTitle")
        busy_top = any((xfrm_of(sh) is not None and xfrm_of(sh).find("a:off", NS) is not None
                        and int(xfrm_of(sh).find("a:off", NS).get("y")) < 0.15 * IN) for sh in shapes(root))
        if cover:
            pics = any(ln(sh) == "pic" for sh in shapes(root))
            left_busy = any((xfrm_of(sh) is not None and xfrm_of(sh).find("a:off", NS) is not None
                             and int(xfrm_of(sh).find("a:off", NS).get("x")) < 0.4 * IN) for sh in shapes(root))
            if not bgimg and not pics and not left_busy:
                add_rect(root, f"{DECO}-band", 0, 0, 0.32 * IN, H, "tx2")
                add_rect(root, f"{DECO}-band2", 0.32 * IN, 0, 0.06 * IN, H, "accent1")
                if not keep_align:   # 띠와 어울리게 표지 글을 왼쪽 정렬
                    for sh in shapes(root):
                        if ph_type(sh) in ("ctrTitle", "title", "subTitle"):
                            for p in sh.iter(q(A, "p")):
                                put(p, "pPr", ["pPr", "r", "br", "fld", "endParaRPr"]).set("algn", "l")
                n += 1; deck.touch(s)
            continue
        if not bgimg and not busy_top:
            add_rect(root, f"{DECO}-top", 0, 0, W, 0.1 * IN, "tx2")
        add_slide_number(root, W, H, preset["muted"])
        n += 1; deck.touch(s)
    return n


def title_box_width(deck, slide, sh):
    xf = xfrm_of(sh)
    if xf is not None and xf.find("a:ext", NS) is not None:
        return int(xf.find("a:ext", NS).get("cx"))
    lay = deck.first(slide, "slideLayout")
    mas = deck.first(lay, "slideMaster") if lay else None
    for part in (lay, mas):
        for lsh in shapes(deck.xml(part)) if part else []:
            x = xfrm_of(lsh)
            if ph_type(lsh) == ph_type(sh) and x is not None and x.find("a:ext", NS) is not None:
                return int(x.find("a:ext", NS).get("cx"))
    return None


def balance_titles(deck, slides, meas):
    """adobe/balance-text 방식: 두 줄로 넘어가는 제목의 마지막 줄에 한두 어절만 남으면, 두 줄 길이가 비슷해지는 어절 사이에서 줄을 나눈다."""
    n = 0
    for s in slides:
        root = deck.xml(s)
        t = title_shape(root)
        if t is None:
            continue
        paras = [p for p in t.findall("p:txBody/a:p", NS) if "".join(x.text or "" for x in p.iter(q(A, "t"))).strip()]
        if len(paras) != 1 or paras[0].find("a:br", NS) is not None:
            continue
        runs = paras[0].findall("a:r", NS)
        if len(runs) != 1:
            continue
        r = runs[0]; rp = r.find("a:rPr", NS); text = r.findtext("a:t", default="", namespaces=NS)
        if rp is None or not rp.get("sz") or " " not in text.strip():
            continue
        width = title_box_width(deck, s, t)
        bp = t.find("p:txBody/a:bodyPr", NS)
        ins = (int(bp.get("lIns", 91440)) + int(bp.get("rIns", 91440))) if bp is not None else 182880
        if not width:
            continue
        avail = width - ins
        size = int(rp.get("sz")) / 100 * 12700
        face = pdk.theme_faces(deck, s).get("mj", "")
        if face not in meas.fonts:
            continue
        wd = lambda x: meas.width(x, face, size) * (1 + float(rp.get("spc", "0")) / 100 / (size / 12700) * 0)  # 자간은 측정에 거의 영향 없음
        words = text.split(" ")
        total = wd(text)
        if total <= avail or total > avail * 1.9:
            continue   # 한 줄이거나 세 줄 이상
        # 지금 줄바꿈(어절 단위)의 마지막 줄 길이
        cur, last = "", ""
        for w_ in words:
            trial = (cur + " " + w_).strip()
            if wd(trial) <= avail:
                cur = trial
            else:
                cur = w_
        last = cur
        if wd(last) > avail * 0.4:
            continue
        best = None
        for k in range(1, len(words)):
            a1, a2 = " ".join(words[:k]), " ".join(words[k:])
            m = max(wd(a1), wd(a2))
            if m <= avail * 0.96 and (best is None or m < best[0]):
                best = (m, k)
        if not best:
            continue
        k = best[1]
        r.find("a:t", NS).text = " ".join(words[:k])
        br = etree.Element(q(A, "br")); br.append(etree.fromstring(etree.tostring(rp)))
        r2 = etree.fromstring(etree.tostring(r)); r2.find("a:t", NS).text = " ".join(words[k:])
        r.addnext(br); br.addnext(r2)
        n += 1; deck.touch(s)
    return n


def plain_template(deck):
    """기본 Office 테마이고 마스터·레이아웃에 그림이나 장식 도형이 없으면 '꾸밈 없는 기본 서식'으로 본다."""
    for n in deck.names:
        if re.match(r"ppt/theme/theme\d+\.xml$", n):
            a1 = deck.xml(n).find(".//a:clrScheme/a:accent1/a:srgbClr", NS)
            if a1 is not None and a1.get("val", "").upper() not in OFFICE_ACCENT1:
                return False
        if re.match(r"ppt/(slideMasters/slideMaster|slideLayouts/slideLayout)\d+\.xml$", n):
            for sh in shapes(deck.xml(n)):
                if ln(sh) in ("pic", "grpSp", "cxnSp") or (ln(sh) == "sp" and ph_type(sh) is None):
                    return False
    return True


def main():
    ap = argparse.ArgumentParser(description="분야별 디자인 프리셋으로 피피티 다듬기")
    ap.add_argument("pptx", nargs="?")
    ap.add_argument("--out", help="저장 경로(없으면 무엇을 바꿀지만 보여 줌)")
    ap.add_argument("--preset", default="auto", help="auto 또는 프리셋 이름(--list)")
    ap.add_argument("--level", choices=["light", "standard", "bold"], default="standard")
    ap.add_argument("--chart-palette", help="그래프만 다른 팔레트로(예: colorblind)")
    ap.add_argument("--keep-align", action="store_true", help="제목 정렬을 바꾸지 않음")
    ap.add_argument("--manifest", help="fetch_fonts.py의 manifest.json(제목 줄 균형을 맞출 때 실제 글꼴 폭을 씀)")
    ap.add_argument("--data", help="data 폴더(오프라인)")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--json", help="바꾼 내역을 JSON으로")
    a = ap.parse_args()
    tokens = load_tokens(a.data)
    if a.list or not a.pptx:
        for k, p in tokens["presets"].items():
            print(f"  {k:<10} {p['name']:<24} 제목 #{p['title']}  강조 {' '.join('#' + c for c in p['accents'])}")
        for k, p in tokens["chartPalettes"].items():
            print(f"  {k:<10} {p['name']} (그래프 전용, --chart-palette)")
        print("출처: " + " / ".join(tokens["sources"].values()))
        return
    if a.preset == "auto":
        key, why = auto_preset(a.pptx, tokens, a.data)
    else:
        if a.preset not in tokens["presets"]:
            sys.exit(f"프리셋 '{a.preset}'이 없습니다. --list로 확인하세요.")
        key, why = a.preset, "지정"
    preset = tokens["presets"][key]
    accents = tokens["chartPalettes"][a.chart_palette]["accents"] if a.chart_palette else preset["accents"]
    deck = Deck(a.pptx)
    slides = deck.slides()
    plain = plain_template(deck)
    notes = [f"프리셋: {key} — {preset['name']} ({why}), 강도 {a.level}"]
    dark = False
    try:
        import analyze_deck as ad
        dark = bool(ad.diagnose(ad.extract(a.pptx))["visual"].get("어두운배경"))
    except Exception:
        pass
    if dark:
        notes.append("어두운 배경 덱이라 글자·바탕 색은 그대로 두고 강조색만 바꿈")
    if a.level != "light":
        style_theme(deck, preset, a.level == "bold", notes, dark)
    style_text(deck, slides, preset, tokens, a.level, a.keep_align, notes, dark)
    if a.level != "light" and a.manifest:
        b = balance_titles(deck, slides, pdk.Measure(a.manifest))
        if b:
            notes.append(f"제목 {b}개: 마지막 줄에 한두 어절만 남지 않게 줄 길이를 고르게 나눔(balance-text 방식)")
    charts, chart_notes, tables, recol = 0, [], 0, 0
    for s in slides:
        rel = deck.rels(s)
        root = deck.xml(s)
        for gf in root.iter(q(P, "graphicFrame")):
            ch = gf.find(".//c:chart", NS)
            if ch is None:
                continue
            tgt = rel.get(ch.get(f"{{{pdk.R}}}id"))
            if tgt and tgt[0] in deck.names and style_chart(deck.xml(tgt[0]), preset, accents, tokens["rules"], chart_notes):
                deck.touch(tgt[0]); charts += 1
        t = style_tables(root, preset)
        if t:
            tables += t; deck.touch(s)
        if a.level == "bold":
            r = recolor_shapes(root, preset)
            if r:
                recol += r; deck.touch(s)
    if charts:
        notes.append(f"그래프 {charts}개: Urban Institute 방식(테두리·값 축선·눈금 표시 없음, 옅은 눈금선, 막대 간격 {tokens['rules']['chart']['gapWidth']}%, 프리셋 색)"
                     + "".join(f"; {m}" for m in sorted(set(chart_notes))))
    if tables:
        notes.append(f"표 {tables}개: 제목 줄 진한 바탕·흰 굵은 글자, 옅은 줄무늬, 가로 선만, 숫자 열 오른쪽 정렬")
    if recol:
        notes.append(f"직접 칠한 도형 색 {recol}곳을 제목색·강조색 중 가까운 쪽으로(색은 세 가지까지)")
    if not dark and (a.level == "bold" or (a.level == "standard" and plain)):
        d = decorate(deck, slides, preset, a.keep_align)
        if d:
            notes.append(f"장식 {d}장: 표지 왼쪽 띠, 내용 슬라이드 위쪽 가는 띠(제목 색)와 쪽 번호")
    print("디자인 다듬기: " + a.pptx)
    for m in notes:
        print("  · " + m)
    if a.out:
        deck.save(a.out)
        print(f"저장: {a.out}")
    if a.json:
        json.dump({"preset": key, "notes": notes}, open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
