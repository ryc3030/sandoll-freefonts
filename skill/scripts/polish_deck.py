#!/usr/bin/env python3
"""피피티의 그래프와 디자인을 진단하고, 거의 언제나 잘못인 것은 자동으로 고친다.

  python polish_deck.py output.pptx                                  # 진단만(파일은 그대로)
  python polish_deck.py output.pptx --fix polished.pptx              # 안전한 보정을 적용해 새 파일로
  python polish_deck.py output.pptx --fix polished.pptx --manifest fonts/manifest.json --fit --contrast --chart-colors
  python polish_deck.py output.pptx --json report.json               # 진단 결과를 JSON으로도

--fix 를 주면 자동으로 고치는 것(거의 언제나 잘못이라 기본으로 고침)
  그래프  · 3D 그래프(막대·선·영역·원) → 같은 종류의 2D 그래프(3D는 값을 읽기 어렵고 앞 막대가 뒤를 가림)
          · 0이 아닌 값에서 시작하는 막대·영역 그래프의 값 축 → 0에서 시작(차이가 부풀려 보이는 것을 막음)
          · 너무 작은 그래프 글자 → 축·범례·값 12pt, 제목 14pt 이상(작은 그래프는 10pt)
          · 짙거나 두꺼운 주 눈금선 → 옅은 회색 가는 선, 보조 눈금선 → 지움
          · 계열이 하나뿐인 그래프의 범례 → 지움(제목과 같은 말을 되풀이함)
  디자인  · 가로세로 비율이 틀어진 사진 → 원래 비율로(가로 폭과 가운데는 그대로)
          · 슬라이드마다 조금씩 어긋난 제목 위치(0.3인치 안) → 가장 많이 쓴 위치로
          · 한 슬라이드 안에서 거의 같은데 조금 어긋난 왼쪽 끝(0.1인치 안) → 많이 쓴 쪽으로
옵션으로만 고치는 것(취향이 갈리거나 추정이 들어감)
  --chart-colors  그래프 계열 색을 덱의 테마 색(강조 1~6)으로 맞춘다
  --contrast      배경과 대비가 낮은 글자(WCAG 기준 미달)를 검정 또는 흰색 중 대비가 큰 쪽으로
  --fit           상자를 넘치는 것으로 추정되는 글자를 상자에 맞게 줄인다(본문 12pt, 제목 20pt 아래로는 줄이지 않음)
진단만 하는 것
  원그래프 조각 7개 이상, 슬라이드 밖으로 나간 개체, 12pt보다 작은 본문 글자, 대비가 낮은 글자, 넘침 추정, 글자 색 가짓수

넘침과 대비는 XML로 추정한 값이다. 보정한 뒤에는 export_pdf.py로 PDF를 만들어 그림으로 다시 확인한다.
"""
import argparse, colorsys, io, json, os, re, shutil, sys, tempfile, zipfile
from collections import Counter, defaultdict
from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": A, "p": P, "c": C, "r": R}
IN = 914400
TITLE_PH = {"title", "ctrTitle"}


def q(ns, tag):
    return f"{{{ns}}}{tag}"


def ln(el):
    return etree.QName(el).localname


# ── 패키지 ──────────────────────────────────────────────────────────────
class Deck:
    def __init__(self, path):
        try:
            self.z = zipfile.ZipFile(path)
        except zipfile.BadZipFile:
            sys.exit(f"'{path}'을(를) 열 수 없습니다. 암호가 걸렸거나 손상된 파일입니다.")
        self.names = set(self.z.namelist())
        self.cache, self.dirty = {}, set()

    def xml(self, name):
        if name not in self.cache:
            self.cache[name] = etree.fromstring(self.z.read(name))
        return self.cache[name]

    def touch(self, name):
        self.dirty.add(name)

    def rels(self, part):
        d, f = part.rsplit("/", 1)
        rp = f"{d}/_rels/{f}.rels"
        out = {}
        if rp not in self.names:
            return out
        for r in self.xml(rp):
            if r.get("TargetMode") == "External":
                continue
            stack = []
            for p in (d + "/" + r.get("Target", "")).split("/"):
                if p == "..":
                    stack and stack.pop()
                elif p and p != ".":
                    stack.append(p)
            out[r.get("Id")] = ("/".join(stack), r.get("Type", "").rsplit("/", 1)[-1])
        return out

    def first(self, part, typ):
        return next((t for t, ty in self.rels(part).values() if ty == typ), None)

    def slides(self):
        pres = self.xml("ppt/presentation.xml")
        rel = self.rels("ppt/presentation.xml")
        out = []
        for s in pres.findall("p:sldIdLst/p:sldId", NS):
            t = rel.get(s.get(q(R, "id")))
            if t and t[0] in self.names:
                out.append(t[0])
        return out

    def size(self):
        sz = self.xml("ppt/presentation.xml").find("p:sldSz", NS)
        return (int(sz.get("cx")), int(sz.get("cy"))) if sz is not None else (9144000, 6858000)

    def save(self, dst):
        fd, tmp = tempfile.mkstemp(suffix=".pptx", dir=os.path.dirname(os.path.abspath(dst)) or ".")
        os.close(fd)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for it in self.z.infolist():
                data = self.z.read(it.filename)
                if it.filename in self.dirty:
                    data = etree.tostring(self.cache[it.filename], xml_declaration=True, encoding="UTF-8", standalone=True)
                out.writestr(it, data)
        os.chmod(tmp, 0o644)
        shutil.move(tmp, dst)


# ── 색 ──────────────────────────────────────────────────────────────────
def rel_lum(hexv):
    def ch(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = (int(hexv[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def contrast(a, b):
    la, lb = sorted((rel_lum(a), rel_lum(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def apply_mods(hexv, el):
    """schemeClr·srgbClr 아래의 lumMod·lumOff·tint·shade를 적용한다(근사)."""
    r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    for m in el:
        v = int(m.get("val", "100000")) / 100000
        t = ln(m)
        if t == "lumMod":
            l *= v
        elif t == "lumOff":
            l += v
        elif t == "tint":
            l = l + (1 - l) * (1 - v)
        elif t == "shade":
            l *= v
    l = min(max(l, 0), 1)
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


class Colors:
    def __init__(self, deck, slide):
        self.deck = deck
        lay = deck.first(slide, "slideLayout")
        mas = deck.first(lay, "slideMaster") if lay else None
        th = deck.first(mas, "theme") if mas else None
        self.parts = [p for p in (slide, lay, mas) if p]
        self.scheme = {}
        if th:
            cs = deck.xml(th).find(".//a:clrScheme", NS)
            for el in (cs if cs is not None else []):
                c = el.find("a:srgbClr", NS); s = el.find("a:sysClr", NS)
                self.scheme[ln(el)] = c.get("val") if c is not None else (s.get("lastClr") if s is not None else None)
        cm = deck.xml(mas).find("p:clrMap", NS) if mas else None
        self.clrmap = dict(cm.attrib) if cm is not None else {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}
        ov = deck.xml(slide).find("p:clrMapOvr/a:overrideClrMapping", NS)
        if ov is not None:
            self.clrmap = dict(ov.attrib)

    def of(self, fill_parent):
        """a:solidFill 등 색 요소를 가진 부모에서 색(hex)을 읽는다. 없거나 그림·그라데이션이면 None."""
        if fill_parent is None:
            return None
        for el in fill_parent.iter(q(A, "srgbClr"), q(A, "schemeClr"), q(A, "sysClr"), q(A, "prstClr")):
            t = ln(el)
            if t == "srgbClr":
                base = el.get("val")
            elif t == "sysClr":
                base = el.get("lastClr")
            elif t == "prstClr":
                base = {"black": "000000", "white": "FFFFFF"}.get(el.get("val"))
            else:
                n = el.get("val")
                base = self.scheme.get(self.clrmap.get(n, n)) if n not in ("phClr",) else None
            return apply_mods(base, el) if base else None
        return None

    def background(self):
        for part in self.parts:
            bg = self.deck.xml(part).find("p:cSld/p:bg", NS)
            if bg is None:
                continue
            if bg.find(".//a:blipFill", NS) is not None or bg.find(".//a:gradFill", NS) is not None:
                return None
            ref = bg.find("p:bgRef", NS)
            return self.of(ref if ref is not None else bg.find("p:bgPr", NS))
        return self.scheme.get(self.clrmap.get("bg1", "lt1"))


# ── 그래프 ──────────────────────────────────────────────────────────────
CHART3D = {"bar3DChart": "barChart", "line3DChart": "lineChart", "area3DChart": "areaChart", "pie3DChart": "pieChart"}


def chart_values(plot):
    vals = []
    for v in plot.iterfind(".//c:val//c:v", NS):
        try:
            vals.append(float(v.text))
        except (TypeError, ValueError):
            pass
    return vals


def fix_chart(root, frame_w, fix, opts, notes):
    changed = False
    chart = root.find("c:chart", NS)
    plot = chart.find("c:plotArea", NS) if chart is not None else None
    if plot is None:
        return False
    # 1) 3D → 2D
    for el in list(plot):
        if ln(el) in CHART3D:
            kind = {"bar": "막대", "line": "선", "area": "영역", "pie": "원"}[ln(el).replace("3DChart", "")]
            notes.append(("그래프", f"3D {kind} 그래프는 값을 읽기 어렵고 앞 요소가 뒤를 가립니다" + (" → 2D로 바꿈" if fix else "")))
            if not fix:
                continue
            new = CHART3D[ln(el)]
            el.tag = q(C, new)
            for t in ("gapDepth", "shape"):
                for x in el.findall(f"c:{t}", NS):
                    el.remove(x)
            for x in el.findall("c:ser/c:shape", NS):
                x.getparent().remove(x)
            ser_ax = {s.find("c:axId", NS).get("val") for s in plot.findall("c:serAx", NS)}
            for a in el.findall("c:axId", NS):
                if a.get("val") in ser_ax:
                    el.remove(a)
            for s in plot.findall("c:serAx", NS):
                plot.remove(s)
            if new == "barChart":
                g = el.find("c:grouping", NS)
                if g is not None and g.get("val") == "standard":
                    g.set("val", "clustered")
                if g is not None and g.get("val") in ("stacked", "percentStacked") and el.find("c:overlap", NS) is None:
                    ov = etree.Element(q(C, "overlap")); ov.set("val", "100")
                    anchor = el.find("c:gapWidth", NS)
                    if anchor is not None:
                        anchor.addnext(ov)
                    else:
                        el.find("c:axId", NS).addprevious(ov)
            for t in ("view3D", "floor", "sideWall", "backWall"):
                for x in chart.findall(f"c:{t}", NS):
                    chart.remove(x)
            changed = True
    # 2) 막대·영역 값 축은 0에서
    vals = chart_values(plot)
    for ct in plot:
        if ln(ct) not in ("barChart", "areaChart", "bar3DChart", "area3DChart"):
            continue
        ids = {a.get("val") for a in ct.findall("c:axId", NS)}
        for ax in plot.findall("c:valAx", NS):
            if ax.find("c:axId", NS).get("val") not in ids:
                continue
            mn = ax.find("c:scaling/c:min", NS)
            if mn is not None and ax.find("c:scaling/c:logBase", NS) is None:
                try:
                    m = float(mn.get("val"))
                except (TypeError, ValueError):
                    continue
                if m > 0 and (not vals or min(vals) >= 0):
                    notes.append(("그래프", f"막대 그래프 값 축이 {m:g}에서 시작해 차이가 부풀려 보입니다" + (" → 0에서 시작" if fix else "")))
                    if fix:
                        mn.set("val", "0"); changed = True
    # 3) 그래프 글자 크기
    small = 1000 if frame_w < 4 * IN else 1200
    tiny = []
    for rp in root.iter(q(A, "defRPr"), q(A, "rPr"), q(A, "endParaRPr")):
        sz = rp.get("sz")
        if not sz:
            continue
        in_title = any(ln(x) == "title" for x in rp.iterancestors())
        need = 1400 if in_title else small
        if int(sz) < need:
            tiny.append(int(sz) / 100)
            if fix:
                rp.set("sz", str(need)); changed = True
    space_txpr = root.find("c:txPr", NS)
    if space_txpr is None:
        if fix:
            tx = etree.fromstring(f'<c:txPr xmlns:c="{C}" xmlns:a="{A}"><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
                                  f'<a:defRPr sz="{small}"/></a:pPr><a:endParaRPr lang="ko-KR"/></a:p></c:txPr>')
            anchor = root.find("c:spPr", NS)
            anchor = anchor if anchor is not None else chart
            anchor.addnext(tx); changed = True
    else:
        d = space_txpr.find(".//a:defRPr", NS)
        if d is not None and not d.get("sz"):
            if fix:
                d.set("sz", str(small)); changed = True
    if tiny:
        notes.append(("그래프", f"그래프 글자가 작습니다({min(tiny):g}pt)" + (f" → {small // 100}pt 이상(제목 14pt)" if fix else "")))
    # 4) 눈금선
    for g in plot.iter(q(C, "minorGridlines")):
        notes.append(("그래프", "보조 눈금선이 그래프를 어지럽힙니다" + (" → 지움" if fix else "")))
        if fix:
            g.getparent().remove(g); changed = True
    for g in plot.iter(q(C, "majorGridlines")):
        line = g.find("c:spPr/a:ln", NS)
        if line is None:
            continue
        w = int(line.get("w", "9525"))
        hexv = None
        s = line.find("a:solidFill/a:srgbClr", NS)
        if s is not None:
            hexv = apply_mods(s.get("val"), s)
        if w > 12700 or (hexv and rel_lum(hexv) < 0.45):
            notes.append(("그래프", "주 눈금선이 짙거나 두껍습니다" + (" → 옅은 회색 가는 선" if fix else "")))
            if fix:
                line.set("w", "6350")
                for x in list(line):
                    if ln(x) in ("solidFill", "noFill", "gradFill", "prstDash"):
                        line.remove(x)
                sf = etree.Element(q(A, "solidFill")); c = etree.SubElement(sf, q(A, "srgbClr")); c.set("val", "D9D9D9")
                line.insert(0, sf); changed = True
    # 5) 계열 하나인 그래프의 범례
    types = [ct for ct in plot if ln(ct).endswith("Chart")]
    nser = sum(len(ct.findall("c:ser", NS)) for ct in types)
    pie = any(ln(ct) in ("pieChart", "pie3DChart", "doughnutChart", "ofPieChart") for ct in types)
    leg = chart.find("c:legend", NS)
    if nser == 1 and not pie and leg is not None:
        notes.append(("그래프", "계열이 하나뿐인데 범례가 있습니다" + (" → 지움" if fix else "")))
        if fix:
            chart.remove(leg); changed = True
    # 6) 그래프 종류 고르기(FT Visual Vocabulary 기준, 진단만)
    for ct in types:
        kind = ln(ct).replace("3DChart", "Chart")
        sers = ct.findall("c:ser", NS)
        npt = max((len(x.findall("c:val//c:pt", NS)) for x in sers), default=0)
        if kind in ("pieChart", "doughnutChart"):
            if npt >= 6:
                notes.append(("그래프", f"원그래프 조각이 {npt}개라 비교하기 어렵습니다. 부분과 전체는 조각 5개 이하일 때만 원그래프가 낫고, "
                                       "그 이상은 큰 순서로 늘어놓은 가로 막대나 '기타' 묶기를 권합니다(FT Visual Vocabulary, 진단만)"))
        elif kind == "lineChart" and 0 < npt <= 3:
            notes.append(("그래프", f"점이 {npt}개뿐인 선 그래프입니다. 시점이 적으면 세로 막대가 비교하기 쉽습니다(FT Visual Vocabulary, 진단만)"))
        elif kind == "barChart":
            bd = ct.find("c:barDir", NS)
            labels = [v.text or "" for v in ct.iterfind("c:ser/c:cat//c:v", NS)][:npt]
            if bd is not None and bd.get("val") == "col" and npt >= 5 and labels and sum(map(len, labels)) / len(labels) >= 6:
                notes.append(("그래프", "항목 이름이 길어 세로 막대 아래에서 줄바꿈되거나 기울어집니다. 순위·비교는 가로 막대가 읽기 쉽습니다(FT Visual Vocabulary, 진단만)"))
            g = ct.find("c:grouping", NS)
            if g is not None and g.get("val") in ("stacked", "percentStacked") and len(sers) > 4:
                notes.append(("그래프", f"쌓은 막대에 계열이 {len(sers)}개입니다. 가운데 조각끼리는 비교가 어려우니 4개 이하로 묶어 주세요(진단만)"))
    live_val = [ax for ax in plot.findall("c:valAx", NS) if not (ax.find("c:delete", NS) is not None and ax.find("c:delete", NS).get("val") == "1")]
    if len(live_val) >= 2 and len(types) >= 2:
        notes.append(("그래프", "값 축이 둘인 그래프입니다. 두 축의 눈금을 어떻게 잡느냐에 따라 관계가 달라 보이니, 가능하면 그래프 두 개로 나누세요(진단만)"))
    # 7) 테마 색(옵션)
    if fix and opts.chart_colors:
        k = 0
        for ct in types:
            for s in ct.findall("c:ser", NS):
                pts = s.findall("c:dPt", NS)
                targets = [(s.find("c:spPr", NS), k)] + [(d.find("c:spPr", NS), int(d.find("c:idx", NS).get("val"))) for d in pts]
                for sp, i in targets:
                    if sp is None:
                        continue
                    for f in sp.iter(q(A, "solidFill")):
                        for x in list(f):
                            f.remove(x)
                        sc = etree.SubElement(f, q(A, "schemeClr")); sc.set("val", f"accent{i % 6 + 1}")
                        changed = True
                k += 1
        notes.append(("그래프", "계열 색을 덱의 테마 강조색으로 맞춤"))
    return changed


# ── 글자 측정(넘침 추정) ────────────────────────────────────────────────
class Measure:
    def __init__(self, manifest):
        self.fonts = {}
        if manifest:
            try:
                from fontTools.ttLib import TTFont
                base = os.path.dirname(os.path.abspath(manifest))
                for m in json.load(open(manifest, encoding="utf-8")):
                    f = m.get("file") or ""
                    if f and not os.path.exists(f):   # manifest가 상대 경로로 적혀 있으면 manifest 폴더 기준으로
                        cand = [os.path.join(base, os.path.basename(f)), os.path.join(os.path.dirname(base), f)]
                        f = next((c for c in cand if os.path.exists(c)), f)
                    if f and os.path.exists(f) and m["family"] not in self.fonts:
                        t = TTFont(f, lazy=True)
                        upm = t["head"].unitsPerEm; hh = t["hhea"]
                        self.fonts[m["family"]] = (t.getBestCmap(), t["hmtx"].metrics, upm,
                                                   (hh.ascent - hh.descent + hh.lineGap) / upm)
            except Exception:
                pass

    def width(self, text, face, size):
        f = self.fonts.get(face)
        if f:
            cmap, hm, upm, _ = f
            return sum(hm.get(cmap.get(ord(ch), ".notdef"), (upm * 0.6, 0))[0] for ch in text) / upm * size
        w = 0
        for ch in text:
            o = ord(ch)
            w += 0.95 if (0xAC00 <= o <= 0xD7A3 or 0x3000 <= o <= 0x9FFF) else (0.3 if ch == " " else 0.56)
        return w * size

    def line_h(self, face, size):
        f = self.fonts.get(face)
        return (f[3] if f else 1.3) * size


def theme_faces(deck, slide):
    lay = deck.first(slide, "slideLayout")
    mas = deck.first(lay, "slideMaster") if lay else None
    th = deck.first(mas, "theme") if mas else None
    out = {}
    if th:
        t = deck.xml(th)
        for tag, key in (("majorFont", "mj"), ("minorFont", "mn")):
            e = t.find(f".//a:{tag}/a:ea", NS); l_ = t.find(f".//a:{tag}/a:latin", NS)
            out[key] = (e.get("typeface") if e is not None and e.get("typeface") else (l_.get("typeface") if l_ is not None else ""))
    return out


def box_need(sp, meas, faces, scale=1.0):
    """상자에 필요한 높이(EMU)와 쓸 수 있는 높이. 크기를 알 수 없는 글자가 있으면 None."""
    body = sp.find("p:txBody", NS)
    xfrm = sp.find("p:spPr/a:xfrm", NS)
    if body is None or xfrm is None or xfrm.find("a:ext", NS) is None:
        return None
    bp = body.find("a:bodyPr", NS)
    if bp is not None and (bp.find("a:normAutofit", NS) is not None or bp.find("a:spAutoFit", NS) is not None or bp.get("wrap") == "none"):
        return None
    if not "".join(t.text or "" for t in body.iter(q(A, "t"))).strip():
        return None   # 글자가 없는 도형(색 띠 등)
    is_ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS) is not None
    ext = xfrm.find("a:ext", NS)
    g = lambda k, d: int(bp.get(k, d)) if bp is not None else d
    avail_w = int(ext.get("cx")) - g("lIns", 91440) - g("rIns", 91440)
    avail_h = int(ext.get("cy")) - g("tIns", 45720) - g("bIns", 45720)
    if avail_w <= 0:
        return None
    total = 0
    for p in body.findall("a:p", NS):
        runs = p.findall("a:r", NS)
        text = "".join((r.findtext("a:t", default="", namespaces=NS)) for r in runs)
        sizes = [int(r.find("a:rPr", NS).get("sz")) for r in runs if r.find("a:rPr", NS) is not None and r.find("a:rPr", NS).get("sz")]
        end = p.find("a:endParaRPr", NS)
        if not sizes:
            if not text.strip():
                es = end.get("sz") if end is not None else None
                sizes = [int(es)] if es else ([] if is_ph else [1800])
                if not sizes:
                    continue
            elif is_ph:
                return None
            else:
                sizes = [1800]
        if text.strip() and len(sizes) < len(runs) and is_ph:
            return None
        size = max(sizes) / 100 * scale * 12700  # EMU
        face = None
        for r in runs:
            rp = r.find("a:rPr", NS)
            e = rp.find("a:ea", NS) if rp is not None else None
            if e is not None:
                tf = e.get("typeface", "")
                face = faces.get("mj" if tf.startswith("+mj") else "mn") if tf.startswith("+") else tf
                break
        face = face or faces.get("mn", "")
        ppr = p.find("a:pPr", NS)
        indent = int(ppr.get("marL", "0")) if ppr is not None else 0
        lw = avail_w - indent
        # 줄 나누기: 어절 단위, 어절이 줄보다 길면 글자 단위
        lines, cur = 1, 0.0
        for word in re.split(r"(\s+)", text):
            if not word:
                continue
            w = meas.width(word, face, size)
            if cur + w <= lw or (cur == 0 and w <= lw):
                cur += w
            elif w > lw:
                for ch in word:
                    cw = meas.width(ch, face, size)
                    if cur + cw > lw:
                        lines += 1; cur = 0
                    cur += cw
            else:
                if word.isspace():
                    continue
                lines += 1; cur = w
        spc = 1.0
        if ppr is not None:
            pct = ppr.find("a:lnSpc/a:spcPct", NS)
            if pct is not None:
                spc = int(pct.get("val")) / 100000
        total += lines * meas.line_h(face, size) * spc
        if ppr is not None:
            for t in ("spcBef", "spcAft"):
                pts = ppr.find(f"a:{t}/a:spcPts", NS)
                if pts is not None:
                    total += int(pts.get("val")) / 100 * 12700
    return total, avail_h


# ── 슬라이드 ────────────────────────────────────────────────────────────
def shapes(root):
    tree = root.find("p:cSld/p:spTree", NS)
    return [] if tree is None else [s for s in tree if ln(s) in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp")]


def xfrm_of(sh):
    if ln(sh) == "graphicFrame":
        return sh.find("p:xfrm", NS)
    if ln(sh) == "grpSp":
        return sh.find("p:grpSpPr/a:xfrm", NS)
    return sh.find("p:spPr/a:xfrm", NS)


def ph_type(sh):
    ph = sh.find("p:nvSpPr/p:nvPr/p:ph", NS)
    return (ph.get("type") or "body") if ph is not None else None


def max_sz(sh):
    z = [int(r.get("sz")) for r in sh.iter(q(A, "rPr"), q(A, "endParaRPr")) if r.get("sz")]
    return max(z) if z else 0


def title_shape(root):
    for sh in shapes(root):
        if ln(sh) == "sp" and ph_type(sh) in TITLE_PH and "".join(sh.itertext()).strip():
            return sh
    best = None
    for sh in shapes(root):
        if ln(sh) == "sp" and ph_type(sh) is None and max_sz(sh) >= 2800 and (best is None or max_sz(sh) > max_sz(best)):
            best = sh
    return best


def inherited_title_pos(deck, slide):
    """제목 자리표시자가 위치를 물려받을 때 레이아웃 → 마스터에서 찾은 위치."""
    lay = deck.first(slide, "slideLayout")
    mas = deck.first(lay, "slideMaster") if lay else None
    for part in (lay, mas):
        for lsh in shapes(deck.xml(part)) if part else []:
            xf = xfrm_of(lsh)
            if ph_type(lsh) in TITLE_PH and xf is not None and xf.find("a:off", NS) is not None:
                o = xf.find("a:off", NS)
                return int(o.get("x")), int(o.get("y"))
    return None


def image_size(data):
    try:
        from PIL import Image
        with Image.open(io.BytesIO(data)) as im:
            return im.size
    except Exception:
        return None



# ── 사진 자르기(smartcrop 방식) ─────────────────────────────────────────
def smart_crop(data, want):
    """jwagner/smartcrop.js(MIT)의 생각을 단순화: 가장자리(윤곽)와 채도가 많은 곳을 중요하다고 보고,
    원하는 비율의 창을 밀어 가며 중요도가 가장 큰 자리를 고른다(가운데를 우대해 주제가 잘리지 않게). 반환: 자를 비율(l, t, r, b)."""
    try:
        from PIL import Image, ImageFilter
        im = Image.open(io.BytesIO(data)).convert("RGB")
    except Exception:
        return 0, 0, 0, 0
    im.thumbnail((240, 240))
    w, h = im.size
    flat = lambda img: list(img.get_flattened_data() if hasattr(img, "get_flattened_data") else img.getdata())
    edge = flat(im.convert("L").filter(ImageFilter.FIND_EDGES))
    sat = flat(im.convert("HSV").split()[1])
    energy = [0.9 * e + 0.1 * s_ for e, s_ in zip(edge, sat)]
    if w / h > want:            # 이미지가 더 넓다 → 가로를 자른다
        cw = max(1, round(h * want))
        cols = [sum(energy[y * w + x] for y in range(h)) for x in range(w)]
        best, bx = -1, 0
        run = sum(cols[:cw])
        for x in range(0, w - cw + 1):
            if x:
                run += cols[x + cw - 1] - cols[x - 1]
            score = run * (1 - 0.8 * abs((x + cw / 2) / w - 0.5))
            if score > best:
                best, bx = score, x
        return bx / w, 0, (w - bx - cw) / w, 0
    ch = max(1, round(w / want))
    rows = [sum(energy[y * w:(y + 1) * w]) for y in range(h)]
    best, by = -1, 0
    run = sum(rows[:ch])
    for y in range(0, h - ch + 1):
        if y:
            run += rows[y + ch - 1] - rows[y - 1]
        score = run * (1 - 0.8 * abs((y + ch / 2) / h - 0.5))
        if score > best:
            best, by = score, y
    return 0, by / h, 0, (h - by - ch) / h


# ── 검사(intern 규칙)·접근성(Microsoft 접근성 검사·Community-Access 규칙) ──
HUE_GROUP = 30   # 색상환에서 30도 안은 같은 색으로 본다


def hue_groups(hexes):
    groups = []
    for hx in hexes:
        r, g, b = (int(hx[i:i + 2], 16) / 255 for i in (0, 2, 4))
        h, l_, sat = colorsys.rgb_to_hls(r, g, b)
        if sat < 0.18 or l_ > 0.93 or l_ < 0.07:
            continue          # 무채색(검정·회색·흰색)은 세지 않는다
        deg = h * 360
        if not any(min(abs(deg - d), 360 - abs(deg - d)) < HUE_GROUP for d in groups):
            groups.append(deg)
    return len(groups)


def lint_and_access(deck, slides, notes, deck_notes, fix, a):
    W, H = deck.size()
    media, titles, sizes, colors = [], [], Counter(), []
    alt_in = {}
    if a.alt:
        for it in json.load(open(a.alt, encoding="utf-8")):
            alt_in[(int(it["slide"]), it["name"])] = it["alt"].strip()
    for i, s in enumerate(slides, 1):
        root = deck.xml(s)
        t = title_shape(root)
        ttext = "".join(t.itertext()).strip() if t is not None else ""
        titles.append(ttext)
        if not ttext:
            notes[i].append(("접근성", "슬라이드 제목이 없습니다. 화면 낭독기는 제목으로 슬라이드를 찾습니다(진단만)"))
        elif re.search(r"[.,:;。]$", ttext):
            notes[i].append(("검사", f"제목 '{ttext[:16]}'이 마침표·쉼표로 끝납니다. 제목은 이름표라 문장부호를 빼는 편이 깔끔합니다(진단만)"))
        # 대체 텍스트
        for el in root.iter(q(P, "pic"), q(P, "graphicFrame")):
            nv = el.find(".//p:cNvPr", NS)
            if nv is None:
                continue
            kind = "그림" if ln(el) == "pic" else ("그래프" if el.find(".//c:chart", NS) is not None else ("표" if el.find(".//a:tbl", NS) is not None else "개체"))
            if kind == "표":
                continue
            deco = any("decorative" in etree.tostring(x).decode() for x in nv.iter(q(A, "ext")))
            key = (i, nv.get("name"))
            if key in alt_in and fix:
                nv.set("descr", alt_in[key]); deck.touch(s)
            d = (nv.get("descr") or "").strip()
            if (not d or re.search(r"\.(png|jpe?g|gif|bmp|emf|wmf|tiff?|svg)$", d, re.I)) and not deco:
                media.append({"slide": i, "name": nv.get("name"), "kind": kind, "now": d})
            elif len(nv.get("descr") or "") > 150:
                notes[i].append(("접근성", f"{kind} '{nv.get('name')}'의 대체 텍스트가 150자를 넘습니다. 짧게 줄이세요(진단만)"))
        # 표: 머리글 행·병합 칸
        for tbl in root.iter(q(A, "tbl")):
            tp = tbl.find("a:tblPr", NS)
            if tp is None or tp.get("firstRow") != "1":
                notes[i].append(("접근성", "표에 머리글 행 표시가 없습니다" + (" → 첫 줄을 머리글로 지정" if fix else "")))
                if fix:
                    if tp is None:
                        tp = etree.Element(q(A, "tblPr")); tbl.insert(0, tp)
                    tp.set("firstRow", "1"); deck.touch(s)
            if any(tc.get(k) for tc in tbl.iter(q(A, "tc")) for k in ("gridSpan", "rowSpan", "hMerge", "vMerge")):
                notes[i].append(("접근성", "표에 병합한 칸이 있습니다. 화면 낭독기가 칸 위치를 잃기 쉬우니 가능하면 병합을 풀어 주세요(진단만)"))
        # 빈 글상자·겹친 글상자·글 다듬기·크기 가짓수·색
        boxes = []
        for sh in list(shapes(root)):
            if ln(sh) != "sp" or sh.find("p:txBody", NS) is None:
                continue
            text = "".join(x.text or "" for x in sh.iter(q(A, "t")))
            if not text.strip():
                sp = sh.find("p:spPr", NS)
                visible = sp is not None and (sp.find("a:solidFill", NS) is not None or sp.find("a:gradFill", NS) is not None
                                              or sp.find("a:blipFill", NS) is not None or (sp.find("a:ln", NS) is not None and sp.find("a:ln/a:noFill", NS) is None))
                styled = sh.find("p:style", NS) is not None
                if ph_type(sh) is None and not visible and not styled:
                    notes[i].append(("검사", "보이지 않는 빈 글상자가 있습니다" + (" → 지움" if fix else "")))
                    if fix:
                        sh.getparent().remove(sh); deck.touch(s)
                continue
            xf = xfrm_of(sh)
            if xf is not None and xf.find("a:off", NS) is not None and xf.find("a:ext", NS) is not None:
                o, e = xf.find("a:off", NS), xf.find("a:ext", NS)
                boxes.append((int(o.get("x")), int(o.get("y")), int(e.get("cx")), int(e.get("cy")), text.strip()[:10]))
            for p_ in sh.findall("p:txBody/a:p", NS):
                pt = "".join(x.text or "" for x in p_.iter(q(A, "t")))
                if re.search(r"\S {2,}\S", pt):
                    notes[i].append(("검사", f"'{pt.strip()[:12]}' 문단 가운데 띄어쓰기가 두 칸 이상입니다(개조식 정렬용이면 그대로 두세요, 진단만)"))
                if sh is not t and ph_type(sh) in ("body", "obj") and len(pt.strip()) > 70:
                    notes[i].append(("검사", f"'{pt.strip()[:12]}…' 글머리 한 줄이 {len(pt.strip())}자입니다. 발표 화면의 요점은 짧을수록 읽힙니다(진단만)"))
            if sh is not t and ph_type(sh) not in ("ctrTitle", "subTitle", "title"):
                for rp in sh.iter(q(A, "rPr")):
                    if rp.get("sz") and int(rp.get("sz")) > 1200:
                        sizes[int(rp.get("sz"))] += 1
            for c in sh.iter(q(A, "srgbClr")):
                colors.append(c.get("val").upper())
        for j in range(len(boxes)):
            for k in range(j + 1, len(boxes)):
                x1, y1, w1, h1, n1 = boxes[j]; x2, y2, w2, h2, n2 = boxes[k]
                ov = max(0, min(x1 + w1, x2 + w2) - max(x1, x2)) * max(0, min(y1 + h1, y2 + h2) - max(y1, y2))
                if ov > 0.15 * min(w1 * h1, w2 * h2):
                    notes[i].append(("검사", f"글상자 '{n1}'와 '{n2}'가 겹칩니다(진단만)"))
        for c in root.iter(q(P, "pic")):
            pass
    # 덱 전체
    dup = [x for x, n in Counter(t for t in titles if t).items() if n > 1]
    if dup:
        deck_notes.append(("접근성", "같은 제목을 쓴 슬라이드가 있습니다: " + ", ".join(f"'{d[:14]}'" for d in dup[:3]) + ". '(계속)'처럼 구별해 주세요(진단만)"))
    if len(sizes) > 3:
        top = ", ".join(f"{k // 100}pt" for k, _ in sizes.most_common())
        deck_notes.append(("검사", f"본문 글자 크기가 {len(sizes)}가지입니다({top}). 2~3가지로 줄이면 위계가 분명해집니다(진단만)"))
    for n in sorted(x for x in deck.names if re.match(r"ppt/charts/chart\d+\.xml$", x)):
        for c in deck.xml(n).iter(q(A, "srgbClr")):
            colors.append(c.get("val").upper())
    groups = hue_groups(set(colors))
    if groups > 3:
        deck_notes.append(("디자인", f"직접 칠한 색이 {groups}가지 색 계열입니다. 제목색·강조색 하나·회색, 세 가지 안으로 줄이면 깔끔합니다"
                           " (style_deck.py --level bold가 가까운 색으로 정리, 진단만)"))
    if media:
        by = Counter(m["kind"] for m in media)
        deck_notes.append(("접근성", "대체 텍스트가 없는 " + ", ".join(f"{k} {v}개" for k, v in by.items())
                           + ". 그림을 보고 설명을 써서 --alt로 넣으세요(목록은 --json의 media)"))
    # 문서 속성: 제목·언어
    cp = "docProps/core.xml"
    if cp in deck.names:
        core = deck.xml(cp)
        DC = "http://purl.org/dc/elements/1.1/"
        tt = core.find(f"{{{DC}}}title")
        if (tt is None or not (tt.text or "").strip()) and titles and titles[0]:
            deck_notes.append(("접근성", "문서 속성에 제목이 없습니다" + (f" → 첫 슬라이드 제목 '{titles[0][:20]}'로" if fix else "")))
            if fix:
                if tt is None:
                    tt = etree.SubElement(core, f"{{{DC}}}title")
                tt.text = titles[0]; deck.touch(cp)
        lg = core.find(f"{{{DC}}}language")
        if lg is None or not (lg.text or "").strip():
            deck_notes.append(("접근성", "문서 언어가 지정되지 않았습니다" + (" → ko-KR" if fix else "")))
            if fix:
                if lg is None:
                    lg = etree.SubElement(core, f"{{{DC}}}language")
                lg.text = "ko-KR"; deck.touch(cp)
    return media


def main():
    ap = argparse.ArgumentParser(description="피피티 그래프·디자인 진단과 보정")
    ap.add_argument("pptx")
    ap.add_argument("--fix", metavar="OUT", help="보정한 파일을 이 경로로 저장(없으면 진단만)")
    ap.add_argument("--manifest", help="fetch_fonts.py의 manifest.json(넘침 추정에 실제 글꼴 폭을 씀)")
    ap.add_argument("--chart-colors", action="store_true", help="그래프 계열 색을 테마 강조색으로")
    ap.add_argument("--contrast", action="store_true", help="대비가 낮은 글자를 검정·흰색 중 대비가 큰 쪽으로")
    ap.add_argument("--fit", action="store_true", help="넘치는 것으로 추정되는 글자를 줄여 상자에 맞춤")
    ap.add_argument("--no-align", action="store_true", help="위치 맞추기(제목 위치·왼쪽 끝)를 하지 않음")
    ap.add_argument("--no-images", action="store_true", help="사진 비율 바로잡기를 하지 않음")
    ap.add_argument("--image-fit", choices=["contain", "crop"], default="contain",
                    help="늘어난 사진: contain=틀 안에 맞춰 넣기(기본), crop=틀을 지키고 중요한 부분을 남겨 자르기(smartcrop 방식)")
    ap.add_argument("--alt", help="대체 텍스트 JSON([{slide, name, alt}])을 그림·그래프에 넣는다(--fix와 함께)")
    ap.add_argument("--json", help="진단 결과를 JSON으로 저장(대체 텍스트가 없는 그림·그래프 목록 media 포함)")
    a = ap.parse_args()
    fix = bool(a.fix)
    deck = Deck(a.pptx)
    W, H = deck.size()
    slides = deck.slides()
    meas = Measure(a.manifest)
    notes = defaultdict(list)   # 슬라이드 번호 → [(분류, 내용)]
    stats = Counter()

    # 그래프
    for i, s in enumerate(slides, 1):
        root = deck.xml(s)
        rel = deck.rels(s)
        for gf in root.iter(q(P, "graphicFrame")):
            ch = gf.find(".//c:chart", NS)
            if ch is None:
                continue
            tgt = rel.get(ch.get(q(R, "id")))
            if not tgt or tgt[0] not in deck.names:
                continue
            x = gf.find("p:xfrm/a:ext", NS)
            fw = int(x.get("cx")) if x is not None else W
            croot = deck.xml(tgt[0])
            n0 = len(notes[i])
            if fix_chart(croot, fw, fix, a, notes[i]):
                deck.touch(tgt[0])
            stats["chart_notes"] += len(notes[i]) - n0
            stats["charts"] += 1

    # 사진 비율
    for i, s in enumerate(slides, 1):
        root = deck.xml(s)
        rel = deck.rels(s)
        for pic in shapes(root):
            if ln(pic) != "pic" or a.no_images:
                continue
            blip = pic.find("p:blipFill/a:blip", NS)
            xf = pic.find("p:spPr/a:xfrm", NS)
            if blip is None or xf is None or xf.find("a:ext", NS) is None or xf.get("rot"):
                continue
            tgt = rel.get(blip.get(q(R, "embed")))
            if not tgt or tgt[0] not in deck.names:
                continue
            px = image_size(deck.z.read(tgt[0]))
            if not px:
                continue
            sr = pic.find("p:blipFill/a:srcRect", NS)
            g = lambda k: int(sr.get(k, "0")) / 100000 if sr is not None else 0
            sw, sh_ = px[0] * (1 - g("l") - g("r")), px[1] * (1 - g("t") - g("b"))
            if sw <= 0 or sh_ <= 0:
                continue
            off, ext = xf.find("a:off", NS), xf.find("a:ext", NS)
            cx, cy = int(ext.get("cx")), int(ext.get("cy"))
            if cy == 0:
                continue
            want, have = sw / sh_, cx / cy
            if abs(have / want - 1) > 0.03:
                notes[i].append(("디자인", f"사진 비율이 {abs(have / want - 1) * 100:.0f}% 틀어져 늘어나 보입니다" + (" → 원래 비율로" if fix else "")))
                if fix and a.image_fit == "crop" and (sr is None or not any(int(sr.get(k, "0")) for k in "ltrb")):
                    l_, t_, r_, b_ = smart_crop(deck.z.read(tgt[0]), have)
                    bf = pic.find("p:blipFill", NS)
                    if sr is None:
                        sr = etree.Element(q(A, "srcRect")); blip.addnext(sr)
                    for k, v in zip("ltrb", (l_, t_, r_, b_)):
                        if v:
                            sr.set(k, str(int(v * 100000)))
                    notes[i][-1] = ("디자인", notes[i][-1][1].replace("→ 원래 비율로", "→ 틀은 그대로 두고 중요한 부분을 남겨 자름(smartcrop 방식)"))
                    deck.touch(s)
                elif fix:   # 원래 틀 안에 맞춰 넣는다(틀보다 커지지 않아 주변 개체를 덮지 않음), 틀 가운데에 둔다
                    if have > want:
                        ncx = round(cy * want); off.set("x", str(int(off.get("x")) + (cx - ncx) // 2)); ext.set("cx", str(ncx))
                    else:
                        ncy = round(cx / want); off.set("y", str(int(off.get("y")) + (cy - ncy) // 2)); ext.set("cy", str(ncy))
                    deck.touch(s)

    # 제목 위치: 여러 슬라이드에서 가장 많이 쓴 위치, 없으면 레이아웃이 정한 위치에 맞춘다
    if not a.no_align:
        tpos, votes, layout_pos = [], Counter(), None
        for i, s in enumerate(slides, 1):
            t = title_shape(deck.xml(s))
            if t is None or ph_type(t) == "ctrTitle" or i == 1:
                continue
            xf = xfrm_of(t)
            if xf is not None and xf.find("a:off", NS) is not None:
                o = xf.find("a:off", NS)
                tpos.append((i, s, o, int(o.get("x")), int(o.get("y"))))
                votes[(int(o.get("x")), int(o.get("y")))] += 1
            elif ph_type(t) in TITLE_PH:   # 위치를 레이아웃·마스터에서 물려받는 제목
                pos = inherited_title_pos(deck, s)
                if pos:
                    votes[pos] += 1; layout_pos = layout_pos or pos
        target = None
        if votes:
            pos, n = votes.most_common(1)[0]
            target = pos if n >= 2 else layout_pos
        if target is None and tpos:   # 모두 직접 위치를 가진 경우: 첫 내용 슬라이드 레이아웃의 제목 위치
            target = inherited_title_pos(deck, tpos[0][1])
        if target and len(tpos) + (1 if layout_pos else 0) >= 2:
            mx, my = target
            for i, s, o, x, y in tpos:
                if (x, y) != (mx, my) and abs(x - mx) <= 0.3 * IN and abs(y - my) <= 0.3 * IN:
                    notes[i].append(("디자인", f"제목 위치가 다른 슬라이드와 {max(abs(x - mx), abs(y - my)) / IN * 2.54:.1f}cm 어긋납니다" + (" → 맞춤" if fix else "")))
                    if fix:
                        o.set("x", str(mx)); o.set("y", str(my)); deck.touch(s)
        # 한 슬라이드 안의 왼쪽 끝
        for i, s in enumerate(slides, 1):
            root = deck.xml(s)
            items = []
            for sh in shapes(root):
                xf = xfrm_of(sh)
                if xf is None or xf.get("rot") or xf.find("a:off", NS) is None:
                    continue
                x = int(xf.find("a:off", NS).get("x"))
                if x > 0:
                    items.append((x, xf.find("a:off", NS)))
            cnt = Counter(x for x, _ in items)
            moved = 0
            for x, o in items:
                near = [v for v in cnt if v != x and abs(v - x) <= 0.1 * IN and (cnt[v] > cnt[x] or (cnt[v] == cnt[x] and v < x))]
                if near:
                    tgt = max(near, key=lambda v: (cnt[v], -v))
                    moved += 1
                    if fix:
                        o.set("x", str(tgt))
            if moved:
                notes[i].append(("디자인", f"왼쪽 끝이 거의 같은데 조금 어긋난 개체 {moved}개" + (" → 맞춤" if fix else "")))
                if fix:
                    deck.touch(s)
            # 위쪽 끝(intern CLOSE_Y): 같은 줄에 놓인 개체의 윗변이 0.1인치 안으로 어긋남
            items = []
            for sh in shapes(root):
                xf = xfrm_of(sh)
                if xf is None or xf.get("rot") or xf.find("a:off", NS) is None or sh is title_shape(root):
                    continue
                items.append((int(xf.find("a:off", NS).get("y")), xf.find("a:off", NS)))
            cnt = Counter(y for y, _ in items)
            moved = 0
            for y, o in items:
                near = [v for v in cnt if v != y and abs(v - y) <= 0.1 * IN and (cnt[v] > cnt[y] or (cnt[v] == cnt[y] and v < y))]
                if near:
                    moved += 1
                    if fix:
                        o.set("y", str(max(near, key=lambda v: (cnt[v], -v))))
            if moved:
                notes[i].append(("디자인", f"위쪽 끝이 거의 같은데 조금 어긋난 개체 {moved}개" + (" → 맞춤" if fix else "")))
                if fix:
                    deck.touch(s)

    # 슬라이드 밖, 작은 글자, 대비, 넘침
    colors_used = Counter()
    for i, s in enumerate(slides, 1):
        root = deck.xml(s)
        col = Colors(deck, s)
        bg = col.background()
        faces = theme_faces(deck, s)
        small, low, out = 0, [], 0
        for sh in shapes(root):
            xf = xfrm_of(sh)
            if xf is not None and xf.find("a:off", NS) is not None and xf.find("a:ext", NS) is not None:
                x, y = int(xf.find("a:off", NS).get("x")), int(xf.find("a:off", NS).get("y"))
                cx, cy = int(xf.find("a:ext", NS).get("cx")), int(xf.find("a:ext", NS).get("cy"))
                if x < -0.1 * IN or y < -0.1 * IN or x + cx > W + 0.1 * IN or y + cy > H + 0.1 * IN:
                    out += 1
                bottom = y > H * 0.85
            else:
                bottom = False
            if ln(sh) != "sp" or sh.find("p:txBody", NS) is None:
                continue
            # 상자 배경
            fill = sh.find("p:spPr/a:solidFill", NS)
            nofill = sh.find("p:spPr/a:noFill", NS) is not None
            ref = sh.find("p:style/a:fillRef", NS)
            box_bg = col.of(fill) if fill is not None else (None if nofill else (col.of(ref) if ref is not None and ref.get("idx") not in (None, "0") else None))
            if fill is None and (nofill or ref is None or ref.get("idx") in (None, "0")):
                box_bg = bg
            fref = sh.find("p:style/a:fontRef", NS)
            default_tx = col.of(fref) if fref is not None else col.scheme.get(col.clrmap.get("tx1", "dk1"))
            for r in sh.iter(q(A, "r")):
                rp = r.find("a:rPr", NS)
                txt = r.findtext("a:t", default="", namespaces=NS)
                if not txt.strip():
                    continue
                sz = int(rp.get("sz")) if rp is not None and rp.get("sz") else None
                if sz and sz < 1200 and not bottom:
                    small += 1
                sf = rp.find("a:solidFill", NS) if rp is not None else None
                tc = col.of(sf) if sf is not None else default_tx
                if sf is not None and sf.find("a:srgbClr", NS) is not None:
                    colors_used[sf.find("a:srgbClr", NS).get("val").upper()] += 1
                if tc and box_bg:
                    big = (sz or 1800) >= 1800 or ((sz or 1800) >= 1400 and rp is not None and rp.get("b") == "1")
                    need = 3.0 if big else 4.5
                    cr = contrast(tc, box_bg)
                    if cr < need:
                        low.append((txt.strip()[:12], cr))
                        if fix and a.contrast:
                            best = "1A1A1A" if contrast("1A1A1A", box_bg) >= contrast("FFFFFF", box_bg) else "FFFFFF"
                            if rp is None:
                                rp = etree.Element(q(A, "rPr")); r.insert(0, rp)
                            for x in rp.findall("a:solidFill", NS):
                                rp.remove(x)
                            nf = etree.Element(q(A, "solidFill")); c = etree.SubElement(nf, q(A, "srgbClr")); c.set("val", best)
                            kids = [k for k in rp if ln(k) in ("ln", "noFill")]
                            (kids[-1].addnext(nf) if kids else rp.insert(0, nf))
                            deck.touch(s)
            # 넘침
            res = box_need(sh, meas, faces)
            if res and res[0] > res[1] * 1.08 + 0.1 * IN:
                pct = res[0] / max(res[1], 1) * 100 - 100
                title = ph_type(sh) in TITLE_PH
                label = "".join(sh.itertext()).strip()[:14]
                if fix and a.fit:
                    floor = 2000 if title else 1200
                    lo, hi = 0.3, 1.0
                    for _ in range(18):
                        mid = (lo + hi) / 2
                        r2 = box_need(sh, meas, faces, mid)
                        if r2 and r2[0] <= r2[1]:
                            lo = mid
                        else:
                            hi = mid
                    scale = lo
                    szs = [int(rp.get("sz")) for rp in sh.iter(q(A, "rPr"), q(A, "endParaRPr")) if rp.get("sz")]
                    floor_scale = max(floor / max(szs), 0.3) if szs else 1
                    applied = max(scale, floor_scale)
                    for rp in sh.iter(q(A, "rPr"), q(A, "endParaRPr")):
                        if rp.get("sz"):
                            rp.set("sz", str(max(int(int(rp.get("sz")) * applied / 50) * 50, 800)))
                    for rr in sh.iter(q(A, "r")):
                        if rr.find("a:rPr", NS) is None:
                            e = etree.Element(q(A, "rPr")); e.set("sz", str(int(1800 * applied / 50) * 50)); rr.insert(0, e)
                    deck.touch(s)
                    msg = f"'{label}' 상자를 글자가 약 {pct:.0f}% 넘치는 것으로 추정됩니다 → 글자를 {applied * 100:.0f}%로 줄임"
                    if applied > scale:
                        msg += f"(하한 {floor // 100}pt에 걸려 다 맞지 않음, 상자를 키우거나 글을 줄이세요)"
                    notes[i].append(("디자인", msg))
                else:
                    notes[i].append(("디자인", f"'{label}' 상자를 글자가 약 {pct:.0f}% 넘치는 것으로 추정됩니다(--fit으로 줄이거나 상자를 키우세요)"))
        if out:
            notes[i].append(("디자인", f"슬라이드 밖으로 나간 개체 {out}개(진단만)"))
        if small:
            notes[i].append(("디자인", f"12pt보다 작은 글자 {small}곳(바닥글 제외). 발표 화면에서 읽기 어렵습니다(진단만)"))
        if low:
            worst = min(low, key=lambda t: t[1])
            notes[i].append(("디자인", f"배경과 대비가 낮은 글자 {len(low)}곳(가장 낮은 곳 '{worst[0]}' {worst[1]:.1f}:1, 기준 4.5:1)"
                             + (" → 검정·흰색 중 대비가 큰 쪽으로" if fix and a.contrast else "")))
    deck_notes = []
    media = lint_and_access(deck, slides, notes, deck_notes, fix, a)

    # 출력
    total = sum(len(v) for v in notes.values()) + len(deck_notes)
    print(f"{'보정' if fix else '진단'}: {a.pptx}  (슬라이드 {len(slides)}장, 그래프 {stats['charts']}개, 지적 {total}건)")
    for i in sorted(notes):
        for kind, msg in notes[i]:
            print(f"  {i:>2}장 [{kind}] {msg}")
    for kind, msg in deck_notes:
        print(f"  전체 [{kind}] {msg}")
    if not total:
        print("  지적할 것이 없습니다.")
    if fix:
        deck.save(a.fix)
        print(f"저장: {a.fix} (바뀐 XML {len(deck.dirty)}개)")
    if a.json:
        json.dump({"slides": {str(k): [m for _, m in v] for k, v in notes.items()}, "deck": [m for _, m in deck_notes], "media": media},
                  open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
