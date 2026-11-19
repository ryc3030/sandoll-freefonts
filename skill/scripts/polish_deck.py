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
    # 6) 원그래프 조각 수(진단만)
    for ct in types:
        if ln(ct) in ("pieChart", "pie3DChart", "doughnutChart"):
            for s in ct.findall("c:ser", NS):
                n = len(s.findall("c:val//c:pt", NS))
                if n >= 7:
                    notes.append(("그래프", f"원그래프 조각이 {n}개라 비교하기 어렵습니다. 막대 그래프나 '기타' 묶기를 권합니다(진단만)"))
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
    ap.add_argument("--json", help="진단 결과를 JSON으로 저장")
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
                if fix:   # 원래 틀 안에 맞춰 넣는다(틀보다 커지지 않아 주변 개체를 덮지 않음), 틀 가운데에 둔다
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
    if len(colors_used) > 6:
        deck_notes.append(("디자인", f"직접 지정한 글자 색이 {len(colors_used)}가지입니다. 강조색은 2~3가지로 줄이는 편이 읽기 쉽습니다(진단만)"))

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
        json.dump({"slides": {str(k): [m for _, m in v] for k, v in notes.items()}, "deck": [m for _, m in deck_notes]},
                  open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
