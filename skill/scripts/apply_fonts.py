#!/usr/bin/env python3
"""PPTX 전체의 글꼴을 제목용·본문용 두 가지로 통일한다.

  # 권장: fetch_fonts.py의 manifest와 '이름:굵기'로 지정(실제 글꼴 이름·굵게 처리·빠진 기호 검사를 알아서 함)
  python apply_fonts.py in.pptx out.pptx --manifest fonts/manifest.json --title 프리텐다드:ExtraBold --body 프리텐다드:Regular
  # 제목이 영문 전용 폰트일 때
  python apply_fonts.py in.pptx out.pptx --manifest fonts/manifest.json --title "Instrument Serif:Regular" --title-ea "마루 부리:Bold" --body "마루 부리:Regular"
  # manifest 없이 PowerPoint 이름을 직접 줄 수도 있다(이때 '이름:굵기'는 쓸 수 없음)
  python apply_fonts.py in.pptx out.pptx --title "Pretendard ExtraBold" --body "Pretendard"

하는 일
  1) 테마의 제목 글꼴(majorFont)·본문 글꼴(minorFont)을 latin·ea(한글)·Hang 항목까지 바꾼다.
  2) 슬라이드·레이아웃·마스터·노트·차트·SmartArt·presentation.xml 기본 스타일·테마 기본 도형에 직접 박힌 글꼴을
     테마 참조(+mj-*/+mn-*)로 바꿔 모든 글자가 1)을 따르게 한다(차트는 이름을 직접 씀).
  3) 제목: 제목 자리표시자(title, ctrTitle) + 자리표시자가 아닌 상자 중 슬라이드에서 가장 큰 글자(28pt 이상)를 가진 상자.
  4) 굵게: 제목 글꼴이 'Pretendard ExtraBold'처럼 굵기가 이름에 든 글꼴이면 제목의 b=1을 꺼 가짜 굵게를 막는다.
     본문의 b=1 글자는 manifest에 본문 가족의 Bold가 있고 그 이름이 따로 있으면(예: 'Paperlogy 7 Bold') 그 글꼴로 바꾼다.
  5) 기호·아이콘·코드 글꼴(Wingdings, Font Awesome, Consolas 등)과 --keep 글꼴은 건드리지 않는다.
  6) 덱에 쓰인 글자 중 고른 글꼴에 없는 기호가 있으면 경고한다.
  7) 'Noto Sans KR Black'처럼 굵기가 이름에 든 글꼴은 굵기를 살려, 새 가족에서 가장 가까운 굵기로 바꾼다(manifest에 있을 때).
  8) 테마와 글자의 cs(복합 문자) 글꼴도 함께 맞춘다. 한글 입력 중 PowerPoint가 글꼴을 잘못 고르는 일을 막는다.
  --theme-xml 을 주면 PowerPoint '디자인 → 글꼴'에 등록할 수 있는 글꼴 테마 파일도 만든다(새 슬라이드에도 같은 글꼴).

참고: 7)·8)과 글꼴 테마 파일은 achimnol/pptx-tool(MIT)의 처리 방식을 참고했다.

--mode force 는 테마 참조 대신 모든 글자에 글꼴 이름을 직접 써 넣는다(Google Slides 등 테마를 무시하는 프로그램용).
"""
import argparse, json, os, re, shutil, sys, tempfile, zipfile
from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}
KEEP_DEFAULT = re.compile(r"^(wingdings.*|symbol|webdings|marlett|segoe (ui (symbol|emoji)|mdl2 assets|fluent icons)|"
                          r"font ?awesome.*|material (icons|symbols).*|consolas|courier( new)?|d2coding.*|nanum ?gothic ?coding|"
                          r"cascadia.*|jetbrains mono|source code pro|menlo|monaco)$", re.I)
RPR_TAGS = {f"{{{A}}}{t}" for t in ("rPr", "defRPr", "endParaRPr")}
TITLE_PH = {"title", "ctrTitle"}
WORDER = ["Thin", "ExtraLight", "Light", "Regular", "Medium", "SemiBold", "Bold", "ExtraBold", "Black"]
WALIAS = {"ultralight": "ExtraLight", "semilight": "Light", "demilight": "Light", "book": "Regular", "normal": "Regular",
          "demibold": "SemiBold", "ultrabold": "ExtraBold", "heavy": "Black"}
WSUFFIX = re.compile(r"\s+(thin|extralight|ultralight|semilight|demilight|light|book|normal|regular|medium|semibold|demibold|"
                     r"bold|extrabold|ultrabold|black|heavy)$", re.I)


def weight_of(face):
    """'Noto Sans KR Black' → 'Black'. 굵기 접미사가 없으면 None."""
    m = WSUFFIX.search(face or "")
    if not m:
        return None
    w = m.group(1).lower()
    return WALIAS.get(w) or next(x for x in WORDER if x.lower() == w)
ORDER = ["ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "effectLst", "effectDag",
         "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]


def q(tag):
    return f"{{{A}}}{tag}"


def reorder(rpr):
    kids = list(rpr)
    kids.sort(key=lambda e: ORDER.index(etree.QName(e).localname) if etree.QName(e).localname in ORDER else 99)
    for k in kids:
        rpr.remove(k)
    for k in kids:
        rpr.append(k)


def set_face(rpr, child, face):
    el = rpr.find(f"a:{child}", NS)
    if el is None:
        el = etree.SubElement(rpr, q(child)); reorder(rpr)
    el.set("typeface", face)


def set_theme(root, fonts):
    n = 0
    for tag, role in (("majorFont", "title"), ("minorFont", "body")):
        for fs in root.iter(q(tag)):
            for i, child in enumerate(("latin", "ea", "cs")):
                el = fs.find(f"a:{child}", NS)
                if el is None:
                    el = etree.Element(q(child))
                    fs.insert(i, el)  # 순서: latin, ea, cs, font*
                if child == "cs":
                    if not el.get("typeface"):  # 비어 있을 때만 한글 글꼴로(아랍·태국 글꼴이 지정돼 있으면 둔다)
                        el.set("typeface", fonts[(role, "ea")]); n += 1
                    continue
                el.set("typeface", fonts[(role, child)]); n += 1
            hang = [f for f in fs.findall("a:font", NS) if f.get("script") == "Hang"]
            if hang:
                hang[0].set("typeface", fonts[(role, "ea")])
            else:
                h = etree.SubElement(fs, q("font")); h.set("script", "Hang"); h.set("typeface", fonts[(role, "ea")])
            n += 1
    return n


def max_size(sp):
    z = [int(r.get("sz")) for r in sp.iter(q("rPr"), q("defRPr"), q("endParaRPr")) if r.get("sz")]
    return max(z) if z else 0


def title_like(root):
    """자리표시자가 아닌 상자 중 슬라이드에서 가장 큰 글자(28pt 이상)를 가진 상자들."""
    best, found = 0, []
    for sp in root.iter(f"{{{P}}}sp"):
        ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
        if ph is not None and ph.get("type") in TITLE_PH and "".join(t.text or "" for t in sp.iter(q("t"))).strip():
            return set()  # 자리표시자 제목이 있는 장은 그것이 제목
    for sp in root.iter(f"{{{P}}}sp"):
        if sp.find("p:nvSpPr/p:nvPr/p:ph", NS) is not None or sp.find("p:txBody", NS) is None:
            continue
        m = max_size(sp)
        if m >= 2800:
            if m > best:
                best, found = m, [sp]
            elif m == best:
                found.append(sp)
    return set(found)


def is_title(el, extra):
    for anc in el.iterancestors():
        if anc.tag == f"{{{P}}}titleStyle":
            return True
        if anc.tag in (f"{{{P}}}bodyStyle", f"{{{P}}}otherStyle"):
            return False
        if anc.tag == f"{{{P}}}sp":
            if anc in extra:
                return True
            ph = anc.find("p:nvSpPr/p:nvPr/p:ph", NS)
            return ph is not None and ph.get("type") in TITLE_PH
    return False


def fix_runs(root, cfg, mode, stats, title_shapes=frozenset()):
    """root 안 모든 글자 서식을 정리한다. mode: theme | force"""
    # 글자(run)에 rPr이 없으면 만들어 둔다(제목 상자 표시나 force 모드에서 글꼴을 넣기 위해).
    if mode == "force" or title_shapes:
        for r in list(root.iter(q("r"), q("fld"))):
            if r.find("a:rPr", NS) is None and (mode == "force" or is_title(r, title_shapes)):
                r.insert(0, etree.Element(q("rPr")))
    for rpr in root.iter(*RPR_TAGS):
        title = is_title(rpr, title_shapes)
        role = "title" if title else "body"
        # 굵게 처리
        if title and cfg["title_nobold"] and rpr.get("b") == "1":
            rpr.set("b", "0"); stats["bold_title"] += 1
        bold_face = None
        if not title and rpr.get("b") == "1":
            if cfg["body_bold_face"]:
                bold_face = cfg["body_bold_face"]; rpr.set("b", "0"); stats["bold_body"] += 1
            else:
                stats["bold_body_fake"] += 1
        # 굵기가 이름에 든 원래 글꼴('Noto Sans KR Black')은 새 가족의 가장 가까운 굵기로
        weighted = None
        if not bold_face and cfg["weights"].get(role):
            for child in ("ea", "latin"):
                el = rpr.find(f"a:{child}", NS)
                cur = el.get("typeface") if el is not None else None
                if cur and not cur.startswith("+") and not KEEP_DEFAULT.match(cur) and cur.lower() not in cfg["keep"]:
                    w = weight_of(cur)
                    if w:
                        weighted = pick_weight(cfg["weights"][role], w, cfg["base"][role])
                    break
        if weighted is not None:
            face, bold = weighted
            if face is not None:
                for child in ("latin", "ea", "cs"):
                    set_face(rpr, child, face)
                if bold:
                    rpr.set("b", "1")
                elif rpr.get("b") == "1":
                    rpr.set("b", "0")
                stats["weight_kept"] += 1
                continue
        for child in ("latin", "ea", "cs"):
            el = rpr.find(f"a:{child}", NS)
            cur = el.get("typeface") if el is not None else None
            if cur and (KEEP_DEFAULT.match(cur) or cur.lower() in cfg["keep"]):
                continue
            slot = {"latin": "lt", "ea": "ea", "cs": "cs"}[child]
            if child == "cs" and el is None:
                continue  # cs는 원래 지정된 곳만 정리
            if bold_face:
                set_face(rpr, child, bold_face); stats["runs"] += 1
            elif mode == "force":
                set_face(rpr, child, cfg["fonts"][(role, "latin" if child == "latin" else "ea")]); stats["runs"] += 1
            elif title and title_shapes and rpr.tag == q("rPr") and is_title(rpr, title_shapes) and not _is_ph_title(rpr):
                set_face(rpr, child, f"+mj-{slot}"); stats["runs"] += 1
            elif cur and not cur.startswith("+"):
                el.set("typeface", f"+{'mj' if title else 'mn'}-{slot}"); stats["runs"] += 1
    return stats


def pick_weight(entries, want, base):
    """entries: 그 역할 가족의 manifest 항목들. 원하는 굵기와 가장 가까운 것을 (글꼴 이름, 굵게 여부)로.
    가장 가까운 것이 역할의 기본 굵기면 (None, _) → 테마 참조로 둔다."""
    have = {e["weight"]: e for e in entries if e["weight"] in WORDER}
    if not have:
        return None
    wi = WORDER.index(want)
    best = min(have, key=lambda w: (abs(WORDER.index(w) - wi), -WORDER.index(w) if wi >= 3 else WORDER.index(w)))
    e = have[best]
    if base and e["weight"] == base["weight"]:
        return (None, False)
    return (e["family"], bool(e["bold"]))


def _is_ph_title(el):
    for anc in el.iterancestors(f"{{{P}}}sp"):
        ph = anc.find("p:nvSpPr/p:nvPr/p:ph", NS)
        return ph is not None and ph.get("type") in TITLE_PH
    return False


def write_theme_xml(path, fonts):
    """Office 글꼴 테마(fontScheme) 파일. Windows: %APPDATA%\\Microsoft\\Templates\\Document Themes\\Theme Fonts\\
    macOS: ~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Themes.localized/Theme Fonts/"""
    name = os.path.splitext(os.path.basename(path))[0]
    root = etree.Element(q("fontScheme"), nsmap={"a": A}); root.set("name", name)
    for tag, role in (("majorFont", "title"), ("minorFont", "body")):
        fs = etree.SubElement(root, q(tag))
        for child, key in (("latin", "latin"), ("ea", "ea"), ("cs", "ea")):
            etree.SubElement(fs, q(child)).set("typeface", fonts[(role, key)])
        h = etree.SubElement(fs, q("font")); h.set("script", "Hang"); h.set("typeface", fonts[(role, "ea")])
    open(path, "wb").write(etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True, pretty_print=True))


def deck_chars(z):
    s = set()
    for n in z.namelist():
        if re.match(r"ppt/(slides/slide|charts/chart|diagrams/data)\d+\.xml$", n):
            for t in etree.fromstring(z.read(n)).iter(q("t")):
                s.update(t.text or "")
    return s


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src"); ap.add_argument("dst")
    ap.add_argument("--title", required=True, help="제목 글꼴. --manifest가 있으면 '이름:굵기'")
    ap.add_argument("--body", required=True, help="본문 글꼴. --manifest가 있으면 '이름:굵기'")
    ap.add_argument("--title-ea", help="제목의 한글 글꼴(제목 글꼴이 영문 전용일 때)")
    ap.add_argument("--body-ea", help="본문의 한글 글꼴(본문 글꼴이 영문 전용일 때)")
    ap.add_argument("--keep", default="", help="건드리지 않을 글꼴, 쉼표 구분(기호·아이콘·코드 글꼴은 기본으로 보존)")
    ap.add_argument("--mode", choices=["theme", "force"], default="theme")
    ap.add_argument("--title-bold", action="store_true", help="manifest 없이 쓸 때, 제목을 굵게 스타일로")
    ap.add_argument("--manifest", help="fetch_fonts.py가 만든 manifest.json")
    ap.add_argument("--theme-xml", help="PowerPoint 글꼴 테마 파일(.xml)도 만든다. 디자인 → 글꼴에 등록하면 새 슬라이드에도 같은 글꼴")
    a = ap.parse_args()

    man = None
    if a.manifest:
        man = json.load(open(a.manifest, encoding="utf-8"))
    elif any(":" in (x or "") for x in (a.title, a.body, a.title_ea, a.body_ea)):
        sys.exit("'이름:굵기' 형식은 --manifest와 함께 써야 합니다. fetch_fonts.py로 받은 뒤 --manifest fonts/manifest.json을 주세요.")

    def norm(s):
        return (s or "").replace(" ", "").lower()

    def entry(spec):
        if not man or ":" not in spec:
            return None
        name, w = spec.rsplit(":", 1)
        keys = lambda m: {norm(m["name"]), norm(m["family"]), norm(m.get("typoFamily")), norm(m["font"].split(":", 1)[-1])}
        for m in man:
            if norm(name) in keys(m) and m["weight"].lower() == w.lower():
                return m
        # 짧은 이름(예: '주아' → '배민 주아체')은 한 가족에만 들어맞을 때 받아들인다
        part = [m for m in man if m["weight"].lower() == w.lower() and any(norm(name) in k for k in keys(m))]
        if len({m["font"] for m in part}) == 1:
            return part[0]
        sys.exit(f"manifest에서 '{spec}'을(를) 찾지 못했습니다. 받은 것: " + ", ".join(f"{m['name']}:{m['weight']}" for m in man))

    te, be = entry(a.title), entry(a.body)
    tea, bea = entry(a.title_ea) if a.title_ea else None, entry(a.body_ea) if a.body_ea else None
    face = lambda e, raw: e["family"] if e else raw
    fonts = {("title", "latin"): face(te, a.title), ("title", "ea"): face(tea, a.title_ea) if a.title_ea else face(te, a.title),
             ("body", "latin"): face(be, a.body), ("body", "ea"): face(bea, a.body_ea) if a.body_ea else face(be, a.body)}
    # 굵게 설정
    title_bold = a.title_bold or bool(te and te["bold"])
    title_nobold = bool(te and not te["bold"])
    body_bold_face = None
    if be:
        cands = [m for m in man if m["font"] == be["font"]]
        bb = next((m for w in ("Bold", "SemiBold", "ExtraBold") for m in cands if m["weight"] == w), None)
        if bb and not bb["bold"]:
            body_bold_face = bb["family"]  # 굵은 글꼴 이름이 따로 있음 → b=1 대신 그 글꼴로
    # 역할별로 받은 굵기들(굵기 이름이 든 원래 글꼴을 옮길 때)
    weights = {}
    if man:
        for role, e in (("title", tea or te), ("body", bea or be)):
            if e:
                weights[role] = [m for m in man if m["font"] == e["font"]]
    cfg = {"fonts": fonts, "keep": {k.strip().lower() for k in a.keep.split(",") if k.strip()},
           "title_nobold": title_nobold, "body_bold_face": body_bold_face,
           "weights": weights, "base": {"title": tea or te, "body": bea or be}}

    try:
        zin = zipfile.ZipFile(a.src)
    except zipfile.BadZipFile:
        sys.exit(f"'{a.src}'을(를) 열 수 없습니다. 암호가 걸렸거나 손상된 파일입니다(PowerPoint에서 암호를 풀고 다시 저장해 주세요).")

    chars = deck_chars(zin)
    part_re = re.compile(r"^ppt/(presentation|theme/theme\d+|slides/slide\d+|slideLayouts/slideLayout\d+|slideMasters/slideMaster\d+|"
                         r"notesMasters/notesMaster\d+|notesSlides/notesSlide\d+|handoutMasters/handoutMaster\d+|"
                         r"charts/chart\d+|diagrams/drawing\d+|diagrams/data\d+)\.xml$")
    stats = {"theme": 0, "runs": 0, "parts": 0, "title_boxes": 0, "bold_title": 0, "bold_body": 0, "bold_body_fake": 0, "weight_kept": 0}
    fd, tmp = tempfile.mkstemp(suffix=".pptx", dir=os.path.dirname(os.path.abspath(a.dst)) or ".")
    os.close(fd)
    with zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if part_re.match(item.filename):
                root = etree.fromstring(data)
                fn = item.filename
                if fn.startswith("ppt/theme/"):
                    stats["theme"] += set_theme(root, fonts)
                    fix_runs(root, cfg, a.mode, stats)  # objectDefaults(새 상자 기본 글꼴)
                elif fn.startswith("ppt/charts/"):
                    fix_runs(root, cfg, "force", stats)  # 차트는 프로그램마다 테마 참조를 다르게 풀어 이름을 직접 쓴다
                elif fn.startswith("ppt/slides/"):
                    tl = title_like(root)
                    stats["title_boxes"] += len(tl)
                    fix_runs(root, cfg, a.mode, stats, tl)
                else:
                    fix_runs(root, cfg, a.mode, stats)
                if title_bold:
                    for el in root.iter(*RPR_TAGS):
                        if is_title(el, frozenset()):
                            el.set("b", "1")
                stats["parts"] += 1
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    os.chmod(tmp, 0o644)
    shutil.move(tmp, a.dst)

    print(f"완료: {a.dst}  (테마 {stats['theme']}곳, 글자 서식 {stats['runs']}곳, XML {stats['parts']}개, 자리표시자 아닌 제목 상자 {stats['title_boxes']}개)")
    print(f"  제목: {fonts[('title', 'latin')]}" + (" (굵게)" if title_bold else "")
          + (f" / 한글 {fonts[('title', 'ea')]}" if fonts[('title', 'ea')] != fonts[('title', 'latin')] else ""))
    print(f"  본문: {fonts[('body', 'latin')]}" + (f" / 한글 {fonts[('body', 'ea')]}" if fonts[('body', 'ea')] != fonts[('body', 'latin')] else ""))
    if stats["bold_title"]:
        print(f"  · 제목의 굵게 {stats['bold_title']}곳을 껐습니다('{fonts[('title', 'latin')]}'에는 굵게 스타일 파일이 따로 없어, 그대로 두면 가짜 굵게가 됨).")
    if stats["weight_kept"]:
        print(f"  · 굵기가 이름에 든 원래 글꼴 {stats['weight_kept']}곳은 새 가족의 가장 가까운 굵기로 바꿨습니다.")
    if a.theme_xml:
        write_theme_xml(a.theme_xml, fonts)
        print(f"  · 글꼴 테마 파일: {a.theme_xml} (PowerPoint 디자인 → 글꼴 → 글꼴 사용자 지정 목록에 나오도록 테마 폴더에 넣으면 새 슬라이드에도 적용)")
    if stats["bold_body"]:
        print(f"  · 본문 굵은 글자 {stats['bold_body']}곳을 '{body_bold_face}' 글꼴로 바꿨습니다.")
    if stats["bold_body_fake"] and be and not any(m["font"] == be["font"] and m["bold"] for m in man):
        print(f"  ! 본문에 굵은 글자 {stats['bold_body_fake']}곳이 있는데 본문 가족의 Bold를 받지 않았습니다 → 가짜 굵게로 보입니다. Bold도 받아 다시 적용하세요.")
    if man:
        used = [e for e in (te, be, tea, bea) if e]
        for e in used:
            miss = [c for c in (e.get("missing") or "") if c in chars]
            if miss:
                print(f"  ! '{e['family']}'에 덱에 쓰인 기호 {''.join(miss)} 가 없습니다 → 그 글자만 다른 글꼴로 보입니다. 기호를 바꾸거나 다른 폰트를 고려하세요.")
            if e.get("hangul", 11172) < 11172 and os.path.exists(e.get("file", "")):
                from fontTools.ttLib import TTFont
                cmap = TTFont(e["file"], lazy=True).getBestCmap() or {}
                gap = sorted({c for c in chars if 0xAC00 <= ord(c) <= 0xD7A3 and ord(c) not in cmap})
                if gap:
                    print(f"  ! '{e['family']}'는 한글 {e['hangul']}자만 있어 덱의 '{''.join(gap[:20])}'{' 등' if len(gap) > 20 else ''}이(가) 다른 글꼴로 보입니다.")




if __name__ == "__main__":
    sys.exit(main())
