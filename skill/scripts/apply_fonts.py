#!/usr/bin/env python3
"""PPTX 전체의 글꼴을 제목용·본문용 두 가지로 통일한다.

  python apply_fonts.py in.pptx out.pptx --title "Pretendard ExtraBold" --body "Pretendard"
  python apply_fonts.py in.pptx out.pptx --title "Instrument Serif" --title-ea "MaruBuri" --body "MaruBuri"

하는 일
  1) 테마(ppt/theme/*.xml)의 제목 글꼴(majorFont)·본문 글꼴(minorFont)을 latin·ea(한글)·Hang 항목까지 바꾼다.
  2) 슬라이드·레이아웃·마스터·차트·SmartArt에 직접 박힌 글꼴 이름을 테마 참조(+mj-*/+mn-*)로 바꿔
     모든 글자가 1)을 따르게 한다. 제목 자리표시자(title, ctrTitle) 안 글자는 제목 글꼴, 나머지는 본문 글꼴.
  3) 기호 글꼴(Wingdings 등)과 --keep 으로 지정한 글꼴(예: 코드용 고정폭)은 건드리지 않는다.

--mode force 를 주면 테마 참조 대신 글꼴 이름을 모든 글자에 직접 써 넣는다
(테마를 무시하는 프로그램이나 Google Slides로 옮길 때).
"""
import argparse, re, shutil, sys, tempfile, zipfile
from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}
SYMBOL = {"wingdings", "wingdings 2", "wingdings 3", "symbol", "webdings", "marlett", "segoe ui symbol", "segoe ui emoji"}
RPR_TAGS = {f"{{{A}}}{t}" for t in ("rPr", "defRPr", "endParaRPr")}
TITLE_PH = {"title", "ctrTitle"}


def set_theme(root, title, title_ea, body, body_ea):
    n = 0
    for tag, lat, ea in (("majorFont", title, title_ea), ("minorFont", body, body_ea)):
        for fs in root.iter(f"{{{A}}}{tag}"):
            for child, val in (("latin", lat), ("ea", ea)):
                el = fs.find(f"a:{child}", NS)
                if el is None:
                    el = etree.SubElement(fs, f"{{{A}}}{child}")
                el.set("typeface", val); n += 1
            for f in fs.findall("a:font", NS):
                if f.get("script") == "Hang":
                    f.set("typeface", ea); n += 1
            if not any(f.get("script") == "Hang" for f in fs.findall("a:font", NS)):
                h = etree.SubElement(fs, f"{{{A}}}font"); h.set("script", "Hang"); h.set("typeface", ea); n += 1
    return n


def is_title(el):
    """el이 제목 자리표시자나 마스터 titleStyle 안에 있는가."""
    for anc in el.iterancestors():
        if anc.tag == f"{{{P}}}titleStyle":
            return True
        if anc.tag in (f"{{{P}}}bodyStyle", f"{{{P}}}otherStyle"):
            return False
        if anc.tag == f"{{{P}}}sp":
            ph = anc.find("p:nvSpPr/p:nvPr/p:ph", NS)
            return ph is not None and ph.get("type") in TITLE_PH
    return False


def fix_runs(root, keep, mode, fonts, title_bold=False):
    n = 0
    for rpr in root.iter(*RPR_TAGS):
        title = is_title(rpr)
        if title and title_bold:
            rpr.set("b", "1")
        for child in ("latin", "ea", "cs"):
            el = rpr.find(f"a:{child}", NS)
            if child == "cs":
                continue  # 복합 문자(아랍어 등) 글꼴은 그대로 둔다
            if mode == "force":
                if el is None:
                    el = etree.SubElement(rpr, f"{{{A}}}{child}")
                    # 스키마 순서: ln, fill..., effect..., highlight, uLnTx.., latin, ea, cs, sym, hlinkClick...
                    _reorder(rpr)
                cur = (el.get("typeface") or "").lower()
                if cur in SYMBOL or cur in keep:
                    continue
                el.set("typeface", fonts[("title" if title else "body", child)]); n += 1
            else:
                if el is None:
                    continue
                cur = el.get("typeface") or ""
                if cur.startswith("+") or cur.lower() in SYMBOL or cur.lower() in keep:
                    continue
                el.set("typeface", f"+{'mj' if title else 'mn'}-{'lt' if child == 'latin' else 'ea'}"); n += 1
    return n


ORDER = ["ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "effectLst", "effectDag",
         "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]


def _reorder(rpr):
    kids = list(rpr)
    kids.sort(key=lambda e: ORDER.index(etree.QName(e).localname) if etree.QName(e).localname in ORDER else 99)
    for k in kids:
        rpr.remove(k)
    for k in kids:
        rpr.append(k)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src"); ap.add_argument("dst")
    ap.add_argument("--title", required=True, help="제목 글꼴(PowerPoint 이름)")
    ap.add_argument("--body", required=True, help="본문 글꼴(PowerPoint 이름)")
    ap.add_argument("--title-ea", help="제목의 한글 글꼴(제목 글꼴이 영문 전용일 때)")
    ap.add_argument("--body-ea", help="본문의 한글 글꼴(본문 글꼴이 영문 전용일 때)")
    ap.add_argument("--keep", default="", help="건드리지 않을 글꼴, 쉼표 구분 (예: Consolas,D2Coding)")
    ap.add_argument("--mode", choices=["theme", "force"], default="theme")
    ap.add_argument("--title-bold", action="store_true",
                    help="제목을 굵게 스타일로(제목 글꼴의 PowerPoint 이름이 Regular와 같고 '(굵게)'로 나온 경우)")
    a = ap.parse_args()
    title_ea, body_ea = a.title_ea or a.title, a.body_ea or a.body
    keep = {k.strip().lower() for k in a.keep.split(",") if k.strip()}
    fonts = {("title", "latin"): a.title, ("title", "ea"): title_ea, ("body", "latin"): a.body, ("body", "ea"): body_ea}

    part_re = re.compile(r"^ppt/(theme/theme\d+|slides/slide\d+|slideLayouts/slideLayout\d+|slideMasters/slideMaster\d+|"
                         r"notesMasters/notesMaster\d+|notesSlides/notesSlide\d+|handoutMasters/handoutMaster\d+|"
                         r"charts/chart\d+|diagrams/drawing\d+|diagrams/data\d+)\.xml$")
    stats = {"theme": 0, "runs": 0, "parts": 0}
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx").name
    with zipfile.ZipFile(a.src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if part_re.match(item.filename):
                root = etree.fromstring(data)
                if item.filename.startswith("ppt/theme/"):
                    stats["theme"] += set_theme(root, a.title, title_ea, a.body, body_ea)
                elif item.filename.startswith("ppt/charts/"):
                    # 차트는 프로그램마다 테마 참조를 다르게 풀어서 글꼴 이름을 직접 쓴다.
                    stats["runs"] += fix_runs(root, keep, "force", fonts)
                else:
                    stats["runs"] += fix_runs(root, keep, a.mode, fonts, a.title_bold)
                stats["parts"] += 1
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    shutil.move(tmp, a.dst)
    print(f"완료: {a.dst}  (테마 항목 {stats['theme']}곳, 글자 서식 {stats['runs']}곳, XML {stats['parts']}개)")
    print(f"  제목: {a.title}" + (f" / 한글 {title_ea}" if title_ea != a.title else ""))
    print(f"  본문: {a.body}" + (f" / 한글 {body_ea}" if body_ea != a.body else ""))


if __name__ == "__main__":
    sys.exit(main())
