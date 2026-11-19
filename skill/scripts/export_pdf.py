#!/usr/bin/env python3
"""폰트를 바꾼 PPTX를 PDF로 내보낸다. 받은 폰트만 보이는 빈 환경에서 변환해, PDF에 그 폰트가 실제로 들어갔는지 확인한다.

  python export_pdf.py output.pptx --fonts fonts                    # → output.pdf
  python export_pdf.py output.pptx --fonts fonts --png after        # 확인용 이미지(after-1.png …)도
  python export_pdf.py output.pptx --fonts fonts --manifest fonts/manifest.json   # 들어가야 할 폰트를 대조

하는 일
  1) 빈 HOME에 --fonts 폴더의 폰트만 넣고 LibreOffice로 변환한다(이미 설치된 폰트가 섞여 거짓 통과가 나지 않게).
  2) LibreOffice가 한글과 숫자·영문 사이를 벌려 놓는 설정('제72회' → '제 72 회')을 문단마다 끄고 변환한다.
     PowerPoint에는 없는 간격이라 그대로 두면 PDF만 어색해진다(파이썬 uno가 없으면 이 단계 없이 변환하고 알린다).
  3) 접근성 태그가 붙은 PDF(Tagged PDF)로 저장한다.
  4) pdffonts로 PDF에 들어간 글꼴을 읽어, 받은 폰트가 아닌 글꼴로 대체된 글자가 있으면 경고한다.

PDF는 LibreOffice가 그린 것이라 줄바꿈이 PowerPoint와 조금 다를 수 있다. 제출용 원본 PDF가 꼭 PowerPoint 모양이어야 하면
PowerPoint에서 '파일 → 다른 이름으로 저장 → PDF'로 다시 만들도록 안내한다.
"""
import argparse, json, os, re, shutil, socket, subprocess, sys, tempfile, time
from pathlib import Path

UNO_SCRIPT = r'''
import sys, time, uno
from com.sun.star.beans import PropertyValue
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
src, dst, port = sys.argv[1], sys.argv[2], sys.argv[3]
local = uno.getComponentContext()
res = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
ctx = None
for _ in range(120):
    try:
        ctx = res.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"); break
    except Exception:
        time.sleep(0.5)
if ctx is None:
    sys.exit("LibreOffice에 연결하지 못했습니다")
desk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(src), "_blank", 0, (pv("Hidden", True),))
n = 0
def fix_text(t):
    global n
    try:
        e = t.getText().createEnumeration()
    except Exception:
        return
    while e.hasMoreElements():
        para = e.nextElement()
        try:
            para.ParaIsCharacterDistance = False; n += 1
        except Exception:
            pass
def walk(sh):
    if sh.supportsService("com.sun.star.drawing.GroupShape"):
        for i in range(sh.getCount()):
            walk(sh.getByIndex(i))
    elif sh.supportsService("com.sun.star.drawing.TableShape"):
        m = sh.Model
        for r in range(m.Rows.Count):
            for c in range(m.Columns.Count):
                fix_text(m.getCellByPosition(c, r))
    elif hasattr(sh, "getText"):
        fix_text(sh)
for pages in (doc.getDrawPages(), doc.getMasterPages()):
    for i in range(pages.getCount()):
        pg = pages.getByIndex(i)
        for j in range(pg.getCount()):
            walk(pg.getByIndex(j))
fd = uno.Any("[]com.sun.star.beans.PropertyValue", (pv("UseTaggedPDF", True), pv("ExportNotesPages", False)))
args = (pv("FilterName", "impress_pdf_Export"), PropertyValue("FilterData", 0, fd, 0))
uno.invoke(doc, "storeToURL", (uno.systemPathToFileUrl(dst), uno.Any("[]com.sun.star.beans.PropertyValue", args)))
doc.close(True)
try:
    desk.terminate()
except Exception:
    pass
print(f"PARAS {n}")
'''


def free_port():
    s = socket.socket(); s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p


def has_uno():
    return subprocess.run([sys.executable, "-c", "import uno"], capture_output=True).returncode == 0


def ps_names(font_files):
    """폰트 파일의 PostScript 이름(PDF에 적히는 이름)."""
    try:
        from fontTools.ttLib import TTFont
    except ImportError:
        return set()
    out = set()
    for f in font_files:
        try:
            t = TTFont(f, lazy=True, fontNumber=0)
            n = t["name"].getDebugName(6)
            if n:
                out.add(n.replace(" ", ""))
        except Exception:
            pass
    return out


def main():
    ap = argparse.ArgumentParser(description="폰트를 바꾼 PPTX를 PDF로 내보내고 폰트가 들어갔는지 확인")
    ap.add_argument("pptx")
    ap.add_argument("--fonts", action="append", default=[], help="폰트 폴더(여러 번 줄 수 있음, 하위 폴더 포함)")
    ap.add_argument("--out", help="PDF 경로(기본: pptx와 같은 이름)")
    ap.add_argument("--png", help="확인용 이미지 접두어(예: after → after-1.png …)")
    ap.add_argument("--dpi", type=int, default=60)
    ap.add_argument("--manifest", help="fetch_fonts.py의 manifest.json(들어가야 할 폰트 대조)")
    ap.add_argument("--timeout", type=int, default=240)
    a = ap.parse_args()

    src = Path(a.pptx).resolve()
    out = Path(a.out).resolve() if a.out else src.with_suffix(".pdf")
    if not shutil.which("soffice"):
        sys.exit("LibreOffice(soffice)가 없어 PDF를 만들 수 없습니다. PowerPoint에서 '다른 이름으로 저장 → PDF'를 안내하세요.")
    font_files = []
    for d in a.fonts:
        font_files += [p for p in Path(d).rglob("*") if p.suffix.lower() in (".ttf", ".otf", ".ttc")]

    with tempfile.TemporaryDirectory() as home:
        fdir = Path(home) / ".fonts"; fdir.mkdir()
        for f in font_files:
            shutil.copy(f, fdir / f.name)
        env = dict(os.environ, HOME=home)
        subprocess.run(["fc-cache", "-f"], env=env, capture_output=True)
        profile = Path(home) / "lo-profile"
        spaced = False
        if has_uno():
            port = free_port()
            office = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", "--nologo",
                                       f"-env:UserInstallation={profile.as_uri()}",
                                       f"--accept=socket,host=127.0.0.1,port={port};urp;"],
                                      env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            script = Path(home) / "conv.py"; script.write_text(UNO_SCRIPT, "utf-8")
            try:
                r = subprocess.run([sys.executable, str(script), str(src), str(out), str(port)], env=env,
                                   capture_output=True, text=True, timeout=a.timeout)
                spaced = r.returncode == 0 and out.exists()
                if not spaced:
                    print("  ! uno 변환 실패, 기본 변환으로 다시 합니다: " + (r.stderr or r.stdout).strip()[-300:])
            except subprocess.TimeoutExpired:
                print("  ! uno 변환 시간 초과, 기본 변환으로 다시 합니다.")
            finally:
                try:
                    office.wait(timeout=20)
                except subprocess.TimeoutExpired:
                    office.kill()
        if not spaced:
            tmpout = Path(home) / "out"; tmpout.mkdir(exist_ok=True)
            subprocess.run(["soffice", "--headless", f"-env:UserInstallation={profile.as_uri()}", "--convert-to", "pdf",
                            "--outdir", str(tmpout), str(src)], env=env, capture_output=True, timeout=a.timeout)
            made = tmpout / (src.stem + ".pdf")
            if not made.exists():
                sys.exit("PDF를 만들지 못했습니다. 파일이 손상됐거나 LibreOffice가 열 수 없는 형식입니다.")
            shutil.move(str(made), out)
            print("  ! 파이썬 uno가 없어 기본 변환으로 만들었습니다. 한글과 숫자 사이가 조금 벌어져 보일 수 있습니다(PowerPoint에서는 생기지 않음).")

    print(f"PDF: {out}  ({out.stat().st_size / 1e6:.1f}MB)")
    # 들어간 글꼴 확인
    if shutil.which("pdffonts"):
        lines = subprocess.run(["pdffonts", str(out)], capture_output=True, text=True).stdout.splitlines()[2:]
        used = sorted({re.sub(r"^[A-Z]{6}\+", "", l.split()[0]) for l in lines if l.strip()})
        expect = ps_names(font_files)
        if a.manifest:
            man = json.load(open(a.manifest, encoding="utf-8"))
            expect |= ps_names([m["file"] for m in man if os.path.exists(m.get("file", ""))])
        ok = [u for u in used if u.replace(" ", "") in expect]
        other = [u for u in used if u not in ok]
        print("  들어간 폰트: " + (", ".join(ok) or "없음"))
        bullets = [u for u in other if re.match(r"(Liberation|DejaVu|OpenSymbol)", u)]
        rest = [u for u in other if u not in bullets]
        if bullets:
            print("  · " + ", ".join(bullets) + ": 글머리표(•)·기호용 기본 글꼴이면 정상입니다.")
        if rest:
            print("  ! 받은 폰트가 아닌 글꼴로 그려진 글자가 있습니다: " + ", ".join(rest)
                  + " → 빠진 기호·드문 한글이거나 바꾸지 않은 글꼴(--keep, 기호·코드 글꼴)입니다. 이미지로 확인하세요.")
        if expect and not ok:
            print("  ! 받은 폰트가 PDF에 하나도 들어가지 않았습니다. 글꼴 이름이나 --fonts 폴더를 확인하세요.")
    if a.png and shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-r", str(a.dpi), "-png", str(out), a.png], check=True)
        print(f"  확인용 이미지: {a.png}-*.png")


if __name__ == "__main__":
    main()
