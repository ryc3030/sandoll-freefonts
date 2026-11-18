#!/usr/bin/env python3
"""무료폰트 색인에서 재배포가 허용된 폰트를 내려받고, PowerPoint가 쓰는 실제 글꼴 이름을 알려 준다.

출처 두 가지
  - 산돌구름 무료폰트 중 OFL 폰트(fontfiles.json): 제작사·네이버·jsDelivr 등 공개 배포처
  - Google Fonts 한국어 폰트(google-fonts.json): Google Fonts가 굵기별로 만든 정적 TTF

사용 예
  python fetch_fonts.py --list
  python fetch_fonts.py 프리텐다드 --weights Regular,Bold,ExtraBold --out fonts --install
  python fetch_fonts.py 본명조 프리텐다드 --weights Regular,Bold --out fonts      # 제목·본문 가족을 한 번에
  python fetch_fonts.py "Gothic A1" 함렛 --out fonts                              # 같은 --out에 이어 받으면 manifest가 합쳐진다
  python fetch_fonts.py --local 서울남산체B.ttf 서울남산체M.ttf --out fonts       # 사용자가 공식 배포처에서 받아 준 파일(기관 전용 서체 등)

결과: <out>/폰트 파일, <out>/LICENSE-*.txt(저작권 표시 + OFL 1.1 원문), <out>/manifest.json
  manifest 항목 = {font, name, weight, file, family(PowerPoint 이름), bold, outline(TrueType/CFF), fsType, embeddable, missing(빠진 기호)}

라이선스 안전장치
  - 목록(fontfiles.json·google-fonts.json)에 있는 OFL·재배포 허용 폰트만 받는다. 그 밖의 폰트는 공식 배포처 링크만 출력한다.
  - 받은 파일의 name 테이블(13·14번)도 확인해, 재배포를 막는 문구가 있으면 지우고 건너뛴다.
"""
import argparse, json, re, shutil, subprocess, sys, urllib.request
from pathlib import Path

REPO_RAW = "https://raw.githubusercontent.com/ryc3030/sandoll-freefonts/main/data/"
UA = {"User-Agent": "sandoll-freefonts-skill/2.0"}
ORDER = ["Thin", "ExtraLight", "Light", "Regular", "Medium", "SemiBold", "Bold", "ExtraBold", "Black", "Heavy"]
GW = {"Thin": 100, "ExtraLight": 200, "Light": 300, "Regular": 400, "Medium": 500, "SemiBold": 600,
      "Bold": 700, "ExtraBold": 800, "Black": 900, "Heavy": 900}
# 한국 보고서·발표에 자주 나오는 기호. 폰트에 없으면 다른 글꼴로 대체돼 한 줄 안에서 글꼴이 섞인다.
SYMBOLS = "□■○●◦◎◆◇▶▷※ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ①②③④⑤⑥⑦⑧⑨⑩·–—→←↑↓₩㎡℃±×÷≤≥βαΔ「」『』〈〉《》“”‘’…"
BAD_LIC = re.compile(r"재배포\s*(?:를\s*)?금지|무단\s*배포|not\s+be\s+redistributed|may\s+not\s+redistribute", re.I)


BUNDLED = Path(__file__).resolve().parent.parent / "data"  # 설치 ZIP에 함께 든 데이터


def load(name, local_dir=None, text=False):
    """--data 폴더 → 저장소 최신(main) → 스킬에 함께 든 데이터 순으로 읽는다(인터넷이 막힌 환경 대비)."""
    if local_dir and (Path(local_dir) / name).exists():
        t = (Path(local_dir) / name).read_text("utf-8")
    else:
        try:
            with urllib.request.urlopen(urllib.request.Request(REPO_RAW + name, headers=UA), timeout=15) as r:
                t = r.read().decode("utf-8")
        except Exception:
            if not (BUNDLED / name).exists():
                raise
            t = (BUNDLED / name).read_text("utf-8")
    return t if text else json.loads(t)


def download(url, dest, ua=None):
    """임시 파일에 받은 뒤 이름을 바꾼다(중간에 끊겨도 깨진 파일이 남지 않게)."""
    if dest.exists() and dest.stat().st_size > 0:
        return
    tmp = dest.with_suffix(dest.suffix + ".part")
    with urllib.request.urlopen(urllib.request.Request(url, headers=ua or UA), timeout=180) as r, open(tmp, "wb") as fp:
        shutil.copyfileobj(r, fp)
    tmp.replace(dest)


def inspect(path):
    """PowerPoint 이름(nameID 1), 굵게 여부, 외곽선 종류, 포함 허용(fsType), 라이선스 표기, 빠진 기호."""
    from fontTools.ttLib import TTFont
    f = TTFont(path, lazy=True, fontNumber=0)
    n = f["name"]

    def g(i):
        rec = n.getName(i, 3, 1, 0x409) or n.getName(i, 1, 0, 0) or n.getName(i, 3, 1, 0x412)
        return str(rec) if rec else None
    fam, sub = g(1), g(2) or "Regular"
    fs = f["OS/2"].fsType if "OS/2" in f else 0
    cmap = f.getBestCmap() or {}
    lic = " ".join(filter(None, [g(13), g(14)]))
    return {
        "family": fam, "subfamily": sub, "typoFamily": g(16) or fam,
        "bold": sub.lower() in ("bold", "bold italic"),
        "outline": "CFF" if ("CFF " in f or "CFF2" in f) else "TrueType",
        "fsType": fs,
        # 0 = 설치 가능, 8 = 편집 가능 → 받는 사람도 편집 가능. 4 = 미리보기·인쇄만, 2 = 포함 금지
        "embeddable": "편집 가능" if (fs & 0x000E) in (0, 8) else ("보기·인쇄만" if fs & 4 else "포함 금지"),
        "license": lic[:200], "copyright": (g(0) or "")[:300],
        "missing": "".join(ch for ch in SYMBOLS if ord(ch) not in cmap),
        "hangul": sum(1 for c in range(0xAC00, 0xD7A4) if c in cmap),
    }


def license_ok(info):
    s = info["license"]
    if re.search(r"Open Font License|scripts\.sil\.org/OFL|openfontlicense|\bOFL\b|Apache License|Ubuntu Font Licen", s, re.I):
        return True, "파일에 OFL 등 재배포 허용 라이선스 표기"
    if s and BAD_LIC.search(s):
        return False, f"파일의 라이선스 표기가 재배포를 막음: {s[:80]}"
    return True, ("파일에 라이선스 표기가 없어 목록의 OFL 표기를 따름" if not s else f"파일 표기 확인 필요: {s[:80]}")


def nearest(have, want):
    """없는 굵기는 숫자로 가장 가까운 것. 동률이면 Regular 이상 요청은 무거운 쪽(Light 본문 방지)."""
    if want in have:
        return want
    if want not in GW or not have:
        return None
    t = GW[want]
    return min(have, key=lambda w: (abs(GW.get(w, 400) - t), -GW.get(w, 400) if t >= 400 else GW.get(w, 400)))


def google_files(g, weights):
    """Google Fonts CSS API에 TTF를 요청해 굵기별 정적 파일 주소를 얻는다(브라우저가 아닌 UA면 TTF를 준다)."""
    have = dict(zip(g["weights"], g["weightValues"]))
    vals = sorted({have[w] for w in weights if w in have})
    css_url = f"https://fonts.googleapis.com/css2?family={g['family'].replace(' ', '+')}:wght@{';'.join(map(str, vals))}"
    with urllib.request.urlopen(urllib.request.Request(css_url, headers={"User-Agent": "Wget/1.21"}), timeout=60) as r:
        css = r.read().decode("utf-8")
    urls = {int(m.group(1)): m.group(2) for m in re.finditer(r"font-weight:\s*(\d+);.*?url\((https://[^)]+\.ttf)\)", css, re.S)}
    return {w: urls.get(have.get(w)) for w in weights}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("fonts", nargs="*", help="산돌구름 ID, 한글 이름 또는 family 이름")
    ap.add_argument("--weights", default="Regular,Bold,ExtraBold", help="쉼표로 구분. all = 전부")
    ap.add_argument("--out", default="fonts")
    ap.add_argument("--install", action="store_true", help="~/.fonts에 설치하고 fc-cache 실행(LibreOffice 렌더링 확인용)")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--data", help="data 폴더 경로(오프라인일 때)")
    ap.add_argument("--local", nargs="+", help="사용자가 직접 준 폰트 파일(.ttf/.otf). 라이선스 확인은 사용자 책임으로, 결과물 묶음에는 넣지 않는다")
    a = ap.parse_args()

    if a.local:
        out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
        mp = out / "manifest.json"
        old = json.loads(mp.read_text("utf-8")) if mp.exists() else []
        new = []
        for fp in a.local:
            src = Path(fp)
            if not src.exists():
                print(f"! {fp}: 파일이 없습니다", file=sys.stderr); continue
            dest = out / "local" / src.name
            dest.parent.mkdir(exist_ok=True)
            shutil.copy(src, dest)
            info = inspect(dest)
            w = info.get("subfamily") or "Regular"
            # 굵기 이름 추정: 파일 이름이나 서브패밀리에서
            for cand in ORDER[::-1]:
                if cand.lower() in (src.stem + " " + info["family"] + " " + w).lower().replace(" ", ""):
                    w = cand; break
            else:
                w = "Bold" if info["bold"] else "Regular"
            fam = info.get("typoFamily") or info["family"]
            new.append({"font": "L:" + fam, "name": fam, "weight": w, "file": str(dest), **info, "latinOnly": False,
                        "licenseCheck": "사용자 제공 파일(재배포하지 말 것)", "local": True})
            print(f"  ✓ {src.name} → '{fam}:{w}' (PowerPoint 이름 '{info['family']}'{', 굵게' if info['bold'] else ''})"
                  + (f"  빠진 기호 {info['missing']}" if info["missing"] else ""))
        merged = {(m["font"], m["weight"]): m for m in old + new}
        mp.write_text(json.dumps(list(merged.values()), ensure_ascii=False, indent=2), "utf-8")
        if a.install and new:
            fd = Path.home() / ".fonts"; fd.mkdir(exist_ok=True)
            for m in new:
                shutil.copy(m["file"], fd)
            subprocess.run(["fc-cache", "-f"], check=False, capture_output=True)
        print(f"manifest: {mp} (이번 {len(new)}개, 누적 {len(merged)}개). apply_fonts.py에는 '이름:굵기'로 지정")
        return

    files = load("fontfiles.json", a.data)["fonts"]
    try:
        catalog = {str(f["id"]): f for f in load("fonts.json", a.data)["fonts"]}
    except Exception:
        catalog = {}
    try:
        google = {g["family"]: g for g in load("google-fonts.json", a.data)["fonts"] if g.get("redistributable")}
    except Exception:
        google = {}
    try:
        ofl_text = load("OFL-1.1.txt", a.data, text=True)
    except Exception:
        ofl_text = "SIL Open Font License, Version 1.1 — https://openfontlicense.org/open-font-license-official-text/\n"
    sandoll_fams = {f["family"].replace(" ", "").lower() for f in files.values()} | {
        (f.get("google") or "").replace(" ", "").lower() for f in files.values()}

    if a.list or not a.fonts:
        print("[산돌구름 OFL]")
        for k, f in files.items():
            print(f"{k:>6}  {f['name']:<14} {f['family']:<18} {f['style']:<8} {'/'.join(f['use']):<11} {','.join(f['files'])}  — {f['tone']}")
        print("\n[Google Fonts 한국어]")
        for g in google.values():
            if g["family"].replace(" ", "").lower() not in sandoll_fams:
                print(f"{'G':>6}  {g['name']:<14} {g['family']:<18} {g['style']:<8} {'/'.join(g['use']):<11} {','.join(g['weights'])}  — {', '.join(g['tags'])}")
        return

    def norm(s):
        return (s or "").lower().replace(" ", "")

    def find(q):
        ql = norm(q)
        for k, f in files.items():
            if ql in (k, norm(f["name"]), norm(f["family"]), norm(f.get("google"))):
                return k
        for fam, g in google.items():
            if ql in (norm(fam), norm(g["name"]), norm(g["name"].split(" (")[0])):
                return "G:" + fam
        for k, c in catalog.items():
            if ql in (k, norm(c.get("name")), norm(c.get("nameEn"))):
                return k
        return None

    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    mp = out / "manifest.json"
    old = json.loads(mp.read_text("utf-8")) if mp.exists() else []
    new, failed = [], []

    for q in a.fonts:
        k = find(q)
        if k is None:
            print(f"! '{q}'을(를) 찾지 못했습니다. --list로 목록을 보세요.", file=sys.stderr); failed.append(q); continue
        if k.startswith("G:"):
            gfam = google[k[2:]]
            name, fam_label, srcs, origin = gfam["name"], gfam["family"], {w: "google" for w in gfam["weights"]}, gfam["specimenUrl"]
        elif k in files:
            f = files[k]
            name, fam_label, srcs = f["name"], f["family"], dict(f["files"])
            gfam = google.get(f.get("google") or f["family"])
            if gfam:  # 산돌 목록에 없는 굵기는 같은 가족의 Google 판에서 보충
                for w in gfam["weights"]:
                    srcs.setdefault(w, "google")
            origin = catalog.get(k, {}).get("makerUrl") or catalog.get(k, {}).get("sandollUrl", "")
        else:
            c = catalog.get(k, {})
            print(f"! {c.get('name', q)}: 자동으로 받을 수 없는 폰트입니다(OFL 아님 또는 파일 경로 미등록). "
                  f"공식 배포처: {c.get('makerUrl') or c.get('sandollUrl')}", file=sys.stderr)
            failed.append(q); continue

        want = [w for w in ORDER if w in srcs] if a.weights == "all" else [w.strip() for w in a.weights.split(",") if w.strip()]
        plan = {}
        for w in want:
            got = nearest(list(srcs), w)
            if not got:
                print(f"  - {name} {w}: 없음", file=sys.stderr); continue
            if got != w and got not in plan:
                print(f"  - {name} {w}: 없어 {got}로 대체", file=sys.stderr)
            plan.setdefault(got, srcs[got])
        g_need = [w for w, u in plan.items() if u == "google"]
        if g_need and gfam:
            try:
                plan.update(google_files(gfam, g_need))
            except Exception as e:
                print(f"  ! {name}: Google Fonts 주소를 얻지 못함({e})", file=sys.stderr)
        copyrights = []
        for w, url in plan.items():
            if not url or url == "google":
                print(f"  ! {name} {w}: 파일 주소 없음", file=sys.stderr); failed.append(f"{name}:{w}"); continue
            ext = ".otf" if url.lower().endswith(".otf") else ".ttf"
            dest = out / f"{fam_label.replace(' ', '')}-{w}{ext}"
            try:
                download(url, dest)
                info = inspect(dest)
            except Exception as e:
                print(f"  ! {name} {w}: 내려받기 실패({e})", file=sys.stderr); failed.append(f"{name}:{w}")
                dest.unlink(missing_ok=True); continue
            ok, why = license_ok(info)
            if not ok:
                print(f"  ! {name} {w}: {why} → 지우고 건너뜀", file=sys.stderr); dest.unlink(missing_ok=True)
                failed.append(f"{name}:{w}"); continue
            if info["copyright"] and info["copyright"] not in copyrights:
                copyrights.append(info["copyright"])
            new.append({"font": k, "name": name, "weight": w, "file": str(dest), **info,
                        "latinOnly": files.get(k, {}).get("latinOnly", False), "licenseCheck": why})
            note = []
            if info["bold"]: note.append("굵게")
            if info["outline"] == "CFF": note.append("CFF: PowerPoint 글꼴 포함 불가")
            if info["embeddable"] != "편집 가능": note.append(f"포함 시 {info['embeddable']}")
            if info["missing"]: note.append(f"빠진 기호 {info['missing']}")
            print(f"  ✓ {name} {w:<10} → PowerPoint 이름 '{info['family']}'" + (f"  ({'; '.join(note)})" if note else ""))
        if copyrights:
            (out / f"LICENSE-{fam_label.replace(' ', '')}.txt").write_text(
                f"{name} ({fam_label})\n출처: {origin}\n\n" + "\n".join(copyrights) + "\n\n" + ofl_text, "utf-8")

    merged = {(m["font"], m["weight"]): m for m in old + new}
    mp.write_text(json.dumps(list(merged.values()), ensure_ascii=False, indent=2), "utf-8")
    if a.install and new:
        fd = Path.home() / ".fonts"; fd.mkdir(exist_ok=True)
        for m in new:
            shutil.copy(m["file"], fd)
        subprocess.run(["fc-cache", "-f"], check=False, capture_output=True)
        print(f"설치: {fd}")
    print(f"manifest: {mp} (이번 {len(new)}개, 누적 {len(merged)}개)")
    if failed:
        print("받지 못한 것: " + ", ".join(failed), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
