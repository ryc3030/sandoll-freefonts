#!/usr/bin/env python3
"""무료폰트 색인에서 폰트 파일을 내려받고, PowerPoint가 쓰는 실제 글꼴 이름을 알려 준다.

출처 두 가지
  - 산돌구름 무료폰트 중 OFL 폰트(fontfiles.json): 제작사·네이버·jsDelivr 등 공개 배포처에서 받는다.
  - Google Fonts 한국어 폰트(google-fonts.json): Google Fonts가 굵기별로 만든 정적 TTF를 받는다.

사용 예
  python fetch_fonts.py --list                      # 받을 수 있는 폰트 목록
  python fetch_fonts.py 프리텐다드 --weights Regular,Bold,ExtraBold --out fonts
  python fetch_fonts.py 16951 264 --out fonts --install
  python fetch_fonts.py "Gothic A1" 함렛 --weights Regular,Black --out fonts

결과: <out>/<파일들> 과 <out>/manifest.json
  manifest의 각 항목 = {font, weight, file, family(PowerPoint에 넣을 이름), bold(굵게 스타일 여부)}

OFL이 아닌 폰트는 받지 않는다(재배포 금지 가능성). 그런 폰트는 공식 배포처 링크만 출력한다.
"""
import argparse, json, os, shutil, subprocess, sys, urllib.request
from pathlib import Path

REPO_RAW = "https://raw.githubusercontent.com/ryc3030/sandoll-freefonts/main/data/"
UA = {"User-Agent": "sandoll-freefonts-skill/1.0"}


def load(name, local_dir=None):
    if local_dir and (Path(local_dir) / name).exists():
        return json.loads((Path(local_dir) / name).read_text("utf-8"))
    with urllib.request.urlopen(urllib.request.Request(REPO_RAW + name, headers=UA), timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def font_names(path):
    """PowerPoint(Windows)는 nameID 1(legacy family)을 글꼴 이름으로 쓴다. Regular/Bold/Italic만 같은 이름을 공유한다."""
    from fontTools.ttLib import TTFont
    f = TTFont(path, lazy=True, fontNumber=0)
    n = f["name"]
    def g(i):
        rec = n.getName(i, 3, 1, 0x409) or n.getName(i, 1, 0, 0) or n.getName(i, 3, 1, 0x412)
        return str(rec) if rec else None
    fam, sub = g(1), g(2) or "Regular"
    return {"family": fam, "subfamily": sub, "typoFamily": g(16) or fam, "bold": sub.lower() in ("bold", "bold italic")}


GW = {"Thin": 100, "ExtraLight": 200, "Light": 300, "Regular": 400, "Medium": 500,
      "SemiBold": 600, "Bold": 700, "ExtraBold": 800, "Black": 900, "Heavy": 900}


def fetch_google(g, weights, out):
    """Google Fonts CSS API에 TTF를 요청해 굵기별 정적 파일을 받는다(브라우저가 아닌 UA면 TTF를 준다)."""
    import re
    have = dict(zip(g["weights"], g["weightValues"]))
    want = list(have) if weights == "all" else [w.strip() for w in weights.split(",")]
    vals = []
    for w in want:
        if w in have:
            vals.append((w, have[w]))
        elif w in GW and have:
            alt = min(have, key=lambda x: abs(have[x] - GW[w]))
            print(f"  - {g['name']} {w}: 없어 {alt}로 대체", file=sys.stderr)
            if (alt, have[alt]) not in vals:
                vals.append((alt, have[alt]))
    if not vals:
        return []
    fam = g["family"].replace(" ", "+")
    css_url = f"https://fonts.googleapis.com/css2?family={fam}:wght@{';'.join(str(v) for _, v in sorted(vals, key=lambda x: x[1]))}"
    with urllib.request.urlopen(urllib.request.Request(css_url, headers={"User-Agent": "Wget/1.21"}), timeout=60) as r:
        css = r.read().decode("utf-8")
    urls = {int(m.group(1)): m.group(2) for m in re.finditer(r"font-weight:\s*(\d+);.*?url\((https://[^)]+\.ttf)\)", css, re.S)}
    res = []
    for w, v in vals:
        url = urls.get(v)
        if not url:
            print(f"  - {g['name']} {w}: 파일 주소를 찾지 못함", file=sys.stderr); continue
        dest = out / f"{g['family'].replace(' ', '')}-{w}.ttf"
        if not dest.exists():
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=180) as r, open(dest, "wb") as fp:
                shutil.copyfileobj(r, fp)
        info = font_names(dest)
        res.append({"font": "G:" + g["family"], "name": g["name"], "weight": w, "file": str(dest), **info, "latinOnly": False})
        print(f"  ✓ {g['name']} {w:<10} → PowerPoint 이름 '{info['family']}'{' (굵게)' if info['bold'] else ''}")
    (out / f"LICENSE-{g['family'].replace(' ', '')}.txt").write_text(
        f"{g['name']} ({g['family']})\n{g['license']} — Google Fonts 배포\n{g['specimenUrl']}\n"
        f"라이선스 원문: https://openfontlicense.org/open-font-license-official-text/\n", "utf-8")
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("fonts", nargs="*", help="산돌구름 ID, 한글 이름 또는 family 이름")
    ap.add_argument("--weights", default="Regular,Bold", help="쉼표로 구분. all = 전부")
    ap.add_argument("--out", default="fonts")
    ap.add_argument("--install", action="store_true", help="~/.fonts에 설치하고 fc-cache 실행(LibreOffice 렌더링 확인용)")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--data", help="data 폴더 경로(오프라인일 때)")
    a = ap.parse_args()

    files = load("fontfiles.json", a.data)["fonts"]
    try:
        catalog = {str(f["id"]): f for f in load("fonts.json", a.data)["fonts"]}
    except Exception:
        catalog = {}
    try:
        google = {g["family"]: g for g in load("google-fonts.json", a.data)["fonts"] if g.get("redistributable")}
    except Exception:
        google = {}
    # 산돌구름 목록에 이미 있는 Google 폰트는 산돌 쪽 항목으로 받는다.
    sandoll_families = {f["family"].replace(" ", "").lower() for f in files.values()}

    if a.list or not a.fonts:
        print("[산돌구름 OFL]")
        for k, f in files.items():
            print(f"{k:>6}  {f['name']:<14} {f['family']:<18} {f['style']:<8} {'/'.join(f['use']):<11} {','.join(f['files'])}  — {f['tone']}")
        print("\n[Google Fonts 한국어]")
        for g in google.values():
            if g["family"].replace(" ", "").lower() in sandoll_families:
                continue
            print(f"{'G':>6}  {g['name']:<14} {g['family']:<18} {g['style']:<8} {'/'.join(g['use']):<11} {','.join(g['weights'])}  — {', '.join(g['tags'])}")
        return

    def find(q):
        ql = q.lower().replace(" ", "")
        for k, f in files.items():
            if ql in (k, f["name"].lower().replace(" ", ""), f["family"].lower().replace(" ", "")):
                return k
        for fam, g in google.items():
            names = {fam.lower().replace(" ", ""), g["name"].lower().replace(" ", ""),
                     g["name"].split(" (")[0].lower().replace(" ", "")}
            if ql in names:
                return "G:" + fam
        for k, c in catalog.items():
            if ql in (k, c.get("name", "").lower().replace(" ", ""), (c.get("nameEn") or "").lower().replace(" ", "")):
                return k
        return None

    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    manifest, skipped = [], []
    for q in a.fonts:
        k = find(q)
        if k is None:
            print(f"! '{q}'을(를) 찾지 못했습니다. --list로 목록을 보세요.", file=sys.stderr); continue
        if k.startswith("G:"):
            manifest += fetch_google(google[k[2:]], a.weights, out)
            continue
        if k not in files:
            c = catalog.get(k, {})
            skipped.append(k)
            print(f"! {c.get('name', q)}: 자동으로 받을 수 없는 폰트입니다(OFL 아님 또는 파일 경로 미등록). "
                  f"공식 배포처: {c.get('makerUrl') or c.get('sandollUrl')}", file=sys.stderr)
            continue
        f = files[k]
        want = list(f["files"]) if a.weights == "all" else [w.strip() for w in a.weights.split(",")]
        for w in want:
            url = f["files"].get(w)
            if not url:
                # 가장 가까운 굵기로 대체
                order = ["Light", "Regular", "Medium", "SemiBold", "Bold", "ExtraBold", "Black", "Heavy"]
                have = [x for x in order if x in f["files"]]
                alt = min(have, key=lambda x: abs(order.index(x) - order.index(w))) if w in order and have else None
                if not alt or any(m["font"] == k and m["weight"] == alt for m in manifest):
                    print(f"  - {f['name']} {w}: 없음", file=sys.stderr); continue
                print(f"  - {f['name']} {w}: 없어 {alt}로 대체", file=sys.stderr); w, url = alt, f["files"][alt]
            dest = out / url.rsplit("/", 1)[1]
            if not dest.exists():
                with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=120) as r, open(dest, "wb") as fp:
                    shutil.copyfileobj(r, fp)
            info = font_names(dest)
            manifest.append({"font": k, "name": f["name"], "weight": w, "file": str(dest), **info,
                             "latinOnly": f.get("latinOnly", False)})
            print(f"  ✓ {f['name']} {w:<10} → PowerPoint 이름 '{info['family']}'{' (굵게)' if info['bold'] else ''}")
        lic = catalog.get(k, {})
        (out / f"LICENSE-{f['family'].replace(' ', '')}.txt").write_text(
            f"{f['name']} ({f['family']})\nSIL Open Font License 1.1\n공식 배포처: {lic.get('makerUrl', '')}\n산돌구름: {lic.get('sandollUrl', '')}\n\n{lic.get('licenseText', '')}\n", "utf-8")

    (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), "utf-8")
    if a.install and manifest:
        fd = Path.home() / ".fonts"; fd.mkdir(exist_ok=True)
        for m in manifest:
            shutil.copy(m["file"], fd)
        subprocess.run(["fc-cache", "-f"], check=False, capture_output=True)
        print(f"설치: {fd}")
    print(f"manifest: {out / 'manifest.json'} ({len(manifest)}개 파일)")


if __name__ == "__main__":
    main()
