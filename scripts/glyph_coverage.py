#!/usr/bin/env python3
"""폰트별 기호·한글 지원 범위를 미리 계산해 data/glyph-coverage.json에 저장한다(분석 스킬이 추천할 때 씀).

각 폰트의 Regular(없으면 가장 가까운 굵기) 파일 하나를 받아 cmap을 본다. 주소가 그대로면 이전 결과를 재사용한다.
"""
import importlib.util, json, sys, tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("ff", ROOT / "skill" / "scripts" / "fetch_fonts.py")
ff = importlib.util.module_from_spec(spec); spec.loader.exec_module(ff)
OUT = ROOT / "data" / "glyph-coverage.json"


def main():
    data = str(ROOT / "data")
    files = ff.load("fontfiles.json", data)["fonts"]
    google = {g["family"]: g for g in ff.load("google-fonts.json", data)["fonts"] if g.get("redistributable")}
    old = json.loads(OUT.read_text("utf-8")).get("fonts", {}) if OUT.exists() else {}
    res = {}
    jobs = []
    for k, f in files.items():
        w = ff.nearest(list(f["files"]), "Regular")
        url = f["files"][w]
        g = google.get(f.get("google") or f["family"])
        jobs.append((k, f["name"], url, w, g))
    norm = lambda x: (x or "").replace(" ", "").lower()
    seen = {norm(f.get("google") or f["family"]) for f in files.values()}
    for fam, g in google.items():
        if norm(fam) not in seen:
            jobs.append(("G:" + fam, g["name"], "google", ff.nearest(g["weights"], "Regular"), g))
    with tempfile.TemporaryDirectory() as td:
        for key, name, url, w, g in jobs:
            if url == "google":
                url = ff.google_files(g, [w]).get(w)
            if key in old and old[key].get("url") == url:
                res[key] = old[key]; continue
            try:
                dest = Path(td) / f"{abs(hash(url))}.font"
                ff.download(url, dest)
                info = ff.inspect(dest)
                res[key] = {"name": name, "url": url, "missing": info["missing"], "hangul": info["hangul"],
                            "outline": info["outline"], "embeddable": info["embeddable"]}
                print(f"{name:<16} 한글 {info['hangul']:>5}  빠진 기호 {info['missing'] or '-'}")
            except Exception as e:
                print(f"! {name}: {e}", file=sys.stderr)
                if key in old:
                    res[key] = old[key]
    OUT.write_text(json.dumps({"symbols": ff.SYMBOLS, "fonts": res}, ensure_ascii=False, indent=1) + "\n", "utf-8")
    print(f"저장: {OUT} ({len(res)}종)")


if __name__ == "__main__":
    main()
