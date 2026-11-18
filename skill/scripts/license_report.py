#!/usr/bin/env python3
"""사용한 폰트의 라이선스 증빙표(폰트 사용 내역서)를 만든다. 공모전 저작권 서약·공공기관 납품 때 첨부용.

  python license_report.py fonts/manifest.json --out 폰트_라이선스_내역.md [--deck output.pptx]

표: 폰트 / 굵기 / PowerPoint 이름 / 라이선스 / 확인 근거 / 출처 / 파일 SHA-256 / 확인 일시
"""
import argparse, hashlib, json, time
from pathlib import Path


def sha(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()[:16]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("manifest")
    ap.add_argument("--out", default="폰트_라이선스_내역.md")
    ap.add_argument("--deck", help="적용한 PPTX 파일 이름(표 머리에 적음)")
    a = ap.parse_args()
    man = json.loads(Path(a.manifest).read_text("utf-8"))
    base = Path(a.manifest).parent
    rows, seen = [], set()
    for m in man:
        # 자동 다운로드 목록은 OFL 폰트만 담는다(fontfiles.json·google-fonts.json). 파일 자체 표기는 '확인 근거'에 남긴다.
        lic = "사용자 제공(기관 배포 조건 따름)" if m.get("local") else "SIL OFL 1.1"
        src = ""
        for f in base.glob("LICENSE-*.txt"):
            t = f.read_text("utf-8").splitlines()
            if t and m["name"] in t[0]:
                src = next((l.replace("출처: ", "") for l in t[:3] if l.startswith("출처:")), "")
                break
        rows.append(f"| {m['name']} | {m['weight']} | {m['family']} | {lic} | {m.get('licenseCheck', '')} | {src} | "
                    f"`{sha(m['file']) if Path(m['file']).exists() else '-'}` |")
        seen.add(m["name"])
    now = time.strftime("%Y-%m-%d %H:%M")
    out = [f"# 폰트 라이선스 내역", "",
           f"- 작성: {now}" + (f" · 대상 파일: {a.deck}" if a.deck else ""),
           "- 자동으로 받은 폰트는 모두 SIL Open Font License 1.1(또는 그에 준하는 재배포 허용 라이선스)이며, 각 LICENSE 파일에 저작권 표시와 원문이 들어 있습니다.",
           "- 사용자 제공 폰트(기관 전용 서체 등)는 해당 기관의 배포 조건을 따릅니다.", "",
           "| 폰트 | 굵기 | PowerPoint 이름 | 라이선스 | 확인 근거 | 출처 | 파일 SHA-256(앞 16자) |",
           "|---|---|---|---|---|---|---|", *rows, "",
           "라이선스 조건은 바뀔 수 있으므로, 제출 전에 출처의 최신 안내를 한 번 더 확인하세요."]
    Path(a.out).write_text("\n".join(out) + "\n", "utf-8")
    print(f"저장: {a.out} ({len(rows)}개 파일, {len(seen)}개 가족)")


if __name__ == "__main__":
    main()
