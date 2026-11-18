#!/usr/bin/env python3
"""Google Fonts에서 한국어(korean subset)를 지원하는 폰트 목록을 받아 data/google-fonts.json에 저장한다.

Google Fonts는 목록(metadata)과 폰트 파일을 공개 배포하며, 라이선스는 OFL·Apache 등 재배포를 허용하는 것뿐이다.
라이선스는 google/fonts 저장소의 폴더(ofl/apache/ufl)로 확인한다.
"""
import json, re, sys, time, urllib.request, urllib.error
from pathlib import Path

META = "https://fonts.google.com/metadata/fonts"
RAW = "https://raw.githubusercontent.com/google/fonts/main/{lic}/{dir}/METADATA.pb"
OUT = Path(__file__).resolve().parent.parent / "data" / "google-fonts.json"
UA = {"User-Agent": "sandoll-freefonts-index/1.0"}
WEIGHT_NAMES = {100: "Thin", 200: "ExtraLight", 300: "Light", 400: "Regular", 500: "Medium",
                600: "SemiBold", 700: "Bold", 800: "ExtraBold", 900: "Black"}
STYLE = {"Sans Serif": "sans", "Serif": "serif", "Display": "display", "Handwriting": "handwriting", "Monospace": "mono"}
LIC_NAME = {"ofl": "OFL", "apache": "Apache 2.0", "ufl": "Ubuntu Font License"}


# 한글 이름과 성격 태그(분석 스킬의 추천에 쓰임). 새 폰트가 들어오면 여기만 보강하면 된다.
# 태그: formal 격식 · neutral 중립 · friendly 친근 · playful 발랄 · cute 귀여움 · handwritten 손글씨 · retro 복고
#       elegant 우아 · traditional 전통 · academic 학술 · tech 기술·IT · impact 강한 제목 · calm 차분
KO = {
    "Noto Sans KR": ("본고딕 (Noto Sans KR)", ["neutral", "formal", "academic"], ["title", "body"]),
    "Nanum Gothic": ("나눔고딕", ["neutral", "formal"], ["body"]),
    "Noto Serif KR": ("본명조 (Noto Serif KR)", ["formal", "academic", "traditional", "elegant"], ["title", "body"]),
    "Nanum Myeongjo": ("나눔명조", ["traditional", "academic", "calm"], ["body"]),
    "Nanum Gothic Coding": ("나눔고딕코딩", ["tech"], ["code"]),
    "Gothic A1": ("고딕 A1", ["neutral", "formal", "tech"], ["title", "body"]),
    "Nanum Pen Script": ("나눔손글씨 펜", ["handwritten", "friendly", "calm"], ["accent"]),
    "Jua": ("배민 주아체", ["cute", "friendly", "playful"], ["title"]),
    "Black Han Sans": ("검은고딕", ["impact"], ["title"]),
    "IBM Plex Sans KR": ("IBM Plex Sans KR", ["tech", "neutral", "formal"], ["title", "body"]),
    "Dongle": ("동글", ["cute", "playful", "friendly"], ["title"]),
    "Do Hyeon": ("배민 도현체", ["retro", "impact", "playful"], ["title"]),
    "Gowun Batang": ("고운바탕", ["elegant", "calm", "traditional"], ["title", "body"]),
    "Gowun Dodum": ("고운돋움", ["friendly", "calm"], ["body"]),
    "Hahmlet": ("함렛", ["elegant", "academic", "formal"], ["title", "body"]),
    "Nanum Brush Script": ("나눔손글씨 붓", ["handwritten", "traditional"], ["accent"]),
    "Gaegu": ("개구", ["handwritten", "cute", "playful"], ["title", "accent"]),
    "Hi Melody": ("하이멜로디", ["handwritten", "cute"], ["accent"]),
    "Sunflower": ("해바라기", ["friendly", "retro"], ["title"]),
    "Asta Sans": ("아스타 산스", ["tech", "neutral", "friendly"], ["title", "body"]),
    "Song Myung": ("송명", ["traditional", "elegant"], ["title"]),
    "Bagel Fat One": ("베이글 팻", ["playful", "impact", "cute"], ["title"]),
    "Single Day": ("싱글데이", ["cute", "playful"], ["accent"]),
    "Gamja Flower": ("감자꽃", ["handwritten", "cute"], ["accent"]),
    "Gasoek One": ("가석", ["impact", "retro"], ["title"]),
    "Dokdo": ("독도", ["handwritten", "playful"], ["accent"]),
    "East Sea Dokdo": ("동해독도", ["handwritten", "impact"], ["accent"]),
    "Gugi": ("구기", ["retro", "playful"], ["title"]),
    "Stylish": ("스타일리시", ["elegant", "retro"], ["title"]),
    "Poor Story": ("서툰이야기", ["handwritten", "cute"], ["accent"]),
    "Yeon Sung": ("연성", ["retro", "friendly"], ["title"]),
    "Cute Font": ("귀여운폰트", ["cute", "handwritten"], ["accent"]),
    "Orbit": ("오르빗", ["tech", "retro"], ["title"]),
    "Diphylleia": ("산하엽", ["elegant", "traditional"], ["title"]),
    "Kirang Haerang": ("기랑해랑", ["retro", "playful"], ["title"]),
    "Black And White Picture": ("흑백사진", ["retro", "traditional"], ["title"]),
    "Grandiflora One": ("그랜디플로라", ["elegant", "traditional"], ["title"]),
    "Moirai One": ("모이라이", ["impact", "playful"], ["title"]),
}


def get(url):
    with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=60) as r:
        return r.read().decode("utf-8")


def license_of(family):
    d = re.sub(r"[^a-z0-9]", "", family.lower())
    for lic in ("ofl", "apache", "ufl"):
        try:
            get(RAW.format(lic=lic, dir=d))
            return lic
        except urllib.error.HTTPError:
            continue
    return None


def main():
    t = get(META)
    data = json.loads(t[t.find("{"):])
    fams = [f for f in data["familyMetadataList"] if "korean" in f.get("subsets", [])]
    if len(fams) < 10:
        print(f"한국어 폰트가 {len(fams)}종뿐이라 중단합니다(응답 형식 변경 의심).", file=sys.stderr)
        sys.exit(1)
    prev = {}
    if OUT.exists():
        prev = {f["family"]: f for f in json.loads(OUT.read_text("utf-8")).get("fonts", [])}
    out = []
    for f in sorted(fams, key=lambda x: x.get("popularity", 9999)):
        fam = f["family"]
        weights = sorted({int(k) for k in f["fonts"] if k.isdigit()})
        rev = {v: k for k, v in LIC_NAME.items()}
        lic = rev.get(prev.get(fam, {}).get("license")) or license_of(fam)
        time.sleep(0.2)
        q = fam.replace(" ", "+")
        ko = KO.get(fam, (fam, [], ["title"]))
        out.append({
            "family": fam,
            "name": ko[0],
            "tags": ko[1],
            "use": ko[2],
            "category": f.get("category"),
            "style": STYLE.get(f.get("category"), "display"),
            "weights": [WEIGHT_NAMES.get(w, str(w)) for w in weights],
            "weightValues": weights,
            "variable": bool(f.get("axes")),
            "designers": f.get("designers", []),
            "popularity": f.get("popularity"),
            "dateAdded": f.get("dateAdded"),
            "license": LIC_NAME.get(lic, lic),
            "redistributable": lic in ("ofl", "apache", "ufl"),
            "css": f"https://fonts.googleapis.com/css2?family={q}:wght@{';'.join(map(str, weights))}&display=swap",
            "specimenUrl": f"https://fonts.google.com/specimen/{q}",
        })
        print(f"{fam:<26} {out[-1]['license']:<6} {','.join(out[-1]['weights'])}")
    old = json.loads(OUT.read_text("utf-8")) if OUT.exists() else {}
    res = {"source": "https://fonts.google.com/?subset=korean", "fonts": out}
    res["updatedAt"] = (time.strftime("%Y-%m-%dT%H:%M:%S+09:00", time.gmtime(time.time() + 9 * 3600))
                        if old.get("fonts") != out else old.get("updatedAt"))
    OUT.write_text(json.dumps(res, ensure_ascii=False, indent=2) + "\n", "utf-8")
    print(f"저장: {OUT} ({len(out)}종)")


if __name__ == "__main__":
    main()
