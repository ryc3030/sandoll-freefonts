#!/usr/bin/env python3
"""PPTX를 읽어 덱의 성격(분야·글의 문체·밀도·색감)을 진단하고, 어울리는 서체 성격과 폰트 짝 후보를 낸다.

  python analyze_deck.py deck.pptx                 # 사람이 읽는 보고서
  python analyze_deck.py deck.pptx --json out.json # 결과를 JSON으로도 저장

판단 근거
  1) 분야: 슬라이드 글에 나온 어휘(정책·사업·연구·강의·IT·홍보·문화·아동·게임)
  2) 글의 문체: 개조식(~함/~임), 합니다체, 해요체, 반말·구어, 느낌표·이모지
  3) 밀도: 슬라이드당 글자 수, 표·차트 수 → 본문은 가독성 우선인지
  4) 색감: 테마·도형 색의 채도와 밝기 → 차분/선명/파스텔/어두움
  5) 현재 글꼴: 이미 쓰인 글꼴과 개수
결과의 '서체 성격 태그'로 fontfiles.json(산돌구름 OFL)과 google-fonts.json(Google Fonts)의 폰트에 점수를 매긴다.
점수는 출발점일 뿐이다. 최종 선택은 슬라이드 이미지를 직접 보고 판단한다.
"""
import argparse, colorsys, json, re, sys, urllib.request, zipfile
from collections import Counter
from pathlib import Path
from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS = {"a": A, "p": P}
REPO_RAW = "https://raw.githubusercontent.com/ryc3030/sandoll-freefonts/main/data/"

GENRES = {
    "공공·정책": ("formal neutral", r"정책|추진|계획|예산|사업|행정|부처|기관|위원회|조례|국민|시민|주민|지자체|공공|개선\s?방안|과제|성과|지원|제도|법령|시행|국정|도정|시정|군정|민원"),
    "기업·비즈니스": ("formal neutral", r"매출|시장|고객|전략|투자|성장|경쟁|수익|KPI|ROI|영업|브랜드|파트너|분기|실적|IR|사업\s?계획|비용|원가|점유율"),
    "학술·연구": ("academic formal", r"연구|분석|가설|변수|모형|선행\s?연구|표본|유의|회귀|이론|방법론|논문|실증|통계|검정|문헌|데이터\s?수집|함의|한계|결론"),
    "교육·강의": ("friendly neutral", r"학습|수업|강의|학생|퀴즈|활동|교육|과정|차시|목표|복습|예제|실습|정리해\s?봅시다|알아봅시다"),
    "IT·스타트업": ("tech neutral", r"플랫폼|AI|인공지능|데이터|앱|사용자|MVP|API|클라우드|솔루션|서비스|스타트업|SaaS|UX|UI|개발|출시|베타|기능"),
    "홍보·행사": ("impact playful", r"이벤트|행사|할인|혜택|참여|축제|캠페인|홍보|오픈|모집|선착순|경품|신청|초대|특가|런칭|기념"),
    "문화·예술·인문": ("elegant calm traditional", r"전시|예술|문화|공연|작가|역사|전통|감성|이야기|작품|시인|음악|미술|유산|기억|삶|여행"),
    "아동·가족": ("cute playful friendly", r"어린이|아이|유아|친구|놀이|동화|키즈|엄마|아빠|가족|우리\s?반|선생님"),
    "게임·레트로": ("retro playful", r"게임|레트로|픽셀|레벨|플레이|스테이지|보스|아이템|퀘스트|8비트"),
}
REGISTER = {
    "개조식": (r"(함|임|음|됨|필요|추진|확대|강화|마련|검토|완료|예정|가능)\s*[.)]?\s*$", "formal"),
    "합니다체": (r"(습니다|합니다|입니다|됩니다|십시오|겠습니다)\s*[.!?]?\s*$", "formal"),
    "해요체": (r"(해요|어요|아요|예요|이에요|네요|죠|세요|까요)\s*[.!?~]*\s*$", "friendly"),
    "반말·구어": (r"(했어|했지|한다|해봐|하자|거든|잖아|볼까|할래|이야|ㅋㅋ|ㅎㅎ)\s*[.!?~]*\s*$", "playful"),
}
EMOJI = re.compile("[\U0001F300-\U0001FAFF☀-➿]")


def xml(z, name):
    return etree.fromstring(z.read(name))


def texts_of(root):
    paras = []
    for p in root.iter(f"{{{A}}}p"):
        t = "".join(x.text or "" for x in p.iter(f"{{{A}}}t")).strip()
        if t:
            paras.append(t)
    return paras


def color_mood(hexes):
    if not hexes:
        return "판단 불가", [], {}
    sats, lights = [], []
    for h in hexes:
        r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
        hh, l, s = colorsys.rgb_to_hls(r, g, b)
        if 0.06 < l < 0.97:  # 흰색·검정 계열은 빼고 본다
            sats.append(s); lights.append(l)
    if not sats:
        return "무채색(흑백)", ["formal", "neutral"], {}
    s, l = sum(sats) / len(sats), sum(lights) / len(lights)
    stat = {"평균채도": round(s, 2), "평균밝기": round(l, 2)}
    if s < 0.18:
        return "차분한 무채색 계열", ["formal", "neutral", "calm"], stat
    if l > 0.75:
        return "밝은 파스텔", ["cute", "friendly", "calm"], stat
    if l < 0.35:
        return "어둡고 묵직함", ["formal", "elegant", "tech"], stat
    if s > 0.6:
        return "선명하고 강한 색", ["impact", "playful"], stat
    return "보통 채도의 색", ["neutral", "friendly"], stat


def analyze(path):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    slides = sorted([n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)], key=lambda n: int(re.findall(r"\d+", n)[-1]))
    pres = xml(z, "ppt/presentation.xml")
    sz = pres.find("p:sldSz", NS)
    ratio = round(int(sz.get("cx")) / int(sz.get("cy")), 2) if sz is not None else None

    per, all_lines, body_lines, fonts, colors = [], [], [], Counter(), []
    tables = charts = pics = 0
    for n in slides:
        r = xml(z, n)
        title, body = [], []
        for sp in r.iter(f"{{{P}}}sp"):
            ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
            t = texts_of(sp)
            (title if ph is not None and ph.get("type") in ("title", "ctrTitle") else body).extend(t)
        for gf in r.iter(f"{{{P}}}graphicFrame"):
            if gf.find(".//a:tbl", NS) is not None:
                tables += 1; body.extend(texts_of(gf))
            if gf.find(f".//{{{C}}}chart", {}) is not None or "chart" in etree.tostring(gf).decode()[:2000]:
                charts += 1
        pics += len(list(r.iter(f"{{{P}}}pic")))
        for el in r.iter(f"{{{A}}}latin", f"{{{A}}}ea"):
            tf = el.get("typeface")
            if tf and not tf.startswith("+"):
                fonts[tf] += 1
        colors += [c.get("val") for c in r.iter(f"{{{A}}}srgbClr") if c.get("val")]
        chars = sum(len(x) for x in title + body)
        per.append({"slide": len(per) + 1, "title": " / ".join(title)[:60], "chars": chars, "lines": len(body)})
        all_lines += title + body
        body_lines += body

    theme_fonts = {}
    for n in names:
        if re.match(r"ppt/theme/theme\d+\.xml$", n):
            t = xml(z, n)
            for tag in ("majorFont", "minorFont"):
                fs = t.find(f".//a:{tag}", NS)
                if fs is not None:
                    theme_fonts[tag] = {c: (fs.find(f"a:{c}", NS).get("typeface") if fs.find(f"a:{c}", NS) is not None else None) for c in ("latin", "ea")}
            clr = t.find(".//a:clrScheme", NS)
            if clr is not None:
                for k in ("accent1", "accent2", "accent3", "accent4", "dk2", "lt2"):
                    e = clr.find(f"a:{k}/a:srgbClr", NS)
                    if e is not None:
                        colors.append(e.get("val"))
            break

    text = "\n".join(all_lines)
    hangul = len(re.findall(r"[가-힣]", text)); latin = len(re.findall(r"[A-Za-z]", text)); digit = len(re.findall(r"\d", text))
    total = max(1, hangul + latin + digit)

    genre_scores = {g: len(re.findall(rx, text, re.I)) for g, (_, rx) in GENRES.items()}
    gsum = sum(genre_scores.values()) or 1
    genres = sorted(((g, round(v / gsum, 2)) for g, v in genre_scores.items() if v), key=lambda x: -x[1])[:3]

    # 본문 줄만 본다(제목은 대개 명사구라 문체 단서가 아님). 서술어 어미가 없으면 명사형 개조식으로 친다.
    reg = Counter()
    for line in body_lines:
        line = EMOJI.sub("", line).strip()
        if len(line) < 6 or not re.search(r"[가-힣]", line):
            continue
        for name, (rx, _) in REGISTER.items():
            if re.search(rx, line):
                reg[name] += 1; break
        else:
            reg["개조식"] += 1
    reg_total = sum(reg.values()) or 1
    register = sorted(((k, round(v / reg_total, 2)) for k, v in reg.items()), key=lambda x: -x[1])
    excl = text.count("!"); emoji = len(EMOJI.findall(text))

    avg = round(sum(p["chars"] for p in per) / max(1, len(per)))
    mood, color_tags, color_stat = color_mood(colors)

    # 서체 성격 태그 점수
    tags = Counter()
    for g, w in genres:
        for t in GENRES[g][0].split():
            tags[t] += 3 * w
    for k, w in register:
        tags[REGISTER[k][1]] += 2 * w
    for t in color_tags:
        tags[t] += 0.8
    if excl + emoji > max(3, len(per)):
        tags["playful"] += 1
    dense = avg > 220 or tables + charts >= max(2, len(per) // 3)
    if dense:
        tags["neutral"] += 1
    return {
        "file": str(path), "slides": len(per), "aspect": ratio,
        "scripts": {"한글": round(hangul / total, 2), "영문": round(latin / total, 2), "숫자": round(digit / total, 2)},
        "genres": genres, "register": register, "exclamations": excl, "emoji": emoji,
        "density": {"슬라이드당평균글자": avg, "가장많은슬라이드": max(per, key=lambda p: p["chars"])["slide"] if per else None,
                    "표": tables, "차트": charts, "이미지": pics, "빽빽함": dense},
        "color": {"인상": mood, **color_stat},
        "fonts_now": {"테마": theme_fonts, "직접지정": dict(fonts.most_common(8))},
        "tone_tags": [t for t, _ in tags.most_common(4)], "tone_scores": {t: round(v, 2) for t, v in tags.most_common()},
        "per_slide": per,
    }


def load(name, data_dir):
    if data_dir and (Path(data_dir) / name).exists():
        return json.loads((Path(data_dir) / name).read_text("utf-8"))
    with urllib.request.urlopen(REPO_RAW + name, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def recommend(res, data_dir=None, k=3):
    files = load("fontfiles.json", data_dir)["fonts"]
    google = load("google-fonts.json", data_dir)["fonts"]
    pool, seen = [], set()
    for fid, f in files.items():
        pool.append({"key": f["name"], "family": f["family"], "src": "산돌 OFL", "style": f["style"], "use": f["use"],
                     "tags": f.get("tags", []), "weights": list(f["files"]), "latinOnly": f.get("latinOnly", False)})
        seen.add(f["family"].replace(" ", "").lower())
    for g in google:
        if g["family"].replace(" ", "").lower() in seen or not g.get("redistributable"):
            continue
        pool.append({"key": g["name"], "family": g["family"], "src": "Google", "style": g["style"], "use": g["use"],
                     "tags": g.get("tags", []), "weights": g["weights"], "latinOnly": False})
    ts = res["tone_scores"]; dense = res["density"]["빽빽함"]; latin_heavy = res["scripts"]["영문"] > 0.5

    def score(f, role):
        if f["latinOnly"] and not latin_heavy:
            return -9
        if role not in f["use"]:
            return -9
        s = sum(ts.get(t, 0) for t in f["tags"])
        if role == "body":
            if f["style"] in ("display", "handwriting", "pixel", "mono"):
                return -9
            s += 0.3 * min(len(f["weights"]), 6) + (1.0 if dense and "neutral" in f["tags"] else 0)
        else:
            heavy = {"ExtraBold", "Black", "Heavy", "Bold"} & set(f["weights"])
            s += 0.4 if heavy else 0
        return round(s, 2)

    titles = sorted(pool, key=lambda f: -score(f, "title"))
    bodies = sorted(pool, key=lambda f: -score(f, "body"))
    pairs = []
    # ① 한 가족으로 제목·본문 (굵기 차이로 위계)
    for f in bodies:
        if "title" in f["use"] and score(f, "body") > -9 and len(f["weights"]) >= 3:
            order = ("Bold", "SemiBold", "ExtraBold") if f["style"] == "serif" else ("ExtraBold", "Black", "Heavy", "Bold")
            heavy = next((w for w in order if w in f["weights"]), "Bold")
            pairs.append({"방식": "한 가족", "제목": f"{f['key']} {heavy}", "본문": f"{f['key']} Regular",
                          "출처": f["src"], "점수": round(score(f, "title") + score(f, "body"), 2)})
            if len(pairs) >= 2:
                break
    # ② 제목 전용 개성 폰트 + 무난한 본문
    for t in titles[:6]:
        if t["key"] in [p["제목"].rsplit(" ", 1)[0] for p in pairs]:
            continue
        b = next(b for b in bodies if b["family"] != t["family"])
        tw = next((w for w in ("ExtraBold", "Black", "Heavy", "Bold", "Regular") if w in t["weights"]), t["weights"][0])
        pairs.append({"방식": "제목 개성 + 본문 안정", "제목": f"{t['key']} {tw}", "본문": f"{b['key']} Regular",
                      "출처": f"{t['src']} + {b['src']}", "점수": round(score(t, "title") + score(b, "body"), 2)})
        if len(pairs) >= k + 1:
            break
    return {"pairs": pairs[:k + 1],
            "title_top": [(f["key"], score(f, "title")) for f in titles[:6]],
            "body_top": [(f["key"], score(f, "body")) for f in bodies[:6]]}


def report(res, rec):
    L = []
    L.append(f"# 덱 진단: {Path(res['file']).name}")
    L.append(f"- 슬라이드 {res['slides']}장, 화면비 {res['aspect']} · 글자 구성 {res['scripts']}")
    L.append(f"- 분야 추정: " + (", ".join(f"{g}({int(w*100)}%)" for g, w in res["genres"]) or "단서 부족"))
    L.append(f"- 글의 문체: " + (", ".join(f"{k}({int(w*100)}%)" for k, w in res["register"]) or "문장 단서 부족")
             + f" · 느낌표 {res['exclamations']} · 이모지 {res['emoji']}")
    d = res["density"]
    L.append(f"- 밀도: 슬라이드당 평균 {d['슬라이드당평균글자']}자(최다 {d['가장많은슬라이드']}번), 표 {d['표']} · 차트 {d['차트']} · 이미지 {d['이미지']}"
             + (" → 본문은 가독성 최우선" if d["빽빽함"] else ""))
    L.append(f"- 색감: {res['color']['인상']} {({k: v for k, v in res['color'].items() if k != '인상'}) or ''}")
    L.append(f"- 지금 글꼴: 테마 {res['fonts_now']['테마']} / 직접 지정 {res['fonts_now']['직접지정']}")
    L.append(f"\n## 어울리는 서체 성격: {', '.join(res['tone_tags'])}")
    L.append(f"  (점수 {res['tone_scores']})")
    L.append("\n## 폰트 짝 후보")
    for i, p in enumerate(rec["pairs"], 1):
        L.append(f"{i}. [{p['방식']}] 제목 {p['제목']} / 본문 {p['본문']}  ({p['출처']}, 점수 {p['점수']})")
    L.append(f"\n제목 상위: {rec['title_top']}\n본문 상위: {rec['body_top']}")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx"); ap.add_argument("--json"); ap.add_argument("--data", help="data 폴더(오프라인)")
    a = ap.parse_args()
    res = analyze(a.pptx)
    rec = recommend(res, a.data)
    print(report(res, rec))
    if a.json:
        Path(a.json).write_text(json.dumps({"analysis": res, "recommendation": rec}, ensure_ascii=False, indent=2), "utf-8")


if __name__ == "__main__":
    main()
