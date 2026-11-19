#!/usr/bin/env python3
"""data/design-tokens.json(분야별 디자인 프리셋)을 만든다. 색 대비를 검사해 기준에 못 미치면 멈춘다.

출처(값을 옮겨 적었고, 각 저장소의 라이선스를 따른다)
  - KRDS 범정부 UI/UX 디자인 시스템 디자인 토큰: github.com/KRDS-uiux/krds-uiux (행정안전부, 공공누리 제1유형·출처표시)
  - Open Color: github.com/yeun/open-color (MIT, © 2016 heeyeun)
  - Urban Institute 그래프 스타일: github.com/vega/vega-themes theme-urbaninstitute (BSD-3-Clause)
  - ColorBrewer Dark2: github.com/axismaps/colorbrewer (Apache-2.0, Cynthia Brewer)
  - Okabe & Ito(2008) 색각 이상 친화 팔레트
  - 색 단계(같은 색의 옅은 단계)는 adobe/leonardo(Apache-2.0)의 '목표 대비율로 색 만들기' 방식을 따랐다.

색 원칙: 한 덱에 색은 세 가지까지 — 제목색(짙은 색) · 강조색 하나 · 무채색(회색).
그래프 계열이 넷 이상이면 새 색을 더하지 않고 강조색의 옅은 단계(강조색 대비의 0.7·0.5·0.36배)를 쓴다.
"""
import colorsys
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def lum(h):
    def ch(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def cr(a, b):
    x, y = sorted((lum(a), lum(b)), reverse=True)
    return (x + 0.05) / (y + 0.05)


def tint_to(hexv, ratio):
    """같은 색상·채도에서 밝기만 올려 흰 바탕 대비가 ratio가 되는 색(leonardo의 목표 대비율 방식을 HLS로 단순화)."""
    r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s_ = colorsys.rgb_to_hls(r, g, b)
    lo, hi = l, 0.98
    for _ in range(40):
        mid = (lo + hi) / 2
        c = "%02X%02X%02X" % tuple(round(x * 255) for x in colorsys.hls_to_rgb(h, mid, s_))
        if cr(c, "FFFFFF") > ratio:
            lo = mid
        else:
            hi = mid
    return "%02X%02X%02X" % tuple(round(x * 255) for x in colorsys.hls_to_rgb(h, (lo + hi) / 2, s_))


def tints(hexv):
    """강조색의 옅은 단계 셋. 목표 대비율은 강조색 자신의 대비에 비례(0.7·0.5·0.36배, 하한 1.5·1.32·1.18)."""
    c0 = cr(hexv, "FFFFFF")
    return [tint_to(hexv, max(lo, c0 * f)) for f, lo in ((0.7, 1.5), (0.5, 1.32), (0.36, 1.18))]


SRC = {
    "krds": "KRDS 디자인 토큰(행정안전부, 공공누리 제1유형) github.com/KRDS-uiux/krds-uiux",
    "oc": "Open Color(MIT) github.com/yeun/open-color",
    "urban": "Urban Institute 그래프 스타일(vega-themes, BSD-3-Clause) github.com/vega/vega-themes",
    "brewer": "ColorBrewer Dark2(Apache-2.0) github.com/axismaps/colorbrewer",
    "okabe": "Okabe & Ito(2008) 색각 이상 친화 팔레트",
    "leonardo": "목표 대비율로 색 단계 만들기(adobe/leonardo, Apache-2.0) github.com/adobe/leonardo",
}

# text: 본문 글자, title: 제목(짙은 색), accent: 강조색 하나, gray: 무채색 계열 그래프·보조 색
# muted: 보조 글자, light: 옅은 바탕(표 줄무늬), line: 선·테두리, grid: 그래프 눈금선 — 모두 무채색이거나 제목·강조색의 옅은 단계
P = {
    "public": dict(name="공공(KRDS)", src=["krds"], text="1E2124", title="063A74", accent="256EF4", gray="8A949E",
                   muted="58616A", light="F4F5F6", line="CDD1D5", grid="E6E8EA", hlink="0B50D0"),
    "memorial": dict(name="추모·의전(KRDS 무채색)", src=["krds"], text="1E2124", title="052B57", accent="346FB2", gray="8A949E",
                     muted="58616A", light="F4F5F6", line="CDD1D5", grid="E6E8EA", hlink="1C589C"),
    "corporate": dict(name="기업 보고(Urban Institute)", src=["urban"], text="262223", title="0A4C6A", accent="1696D2", gray="9D9D9D",
                      muted="5C5859", light="F5F5F5", line="D2D2D2", grid="DEDDDD", hlink="12719E"),
    "tech": dict(name="IT·스타트업(Open Color)", src=["oc"], text="212529", title="5F3DC4", accent="15AABF", gray="868E96",
                 muted="495057", light="F8F9FA", line="DEE2E6", grid="E9ECEF", hlink="7048E8"),
    "academic": dict(name="학술(ColorBrewer Dark2)", src=["brewer", "oc"], text="212529", title="1B2A4A", accent="1B9E77", gray="868E96",
                     muted="495057", light="F8F9FA", line="DEE2E6", grid="E9ECEF", hlink="364FC7"),
    "culture": dict(name="문화·예술(Open Color)", src=["oc"], text="343A40", title="862E9C", accent="E8590C", gray="868E96",
                    muted="495057", light="F8F9FA", line="DEE2E6", grid="E9ECEF", hlink="9C36B5"),
    "kids": dict(name="아동·가족(Open Color)", src=["oc"], text="343A40", title="1864AB", accent="FF922B", gray="ADB5BD",
                 muted="495057", light="FFF9DB", line="FFE066", grid="F1F3F5", hlink="1971C2"),
    "promo": dict(name="홍보·행사(Open Color)", src=["oc"], text="212529", title="C92A2A", accent="FD7E14", gray="868E96",
                  muted="495057", light="FFF5F5", line="DEE2E6", grid="E9ECEF", hlink="1864AB"),
    "edu": dict(name="교육·강의(Open Color)", src=["oc"], text="212529", title="0B7285", accent="F76707", gray="868E96",
                muted="495057", light="E3FAFC", line="DEE2E6", grid="E9ECEF", hlink="0B7285"),
    "game": dict(name="게임·레트로(Open Color)", src=["oc"], text="212529", title="5F3DC4", accent="22B8CF", gray="ADB5BD",
                 muted="495057", light="F3F0FF", line="D0BFFF", grid="E9ECEF", hlink="7048E8"),
}
for p in P.values():
    # 그래프·테마 강조 1~6: 강조색 → 제목색 → 회색 → 강조색 옅은 단계(대비 3:1, 2:1, 1.5:1)
    p["tints"] = tints(p["accent"])
    p["accents"] = [p["accent"], p["title"], p["gray"]] + p["tints"]
CHART_PALETTES = {
    "colorblind": dict(name="색각 이상 친화(Okabe-Ito 파랑·주황 + 회색)", src=["okabe"], accents=["0072B2", "E69F00", "8A8A8A"]),
}
for p in CHART_PALETTES.values():
    p["accents"] += tints(p["accents"][0])
GENRE = {"공공·정책": "public", "기념·의전": "memorial", "기업·비즈니스": "corporate", "IT·스타트업": "tech", "학술·연구": "academic",
         "문화·예술·인문": "culture", "아동·가족": "kids", "홍보·행사": "promo", "교육·강의": "edu", "게임·레트로": "game"}
# 글자 크기 단계(pt). KRDS의 제목·본문 비율(헤딩 32·24·19, 본문 17)을 발표 화면 크기로 옮겼다.
TYPE = {"16:9": {"cover": 48, "title": 32, "subtitle": 24, "body": 20, "small": 16, "caption": 12},
        "4:3": {"cover": 40, "title": 28, "subtitle": 22, "body": 18, "small": 14, "caption": 11}}
RULES = {
    "maxColors": 3,
    "lineSpacing": 1.2, "paraSpaceBefore": 6, "titleTracking": -0.02,
    "chart": {"gapWidth": 60, "lineWidthPt": 2.25, "gridWidthPt": 0.75, "dataLabelsMaxPoints": 12, "pieMaxSlices": 6,
              "source": "Urban Institute 스타일: 값 축선 없음, 옅은 눈금선, 눈금 표시 없음, 제목 왼쪽, 막대 간격 좁게"},
    "table": {"header": "title", "band": "light", "border": "line", "numbersRight": True},
}

errs = []
for k, p in P.items():
    for role, need in (("text", 7.0), ("title", 4.5), ("muted", 4.5), ("hlink", 4.5), ("accent", 2.0)):
        c = cr(p[role], "FFFFFF")
        if c < need:
            errs.append(f"{k}.{role} {p[role]} 대비 {c:.1f} < {need}")
    if cr(p["title"], p["light"]) < 4.5:
        errs.append(f"{k}.title on light {cr(p['title'], p['light']):.1f}")
    p["contrast"] = {r: round(cr(p[r], "FFFFFF"), 1) for r in ("text", "title", "muted")}
if errs:
    raise SystemExit("대비 기준 미달:\n" + "\n".join(errs))
out = {"_note": "분야별 디자인 프리셋. scripts/build_design_tokens.py로 만든다. 출처는 sources를 보라.",
       "sources": SRC, "presets": P, "chartPalettes": CHART_PALETTES, "genre": GENRE, "type": TYPE, "rules": RULES}
(ROOT / "data" / "design-tokens.json").write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n", "utf-8")
print("presets", len(P), {k: v["contrast"]["title"] for k, v in P.items()})
