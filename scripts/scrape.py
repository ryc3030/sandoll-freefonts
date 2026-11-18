#!/usr/bin/env python3
"""산돌구름 무료폰트(https://www.sandollcloud.com/freefonts) 목록과 상세정보를 수집해 data/fonts.json에 저장한다.

폰트 파일은 내려받거나 재배포하지 않는다. 목록·라이선스 요약·공식 다운로드(제작사) 링크만 모은다.
"""
import json, re, sys, time, urllib.request
from pathlib import Path
from bs4 import BeautifulSoup

BASE = "https://www.sandollcloud.com"
LIST_URL = BASE + "/freefonts"
UA = "Mozilla/5.0 (compatible; sandoll-freefonts-index/1.0; +https://github.com/)"
OUT = Path(__file__).resolve().parent.parent / "data" / "fonts.json"
SECTIONS = {"animation01": "steady", "animation02": "new", "animation03": "ofl"}


def get(url, tries=3):
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Language": "ko-KR"})
            with urllib.request.urlopen(req, timeout=30) as r:
                return r.read().decode("utf-8", "replace")
        except Exception as e:  # noqa
            if i == tries - 1:
                raise
            time.sleep(2 * (i + 1))


def clean(s):
    return re.sub(r"\s+", " ", (s or "")).strip()


def parse_list(html):
    soup = BeautifulSoup(html, "html.parser")
    fonts = {}
    for sec_id, key in SECTIONS.items():
        sec = soup.find("section", id=sec_id)
        if not sec:
            continue
        for a in sec.find_all("a", href=re.compile(r"^/free-font/\d+/")):
            m = re.match(r"^/free-font/(\d+)/([^/?#]+)", a["href"])
            fid, slug = m.group(1), m.group(2)
            f = fonts.setdefault(fid, {"id": int(fid), "slug": slug, "sections": [], "instant": False, "isNew": False})
            if key not in f["sections"]:
                f["sections"].append(key)
            txt = a.get_text(" ")
            if "즉시사용" in txt:
                f["instant"] = True
            if "신규폰트" in txt:
                f["isNew"] = True
            img = a.find("img", class_="img_before") or a.find("img")
            if img and img.get("alt") and "name" not in f:
                f["name"] = clean(img["alt"].split("|")[0])
            fn = a.find(class_="fontname")
            if fn:
                f["name"] = clean(fn.get("ko-fontname") or fn.get_text())
                if fn.get("en-fontname"):
                    f["nameEn"] = clean(fn["en-fontname"])
    return list(fonts.values())


def parse_detail(html):
    soup = BeautifulSoup(html, "html.parser")
    d = {}
    h1 = soup.find("h1", id="fg_name")
    if h1:
        d["name"] = clean(h1.get_text())
    og = soup.find("meta", property="og:image")
    if og and og.get("content"):
        d["previewImage"] = og["content"]
    # 상단 태그 (OFL, 상업용, 임베딩, 로고•CI ...)
    # 'nonechk' 클래스가 붙은 태그는 해당 없음(회색 표시)이다.
    order = ["OFL", "상업용", "임베딩", "로고•CI"]
    tags = {clean(t.get_text()) for t in soup.select("span.tag-chk")
            if clean(t.get_text()) and "nonechk" not in (t.get("class") or [])}
    d["tags"] = sorted(tags, key=lambda x: (order.index(x.split()[0]) if x.split()[0] in order else 9, x))
    d["moods"] = [clean(t.get_text()) for t in soup.select("div.after-bar span.tag") if clean(t.get_text())]
    # 정보 목록
    info = {}
    for li in soup.select("ul.font_info_list li"):
        lab, val = li.find("label"), li.find("span")
        if lab and val:
            info[clean(lab.get_text())] = clean(val.get_text())
    keymap = {"종수": "weights", "출시년도": "year", "분류": "category", "포맷": "format",
              "파일명": "fileName", "지원언어": "languages", "제작사": "maker", "저작권": "copyright", "스펙": "spec"}
    for k, v in keymap.items():
        if info.get(k):
            d[v] = info[k]
    # 라이선스 원문(텍스트)
    for h4 in soup.find_all("h4"):
        if "라이선스 원문" in h4.get_text():
            box = h4.find_next_sibling("div")
            if box:
                paras = [clean(p.get_text(" ")) for p in box.find_all(["p", "li"])]
                paras = [p for p in paras if p]
                d["licenseText"] = "\n".join(paras) if paras else clean(box.get_text(" "))
            break
    # 라이선스 요약표
    summary = []
    for h4 in soup.find_all("h4"):
        if "라이선스 요약표" in h4.get_text():
            table = h4.find_next("table")
            if table:
                for tr in table.select("tbody tr"):
                    tds = tr.find_all("td")
                    if len(tds) >= 3:
                        summary.append({"category": clean(tds[0].get_text()), "scope": clean(tds[1].get_text()),
                                        "allowed": clean(tds[2].get_text())})
            break
    if summary:
        d["licenseSummary"] = summary
    # 제작사(공식 배포) 사이트
    for a in soup.find_all("a", href=re.compile(r"^https?://")):
        if "제작사 사이트" in a.get_text():
            d["makerUrl"] = a["href"]
            break
    return d


def main():
    html = get(LIST_URL)
    fonts = parse_list(html)
    if len(fonts) < 10:
        print(f"목록 파싱 결과가 {len(fonts)}개뿐이라 중단합니다(페이지 구조 변경 의심).", file=sys.stderr)
        sys.exit(1)
    prev = {}
    if OUT.exists():
        prev = {f["id"]: f for f in json.loads(OUT.read_text("utf-8")).get("fonts", [])}
    for i, f in enumerate(fonts, 1):
        url = f"{BASE}/free-font/{f['id']}/{f['slug']}"
        f["sandollUrl"] = url
        try:
            det = parse_detail(get(url))
            name = f.get("name")
            f.update(det)
            if name and not det.get("name"):
                f["name"] = name
        except Exception as e:
            print(f"[{f['id']}] 상세 수집 실패: {e}", file=sys.stderr)
            if f["id"] in prev:
                f = {**prev[f["id"]], **f}
                fonts[i - 1] = f
        f["ofl"] = "OFL" in f.get("tags", [])
        print(f"{i:>2}/{len(fonts)} {f.get('name')} {f.get('tags')} {f.get('makerUrl', '-')}")
        time.sleep(0.5)
    data = {"source": LIST_URL, "fonts": fonts}
    # 내용이 바뀌었을 때만 날짜 갱신
    old = json.loads(OUT.read_text("utf-8")) if OUT.exists() else {}
    if old.get("fonts") != fonts:
        data["updatedAt"] = time.strftime("%Y-%m-%dT%H:%M:%S+09:00", time.gmtime(time.time() + 9 * 3600))
    else:
        data["updatedAt"] = old.get("updatedAt")
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", "utf-8")
    print(f"저장: {OUT} ({len(fonts)}종)")


if __name__ == "__main__":
    main()
