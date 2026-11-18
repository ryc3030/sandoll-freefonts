#!/usr/bin/env python3
"""PPTX를 읽어 덱의 성격을 진단하고, 어울리는 서체 성격과 폰트 짝 후보를 낸다.

  python analyze_deck.py deck.pptx                 # 사람이 읽는 보고서
  python analyze_deck.py deck.pptx --json out.json # 결과를 JSON으로도 저장

보는 변수(근거는 보고서의 '판단 근거'에 모두 적힌다)
  발행  문서 속성의 회사·작성자, 바닥글, 기관명·직함 → '공공 발행' 여부(공공이면 목적과 상관없이 프리텐다드)
  내용  분야 어휘(한·영) · 공공 보고서 기호(□ ○ ※ Ⅰ.) · 학술 인용·통계 표기 · 금액·비율 · 한자·영문(어절) 비율
        글의 문체(개조식·합니다체·한다체·해요체·반말) · 청중(2인칭·권유) · 문장 길이 · 느낌표·물음표·이모지
        발표자 노트·SmartArt·차트 글(가중치 0.5)
  구조  슬라이드 수(숨김 제외) · 화면비 · 슬라이드당 글자 수(발표형/읽는 보고서형) · 자리표시자 제목 비율
        글머리 깊이 · 글자 크기 분포 · 굵은 글자 비율 · 표·차트·SmartArt·이미지 수와 면적
  시각  배경 밝기(슬라이드→레이아웃→마스터, clrMap 반영) · 채운 도형 색의 면적 가중 채도·밝기 · 지금 글꼴 성격
  글꼴  덱에 쓰인 기호·드문 한글이 후보 폰트에 있는지(data/glyph-coverage.json)
분야별 기본 짝이 있다(공공·정책 = 프리텐다드). 점수는 출발점일 뿐이며 최종 선택은 슬라이드 이미지를 보고 판단한다.
"""
import argparse, colorsys, json, math, re, statistics, urllib.request, zipfile
from collections import Counter
from pathlib import Path
from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": A, "p": P, "r": RNS}
REPO_RAW = "https://raw.githubusercontent.com/ryc3030/sandoll-freefonts/main/data/"

GENRES = {
    "공공·정책": ("formal neutral",
              r"정책|추진|계획|예산|사업|행정|부처|기관|위원회|조례|국민|시민|주민|지자체|공공|개선\s?방안|과제|제도|법령|시행|"
              r"국정|도정|시정|군정|구정|민원|협의|관계\s?기관|소관|조치|이행|점검|현황|향후\s?계획|기대\s?효과|건의|협조|예산안|국비|도비|시비|"
              r"행정복지센터|주민센터|누리집|공모|보조금|복지"),
    "기업·비즈니스": ("formal neutral",
                r"매출|시장|고객|전략|투자|성장|경쟁|수익|영업|브랜드|파트너|분기|실적|비용|원가|점유율|"
                r"영업이익|밸류|주주|런웨이|손익|세일즈|거래처|계약|매입|마진|"
                r"\brevenue\b|\bgrowth\b|\bcustomers?\b|\bmarket\b|\binvest\w*|\bmargin\b|\bsales\b|\bpipeline\b|\bquarter\b|\bprofit\b"),
    "학술·연구": ("academic formal",
              r"연구|분석|가설|변수|모형|선행\s?연구|표본|유의|회귀|이론|방법론|논문|실증|통계|검정|문헌|함의|한계|결론|"
              r"독립변수|종속변수|통제변수|신뢰도|타당도|계수|추정|패널|설문|사례\s?연구|연구\s?문제|분석\s?틀|"
              r"이론적|관점|[가-힣]{2,}주의(?![가-힣])|개념|담론|고찰|시사점|"
              r"\bhypothes[ie]s\b|\bregression\b|\bstudy\b|\bliterature\b|\bfindings?\b|\bmethodology\b|\bsignifican\w*"),
    "교육·강의": ("friendly neutral",
              r"학습|수업|강의|학생|퀴즈|교육|과정|차시|학습\s?목표|복습|예제|실습|정리해\s?봅시다|알아봅시다|"
              r"생각해\s?봅시다|문제\s?풀이|오늘의|단원|교과|수강생|워크숍|토론|과제\s?제출|다음\s?시간"),
    "IT·스타트업": ("tech neutral",
                r"플랫폼|인공지능|데이터|앱|사용자|클라우드|솔루션|스타트업|개발|출시|베타|기능|"
                r"아키텍처|서버|모델|알고리즘|자동화|대시보드|로드맵|스프린트|리텐션|트래픽|오픈소스|"
                r"\blatency\b|\binference\b|\bbenchmarks?\b|\bmodels?\b|\bruntime\b|\bkernels?\b|\bGPU\b|"
                r"\bthroughput\b|\bscal(?:e|ing)\b|\bdeploy\w*|\bopen source\b|\bagents?\b|\bplatform\b"),
    "홍보·행사": ("impact playful",
              r"이벤트|행사|할인|혜택|축제|캠페인|홍보|오픈|모집|선착순|경품|초대|특가|런칭|기념|"
              r"지금\s?바로|놓치지|무료|선물|당첨|사전\s?예약|한정|공연\s?일정"),
    "문화·예술·인문": ("elegant calm traditional",
                  r"전시|예술|문화|공연|작가|역사|전통|감성|이야기|작품|시인|음악|미술|유산|기억|삶|여행|"
                  r"풍경|계절|문학|철학|사유|빛|기록|아카이브|큐레이션|사진전|개인전"),
    "아동·가족": ("cute playful friendly",
              r"어린이|아이(?!디어|템|콘|폰|패드|돌|디|스)|유아|놀이|동화|키즈|엄마|아빠|가족|우리\s?반|유치원|어린이집|색칠|그림책|간식|햇살반|친구들"),
    "게임·레트로": ("retro playful", r"게임|레트로|픽셀|레벨|플레이|스테이지|보스|퀘스트|8비트|아케이드|e스포츠|길드|레이드"),
}
# 대소문자를 구분하는 영문 약어(부분 일치 방지)
ACRONYM = {"기업·비즈니스": r"\b(?:KPI|ROI|IR|EBITDA|BEP|B2B|B2C|YoY|QoQ|ARR|MRR|CAC|LTV)\b",
           "IT·스타트업": r"\b(?:AI|API|MVP|UX|UI|SaaS|LLM|MAU|DAU|SDK|GitHub|DevOps|ML)\b"}
GENRE_DEFAULT = {
    "공공·정책": ("프리텐다드", "ExtraBold", "프리텐다드", "Regular"),
    "기업·비즈니스": ("프리텐다드", "ExtraBold", "프리텐다드", "Regular"),
    "IT·스타트업": ("Wanted Sans", "Black", "Wanted Sans", "Regular"),
    "학술·연구": ("본명조", "Bold", "프리텐다드", "Regular"),
    "교육·강의": ("페이퍼로지", "ExtraBold", "페이퍼로지", "Regular"),
    "홍보·행사": ("G마켓 산스", "Bold", "프리텐다드", "Regular"),
    "문화·예술·인문": ("마루 부리", "Bold", "마루 부리", "Regular"),
    "아동·가족": ("배민 주아체", "Regular", "나눔스퀘어라운드", "Regular"),
    "게임·레트로": ("갈무리9", "Regular", "프리텐다드", "Regular"),
}
# 공공 발행 덱에서도 이 분야 성격이 뚜렷하면 프리텐다드 대신 그 분야 서체를 쓴다
CONTENT_OVERRIDE = ["홍보·행사", "문화·예술·인문", "아동·가족", "게임·레트로", "학술·연구", "교육·강의"]
# 점수가 같을 때 분야별로 먼저 고를 폰트
PRIORITY = {g: [v[0], v[2]] for g, v in GENRE_DEFAULT.items()}
PRIORITY["교육·강의"] += ["프리젠테이션", "나눔스퀘어라운드"]
PRIORITY["IT·스타트업"] += ["IBM Plex Sans KR", "프리텐다드"]

PUBLIC_MARK = re.compile(r"^\s*([□■○◦❍ㅇ※◎▶◇◆]|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]\s*\.|[가-하]\.\s)")
ORG = re.compile(r"(?<![가-힣\d])(?:[가-힣]{2,10}(?:부|처|청)|[가-힣]{1,10}(?:공사|공단|재단|진흥원|연구원|위원회|센터|시청|도청|군청|구청|교육청|"
                 r"행정복지센터|의회))(?![가-힣])|(?:주무관|사무관|서기관|담당관|위원장|본부장|시장님|구청장|군수|도지사)")
ORG_STOP = {"지원", "신청", "요청", "출처", "대처", "일부", "전부", "내부", "외부", "세부", "본부", "민원", "간부", "정부"}
SIDO = r"(?:서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주|충청북|충청남|전라북|전라남|경상북|경상남)"
PUB_PUBLISHER = re.compile(
    SIDO + r"(?:특별시|광역시|특별자치시|특별자치도|도)?(?:청|교육청|의회)?(?![가-힣])|"          # 서울특별시, 경기도, 전남교육청
    r"[가-힣]{2,10}(?:부|처|청|공사|공단|재단|진흥원|연구원|위원회|의회|교육청|보훈청|센터)(?![가-힣])|"  # 행정안전부, OO교육청
    r"[가-힣]{2,4}[시군구](?=\s?[가-힣]{0,8}(?:과|국|실|팀|청|본부|센터|의회)(?![가-힣]))|"       # 수원시 OO과
    r"정부|Government|Ministry|City of|Metropolitan")
CITE = re.compile(r"\([가-힣A-Z][^()]{0,40}?(?:외|et al\.?)?,?\s?(?:19|20)\d{2}[a-z]?\)|참고\s?문헌|References|"
                  r"\bp\s?[<>=]\s?0?\.\d+|\bR\^?2\b|β|\bN\s?=\s?\d+")
MONEY = re.compile(r"\d[\d,.]*\s?(?:억|조|만|천)?\s?원|\$\s?\d|\d+(?:\.\d+)?\s?%|\d+\s?배")
AUDIENCE = re.compile(r"여러분|하세요|해\s?주세요|하십시오|보세요|오세요|드립니다|드려요|함께해요|참여하세요")

NOUN_YO = r"(?:필요|중요|주요|개요|수요|소요|강요|요요|동요|불필요)"
REGISTER = {
    "합니다체": (r"(습니다|합니다|입니다|됩니다|십시오|겠습니다|습니까)\s*[.!?]?\s*$", "formal"),
    "한다체": (r"(한다|된다|는다|ㄴ다|이다|있다|없다|았다|었다|였다|하였다)\s*\.?\s*$", "formal"),
    "해요체": (r"(?:[아어여해에예네군까나가래대게세봐와워줘돼지고며]요|죠)\s*[.!?~]*\s*$", "friendly"),
    "반말·구어": (r"(했어|했지|해봐|하자|거든|잖아|볼까|할래|이야|ㅋㅋ|ㅎㅎ)\s*[.!?~]*\s*$", "playful"),
    "개조식": (r"(함|임|음|됨|필요|추진|확대|강화|마련|검토|완료|예정|가능|구축|운영|개선|도입|시행|중요|주요)\s*[.)]?\s*$", "formal"),
}
EMOJI = re.compile("[\U0001F300-\U0001FAFF☀-➿⭐]")
SERIF_NAMES = re.compile(r"명조|바탕|부리|궁서|Myeongjo|Batang|Mincho|Times|Georgia|Garamond|Hahmlet|MaruBuri|Song ?Myung|(?<!Sans )Serif", re.I)
HAND_NAMES = re.compile(r"손글씨|\bPen\b|Brush|Script|Handwriting|Comic|Gaegu|Dokdo|Nanum ?Pen", re.I)


def X(z, name):
    return etree.fromstring(z.read(name))


def lum(hexv):
    r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def rels(z, part):
    d, f = part.rsplit("/", 1)
    rp = f"{d}/_rels/{f}.rels"
    if rp not in z.namelist():
        return {}
    out = {}
    for r in X(z, rp):
        tgt = r.get("Target", "")
        if r.get("TargetMode") == "External":
            continue
        parts = (d + "/" + tgt).split("/")
        stack = []
        for p in parts:
            if p == "..":
                stack.pop()
            elif p and p != ".":
                stack.append(p)
        out.setdefault(r.get("Type", "").rsplit("/", 1)[-1], []).append("/".join(stack))
    return out


def scheme_map(theme):
    m = {}
    clr = theme.find(".//a:clrScheme", NS) if theme is not None else None
    if clr is not None:
        for el in clr:
            c = el.find("a:srgbClr", NS); s = el.find("a:sysClr", NS)
            m[etree.QName(el).localname] = c.get("val") if c is not None else (s.get("lastClr") if s is not None else None)
    return m


def resolve(name, smap, clrmap):
    name = clrmap.get(name, name)
    return smap.get(name)


def bg_of(root, smap, clrmap):
    bg = root.find("p:cSld/p:bg", NS)
    if bg is None:
        return None
    if bg.find(".//a:blipFill", NS) is not None:
        return "image"
    c = bg.find(".//a:srgbClr", NS)
    if c is not None:
        return c.get("val")
    s = bg.find(".//a:schemeClr", NS)
    return resolve(s.get("val"), smap, clrmap) if s is not None else None


def extract(path):
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise SystemExit(f"'{path}'을(를) 열 수 없습니다. 암호가 걸렸거나 손상된 파일입니다.")
    names = set(z.namelist())
    pres = X(z, "ppt/presentation.xml")
    sz = pres.find("p:sldSz", NS)
    cx, cy = (int(sz.get("cx")), int(sz.get("cy"))) if sz is not None else (9144000, 6858000)  # 스키마 기본값 4:3
    # 슬라이드 순서는 sldIdLst를 따른다
    prel = {r.get("Id"): r.get("Target") for r in X(z, "ppt/_rels/presentation.xml.rels")}
    order = []
    for s in pres.findall("p:sldIdLst/p:sldId", NS):
        t = prel.get(s.get(f"{{{RNS}}}id"), "")
        t = "ppt/" + t.lstrip("/").replace("ppt/", "") if not t.startswith("/ppt/") else t.lstrip("/")
        if t in names:
            order.append(t)
    if not order:
        order = sorted([n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)], key=lambda n: int(re.findall(r"\d+", n)[-1]))

    F = {"slides": 0, "hidden": 0, "aspect": round(cx / cy, 2), "per": [], "lines_title": [], "lines_body": [], "lines_extra": [],
         "levels": Counter(), "sizes": [], "fonts": Counter(), "fills": [], "bgs": [], "ph_titles": 0, "box_titles": 0,
         "bold_runs": 0, "runs": 0, "tables": 0, "charts": 0, "smartart": 0, "pics": 0, "pic_area": 0.0, "notes": 0,
         "num_cells": 0, "cells": 0, "chars": set(), "footer": [], "company": "", "author": ""}
    cache = {}

    def chain(slide):
        """슬라이드 → 레이아웃 → 마스터 → 테마"""
        lay = (rels(z, slide).get("slideLayout") or [None])[0]
        mas = (rels(z, lay).get("slideMaster") or [None])[0] if lay else None
        th = (rels(z, mas).get("theme") or [None])[0] if mas else None
        return lay, mas, th

    for n in order:
        r = X(z, n)
        if r.get("show") == "0":
            F["hidden"] += 1; continue
        F["slides"] += 1
        lay, mas, th = chain(n)
        if th and th not in cache:
            tr = X(z, th); cache[th] = (scheme_map(tr), tr)
        smap, troot = cache.get(th, ({}, None))
        mr = X(z, mas) if mas else None
        cm = mr.find("p:clrMap", NS) if mr is not None else None
        clrmap = dict(cm.attrib) if cm is not None else {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}
        ov = r.find("p:clrMapOvr/a:overrideClrMapping", NS)
        if ov is not None:
            clrmap = dict(ov.attrib)
        bg = bg_of(r, smap, clrmap) or (bg_of(X(z, lay), smap, clrmap) if lay else None) or (bg_of(mr, smap, clrmap) if mr is not None else None) \
            or resolve("bg1", smap, clrmap) or "FFFFFF"
        F["bgs"].append(bg)
        if troot is not None:
            for tag in ("majorFont", "minorFont"):
                fs = troot.find(f".//a:{tag}", NS)
                if fs is not None:
                    for c in ("latin", "ea"):
                        e = fs.find(f"a:{c}", NS)
                        if e is not None and e.get("typeface"):
                            F["fonts"][e.get("typeface")] += 1

        title, body = [], []
        sizes_by_sp = []
        for sp in r.iter(f"{{{P}}}sp"):
            ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
            is_t = ph is not None and ph.get("type") in ("title", "ctrTitle")
            if ph is not None and ph.get("type") in ("ftr", "sldNum", "dt"):
                F["footer"].append("".join(t.text or "" for t in sp.iter(f"{{{A}}}t"))); continue
            z_ = [int(x.get("sz")) for x in sp.iter(f"{{{A}}}rPr") if x.get("sz")]
            if ph is None and z_:
                sizes_by_sp.append((max(z_), sp))
            if is_t:
                F["ph_titles"] += 1
            # 채운 도형 색(면적 가중)
            fill = sp.find("p:spPr/a:solidFill", NS)
            ext = sp.find("p:spPr/a:xfrm/a:ext", NS)
            if fill is not None and ext is not None:
                c = fill.find("a:srgbClr", NS); s = fill.find("a:schemeClr", NS)
                col = c.get("val") if c is not None else (resolve(s.get("val"), smap, clrmap) if s is not None else None)
                if col:
                    F["fills"].append((col, int(ext.get("cx", 0)) * int(ext.get("cy", 0)) / (cx * cy)))
            for p in sp.iter(f"{{{A}}}p"):
                t = "".join(x.text or "" for x in p.iter(f"{{{A}}}t")).strip()
                if not t:
                    continue
                (title if is_t else body).append(t)
                if not is_t:
                    ppr = p.find("a:pPr", NS)
                    F["levels"][int(ppr.get("lvl", 0)) if ppr is not None else 0] += 1
        # 자리표시자 제목이 없으면 가장 큰 글자(28pt 이상) 상자를 제목으로 본다
        if not any(sp.find("p:nvSpPr/p:nvPr/p:ph", NS) is not None and sp.find("p:nvSpPr/p:nvPr/p:ph", NS).get("type") in ("title", "ctrTitle")
                   for sp in r.iter(f"{{{P}}}sp")) and sizes_by_sp:
            m, sp = max(sizes_by_sp, key=lambda x: x[0])
            if m >= 2800:
                t = ["".join(x.text or "" for x in p.iter(f"{{{A}}}t")).strip() for p in sp.iter(f"{{{A}}}p")]
                t = [x for x in t if x]
                for x in t:
                    if x in body:
                        body.remove(x)
                title += t; F["box_titles"] += 1
        for gf in r.iter(f"{{{P}}}graphicFrame"):
            gd = gf.find(".//a:graphicData", NS)
            uri = gd.get("uri", "") if gd is not None else ""
            if "table" in uri:
                F["tables"] += 1
                for tc in gf.iter(f"{{{A}}}tc"):
                    cell = "".join(x.text or "" for x in tc.iter(f"{{{A}}}t")).strip()
                    if cell:
                        body.append(cell); F["cells"] += 1
                        F["num_cells"] += bool(re.fullmatch(r"[\d,.%\-+()원억만천명건\s]+", cell))
            elif "chart" in uri:
                F["charts"] += 1
            elif "diagram" in uri:
                F["smartart"] += 1
        for pic in r.iter(f"{{{P}}}pic"):
            F["pics"] += 1
            ext = pic.find(".//a:xfrm/a:ext", NS)
            if ext is not None:
                F["pic_area"] += int(ext.get("cx", 0)) * int(ext.get("cy", 0)) / (cx * cy)
        for run in r.iter(f"{{{A}}}r"):
            F["runs"] += 1
            rp = run.find("a:rPr", NS)
            if rp is not None and rp.get("b") == "1":
                F["bold_runs"] += 1
        for rpr in r.iter(f"{{{A}}}rPr", f"{{{A}}}defRPr"):
            if rpr.get("sz"):
                F["sizes"].append(int(rpr.get("sz")) / 100)
        for el in r.iter(f"{{{A}}}latin", f"{{{A}}}ea"):
            tf = el.get("typeface")
            if tf and not tf.startswith("+"):
                F["fonts"][tf] += 1
        body = [b for b in body if b]
        F["per"].append({"slide": len(F["per"]) + 1, "title": " / ".join(title)[:60], "chars": sum(map(len, title + body))})
        F["lines_title"] += title
        F["lines_body"] += body
        rr = rels(z, n)
        for key in ("notesSlide", "chart", "diagramData"):
            for part in rr.get(key, []):
                if part not in names:
                    continue
                pr = X(z, part)
                if key == "notesSlide":
                    txt = ""
                    for sp in pr.iter(f"{{{P}}}sp"):
                        ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
                        if ph is not None and ph.get("type") == "body":
                            txt += " ".join(t.text or "" for t in sp.iter(f"{{{A}}}t"))
                    if len(txt.strip()) > 10:
                        F["notes"] += 1; F["lines_extra"].append(txt.strip())
                else:
                    F["lines_extra"] += [t.text for t in pr.iter(f"{{{A}}}t") if t.text and t.text.strip()]
    for t in F["lines_title"] + F["lines_body"]:
        F["chars"].update(t)
    for part, key in (("docProps/app.xml", "company"), ("docProps/core.xml", "author")):
        if part in names:
            x = X(z, part)
            for el in x.iter():
                ln = etree.QName(el).localname
                if (key == "company" and ln == "Company") or (key == "author" and ln in ("creator", "lastModifiedBy")):
                    if el.text:
                        F[key] = (F[key] + " " + el.text).strip()
    return F


def color_mood(fills):
    """채운 도형 색의 면적 가중 chroma(max-min)와 밝기."""
    tot = c_sum = l_sum = 0.0
    for h, w in fills:
        if not h or len(h) != 6 or w <= 0:
            continue
        r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
        _, l, _ = colorsys.rgb_to_hls(r, g, b)
        if l > 0.97 or l < 0.04:
            continue
        w = min(w, 1.0)
        tot += w; c_sum += (max(r, g, b) - min(r, g, b)) * w; l_sum += l * w
    if tot < 0.02:
        return "판단할 색 없음(흰 바탕 위주)", ["neutral"], {}
    c, l = c_sum / tot, l_sum / tot
    stat = {"색강도": round(c, 2), "밝기": round(l, 2)}
    if c < 0.12:
        return "차분한 무채색 계열", ["formal", "neutral", "calm"], stat
    if l > 0.72 and c < 0.45:
        return "밝은 파스텔", ["cute", "friendly", "calm"], stat
    if l < 0.3:
        return "어둡고 묵직함", ["formal", "elegant", "tech"], stat
    if c > 0.55:
        return "선명하고 강한 색", ["impact", "playful"], stat
    return "보통 강도의 색", ["neutral", "friendly"], stat


def diagnose(F):
    main_text = "\n".join(F["lines_title"] + F["lines_body"])
    extra_text = "\n".join(F["lines_extra"])
    n = max(1, F["slides"])
    why = []

    han = len(re.findall(r"[가-힣]", main_text)); dig = len(re.findall(r"\d", main_text))
    hanja = len(re.findall(r"[一-鿿]", main_text))
    words = re.findall(r"[가-힣A-Za-z]+", main_text)
    en_words = sum(1 for w in words if re.fullmatch(r"[A-Za-z]+", w))
    scripts = {"한글(글자)": round(han / max(1, han + dig + hanja), 2), "영문(어절)": round(en_words / max(1, len(words)), 2),
               "숫자(글자)": round(dig / max(1, han + dig + hanja + en_words * 5), 2), "한자": hanja}

    marks = sum(1 for l in F["lines_body"] + F["lines_title"] if PUBLIC_MARK.search(l))
    orgs = sum(1 for m in ORG.finditer(main_text) if m.group(0) not in ORG_STOP)
    cites = len(CITE.findall(main_text)); money = len(MONEY.findall(main_text))
    aud = len(AUDIENCE.findall(main_text))
    excl, ques, emoji = main_text.count("!"), main_text.count("?"), len(EMOJI.findall(main_text))
    meta = " ".join([F["company"], F["author"], *F["footer"]])

    # 분야 점수: 단어마다 1+log(반복수)로 세어 한 단어가 점수를 독식하지 않게 한다. 노트·차트·SmartArt 글은 0.5배.
    def lex(text, rx, flags=re.I):
        c = Counter(m.group(0).lower() for m in re.finditer(rx, text, flags))
        return sum(1 + math.log(v) for v in c.values()), len(c)
    raw, distinct = {}, {}
    for g, (_, rx) in GENRES.items():
        s1, d1 = lex(main_text, rx); s2, _ = lex(extra_text, rx)
        if g in ACRONYM:
            s3, d3 = lex(main_text, ACRONYM[g], 0); s1 += s3; d1 += d3
        raw[g], distinct[g] = s1 + 0.5 * s2, d1
    raw["공공·정책"] += 1.5 * min(marks, 8) + 2.0 * min(orgs, 5) + (3 if PUB_PUBLISHER.search(meta) else 0)
    raw["학술·연구"] += 3.0 * min(cites, 5) + (2.0 if hanja > 20 else 0)
    raw["기업·비즈니스"] += 0.7 * min(money, 8)
    raw["홍보·행사"] += 0.8 * min(excl, 6) + 1.0 * min(emoji, 5)
    raw["아동·가족"] += 0.6 * min(emoji, 5)
    raw["교육·강의"] += 0.8 * min(ques, 5)
    gsum = sum(raw.values()) or 1
    top_raw = max(raw, key=raw.get) if raw else None
    # 공공 발행: 문서 속성·바닥글의 기관명, 공공 보고서 기호, 또는 (기관명·직함 + 공공·홍보가 1위)
    publisher_public = bool(PUB_PUBLISHER.search(meta)) or marks >= 3 or (orgs >= 2 and top_raw in ("공공·정책", "홍보·행사"))
    genres = sorted(((g, round(v / gsum, 2)) for g, v in raw.items() if v > 0), key=lambda x: (-x[1], list(GENRES).index(x[0])))
    top = genres[0] if genres else (None, 0)
    margin = top[1] - (genres[1][1] if len(genres) > 1 else 0)
    confident = bool(top[0]) and gsum >= 4 and distinct.get(top[0], 0) + (marks > 0) + (cites > 0) >= 3 \
        and top[1] >= 0.45 and margin >= 0.15
    if marks: why.append(f"공공 보고서 기호(□ ○ ※ Ⅰ. 등) {marks}줄")
    if orgs: why.append(f"기관명·직함 {orgs}회")
    if PUB_PUBLISHER.search(meta): why.append(f"문서 속성·바닥글에 공공기관({PUB_PUBLISHER.search(meta).group(0)})")
    if cites: why.append(f"학술 인용·통계 표기 {cites}회")
    if money: why.append(f"금액·비율 표기 {money}회")
    if hanja: why.append(f"한자 {hanja}자")
    if aud: why.append(f"청중에게 말 거는 표현 {aud}회")
    if F["lines_extra"]: why.append(f"노트·차트·SmartArt 글 {len(F['lines_extra'])}건 반영(0.5배)")

    reg = Counter(); lens = []
    for line in F["lines_body"]:
        line = EMOJI.sub("", line).strip()
        if len(line) < 6 or not re.search(r"[가-힣]", line):
            continue
        lens.append(len(line))
        if re.search(NOUN_YO + r"\s*[.)]?\s*$", line):
            reg["개조식"] += 1; continue
        for name, (rx, _) in REGISTER.items():
            if re.search(rx, line):
                reg[name] += 1; break
        else:
            reg["명사형(판단 보류)"] += 1
    rt = sum(reg.values()) or 1
    register = sorted(((k, round(v / rt, 2)) for k, v in reg.items()), key=lambda x: -x[1])

    chars = [p["chars"] for p in F["per"]] or [0]
    avg = round(statistics.mean(chars)); mx = max(chars)
    deck_type = "읽는 보고서형" if (avg > 250 or (mx > 500 and avg > 150)) else "발표형"
    if F["notes"] >= max(1, n // 3):
        why.append(f"발표자 노트 {F['notes']}장")
        if avg <= 350:
            deck_type = "발표형"
    sizes = sorted(F["sizes"])
    size_stat = {"최소": sizes[0], "중앙": statistics.median(sizes),
                 "14pt미만비율": round(sum(s < 14 for s in sizes) / len(sizes), 2)} if sizes else {}
    depth = max(F["levels"]) if F["levels"] else 0
    visual = F["tables"] + F["charts"] + F["smartart"]
    num_ratio = round(F["num_cells"] / F["cells"], 2) if F["cells"] else 0
    dense = bool(avg > 220 or visual >= max(2, n // 3) or (size_stat and size_stat["14pt미만비율"] > 0.4))
    bold_ratio = round(F["bold_runs"] / F["runs"], 2) if F["runs"] else 0

    mood, ctags, cstat = color_mood(F["fills"] + [(b, 1.0) for b in F["bgs"] if b and b != "image" and len(b) == 6 and lum(b) < 0.9])
    bgl = [lum(b) for b in F["bgs"] if b and b != "image" and len(b) == 6]
    dark = bool(bgl) and statistics.mean(bgl) < 0.35
    img_bg = sum(1 for b in F["bgs"] if b == "image")
    pic_area = round(F["pic_area"] / n, 2)
    image_led = pic_area > 0.45 and avg < 60
    fonts_now = dict(F["fonts"].most_common(8))
    tot_f = sum(F["fonts"].values()) or 1
    serif_now = sum(v for k, v in F["fonts"].items() if SERIF_NAMES.search(k)) / tot_f
    hand_now = sum(v for k, v in F["fonts"].items() if HAND_NAMES.search(k)) / tot_f

    tags = Counter()
    for g, w in genres[:3]:
        for t in GENRES[g][0].split():
            tags[t] += 3 * w
    for k, w in register:
        tags[(REGISTER.get(k) or (None, "neutral"))[1]] += 2 * w
    for t in ctags:
        tags[t] += 0.8
    if excl + emoji > max(3, n):
        tags["playful"] += 1
    if aud >= max(2, n // 2):
        tags["friendly"] += 0.8
    if dense:
        tags["neutral"] += 1
    if dark:
        tags["tech"] += 0.5; tags["elegant"] += 0.3
    if hanja > 20:
        tags["traditional"] += 0.5; tags["academic"] += 0.5
    if serif_now > 0.4:
        tags["elegant"] += 0.4; why.append("지금 덱이 명조 계열을 주로 씀")
    if hand_now > 0.3:
        tags["handwritten"] += 0.4
    if image_led:
        tags["elegant"] += 0.5; tags["calm"] += 0.5; why.append(f"사진 위주(장당 이미지 면적 {pic_area})")

    advice = []
    if dark:
        advice.append("어두운 배경: 본문은 Regular~Medium, Light 금지")
    if size_stat and size_stat["최소"] < 12:
        advice.append(f"가장 작은 글자 {size_stat['최소']}pt: 본문은 글자 폭 넉넉한 고딕, 명조·손글씨 본문 피함")
    advice.append("읽는 보고서형: 제목은 Bold~ExtraBold로 절제, 본문은 Regular" if deck_type == "읽는 보고서형"
                  else "발표형: 제목은 ExtraBold~Black으로 크게, 본문은 Regular~Medium")
    if F["aspect"] < 1.5:
        advice.append("4:3 화면: 폭이 좁은 고딕(SUIT·프리텐다드)이 줄바꿈에 유리")
    if bold_ratio > 0.15:
        advice.append(f"굵은 글자 비율 {int(bold_ratio * 100)}%: 본문 가족의 Bold도 반드시 받기")
    if num_ratio > 0.4:
        advice.append("숫자 표가 많음: 숫자 폭이 고른 고딕(프리텐다드·SUIT)이 유리")
    if F["box_titles"] and F["ph_titles"] == 0:
        advice.append("제목이 자리표시자가 아닌 텍스트 상자: apply_fonts가 가장 큰 글자 상자를 제목으로 처리함(결과 이미지로 확인)")

    special = "".join(sorted(c for c in F["chars"] if not re.match(r"[가-힣A-Za-z0-9\s.,:;!?'\"()\[\]/\-+%&@#*=_~<>{}|\\^$]", c) and not EMOJI.match(c)))
    rare_hangul = "".join(sorted(c for c in F["chars"] if "가" <= c <= "힣" and not _ksx(c)))

    return {
        "slides": F["slides"], "hidden": F["hidden"], "aspect": F["aspect"], "scripts": scripts,
        "publisher_public": publisher_public, "publisher_meta": meta.strip()[:80],
        "genres": genres[:4], "genre_confident": confident, "register": register,
        "genre_raw": {g: round(v, 2) for g, v in raw.items()}, "genre_distinct": distinct,
        "meta_text": meta, "body_text": main_text[:20000],
        "sentence_len": round(statistics.mean(lens)) if lens else None,
        "signals": {"공공기호": marks, "기관직함": orgs, "학술인용": cites, "금액비율": money, "청중호칭": aud,
                    "느낌표": excl, "물음표": ques, "이모지": emoji},
        "structure": {"덱유형": deck_type, "슬라이드당평균글자": avg, "최다글자": mx, "글머리깊이": depth,
                      "자리표시자제목": F["ph_titles"], "상자제목": F["box_titles"], "글자크기": size_stat, "굵은글자비율": bold_ratio,
                      "표": F["tables"], "숫자셀비율": num_ratio, "차트": F["charts"], "SmartArt": F["smartart"],
                      "이미지": F["pics"], "이미지면적(장당)": pic_area, "노트": F["notes"], "빽빽함": dense},
        "visual": {"색감": mood, **cstat, "어두운배경": dark, "이미지배경": img_bg},
        "fonts_now": fonts_now, "special_chars": special, "rare_hangul": rare_hangul,
        "tone_tags": [t for t, _ in tags.most_common(4)],
        "tone_scores": {t: round(v, 2) for t, v in tags.most_common()}, "advice": advice, "why": why,
        "per_slide": F["per"],
    }


def _ksx(ch):
    try:
        ch.encode("euc_kr"); return True
    except UnicodeEncodeError:
        return False


def load(name, data_dir):
    if data_dir and (Path(data_dir) / name).exists():
        return json.loads((Path(data_dir) / name).read_text("utf-8"))
    with urllib.request.urlopen(REPO_RAW + name, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def recommend(d, data_dir=None, k=3):
    files = load("fontfiles.json", data_dir)["fonts"]
    google = load("google-fonts.json", data_dir)["fonts"]
    try:
        cov = load("glyph-coverage.json", data_dir)["fonts"]
    except Exception:
        cov = {}
    norm = lambda s: (s or "").replace(" ", "").lower()
    pool, seen = [], set()
    for fid, f in files.items():
        pool.append({"key": f["name"], "family": f["family"], "src": "산돌 OFL", "style": f["style"], "use": f["use"],
                     "tags": f.get("tags", []), "weights": list(f["files"]), "latinOnly": f.get("latinOnly", False), "cov": cov.get(fid, {})})
        seen |= {norm(f["family"]), norm(f.get("google"))}
    for g in google:
        if norm(g["family"]) in seen or not g.get("redistributable"):
            continue
        pool.append({"key": g["name"], "family": g["family"], "src": "Google", "style": g["style"], "use": g["use"],
                     "tags": g.get("tags", []), "weights": g["weights"], "latinOnly": False, "cov": cov.get("G:" + g["family"], {})})
    by = {f["key"]: f for f in pool}
    ts, dense = d["tone_scores"], d["structure"]["빽빽함"]
    latin_heavy = d["scripts"]["영문(어절)"] > 0.6
    small = bool(d["structure"]["글자크기"]) and d["structure"]["글자크기"]["최소"] < 12
    report_type = d["structure"]["덱유형"] == "읽는 보고서형"
    special, rare = d["special_chars"], d["rare_hangul"]
    g0 = d["genres"][0][0] if d["genres"] else None
    prio = PRIORITY.get(g0, ["프리텐다드"])

    def glyph_gap(f):
        c = f["cov"]
        miss = "".join(ch for ch in special if ch in c.get("missing", ""))
        hangul_gap = bool(rare) and c.get("hangul", 11172) < 11172
        return miss, hangul_gap

    def score(f, role):
        if (f["latinOnly"] and not latin_heavy) or role not in f["use"]:
            return -9
        s = sum(ts.get(t, 0) for t in f["tags"])
        miss, hg = glyph_gap(f)
        s -= 0.6 * len(miss) + (3 if hg else 0)
        if role == "body":
            if f["style"] in ("display", "handwriting", "pixel", "mono"):
                return -9
            s += 0.3 * min(len(f["weights"]), 6) + (1.0 if dense and "neutral" in f["tags"] else 0)
            if small and f["style"] == "serif":
                s -= 1.5
        else:
            s += 0.4 if {"ExtraBold", "Black", "Heavy", "Bold"} & set(f["weights"]) else 0
        return round(s, 2)

    def rank_key(role):
        return lambda f: (-score(f, role), prio.index(f["key"]) if f["key"] in prio else 99)

    def heavy(f):
        if f["style"] == "serif":
            order = ("Bold", "SemiBold", "ExtraBold")
        elif report_type:
            order = ("Bold", "ExtraBold", "SemiBold")
        else:
            order = ("ExtraBold", "Black", "Heavy", "Bold")
        return next((w for w in order if w in f["weights"]), f["weights"][-1])

    def note(*fs):
        out = []
        for f in fs:
            miss, hg = glyph_gap(f)
            if miss: out.append(f"{f['key']}에 {miss} 없음")
            if hg: out.append(f"{f['key']}는 드문 한글({rare[:5]}) 없음")
        return "; ".join(out)

    pairs = []

    def add(kind, t, tw, b, bw):
        pairs.append({"방식": kind, "제목": f"{t['key']} {tw}", "본문": f"{b['key']} {bw}",
                      "출처": t["src"] if t is b else f"{t['src']} + {b['src']}",
                      "점수": round(score(t, "title") + score(b, "body"), 2), "주의": note(t) if t is b else note(t, b)})

    def genre_pair(g, kind):
        tn, tw, bn, bw = GENRE_DEFAULT[g]
        if g == "학술·연구" and not dense and not small:
            bn = "본명조"
        t, b = by[tn], by[bn]
        if report_type and t["style"] != "serif" and tw in ("Black", "ExtraBold"):
            tw = "Bold"
        add(kind, t, tw, b, bw)

    raw, dist = d["genre_raw"], d["genre_distinct"]
    # 공공 발행 덱에서도 내용 성격이 뚜렷하면(축제·전시·아동·학술 등) 그 서체를 쓴다.
    # 기준: 그 분야 점수가 공공 점수의 절반 이상이고, 서로 다른 단서 단어가 3개 이상.
    content = None
    if d["publisher_public"] or g0 == "공공·정책":
        cands = [(g, raw[g]) for g in CONTENT_OVERRIDE if raw.get(g, 0) >= max(3.0, 0.5 * raw.get("공공·정책", 0)) and dist.get(g, 0) >= 3]
        if cands:
            content = max(cands, key=lambda x: x[1])[0]
    pub = by["프리텐다드"]
    if content:
        genre_pair(content, f"공공 발행 + 내용 우선({content})")
        add("공공 기본값", pub, heavy(pub), pub, "Regular")
    elif d["publisher_public"] or (g0 == "공공·정책" and d["genres"][0][1] >= 0.3):
        add("공공 기본값", pub, heavy(pub), pub, "Regular")
    elif g0 and d["genre_confident"] and g0 in GENRE_DEFAULT:
        genre_pair(g0, f"분야 기본값({g0})")
    else:
        add("판단 불확실: 무난한 기본값", pub, heavy(pub), pub, "Regular")

    titles = sorted(pool, key=rank_key("title"))
    bodies = sorted(pool, key=rank_key("body"))
    used = {p["제목"].rsplit(" ", 1)[0] for p in pairs}
    for f in bodies:
        if len(pairs) >= 3:
            break
        if f["key"] in used or "title" not in f["use"] or score(f, "body") <= -9 or len(f["weights"]) < 3:
            continue
        add("한 가족", f, heavy(f), f, "Regular"); used.add(f["key"])
    for t in titles[:10]:
        if len(pairs) >= k + 1:
            break
        if t["key"] in used or score(t, "title") <= -9:
            continue
        b = next(b for b in bodies if b["family"] != t["family"])
        # 제목과 본문이 비슷한 고딕 둘이면 '개성' 짝이 아니다 → 건너뜀
        if t["style"] == b["style"] == "sans" and set(t["tags"]) & {"neutral", "formal"} and set(b["tags"]) & {"neutral", "formal"}:
            continue
        tw = heavy(t) if len(t["weights"]) > 1 else t["weights"][0]
        add("제목 개성 + 본문 안정", t, tw, b, "Regular"); used.add(t["key"])
    # 기관 전용 서체
    agencies = []
    try:
        for ag in load("agency-fonts.json", data_dir)["agencies"]:
            hits = 3 * len(re.findall(ag["match"], d["meta_text"])) + len(re.findall(ag["match"], d["body_text"]))
            if hits >= 2:
                agencies.append({**ag, "hits": hits})
    except Exception:
        pass
    agencies.sort(key=lambda a: -a["hits"])
    return {"pairs": pairs[:k + 1], "agency_fonts": agencies[:2], "content_override": content, "title_top": [(f["key"], score(f, "title")) for f in titles[:6]],
            "body_top": [(f["key"], score(f, "body")) for f in bodies[:6]]}


def report(d, rec, path):
    s, v = d["structure"], d["visual"]
    L = [f"# 덱 진단: {Path(path).name}",
         f"- 슬라이드 {d['slides']}장(숨김 {d['hidden']}) · 화면비 {d['aspect']} · 글자 구성 {d['scripts']}",
         f"- 발행 주체: {'공공기관으로 보임' if d['publisher_public'] else '공공 신호 없음'}" + (f" (문서 속성·바닥글: {d['publisher_meta']})" if d['publisher_meta'] else ""),
         "- 분야: " + (", ".join(f"{g}({int(w * 100)}%)" for g, w in d["genres"]) or "단서 부족")
         + (" · 판단 확실" if d["genre_confident"] else " · 판단 불확실(이미지·노트를 보고 보정할 것)"),
         "- 글의 문체: " + (", ".join(f"{k}({int(w * 100)}%)" for k, w in d["register"]) or "문장 단서 부족")
         + (f" · 평균 문장 {d['sentence_len']}자" if d["sentence_len"] else "")
         + f" · 청중 호칭 {d['signals']['청중호칭']} · 느낌표 {d['signals']['느낌표']} · 물음표 {d['signals']['물음표']} · 이모지 {d['signals']['이모지']}",
         f"- 구조: {s['덱유형']}, 슬라이드당 평균 {s['슬라이드당평균글자']}자(최다 {s['최다글자']}자), 제목 자리표시자 {s['자리표시자제목']}장·상자 제목 {s['상자제목']}장, "
         f"글머리 깊이 {s['글머리깊이']}, 굵은 글자 {int(s['굵은글자비율'] * 100)}%",
         f"- 시각 요소: 표 {s['표']}(숫자 셀 {int(s['숫자셀비율'] * 100)}%)·차트 {s['차트']}·SmartArt {s['SmartArt']}·이미지 {s['이미지']}(장당 면적 {s['이미지면적(장당)']}), 노트 {s['노트']}장"
         + (" → 본문은 가독성 최우선" if s["빽빽함"] else ""),
         f"- 글자 크기: {s['글자크기'] or '지정 없음(템플릿 기본)'}",
         f"- 색: {v['색감']} {({k: v[k] for k in ('색강도', '밝기') if k in v}) or ''}" + (" · 어두운 배경" if v["어두운배경"] else "")
         + (f" · 이미지 배경 {v['이미지배경']}장" if v["이미지배경"] else ""),
         f"- 지금 글꼴: {d['fonts_now']}",
         f"- 특수 기호: {d['special_chars'] or '없음'}" + (f" · 드문 한글: {d['rare_hangul'][:20]}" if d["rare_hangul"] else ""),
         "- 판단 근거: " + ("; ".join(d["why"]) or "어휘만으로 판단"),
         f"\n## 어울리는 서체 성격: {', '.join(d['tone_tags'])}  {d['tone_scores']}",
         "## 굵기·크기 조언: " + " / ".join(d["advice"]),
         "\n## 폰트 짝 후보"]
    for i, p in enumerate(rec["pairs"], 1):
        L.append(f"{i}. [{p['방식']}] 제목 {p['제목']} / 본문 {p['본문']}  ({p['출처']}, 점수 {p['점수']})" + (f"  ⚠ {p['주의']}" if p["주의"] else ""))
    for ag in rec.get("agency_fonts", []):
        L.append(f"\n## 기관 전용 서체 후보: {ag['agency']} → {', '.join(ag['fonts'])}\n   {ag['use']} · 라이선스: {ag['license']} · 공식: {ag['url']}"
                 "\n   (자동으로 받지 않음. 사용자가 받아 주면 fetch_fonts.py --local 파일로 적용)")
    L.append(f"\n제목 상위: {rec['title_top']}\n본문 상위: {rec['body_top']}")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx"); ap.add_argument("--json"); ap.add_argument("--data", help="data 폴더(오프라인)")
    a = ap.parse_args()
    d = diagnose(extract(a.pptx))
    rec = recommend(d, a.data)
    print(report(d, rec, a.pptx))
    if a.json:
        Path(a.json).write_text(json.dumps({"analysis": d, "recommendation": rec}, ensure_ascii=False, indent=2), "utf-8")


if __name__ == "__main__":
    main()
