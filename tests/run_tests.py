#!/usr/bin/env python3
"""스킬 회귀 시험: 성격이 다른 PPTX를 만들어 analyze_deck.py의 1순위 폰트와 경고를 검사한다.

  python tests/run_tests.py          # 실패가 있으면 종료 코드 1
렌더링(LibreOffice) 없이 분석만 하므로 CI에서 몇 초 안에 끝난다.
"""
import importlib.util, sys, tempfile
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("ad", ROOT / "skill/scripts/analyze_deck.py")
ad = importlib.util.module_from_spec(spec); spec.loader.exec_module(ad)
DATA = str(ROOT / "data")

# (이름, 작성자, 슬라이드[(제목, [줄...])], 옵션, 기대 1순위 제목 폰트 집합, 기대 방식 일부 문자열)
CASES = [
    ("공공 보고", "행정안전부 지방재정과", [("Ⅰ. 추진 배경", ["□ 지방재정 여건 악화로 세출 구조조정 필요", "  ○ 의무지출 비중 68%", "※ 관계 부서 협의"]),
                                         ("Ⅱ. 추진 계획", ["□ 성과평가 결과에 따른 예산 조정", "□ 향후 계획: 국회 보고"])], {}, {"프리텐다드"}, "공공"),
    ("공공 축제 홍보", "경기도 문화체육관광국", [("경기 전통문화 축제에 초대합니다!", ["가족과 함께 즐기는 공연과 체험 행사", "선착순 500명 기념품 증정!", "지금 바로 사전 예약하세요"]),
                                          ("행사 안내", ["무료 입장, 누구나 참여 가능", "장소: 경기도청 광장"])], {}, {"G마켓 산스"}, "내용 우선"),
    ("현충일 추념식", "서울지방보훈청", [("제71회 현충일 추념식", ["나라를 위해 희생하신 분들을 기억합니다"]),
                                    ("식순", ["개식", "국민의례", "헌화 및 분향", "추념사", "폐식"])], {}, {"본명조"}, "기념·의전"),
    ("교육청 어린이 안전", "서울특별시교육청 학생안전과", [("어린이 교통안전, 이렇게 지켜요", ["횡단보도에서는 손을 들고 건너요", "등하굣길에는 보호자와 함께해요"]),
                                                ("우리 반 안전 약속", ["친구들과 놀이할 때는 차도에서 멀리 떨어져요"])], {}, {"배민 주아체", "동글"}, "아동"),
    ("학술 발표", "", [("연구 문제와 가설", ["선행연구(김철수, 2021; Lee et al., 2019)를 검토함", "가설: 의석 균형은 교착 확률을 높인다"]),
                     ("분석 결과", ["로지스틱 회귀모형 추정 결과 β=0.82, p<.01", "통제변수: 의원 수, 정당 수"])], {}, {"본명조", "함렛"}, "학술"),
    ("기업 IR", "에이크미 플랫폼사업부", [("FY2026 실적 요약", ["매출 1,820억원 (YoY +34%)", "영업이익률 12.4%", "B2B 고객 수 2배 성장"]),
                                     ("성장 전략", ["해외 시장 점유율 확대", "시리즈 C 투자 유치"])], {}, {"프리텐다드"}, "기업"),
    ("IT 피치(사업부 작성)", "에이크미 플랫폼사업부", [("AI Agent Platform", ["LLM 기반 자동화 솔루션", "API로 기존 서비스와 연동", "MAU 12만, 리텐션 41%"]),
                                                ("Architecture", ["클라우드 네이티브 서버", "오픈소스 모델과 대시보드"])], {"bg": "111827"}, {"Wanted Sans", "IBM Plex Sans KR"}, "IT"),
    ("전시", "", [("빛과 기억의 공간", ["작가 이OO 개인전", "계절의 풍경을 기록한 작품 40점"]),
                 ("전시 이야기", ["시간이 머문 자리, 삶의 흔적을 사유하는 전시"])], {}, {"마루 부리", "고운바탕"}, "문화"),
    ("게임", "", [("길드 대항전 시즌 3", ["레벨 50 이상 플레이어 참여", "보스 레이드 스테이지 신규 추가"]),
                 ("보상 안내", ["레전더리 아이템 상자", "8비트 레트로 스킨"])], {"bg": "1A1A2E"}, {"갈무리9"}, "게임"),
    ("단서 부족", "", [("회의 안건", ["지난 회의 내용 확인", "다음 일정"])], {}, {"프리텐다드"}, "불확실"),
    ("세로 A4", "", [("주민참여예산 안내", ["□ 제안 접수: 9월 30일까지", "□ 문의: 예산과"])], {"portrait": True}, {"프리텐다드"}, ""),
]
CHECKS = [  # (이름, 덱 옵션, 검사 함수, 설명)
    ("드문 한글 감지", [("똠방각하 뷁 쀍", ["햏 똠 쀍 시험"])], lambda d, r: "똠" in d["rare_hangul"], "KS X 1001 밖 한글을 찾아야 함"),
    ("세로 문서 조언", [("안내", ["본문"])], lambda d, r: any("세로" in a for a in d["advice"]), "세로형이면 4:3 조언 대신 세로 문서 조언"),
]


def make(path, author, slides, opt):
    p = Presentation()
    if opt.get("portrait"):
        p.slide_width, p.slide_height = Inches(8.27), Inches(11.69)
    else:
        p.slide_width, p.slide_height = Inches(13.333), Inches(7.5)
    p.core_properties.author = author or "tester"
    for t, lines in slides:
        s = p.slides.add_slide(p.slide_layouts[1]); s.shapes.title.text = t
        if opt.get("bg"):
            s.background.fill.solid(); s.background.fill.fore_color.rgb = RGBColor.from_string(opt["bg"])
        tf = s.placeholders[1].text_frame
        for j, l in enumerate(lines):
            para = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            if l.startswith("  "):
                para.level = 1
            para.text = l.strip()
    p.save(path)


def main():
    fails = 0
    with tempfile.TemporaryDirectory() as td:
        for name, author, slides, opt, want, kind in CASES:
            f = Path(td) / "d.pptx"; make(f, author, slides, opt)
            d = ad.diagnose(ad.extract(str(f))); r = ad.recommend(d, DATA)
            first = r["pairs"][0]
            t = first["제목"].rsplit(" ", 1)[0]
            ok = t in want and (kind in first["방식"] or kind in (d["genres"][0][0] if d["genres"] else "") or not kind
                                or (kind == "불확실" and not d["genre_confident"]))
            fails += not ok
            print(f"{'✓' if ok else '✗'} {name:<18} → {first['방식']}: {first['제목']} / {first['본문']}" + ("" if ok else f"   (기대: {want}, {kind})"))
        for name, slides, fn, desc in CHECKS:
            f = Path(td) / "c.pptx"; make(f, "", slides, {"portrait": "세로" in name})
            d = ad.diagnose(ad.extract(str(f))); r = ad.recommend(d, DATA)
            ok = fn(d, r); fails += not ok
            print(f"{'✓' if ok else '✗'} {name:<18} — {desc}")
    print(f"\n실패 {fails}건")
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
