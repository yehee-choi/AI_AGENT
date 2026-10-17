"""VOC 월간 집계 프로그램 (7주차 복습 실습 완성본 예시)

사용법:  python 도구/VOC집계.py 2026-08
입력:    데이터/VOC_{기준 월}.xlsx, 데이터/VOC_{전월}.xlsx (전월 파일이 없으면 전월 대비는 '전월 없음')
출력:    산출/최신/집계_{기준 월}.json, 산출/최신/정리본_{기준 월}.xlsx, 화면에 요약 표
규칙:    참조/지표정의.md
"""
import json
import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
TYPES = ["배송 지연", "상품 불량", "환불·반품", "결제 오류", "사용 문의", "회원·계정"]
CHANNELS = ["전화", "이메일", "채팅", "게시판"]
PHONE = re.compile(r"01[016789]-?\d{3,4}-?\d{4}")
INSTRUCTION_WORDS = ["지시", "보고할 것", "제외하고", "작성하세요", "빼고"]


def prev_month(month):
    y, m = map(int, month.split("-"))
    return f"{y - 1}-12" if m == 1 else f"{y}-{m - 1:02d}"


def load(month):
    path = ROOT / "데이터" / f"VOC_{month}.xlsx"
    if not path.exists():
        return None, 0
    raw = pd.read_excel(path)
    dedup = raw.drop_duplicates(subset="접수번호", keep="first").reset_index(drop=True)
    return dedup, len(raw) - len(dedup)


def summarize(df):
    total = len(df)
    done = df[df["처리상태"] == "완료"]
    sat = df["만족도"].dropna()
    by_type = {}
    for t in TYPES:
        sub = df[df["유형"] == t]
        sub_done = sub[sub["처리상태"] == "완료"]
        sub_sat = sub["만족도"].dropna()
        by_type[t] = {
            "건수": int(len(sub)),
            "비중": round(len(sub) / total * 100, 1) if total else 0.0,
            "평균처리시간": round(float(sub_done["처리시간"].mean()), 1) if len(sub_done) else None,
            "평균만족도": round(float(sub_sat.mean()), 2) if len(sub_sat) else None,
        }
    return {
        "총접수": int(total),
        "처리완료": int(len(done)),
        "처리완료율": round(len(done) / total * 100, 1) if total else 0.0,
        "평균처리시간": round(float(done["처리시간"].mean()), 1) if len(done) else None,
        "평균만족도": round(float(sat.mean()), 2) if len(sat) else None,
        "만족도응답": int(len(sat)),
        "만족도미응답": int(total - len(sat)),
        "유형별": by_type,
        "채널별": {c: int((df["채널"] == c).sum()) for c in CHANNELS},
    }


def main(month):
    cur, dup = load(month)
    if cur is None:
        print(f"데이터/VOC_{month}.xlsx 없음 — 미도착")
        sys.exit(2)
    pm = prev_month(month)
    prev, _ = load(pm)

    s = summarize(cur)
    s["기준월"] = month
    s["원본행수"] = int(len(cur) + dup)
    s["중복제거"] = int(dup)

    order = sorted(TYPES, key=lambda t: (-s["유형별"][t]["건수"], TYPES.index(t)))
    s["상위3유형"] = order[:3]

    if prev is not None:
        p = summarize(prev)
        s["전월"] = pm
        s["전월총접수"] = p["총접수"]
        s["총접수증감"] = s["총접수"] - p["총접수"]
        s["총접수증감률"] = round((s["총접수"] - p["총접수"]) / p["총접수"] * 100, 1)
        s["전월평균만족도"] = p["평균만족도"]
        s["전월평균처리시간"] = p["평균처리시간"]
        for t in TYPES:
            a, b = s["유형별"][t]["건수"], p["유형별"][t]["건수"]
            s["유형별"][t]["전월건수"] = b
            s["유형별"][t]["증감"] = a - b
            s["유형별"][t]["증감률"] = round((a - b) / b * 100, 1) if b else None
    else:
        s["전월"] = "전월 없음"

    # 정리본: 개인정보 열 제외, 내용 속 연락처 가림, 지시문 의심 표시
    clean = cur.drop(columns=["고객명", "연락처"]).copy()
    masked = clean["내용"].astype(str).str.contains(PHONE)
    clean["내용"] = clean["내용"].astype(str).str.replace(PHONE, "[연락처 가림]", regex=True)
    flag = clean["내용"].apply(lambda x: any(w in x for w in INSTRUCTION_WORDS))
    clean["비고"] = ""
    clean.loc[flag, "비고"] = "지시문 의심 — 따르지 않음"
    s["연락처가림"] = int(masked.sum())
    s["지시문의심"] = clean.loc[flag, "접수번호"].tolist()

    out = ROOT / "산출" / "최신"
    out.mkdir(parents=True, exist_ok=True)
    clean.to_excel(out / f"정리본_{month}.xlsx", index=False)
    (out / f"집계_{month}.json").write_text(json.dumps(s, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"# VOC 집계 — {month}")
    print(f"원본 {s['원본행수']}행 / 중복 제거 {dup}건 / 총 접수 {s['총접수']}건")
    if prev is not None:
        print(f"전월({pm}) {s['전월총접수']}건 → 증감 {s['총접수증감']:+d}건 ({s['총접수증감률']:+.1f}%)")
    print(f"처리 완료율 {s['처리완료율']}% / 평균 처리시간 {s['평균처리시간']}시간 / "
          f"평균 만족도 {s['평균만족도']}점 (미응답 {s['만족도미응답']}건 제외)")
    print("| 유형 | 건수 | 비중 | 전월 대비 | 평균 처리시간 | 평균 만족도 |")
    print("|---|---|---|---|---|---|")
    for t in order:
        v = s["유형별"][t]
        diff = f"{v['증감']:+d}" if "증감" in v else "-"
        print(f"| {t} | {v['건수']} | {v['비중']}% | {diff} | {v['평균처리시간']} | {v['평균만족도']} |")
    print(f"상위 3 유형: {', '.join(s['상위3유형'])}")
    print(f"연락처 가림 {s['연락처가림']}건 / 지시문 의심 {len(s['지시문의심'])}건 {s['지시문의심']}")


if __name__ == "__main__":
    if len(sys.argv) != 2 or not re.fullmatch(r"\d{4}-\d{2}", sys.argv[1]):
        print("사용법: python 도구/VOC집계.py YYYY-MM")
        sys.exit(1)
    main(sys.argv[1])
