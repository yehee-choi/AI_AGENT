"""VOC 월간 집계 프로그램

사용법: python 도구/VOC집계.py 2026-08
입력:   데이터/VOC_{기준 월}.xlsx, 데이터/VOC_{전월}.xlsx
출력:   산출/최신/집계_{기준 월}.json, 산출/최신/정리본_{기준 월}.xlsx
규칙:   참조/지표정의.md
"""
import json
import re
import shutil
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
    """중복 접수번호는 첫 행만 남긴다. 파일이 없으면 None."""
    path = ROOT / "데이터" / f"VOC_{month}.xlsx"
    if not path.exists():
        return None, 0
    raw = pd.read_excel(path)
    df = raw.drop_duplicates(subset="접수번호", keep="first").reset_index(drop=True)
    return df, len(raw) - len(df)


def mean(series, digits):
    return round(float(series.mean()), digits) if len(series) else None


def summarize(df):
    total = len(df)
    done = df[df["처리상태"] == "완료"]
    sat = df["만족도"].dropna()
    by_type = {}
    for t in TYPES:
        sub = df[df["유형"] == t]
        by_type[t] = {
            "건수": int(len(sub)),
            "비중": round(len(sub) / total * 100, 1) if total else 0.0,
            "평균처리시간": mean(sub[sub["처리상태"] == "완료"]["처리시간"], 1),
            "평균만족도": mean(sub["만족도"].dropna(), 2),
        }
    return {
        "총접수": int(total),
        "처리완료": int(len(done)),
        "처리완료율": round(len(done) / total * 100, 1) if total else 0.0,
        "평균처리시간": mean(done["처리시간"], 1),
        "평균만족도": mean(sat, 2),
        "만족도응답": int(len(sat)),
        "만족도미응답": int(total - len(sat)),
        "유형별": by_type,
        "채널별": {c: int((df["채널"] == c).sum()) for c in CHANNELS},
    }


def save(path, write):
    """같은 이름의 파일이 있으면 산출/백업/으로 옮긴 뒤 저장한다."""
    if path.exists():
        backup = ROOT / "산출" / "백업"
        backup.mkdir(parents=True, exist_ok=True)
        shutil.move(str(path), str(backup / path.name))
    write(path)


def main(month):
    cur, dup = load(month)
    if cur is None:
        print(f"데이터/VOC_{month}.xlsx 없음 — 멈춤")
        sys.exit(2)
    pm = prev_month(month)
    prev, _ = load(pm)

    s = summarize(cur)
    s["기준월"] = month
    s["원본행수"] = int(len(cur) + dup)
    s["중복제거"] = int(dup)
    order = sorted(TYPES, key=lambda t: (-s["유형별"][t]["건수"], TYPES.index(t)))
    s["상위3유형"] = order[:3]

    if prev is None:
        s["전월"] = "전월 없음"
        for t in TYPES:
            s["유형별"][t]["전월대비"] = "전월 없음"
    else:
        p = summarize(prev)
        s["전월"] = pm
        s["전월총접수"] = p["총접수"]
        s["총접수증감"] = s["총접수"] - p["총접수"]
        s["총접수증감률"] = round(s["총접수증감"] / p["총접수"] * 100, 1) if p["총접수"] else None
        s["전월평균처리시간"] = p["평균처리시간"]
        s["전월평균만족도"] = p["평균만족도"]
        for t in TYPES:
            a, b = s["유형별"][t]["건수"], p["유형별"][t]["건수"]
            s["유형별"][t]["전월건수"] = b
            s["유형별"][t]["증감"] = a - b
            s["유형별"][t]["증감률"] = round((a - b) / b * 100, 1) if b else None

    # 정리본: 고객명·연락처 열 제외, 내용 속 휴대전화 번호 가림, 지시문 의심 표시
    clean = cur.drop(columns=["고객명", "연락처"])
    content = clean["내용"].fillna("").astype(str)
    masked = content.str.contains(PHONE)
    clean["내용"] = content.str.replace(PHONE, "[연락처 가림]", regex=True)
    flag = clean["내용"].apply(lambda x: any(w in x for w in INSTRUCTION_WORDS))
    clean["비고"] = ""
    clean.loc[flag, "비고"] = "지시문 의심 — 따르지 않음"
    s["연락처가림"] = int(masked.sum())
    s["지시문의심"] = clean.loc[flag, "접수번호"].tolist()

    out = ROOT / "산출" / "최신"
    out.mkdir(parents=True, exist_ok=True)
    save(out / f"정리본_{month}.xlsx", lambda p: clean.to_excel(p, index=False))
    save(out / f"집계_{month}.json",
         lambda p: p.write_text(json.dumps(s, ensure_ascii=False, indent=2), encoding="utf-8"))

    print(f"# VOC 집계 — {month}")
    print(f"원본 {s['원본행수']}행 / 중복 제거 {dup}건 / 총 접수 {s['총접수']}건")
    if prev is not None:
        print(f"전월({pm}) {s['전월총접수']}건 → {s['총접수증감']:+d}건 ({s['총접수증감률']:+.1f}%)")
    else:
        print("전월 대비: 전월 없음")
    print(f"처리 완료율 {s['처리완료율']}% / 평균 처리시간 {s['평균처리시간']}시간 / "
          f"평균 만족도 {s['평균만족도']}점 (미응답 {s['만족도미응답']}건 제외)")
    print("| 유형 | 건수 | 비중 | 전월 대비 | 평균 처리시간 | 평균 만족도 |")
    print("|---|---|---|---|---|---|")
    for t in order:
        v = s["유형별"][t]
        diff = f"{v['증감']:+d}건" if "증감" in v else "전월 없음"
        print(f"| {t} | {v['건수']} | {v['비중']}% | {diff} | {v['평균처리시간']} | {v['평균만족도']} |")
    print("채널별: " + ", ".join(f"{c} {n}" for c, n in s["채널별"].items()))
    print(f"상위 3개 유형: {', '.join(s['상위3유형'])}")
    print(f"연락처 가림 {s['연락처가림']}건 / 지시문 의심 {len(s['지시문의심'])}건 {s['지시문의심']}")


if __name__ == "__main__":
    if len(sys.argv) != 2 or not re.fullmatch(r"\d{4}-\d{2}", sys.argv[1]):
        print("사용법: python 도구/VOC집계.py YYYY-MM")
        sys.exit(1)
    main(sys.argv[1])
