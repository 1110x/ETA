"""
시험기록부 결과값 -> 비용부담금_결과 UPDATE

매칭 규칙:
  - 분석일 ≠ 채수일 (SN의 MM-DD 기반)
  - 채수일 BETWEEN (분석일 - 40일) AND 분석일
  - SN 일치

대상 항목:
  BOD      <- 생물학적_산소요구량_시험기록부.결과
  SS       <- 부유물질_시험기록부.결과
  T-N      <- 총질소_시험기록부.결과
  T-P      <- 총인_시험기록부.결과
  N-Hexan  <- 노말헥산_추출물질_시험기록부.결과
  Phenols  <- 페놀류_직접법_시험기록부.결과 + 페놀류_추출법_시험기록부.결과
  TOC      <- 총_유기탄소_TCIC_시험기록부.결과
"""
import pymysql
from datetime import datetime, timedelta

DB = dict(host="1110s.synology.me", port=3306, db="eta_db",
          user="eta_user", password="1212xx!!AA", charset="utf8mb4")

# 시험기록부 테이블명 -> 비용부담금_결과 컬럼명
MAPPINGS = [
    ("생물학적_산소요구량_시험기록부", "BOD"),
    ("부유물질_시험기록부",             "SS"),
    ("총질소_시험기록부",               "T-N"),
    ("총인_시험기록부",                 "T-P"),
    ("노말헥산_추출물질_시험기록부",    "N-Hexan"),
    ("페놀류_직접법_시험기록부",        "Phenols"),
    ("페놀류_추출법_시험기록부",        "Phenols"),
    ("총_유기탄소_TCIC_시험기록부",    "TOC"),
]


def get_date_col(conn, tbl):
    with conn.cursor() as cur:
        cur.execute("SHOW COLUMNS FROM `" + tbl + "`")
        cols = [row[0] for row in cur.fetchall()]
    return next(c for c in cols if '일' in c and '등록' not in c)


def update_항목(conn, src_tbl, dst_col):
    col_date = get_date_col(conn, src_tbl)
    updated = 0
    skipped = 0
    missing = 0

    with conn.cursor() as cur:
        # 시험기록부 전체 행 조회
        cur.execute(
            "SELECT `" + col_date + "`, SN, 결과 FROM `" + src_tbl + "` "
            "WHERE 결과 IS NOT NULL AND 결과 != '' AND SN IS NOT NULL AND SN != ''"
        )
        rows = cur.fetchall()

    print("  %s: %d건 대상" % (src_tbl, len(rows)))

    for 분석일_raw, sn, 결과 in rows:
        날짜 = str(분석일_raw)[:10]
        try:
            dt = datetime.strptime(날짜, "%Y-%m-%d")
        except Exception:
            skipped += 1
            continue
        start = (dt - timedelta(days=40)).strftime("%Y-%m-%d")
        end   = dt.strftime("%Y-%m-%d")

        with conn.cursor() as cur:
            cur.execute(
                "SELECT id FROM `비용부담금_결과` "
                "WHERE SN=%s AND 채수일 BETWEEN %s AND %s LIMIT 1",
                (sn, start, end)
            )
            row = cur.fetchone()
            if not row:
                missing += 1
                continue
            cur.execute(
                "UPDATE `비용부담금_결과` SET `" + dst_col + "`=%s WHERE id=%s",
                (결과, row[0])
            )
            updated += 1

        if updated % 500 == 0 and updated > 0:
            conn.commit()
            print("    %s -> %s: %d건 업데이트 중..." % (src_tbl, dst_col, updated))

    conn.commit()
    print("  %s -> %s: 업데이트 %d건, 매칭실패 %d건, 스킵 %d건" % (
        src_tbl, dst_col, updated, missing, skipped))
    return updated


def main():
    print("비용부담금_결과 UPDATE 시작\n")
    conn = pymysql.connect(**DB, autocommit=True)
    total = 0
    try:
        for src, dst in MAPPINGS:
            print("[%s -> %s]" % (src, dst))
            total += update_항목(conn, src, dst)
            print()
        print("\n전체 업데이트: %d건" % total)
    finally:
        conn.close()


if __name__ == "__main__":
    main()
