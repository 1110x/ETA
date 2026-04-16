"""
TOC 결과값 -> 비용부담금_결과.TOC UPDATE

공정시험기준 ES 04311.1d 총 유기탄소 - 고온연소산화법:
  1) 기본: TC-IC (가감법) 우선 적용
  2) IC 측정값 > TC의 50% 초과 시: TC-IC 결과 사용 불가
     → NPOC 방법 적용 (NPOC 시험기록부의 결과값 사용)

매칭 규칙:
  - 분석일 ≠ 채수일 (SN의 MM-DD 기반)
  - 채수일 BETWEEN (분석일 - 40일) AND 분석일
  - SN 일치

로직:
  각 TCIC 레코드에 대해:
    TC=TCcon, IC=ICcon 읽기
    만약 IC > TC * 0.5 이면:
      같은 분석일+SN의 NPOC 레코드 찾아서 결과값 사용
      없으면 스킵 (NPOC 없는 케이스)
    아니면:
      TCIC의 결과_TCIC 사용... 은 아니고 TC-IC 가감법 결과
      → TCIC 테이블 "결과" 컬럼이 이미 그 값이므로 그대로 사용
  → 비용부담금_결과.TOC UPDATE
"""
import pymysql
from datetime import datetime, timedelta

DB = dict(host="1110s.synology.me", port=3306, db="eta_db",
          user="eta_user", password="1212xx!!AA", charset="utf8mb4")


def to_float(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except Exception:
        return None


def get_date_col(conn, tbl):
    with conn.cursor() as cur:
        cur.execute("SHOW COLUMNS FROM `" + tbl + "`")
        cols = [row[0] for row in cur.fetchall()]
    return next(c for c in cols if '일' in c and '등록' not in c)


def load_npoc_map(conn, col_date_npoc):
    """(분析일, SN) -> 결과 매핑"""
    with conn.cursor() as cur:
        cur.execute(
            "SELECT `" + col_date_npoc + "`, SN, 결과 FROM `총_유기탄소_NPOC_시험기록부` "
            "WHERE 결과 IS NOT NULL AND 결과 != '' AND SN IS NOT NULL"
        )
        rows = cur.fetchall()
    m = {}
    for 날짜, sn, 결과 in rows:
        key = (str(날짜)[:10], sn)
        m[key] = 결과
    return m


def main():
    print("TOC UPDATE 시작 (TC-IC 50% 규칙 적용)\n")
    conn = pymysql.connect(**DB)

    col_date_tcic = get_date_col(conn, "총_유기탄소_TCIC_시험기록부")
    col_date_npoc = get_date_col(conn, "총_유기탄소_NPOC_시험기록부")
    print("TCIC 날짜컬럼:", col_date_tcic)
    print("NPOC 날짜컬럼:", col_date_npoc)

    # 1단계: 기존 TOC 값 모두 초기화
    with conn.cursor() as cur:
        cur.execute("UPDATE `비용부담금_결과` SET `TOC`=NULL")
    conn.commit()
    print("기존 비용부담금_결과.TOC 초기화 완료\n")

    # NPOC 매핑 사전 로드
    npoc_map = load_npoc_map(conn, col_date_npoc)
    print("NPOC 매핑 로드: %d건\n" % len(npoc_map))

    # TCIC 전체 조회
    with conn.cursor() as cur:
        cur.execute(
            "SELECT `" + col_date_tcic + "`, SN, TCcon, ICcon, 결과 "
            "FROM `총_유기탄소_TCIC_시험기록부` "
            "WHERE SN IS NOT NULL AND SN != ''"
        )
        rows = cur.fetchall()
    print("TCIC 대상: %d건\n" % len(rows))

    updated_tcic = 0
    updated_npoc = 0
    skipped_no_npoc = 0
    missing_waste = 0

    for 분석일_raw, sn, tccon, iccon, 결과_tcic in rows:
        날짜 = str(분석일_raw)[:10]
        try:
            dt = datetime.strptime(날짜, "%Y-%m-%d")
        except Exception:
            continue
        start = (dt - timedelta(days=40)).strftime("%Y-%m-%d")
        end   = dt.strftime("%Y-%m-%d")

        tc = to_float(tccon)
        ic = to_float(iccon)

        # TC-IC 50% 규칙 판단
        사용값 = None
        사용방법 = None
        if tc is not None and ic is not None and tc > 0 and ic > tc * 0.5:
            # NPOC 사용
            key = (날짜, sn)
            npoc_결과 = npoc_map.get(key)
            if npoc_결과:
                사용값 = npoc_결과
                사용방법 = "NPOC"
            else:
                skipped_no_npoc += 1
                continue
        else:
            # TCIC 사용 (TC-IC 가감법 결과)
            if not 결과_tcic:
                continue
            사용값 = 결과_tcic
            사용방법 = "TCIC"

        # 비용부담금_결과 UPDATE
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id FROM `비용부담금_결과` "
                "WHERE SN=%s AND 채수일 BETWEEN %s AND %s LIMIT 1",
                (sn, start, end)
            )
            row = cur.fetchone()
            if not row:
                missing_waste += 1
                continue
            cur.execute(
                "UPDATE `비용부담금_결과` SET `TOC`=%s WHERE id=%s",
                (사용값, row[0])
            )
            if 사용방법 == "TCIC":
                updated_tcic += 1
            else:
                updated_npoc += 1

        total = updated_tcic + updated_npoc
        if total % 500 == 0 and total > 0:
            conn.commit()
            print("  %d건 업데이트 중... (TCIC %d / NPOC %d)" % (total, updated_tcic, updated_npoc))

    conn.commit()

    # NPOC만 있는 케이스 처리 (TCIC 레코드가 없지만 NPOC는 있는 SN)
    with conn.cursor() as cur:
        cur.execute("""
            SELECT n.`""" + col_date_npoc + """`, n.SN, n.결과
            FROM `총_유기탄소_NPOC_시험기록부` n
            LEFT JOIN `총_유기탄소_TCIC_시험기록부` t
              ON LEFT(n.`""" + col_date_npoc + """`,10) = LEFT(t.`""" + col_date_tcic + """`,10)
             AND n.SN = t.SN
            WHERE t.id IS NULL
              AND n.결과 IS NOT NULL AND n.결과 != ''
              AND n.SN IS NOT NULL AND n.SN != ''
        """)
        npoc_only = cur.fetchall()

    print("\nTCIC 없이 NPOC만 있는 케이스: %d건" % len(npoc_only))

    updated_npoc_only = 0
    for 분석일_raw, sn, 결과 in npoc_only:
        날짜 = str(분석일_raw)[:10]
        try:
            dt = datetime.strptime(날짜, "%Y-%m-%d")
        except Exception:
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
                missing_waste += 1
                continue
            cur.execute(
                "UPDATE `비용부담금_결과` SET `TOC`=%s WHERE id=%s",
                (결과, row[0])
            )
            updated_npoc_only += 1

    conn.commit()

    print("\n=== TOC UPDATE 결과 ===")
    print("  TCIC 기준(TC-IC 가감법):    %d건" % updated_tcic)
    print("  NPOC 사용(IC>TC*50%%):      %d건" % updated_npoc)
    print("  NPOC only (TCIC 없음):      %d건" % updated_npoc_only)
    print("  IC>TC*50% 인데 NPOC 없음:   %d건 (스킵)" % skipped_no_npoc)
    print("  비용부담금_결과 매칭 실패: %d건" % missing_waste)
    print("  전체 업데이트:              %d건" % (updated_tcic + updated_npoc + updated_npoc_only))

    conn.close()


if __name__ == "__main__":
    main()
