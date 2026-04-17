#!/usr/bin/env python3
"""
총_유기탄소_TCIC/NPOC_시험기록부 → 비용부담금_결과.TOC UPDATE
SN + 구분 기반 매칭
"""

import mysql.connector

DB_CONFIG = {
    'host': '1110s.synology.me',
    'user': 'eta_user',
    'password': '1212xx!!AA',
    'database': 'eta_db',
    'charset': 'utf8mb4'
}

try:
    print("TOC 시험기록부에서 데이터 추출 중...")
    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor(dictionary=True)

    records = []

    # TCIC 테이블
    cursor.execute("""
        SELECT SN, 구분, 결과 FROM `총_유기탄소_TCIC_시험기록부`
        WHERE 결과 IS NOT NULL AND 결과 != ''
    """)
    tcic_records = cursor.fetchall()
    print(f"✓ 총_유기탄소_TCIC_시험기록부: {len(tcic_records)}개 레코드")
    records.extend(tcic_records)

    # NPOC 테이블
    cursor.execute("""
        SELECT SN, 구분, 결과 FROM `총_유기탄소_NPOC_시험기록부`
        WHERE 결과 IS NOT NULL AND 결과 != ''
    """)
    npoc_records = cursor.fetchall()
    print(f"✓ 총_유기탄소_NPOC_시험기록부: {len(npoc_records)}개 레코드")
    records.extend(npoc_records)

    print(f"✓ 총 {len(records)}개 레코드")

    if records:
        print(f"\n비용부담금_결과 UPDATE 중...")
        updated = 0

        for rec in records:
            sn = rec['SN']
            division = rec['구분']
            result = rec['결과']

            if not sn or not division:
                continue

            try:
                query = """
                    UPDATE `비용부담금_결과`
                    SET TOC = %s
                    WHERE SN = %s
                      AND 구분 = %s
                      AND (TOC IS NULL OR TOC = '')
                """
                cursor.execute(query, (result, sn, division))
                updated += cursor.rowcount

                if updated % 100 == 0:
                    print(f"  {updated}개 업데이트...")
            except Exception as e:
                pass

        conn.commit()
        print(f"✓ {updated}개 행 UPDATE 완료")
        cursor.close()
        conn.close()

except Exception as e:
    print(f"❌ 에러: {e}")
    import traceback
    traceback.print_exc()
