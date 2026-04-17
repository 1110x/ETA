#!/usr/bin/env python3
"""
페놀류_직접법/추출법_시험기록부 → 비용부담금_결과.Phenols UPDATE
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
    print("페놀류 시험기록부에서 데이터 추출 중...")
    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor(dictionary=True)

    records = []

    # 직접법 테이블
    cursor.execute("""
        SELECT SN, 구분, 결과 FROM `페놀류_직접법_시험기록부`
        WHERE 결과 IS NOT NULL AND 결과 != ''
    """)
    direct_records = cursor.fetchall()
    print(f"✓ 페놀류_직접법_시험기록부: {len(direct_records)}개 레코드")
    records.extend(direct_records)

    # 추출법 테이블
    cursor.execute("""
        SELECT SN, 구분, 결과 FROM `페놀류_추출법_시험기록부`
        WHERE 결과 IS NOT NULL AND 결과 != ''
    """)
    extract_records = cursor.fetchall()
    print(f"✓ 페놀류_추출법_시험기록부: {len(extract_records)}개 레코드")
    records.extend(extract_records)

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
                    SET Phenols = %s
                    WHERE SN = %s
                      AND 구분 = %s
                      AND (Phenols IS NULL OR Phenols = '')
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
