"""
방류기준표 transpose (행열 전환)

현재: 행=분석항목, 열=방류기준
변경: 행=방류기준, 열=분석항목

방류기준_매핑은 영향 없음
"""
import pymysql

DB = dict(host="1110s.synology.me", port=3306, db="eta_db",
          user="eta_user", password="1212xx!!AA", charset="utf8mb4")


def main():
    print("=== 방류기준표 행열 전환 ===\n")
    conn = pymysql.connect(**DB)
    cur = conn.cursor()

    # ===== 1. 현재 데이터 읽기 =====
    cur.execute("SHOW COLUMNS FROM `방류기준표`")
    cols = [row[0] for row in cur.fetchall()]
    # _id, 구분 제외한 나머지가 방류기준 컬럼
    std_cols = [c for c in cols if c not in ("_id", "구분")]
    print(f"방류기준 컬럼 수: {len(std_cols)}")

    cur.execute(f"SELECT 구분, {', '.join(f'`{c}`' for c in std_cols)} FROM `방류기준표`")
    rows = cur.fetchall()
    print(f"분석항목 행 수: {len(rows)}")

    # 피벗: (항목, 방류기준) → 값
    # 새 구조: 행=방류기준, 각 행의 컬럼=항목
    # data[방류기준][항목] = 값
    items = [r[0] for r in rows]  # 분석항목 목록
    data = {std: {} for std in std_cols}
    for row in rows:
        항목 = row[0]
        for i, std in enumerate(std_cols):
            val = row[i + 1]  # row[0]은 구분, row[1]부터 기준값
            data[std][항목] = val

    # ===== 2. 기존 테이블 DROP, 새로 CREATE =====
    cur.execute("DROP TABLE `방류기준표`")

    col_defs = ["`_id` INT AUTO_INCREMENT PRIMARY KEY",
                "`구분` VARCHAR(100) NOT NULL"]
    for 항목 in items:
        col_defs.append(f"`{항목}` TEXT NULL DEFAULT NULL")

    create_sql = (f"CREATE TABLE `방류기준표` ({', '.join(col_defs)}) "
                  f"DEFAULT CHARSET=utf8mb4 ROW_FORMAT=DYNAMIC")
    cur.execute(create_sql)
    print(f"\n새 테이블 CREATE 완료: 컬럼 {len(items) + 2}개")

    # ===== 3. 데이터 INSERT =====
    inserted = 0
    col_list = "`구분`, " + ", ".join(f"`{i}`" for i in items)
    placeholders = ",".join(["%s"] * (1 + len(items)))
    for std in std_cols:
        values = [data[std].get(i) for i in items]
        cur.execute(
            f"INSERT INTO `방류기준표` ({col_list}) VALUES ({placeholders})",
            tuple([std] + values)
        )
        inserted += 1

    conn.commit()
    print(f"INSERT 완료: {inserted}행")

    # 결과 확인
    cur.execute("SELECT COUNT(*) FROM `방류기준표`")
    print(f"\n최종 확인 → 방류기준표: {cur.fetchone()[0]}행")
    cur.execute("SHOW COLUMNS FROM `방류기준표`")
    print(f"최종 컬럼 수: {len(cur.fetchall())}")

    conn.close()
    print("\n=== 완료 ===")


if __name__ == "__main__":
    main()
