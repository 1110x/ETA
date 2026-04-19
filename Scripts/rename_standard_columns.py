"""
방류기준표 컬럼명 정리 — 괄호 약어 제거

분석정보 테이블 항목명과 매칭되도록 `시안(CN)` → `시안`,
`폴리클로리네이티드비페닐(PCB)` → `폴리클로리네이티드비페닐`,
`다이에틸헥실프탈레이트(DEHP)` → `다이에틸헥실프탈레이트` 식으로
방류기준표의 컬럼 끝 괄호(...) 접미사를 제거한다.

사용법:
    python Scripts/rename_standard_columns.py          # DRY RUN (계획만 출력)
    python Scripts/rename_standard_columns.py --apply  # 실제 적용
"""
import re
import sys

import pymysql

DB = dict(host="1110s.synology.me", port=3306, db="eta_db",
          user="eta_user", password="1212xx!!AA", charset="utf8mb4")

TABLE = "방류기준표"
TRAIL_PAREN = re.compile(r"\s*\([^)]*\)\s*$")


def strip_paren(name: str) -> str:
    """컬럼명 끝의 ' (XXX)' 또는 '(XXX)' 제거."""
    prev = None
    cur = name
    while prev != cur:
        prev = cur
        cur = TRAIL_PAREN.sub("", cur).strip()
    return cur


def get_columns(cur) -> list:
    cur.execute(
        "SELECT COLUMN_NAME, COLUMN_TYPE, IS_NULLABLE, COLUMN_DEFAULT, EXTRA "
        "FROM information_schema.COLUMNS "
        "WHERE TABLE_SCHEMA = %s AND TABLE_NAME = %s "
        "ORDER BY ORDINAL_POSITION",
        (DB["db"], TABLE),
    )
    return list(cur.fetchall())


def main():
    apply = "--apply" in sys.argv

    print(f"=== 방류기준표 컬럼명 정리 ({'APPLY' if apply else 'DRY RUN'}) ===\n")

    conn = pymysql.connect(**DB)
    cur = conn.cursor()

    cols = get_columns(cur)
    if not cols:
        print(f"⚠ 테이블 `{TABLE}` 를 찾지 못했습니다.")
        return

    existing = {c[0] for c in cols}
    col_meta = {c[0]: c for c in cols}
    plan = []      # (old, new, type, nullable, default, extra)
    drops = []     # 데이터 없는 충돌 컬럼 DROP
    skipped = []

    def non_null_count(col: str) -> int:
        cur.execute(
            f"SELECT COUNT(*) FROM `{TABLE}` "
            f"WHERE `{col}` IS NOT NULL AND TRIM(`{col}`) <> ''"
        )
        return int(cur.fetchone()[0])

    for name, ctype, nullable, default, extra in cols:
        new = strip_paren(name)
        if new == name:
            continue
        if not new:
            skipped.append((name, "(빈 이름 생성)"))
            continue
        if new in existing and new != name:
            # 충돌 — 데이터 존재 여부로 우선순위 결정
            old_cnt = non_null_count(name)
            new_cnt = non_null_count(new)
            if old_cnt > 0 and new_cnt == 0:
                # 약어 컬럼만 데이터 있음: 빈 컬럼 DROP 후 rename
                drops.append((new, "빈 컬럼"))
                existing.discard(new)
                plan.append((name, new, ctype, nullable, default, extra))
                existing.discard(name)
                existing.add(new)
            elif old_cnt == 0 and new_cnt > 0:
                # 정제된 컬럼에 데이터 있음: 약어 컬럼 DROP
                drops.append((name, "빈 컬럼(약어)"))
                existing.discard(name)
            elif old_cnt == 0 and new_cnt == 0:
                # 둘 다 비어 있음: 약어 컬럼 DROP
                drops.append((name, "둘 다 빈 컬럼(약어쪽 제거)"))
                existing.discard(name)
            else:
                skipped.append((name, f"→ `{new}` 충돌: 양쪽 모두 데이터 존재 ({name}={old_cnt}건, {new}={new_cnt}건)"))
            continue
        plan.append((name, new, ctype, nullable, default, extra))
        existing.discard(name)
        existing.add(new)

    print(f"RENAME: {len(plan)}건 / DROP: {len(drops)}건 / SKIP: {len(skipped)}건\n")
    if plan:
        print("[RENAME]")
        for old, new, *_ in plan:
            print(f"  `{old}`  →  `{new}`")
    if drops:
        print("\n[DROP]")
        for col, why in drops:
            print(f"  `{col}`  ({why})")
    if skipped:
        print("\n⚠ 건너뜀:")
        for old, why in skipped:
            print(f"  `{old}`  {why}")

    if not apply:
        print("\n(DRY RUN) — 실제 적용하려면 `--apply` 옵션을 붙여 다시 실행하세요.")
        return

    if not plan and not drops:
        print("\n변경할 컬럼이 없습니다.")
        return

    # 1단계: DROP 먼저 (충돌 제거)
    if drops:
        print("\n--- DROP 실행 ---")
        for col, why in drops:
            try:
                cur.execute(f"ALTER TABLE `{TABLE}` DROP COLUMN `{col}`")
                print(f"  ✓ DROP `{col}` ({why})")
            except Exception as e:
                print(f"  ✗ DROP `{col}` 실패: {e}")
        conn.commit()

    # 2단계: RENAME
    if plan:
        print("\n--- RENAME 실행 ---")
        for old, new, ctype, nullable, default, extra in plan:
            null_sql = "NULL" if nullable == "YES" else "NOT NULL"
            default_sql = ""
            if default is not None:
                default_sql = f" DEFAULT '{default}'"
            elif nullable == "YES":
                default_sql = " DEFAULT NULL"
            extra_sql = f" {extra}" if extra else ""
            sql = (
                f"ALTER TABLE `{TABLE}` "
                f"CHANGE COLUMN `{old}` `{new}` {ctype} {null_sql}{default_sql}{extra_sql}"
            )
            try:
                cur.execute(sql)
                print(f"  ✓ `{old}` → `{new}`")
            except Exception as e:
                print(f"  ✗ `{old}` → `{new}` 실패: {e}")
        conn.commit()

    conn.close()
    print("\n=== 완료 ===")


if __name__ == "__main__":
    main()
