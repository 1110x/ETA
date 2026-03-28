#!/usr/bin/env python3
"""
ETA SQLite → MariaDB 데이터 이전 스크립트
사전조건:
  pip install mysql-connector-python  (또는 pymysql)
  시놀로지 MariaDB가 appsettings.json과 동일한 접속 정보로 구동 중이어야 함

사용법:
  python3 migrate_to_mariadb.py
"""

import sqlite3, json, os, sys
from pathlib import Path

# ── 설정 ────────────────────────────────────────────────────────────────────
SQLITE_PATH = Path.home() / "Documents" / "ETA" / "Data" / "eta.db"
APPSETTINGS = Path(__file__).parent / "appsettings.json"

def load_mariadb_config():
    if not APPSETTINGS.exists():
        print(f"[오류] appsettings.json을 찾을 수 없습니다: {APPSETTINGS}")
        sys.exit(1)
    with open(APPSETTINGS, encoding='utf-8') as f:
        cfg = json.load(f)
    m = cfg.get("MariaDb", {})
    return dict(
        host=m.get("Server", "localhost"),
        port=int(m.get("Port", 3306)),
        database=m.get("Database", "eta_db"),
        user=m.get("User", "eta_user"),
        password=m.get("Password", ""),
        charset="utf8mb4",
    )

try:
    import mysql.connector as mariadb
except ImportError:
    print("[오류] mysql-connector-python 패키지가 필요합니다.")
    print("       pip install mysql-connector-python")
    sys.exit(1)

# ── SQLite → MariaDB 타입 변환 ───────────────────────────────────────────────
def sqlite_to_mariadb_ddl(sqlite_ddl: str) -> str:
    ddl = sqlite_ddl
    ddl = ddl.replace("INTEGER PRIMARY KEY AUTOINCREMENT", "INT AUTO_INCREMENT PRIMARY KEY")
    ddl = ddl.replace("INTEGER PRIMARY KEY", "INT PRIMARY KEY")
    ddl = ddl.replace("AUTOINCREMENT", "AUTO_INCREMENT")
    # SQLite boolean → MariaDB TINYINT
    ddl = ddl.replace(" BOOLEAN", " TINYINT(1)")
    # Remove SQLite-specific clauses
    import re
    ddl = re.sub(r'\bWITHOUT\s+ROWID\b', '', ddl, flags=re.IGNORECASE)
    return ddl

def main():
    cfg = load_mariadb_config()
    print(f"[연결] SQLite: {SQLITE_PATH}")
    print(f"[연결] MariaDB: {cfg['host']}:{cfg['port']}/{cfg['database']}")

    if not SQLITE_PATH.exists():
        print(f"[오류] SQLite 파일을 찾을 수 없습니다: {SQLITE_PATH}")
        sys.exit(1)

    sqlite_conn = sqlite3.connect(SQLITE_PATH)
    sqlite_conn.row_factory = sqlite3.Row

    maria_conn = mariadb.connect(**cfg)
    maria_cur = maria_conn.cursor()
    maria_cur.execute("SET NAMES utf8mb4")
    maria_cur.execute("SET FOREIGN_KEY_CHECKS=0")

    # 테이블 목록 조회
    tables = sqlite_conn.execute(
        "SELECT name, sql FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
    ).fetchall()

    print(f"\n[정보] 이전할 테이블 수: {len(tables)}")

    for (tname, create_sql) in tables:
        print(f"\n── {tname} ──")
        if not create_sql:
            print(f"  [SKIP] CREATE SQL 없음")
            continue

        # MariaDB용 DDL로 변환
        maria_ddl = sqlite_to_mariadb_ddl(create_sql)
        maria_ddl = maria_ddl.rstrip(";")

        # MariaDB 엔진/문자셋 옵션 추가
        maria_ddl += " ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci"

        try:
            maria_cur.execute(f"DROP TABLE IF EXISTS `{tname}`")
            maria_cur.execute(maria_ddl)
            print(f"  [OK] 테이블 생성")
        except Exception as e:
            print(f"  [WARN] 테이블 생성 실패: {e}")
            print(f"  DDL: {maria_ddl[:200]}")
            continue

        # 데이터 이전
        rows = sqlite_conn.execute(f'SELECT * FROM "{tname}"').fetchall()
        if not rows:
            print(f"  [--] 데이터 없음")
            continue

        cols = rows[0].keys()
        col_list = ", ".join(f"`{c}`" for c in cols)
        placeholders = ", ".join(["%s"] * len(cols))
        insert_sql = f"INSERT INTO `{tname}` ({col_list}) VALUES ({placeholders})"

        batch = [tuple(row) for row in rows]
        try:
            maria_cur.executemany(insert_sql, batch)
            maria_conn.commit()
            print(f"  [OK] {len(batch)}행 이전 완료")
        except Exception as e:
            maria_conn.rollback()
            print(f"  [ERR] 데이터 삽입 오류: {e}")

    maria_cur.execute("SET FOREIGN_KEY_CHECKS=1")
    sqlite_conn.close()
    maria_conn.close()
    print("\n✅ 마이그레이션 완료!")

if __name__ == "__main__":
    main()
