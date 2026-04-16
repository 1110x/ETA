#!/usr/bin/env python3
"""
Phase 1 DB Migration: 테이블명 통일 + 레거시 테이블 DROP
- RENAME: 분석의뢰및결과 → 수질분석센터_결과
- RENAME: 폐수의뢰및결과 → 비용부담금_결과
- RENAME: 처리시설_측정결과 → 처리시설_결과
- DROP: 8개 레거시 *_DATA 테이블
"""

import subprocess
import json
import sys
from pathlib import Path

def load_connection_config():
    """appsettings.json에서 MariaDB 연결 정보 읽기"""
    config_path = Path(__file__).parent.parent / "appsettings.json"
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    db_config = config.get("MariaDb", {})
    return {
        'host': db_config.get('Server', 'localhost'),
        'port': db_config.get('Port', '3306'),
        'database': db_config.get('Database', 'eta_db'),
        'user': db_config.get('User', 'root'),
        'password': db_config.get('Password', ''),
    }

def execute_sql(config, sql_statements):
    """mysql CLI를 통해 SQL 실행"""
    try:
        # 모든 명령을 한 번에 실행
        full_sql = ";\n".join(sql_statements) + ";"

        cmd = [
            'mysql',
            '-h', config['host'],
            '-P', config['port'],
            '-u', config['user'],
            f'-p{config["password"]}',
            config['database']
        ]

        result = subprocess.run(
            cmd,
            input=full_sql.encode('utf-8'),
            capture_output=True,
            timeout=30
        )

        if result.returncode != 0:
            return False, result.stderr.decode('utf-8', errors='ignore')
        return True, result.stdout.decode('utf-8', errors='ignore')
    except Exception as e:
        return False, str(e)

def execute_migration():
    """Phase 1 마이그레이션 실행"""
    config = load_connection_config()

    print(f"MariaDB 연결 중... ({config['host']}:{config['port']}/{config['database']})")

    # 1. RENAME 명령 생성
    print("\n[Phase 1] 테이블명 통일 (RENAME)")
    print("=" * 60)

    rename_sqls = [
        "RENAME TABLE `분석의뢰및결과` TO `수질분석센터_결과`",
        "RENAME TABLE `폐수의뢰및결과` TO `비용부담금_결과`",
        "RENAME TABLE `처리시설_측정결과` TO `처리시설_결과`",
    ]

    success, output = execute_sql(config, rename_sqls)
    if success:
        print("✓ 테이블명 통일 완료")
        for rename_sql in rename_sqls:
            parts = rename_sql.split(" TO ")
            old_name = parts[0].replace("RENAME TABLE `", "").strip("`")
            new_name = parts[1].strip("`").replace("`", "")
            print(f"  • {old_name} → {new_name}")
    else:
        print(f"✗ 테이블명 통일 실패: {output}")
        return 1

    # 2. DROP 명령 생성
    print("\n[Phase 1] 레거시 *_DATA 테이블 DROP")
    print("=" * 60)

    legacy_tables = [
        "BOD_DATA",
        "SS_DATA",
        "NHexan_DATA",
        "TN_DATA",
        "TP_DATA",
        "Phenols_DATA",
        "TOC_TCIC_DATA",
        "TOC_NPOC_DATA",
    ]

    drop_sqls = [f"DROP TABLE IF EXISTS `{table}`" for table in legacy_tables]

    success, output = execute_sql(config, drop_sqls)
    if success:
        print("✓ 레거시 테이블 DROP 완료")
        for table in legacy_tables:
            print(f"  • DROP {table}")
    else:
        print(f"✗ 레거시 테이블 DROP 실패: {output}")
        return 1

    print("\n✓ Phase 1 마이그레이션 완료")
    return 0

if __name__ == "__main__":
    sys.exit(execute_migration())
