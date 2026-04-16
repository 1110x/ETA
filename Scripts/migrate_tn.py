"""
TN-DATA -> 총질소_시험기록부 마이그레이션

시트: TN-DATA

고정영역:
  col1: 분석일
  col2: TN-분석자
  col3~7:   ST-01~05 mgL
  col8~12:  ST-01~05 abs
  col13~15: 기울기/절편/R2

샘플 블록 (col30부터 8컬럼 간격):
  +0=시료명, +1=시료량, +2=흡광도, +3=희석배수,
  +4=검량선으로 구한 a, +5=농도, +6=결과(빈값 가능), +7=SN

결과값이 비어있으면: 농도 × 희석배수 로 계산

UPSERT 기준: 분석일 + SN
"""
import openpyxl
import pymysql
from datetime import datetime, date, timedelta

XLSM_PATH = "Docs/CHUNGHA-김가린.xlsm"
DB = dict(host="1110s.synology.me", port=3306, db="eta_db",
          user="eta_user", password="1212xx!!AA", charset="utf8mb4")

TN_TBL = "총질소_시험기록부"


def get_val(cell):
    v = cell.value
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    return "" if s == "-" else s


def try_parse_date(v):
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    try:
        days = int(float(s))
        if days > 59:
            days -= 1
        return (datetime(1899, 12, 30) + timedelta(days=days)).strftime("%Y-%m-%d")
    except Exception:
        pass
    if len(s) == 10 and s[4] == '-':
        return s
    try:
        s2 = s.split(". ")[0].strip().rstrip(".")
        return datetime.strptime(s2, "%Y. %m. %d").strftime("%Y-%m-%d")
    except Exception:
        pass
    return None


def normalize_sn(sn):
    parts = sn.split("-")
    if len(parts) != 3:
        return sn
    try:
        return "%02d-%02d-%02d" % (int(parts[0]), int(parts[1]), int(parts[2]))
    except Exception:
        return sn


def parse_sn(raw):
    s = raw.strip()
    if not s:
        return None, None
    gub = "여수"
    if s.startswith("[율촌]"):
        gub = "율촌"
        s = s[5:]
    elif s.startswith("[세풍]"):
        gub = "세풍"
        s = s[5:]
    try:
        float(s)
        return None, None
    except Exception:
        pass
    return normalize_sn(s), gub


def calc_result(농도, 희석):
    try:
        return "%.4f" % (float(농도) * float(희석))
    except Exception:
        return ""


def get_date_col(conn):
    with conn.cursor() as cur:
        cur.execute("SHOW COLUMNS FROM `" + TN_TBL + "`")
        cols = [row[0] for row in cur.fetchall()]
    return next(c for c in cols if '일' in c and '등록' not in c)


def upsert_tn(conn, col_date, 날짜, sn, 업체명, 구분, 분석자,
              시료량, 흡광도, 희석, 검량선a, 농도, 결과,
              st_mgl, st_abs, curve):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sel = ("SELECT id FROM `" + TN_TBL + "` "
           "WHERE LEFT(`" + col_date + "`,10)=%s AND SN=%s LIMIT 1")
    with conn.cursor() as cur:
        cur.execute(sel, (날짜, sn))
        row = cur.fetchone()
        if row:
            upd = ("UPDATE `" + TN_TBL + "` SET "
                   "업체명=%s, 시료명=%s, 구분=%s, 소스구분=%s, 분석자=%s, "
                   "시료량=%s, 흡광도=%s, 희석배수=%s, 검량선_a=%s, 농도=%s, 결과=%s, "
                   "기울기=%s, 절편=%s, R2=%s, "
                   "ST01_mgL=%s, ST02_mgL=%s, ST03_mgL=%s, ST04_mgL=%s, ST05_mgL=%s, "
                   "ST01_abs=%s, ST02_abs=%s, ST03_abs=%s, ST04_abs=%s, ST05_abs=%s, "
                   "등록일시=%s WHERE id=%s")
            cur.execute(upd, (
                업체명, 업체명, 구분, "폐수배출업소", 분석자,
                시료량, 흡광도, 희석, 검량선a, 농도, 결과,
                curve[0], curve[1], curve[2],
                st_mgl[0], st_mgl[1], st_mgl[2], st_mgl[3], st_mgl[4],
                st_abs[0], st_abs[1], st_abs[2], st_abs[3], st_abs[4],
                now, row[0]
            ))
        else:
            ins = ("INSERT INTO `" + TN_TBL + "` "
                   "(`" + col_date + "`, SN, 업체명, 시료명, 구분, 소스구분, 분석자, "
                   "시료량, 흡광도, 희석배수, 검량선_a, 농도, 결과, "
                   "기울기, 절편, R2, "
                   "ST01_mgL, ST02_mgL, ST03_mgL, ST04_mgL, ST05_mgL, "
                   "ST01_abs, ST02_abs, ST03_abs, ST04_abs, ST05_abs, "
                   "등록일시) "
                   "VALUES (%s,%s,%s,%s,%s,%s,%s,"
                   "%s,%s,%s,%s,%s,%s,"
                   "%s,%s,%s,"
                   "%s,%s,%s,%s,%s,"
                   "%s,%s,%s,%s,%s,"
                   "%s)")
            cur.execute(ins, (
                날짜, sn, 업체명, 업체명, 구분, "폐수배출업소", 분석자,
                시료량, 흡광도, 희석, 검량선a, 농도, 결과,
                curve[0], curve[1], curve[2],
                st_mgl[0], st_mgl[1], st_mgl[2], st_mgl[3], st_mgl[4],
                st_abs[0], st_abs[1], st_abs[2], st_abs[3], st_abs[4],
                now
            ))


def load_tn(wb, conn, col_date):
    ws = wb["TN-DATA"]
    loaded = 0
    skipped = 0
    START_DATE = "2024-04-06"

    for row_num, r in enumerate(ws.iter_rows(min_row=2), start=2):
        날짜 = try_parse_date(r[0].value)
        if not 날짜:
            continue

        # 2024-04-06 이전 데이터는 건너뜀 (이미 저장됨)
        if 날짜 < START_DATE:
            continue

        분석자 = get_val(r[1]) if len(r) > 1 else ""

        # 고정영역: 검정곡선
        st_mgl = [get_val(r[i]) for i in range(2, 7)]
        st_abs = [get_val(r[i]) for i in range(7, 12)]
        curve  = [get_val(r[12]), get_val(r[13]), get_val(r[14])]

        # 샘플 블록 (col30~, 8컬럼 간격)
        for block in range(80):
            idx = 29 + block * 8
            if idx >= len(r):
                break
            시료명 = get_val(r[idx])
            if not 시료명:
                break
            sn_raw = get_val(r[idx + 7]) if idx + 7 < len(r) else ""

            시료량 = get_val(r[idx + 1]) if idx + 1 < len(r) else ""
            흡광도 = get_val(r[idx + 2]) if idx + 2 < len(r) else ""
            희석   = get_val(r[idx + 3]) if idx + 3 < len(r) else ""
            검량선a = get_val(r[idx + 4]) if idx + 4 < len(r) else ""
            농도   = get_val(r[idx + 5]) if idx + 5 < len(r) else ""
            결과   = get_val(r[idx + 6]) if idx + 6 < len(r) else ""

            # 결과값이 없으면 농도 × 희석배수로 계산
            if not 결과:
                결과 = calc_result(농도, 희석)

            if not 결과 or 결과 == "0":
                skipped += 1
                continue

            sn, 구분 = parse_sn(sn_raw)
            if sn is None:
                skipped += 1
                continue

            try:
                upsert_tn(conn, col_date, 날짜, sn, 시료명, 구분, 분석자,
                          시료량, 흡광도, 희석, 검량선a, 농도, 결과,
                          st_mgl, st_abs, curve)
                loaded += 1
                if loaded % 200 == 0:
                    conn.commit()
                    print("    [TN] %d건 처리 중... (행%d)" % (loaded, row_num))
            except Exception as e:
                print("    오류 행%d 블록%d: %s" % (row_num, block, e))

    conn.commit()
    print("  [TN] %d건 저장, %d건 스킵" % (loaded, skipped))
    return loaded


def main():
    print("TN 마이그레이션 시작")
    conn = pymysql.connect(**DB)
    col_date = get_date_col(conn)
    print("날짜 컬럼: %s" % repr(col_date))

    print("기존 데이터 유지, 2024-04-06 이후만 처리")

    wb = openpyxl.load_workbook(XLSM_PATH, data_only=True, read_only=True)
    try:
        total = load_tn(wb, conn, col_date)
        print("\n완료: 총 %d건" % total)
    finally:
        conn.close()
        wb.close()


if __name__ == "__main__":
    main()
