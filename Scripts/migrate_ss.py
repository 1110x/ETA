"""
SS-DATA -> 부유물질_시험기록부 마이그레이션

시트: SS-DATA
블록 (col30부터 8컬럼 간격):
  +0=시료명, +1=시료량, +2=전무게, +3=후무게, +4=(미사용),
  +5=희석배수, +6=결과, +7=SN(Remark)

무게차 = 후무게 - 전무게 (계산)

UPSERT 기준: 분석일 + SN
"""
import openpyxl
import pymysql
from datetime import datetime, date, timedelta

XLSM_PATH = "Docs/CHUNGHA-김가린.xlsm"
DB = dict(host="1110s.synology.me", port=3306, db="eta_db",
          user="eta_user", password="1212xx!!AA", charset="utf8mb4")

SS_TBL = "부유물질_시험기록부"


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


def calc_weight_diff(전, 후):
    try:
        return "%.4f" % (float(후) - float(전))
    except Exception:
        return ""


def get_date_col(conn):
    with conn.cursor() as cur:
        cur.execute("SHOW COLUMNS FROM `" + SS_TBL + "`")
        cols = [row[0] for row in cur.fetchall()]
    return next(c for c in cols if '일' in c and '등록' not in c)


def upsert_ss(conn, col_date, 날짜, sn, 업체명, 구분, 분석자,
              시료량, 전무게, 후무게, 무게차, 희석, 결과):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sel = ("SELECT id FROM `" + SS_TBL + "` "
           "WHERE LEFT(`" + col_date + "`,10)=%s AND SN=%s LIMIT 1")
    with conn.cursor() as cur:
        cur.execute(sel, (날짜, sn))
        row = cur.fetchone()
        if row:
            upd = ("UPDATE `" + SS_TBL + "` SET "
                   "업체명=%s, 시료명=%s, 구분=%s, 소스구분=%s, 분석자=%s, "
                   "시료량=%s, 전무게=%s, 후무게=%s, 무게차=%s, 희석배수=%s, 결과=%s, "
                   "등록일시=%s WHERE id=%s")
            cur.execute(upd, (
                업체명, 업체명, 구분, "폐수배출업소", 분석자,
                시료량, 전무게, 후무게, 무게차, 희석, 결과,
                now, row[0]
            ))
        else:
            ins = ("INSERT INTO `" + SS_TBL + "` "
                   "(`" + col_date + "`, SN, 업체명, 시료명, 구분, 소스구분, 분석자, "
                   "시료량, 전무게, 후무게, 무게차, 희석배수, 결과, 등록일시) "
                   "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)")
            cur.execute(ins, (
                날짜, sn, 업체명, 업체명, 구분, "폐수배출업소", 분석자,
                시료량, 전무게, 후무게, 무게차, 희석, 결과,
                now
            ))


def load_ss(wb, conn, col_date):
    ws = wb["SS-DATA"]
    loaded = 0
    skipped = 0

    for row_num, r in enumerate(ws.iter_rows(min_row=2), start=2):
        날짜 = try_parse_date(r[0].value)
        if not 날짜:
            continue

        분석자 = get_val(r[1]) if len(r) > 1 else ""

        for block in range(80):
            idx = 29 + block * 8
            if idx >= len(r):
                break
            시료명 = get_val(r[idx])
            if not 시료명:
                break
            결과   = get_val(r[idx + 6]) if idx + 6 < len(r) else ""
            sn_raw = get_val(r[idx + 7]) if idx + 7 < len(r) else ""
            if not 결과 or 결과 == "0":
                skipped += 1
                continue
            sn, 구분 = parse_sn(sn_raw)
            if sn is None:
                skipped += 1
                continue
            시료량 = get_val(r[idx + 1]) if idx + 1 < len(r) else ""
            전무게 = get_val(r[idx + 2]) if idx + 2 < len(r) else ""
            후무게 = get_val(r[idx + 3]) if idx + 3 < len(r) else ""
            희석   = get_val(r[idx + 5]) if idx + 5 < len(r) else ""
            무게차 = calc_weight_diff(전무게, 후무게)
            try:
                upsert_ss(conn, col_date, 날짜, sn, 시료명, 구분, 분석자,
                          시료량, 전무게, 후무게, 무게차, 희석, 결과)
                loaded += 1
                if loaded % 200 == 0:
                    conn.commit()
                    print("    [SS] %d건 처리 중... (행%d)" % (loaded, row_num))
            except Exception as e:
                print("    오류 행%d 블록%d: %s" % (row_num, block, e))

    conn.commit()
    print("  [SS] %d건 저장, %d건 스킵" % (loaded, skipped))
    return loaded


def main():
    print("SS 마이그레이션 시작")
    conn = pymysql.connect(**DB)
    col_date = get_date_col(conn)
    print("날짜 컬럼: %s" % repr(col_date))

    with conn.cursor() as cur:
        cur.execute("DELETE FROM `" + SS_TBL + "`")
    conn.commit()
    print("기존 데이터 삭제 완료")

    wb = openpyxl.load_workbook(XLSM_PATH, data_only=True, read_only=True)
    try:
        total = load_ss(wb, conn, col_date)
        print("\n완료: 총 %d건" % total)
    finally:
        conn.close()
        wb.close()


if __name__ == "__main__":
    main()
