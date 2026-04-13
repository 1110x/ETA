#!/usr/bin/env python3
"""시약 문서와 DB CSV를 비교하여 일치/비일치 결과를 출력합니다.

사용법:
  python Scripts/reagent_compare.py --db-csv path/to/reagent_db.csv --out-csv result.csv Docs/시약정리_2일반항목.md Docs/시약정리_3이온류.md Docs/시약정리_4금속류.md Docs/시약정리_5유기물질.md

요구패키지: pandas, rapidfuzz
"""
import argparse
import csv
import re
from pathlib import Path
from rapidfuzz import process, fuzz


def extract_names_from_md(path):
    names = set()
    rx_dash = re.compile(r'^\s*[-•]\s*(.+)')
    with open(path, encoding='utf-8') as f:
        for line in f:
            m = rx_dash.match(line)
            if not m:
                continue
            txt = m.group(1).strip()
            # remove parenthetical explanatory phrases but keep chemical formulas
            txt = re.sub(r'\([^)]*\)', '', txt)
            # split on common separators
            parts = re.split(r"[,;/•]", txt)
            for p in parts:
                p = p.strip()
                if not p:
                    continue
                # ignore very short tokens
                if len(p) < 2:
                    continue
                names.add(p)
    return sorted(names)


def detect_name_column(header):
    keys = [h.lower() for h in header]
    for cand in ('품목명', '품명', 'name', 'item', '영문명', 'item_no', 'itemno'):
        for i, h in enumerate(keys):
            if cand in h:
                return header[i]
    # fallback to first non-id column
    for h in header:
        if 'id' not in h.lower():
            return h
    return header[0]


def load_db_items(db_csv_path):
    with open(db_csv_path, encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        header = reader.fieldnames
        name_col = detect_name_column(header)
        rows = list(reader)
    names = [r.get(name_col, '').strip() for r in rows]
    return rows, names, name_col


def normalize(s):
    return re.sub(r"[^0-9a-zA-Z가-힣\-\s\.\(\)]", ' ', s).strip().lower()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--db-csv', required=True)
    ap.add_argument('--out-csv', required=True)
    ap.add_argument('--threshold', type=float, default=80.0)
    ap.add_argument('md_files', nargs='+')
    args = ap.parse_args()

    # extract names from MD files
    doc_names = []
    for p in args.md_files:
        doc_names.extend(extract_names_from_md(p))
    # unique preserve order
    seen = set()
    doc_names_u = []
    for n in doc_names:
        nn = n.strip()
        if nn and nn not in seen:
            seen.add(nn)
            doc_names_u.append(nn)

    db_rows, db_names, db_name_col = load_db_items(args.db_csv)
    db_names_norm = [normalize(n) for n in db_names]

    # build mapping
    out_rows = []
    for doc in doc_names_u:
        best, score, idx = process.extractOne(normalize(doc), db_names_norm, scorer=fuzz.WRatio, score_cutoff=0)
        matched = False
        match_row = {}
        match_name = ''
        match_score = 0
        if best is not None:
            # find first matching index with that normalized name
            try:
                i = db_names_norm.index(best)
                match_row = db_rows[i]
                match_name = db_rows[i].get(db_name_col, '')
                match_score = score
                matched = score >= args.threshold
            except ValueError:
                pass

        out_rows.append({
            'doc_name': doc,
            'match_name': match_name,
            'match_score': match_score,
            'matched': 'yes' if matched else 'no',
            'db_item_no': match_row.get('ITEM_NO','') if match_row else '',
            'db_id': match_row.get('Id','') if match_row else ''
        })

    # write out CSV
    with open(args.out_csv, 'w', newline='', encoding='utf-8') as f:
        fieldnames = ['doc_name','match_name','match_score','matched','db_item_no','db_id']
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in out_rows:
            w.writerow(r)

    print(f'완료: {len(out_rows)}개 문서 시약을 비교하여 {args.out_csv} 생성')


if __name__ == '__main__':
    main()
