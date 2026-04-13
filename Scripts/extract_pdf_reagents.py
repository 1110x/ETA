#!/usr/bin/env python3
"""PDF에서 시약 사용량(숫자+단위)을 추출해 CSV로 저장합니다.

의존: pdfminer.six

출력: Docs/reagent_amounts_extracted.csv
"""
import re
import csv
from pathlib import Path
from pdfminer.high_level import extract_text

ROOT = Path(__file__).resolve().parents[1]
PDF_DIR = ROOT / 'Docs' / '수질오염공정시험기준 전문(251226 개정)'
OUT_CSV = ROOT / 'Docs' / 'reagent_amounts_extracted.csv'

unit_rx = re.compile(r'(?P<amount>\d+[\.,]?\d*)\s*(?P<unit>mL|ml|㎖|L|g|mg|μg|ug|µL|uL|cc)\b', re.IGNORECASE)
es_rx = re.compile(r'ES\s*0?\d{3,4}', re.IGNORECASE)

def extract_from_pdf(path: Path):
    try:
        txt = extract_text(str(path))
    except Exception as e:
        return []
    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    results = []
    current_es = ''
    for i, line in enumerate(lines):
        m_es = es_rx.search(line)
        if m_es:
            current_es = m_es.group(0)
        for m in unit_rx.finditer(line):
            amount = m.group('amount')
            unit = m.group('unit')
            # capture surrounding context (prev and next tokens)
            prev = lines[i-1] if i>0 else ''
            nxt = lines[i+1] if i+1<len(lines) else ''
            context = ' '.join([prev, line, nxt]).strip()
            results.append({'pdf': path.name, 'es': current_es, 'line': line, 'context': context, 'amount': amount, 'unit': unit})
    return results

def main():
    pdfs = sorted(PDF_DIR.glob('**/*.pdf'))
    allrows = []
    for p in pdfs:
        rows = extract_from_pdf(p)
        allrows.extend(rows)

    with open(OUT_CSV, 'w', newline='', encoding='utf-8') as f:
        fieldnames = ['pdf','es','line','context','amount','unit']
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in allrows:
            w.writerow(r)

    print(f'추출 완료: {len(allrows)}개 항목 -> {OUT_CSV}')

if __name__ == '__main__':
    main()
