#!/usr/bin/env python3
"""Generate a sample-set directory with PDFs and an Excel roster.

Usage:
    python scripts/create_sample_set.py <gmail-address>

The Gmail address is only used for its local part. The script will create
`sample-set/feedback-files` filled with demo PDFs plus a matching
`sample-set/student-sampleset.xlsx` roster.
"""
from __future__ import annotations

import argparse
import random
import shutil
from itertools import product
from pathlib import Path
from xml.sax.saxutils import escape
import zipfile

FIRST_NAMES = [
    'Liam', 'Noah', 'Olivia', 'Emma', 'Ava', 'Sophia', 'Mason', 'Ethan',
    'Isabella', 'Mia', 'Lucas', 'Logan', 'Harper', 'Charlotte', 'Amelia',
    'Evelyn', 'Henry', 'Sebastian', 'Luna', 'Ella'
]

LAST_NAMES = [
    'Anderson', 'Bennett', 'Carter', 'Diaz', 'Edwards', 'Foster', 'Garcia',
    'Harrison', 'Iverson', 'Jacobs', 'Kensington', 'Lopez', 'Montgomery',
    'Novak', 'Owens', 'Patel', 'Quincy', 'Reynolds', 'Santiago', 'Turner'
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description='Generate sample student set')
    parser.add_argument('gmail', help='Gmail address (or local part) used for plus addressing')
    return parser.parse_args()


def normalize_email_base(value: str) -> tuple[str, str]:
    if '@' in value:
        local, domain = value.split('@', 1)
    else:
        local, domain = value, 'gmail.com'
    local = local.strip()
    domain = domain.strip() or 'gmail.com'
    if not local:
        raise ValueError('Invalid Gmail address provided')
    return local, domain


def slugify(value: str) -> str:
    return ''.join(ch if ch.isalnum() else '-' for ch in value.lower()).strip('-').replace('--', '-')


def create_pdf(path: Path, text: str) -> None:
    content = f"BT /F1 24 Tf 72 720 Td ({text}) Tj ET"
    content_bytes = content.encode('utf-8')
    objects = [
        b"1 0 obj<< /Type /Catalog /Pages 2 0 R >>endobj\n",
        b"2 0 obj<< /Type /Pages /Kids [3 0 R] /Count 1 >>endobj\n",
        b"3 0 obj<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>endobj\n",
        b"5 0 obj<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>endobj\n",
    ]
    stream = b"4 0 obj<< /Length " + str(len(content_bytes)).encode() + b" >>stream\n" + content_bytes + b"\nendstream endobj\n"
    objects.insert(3, stream)

    pdf = bytearray()
    header = b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"
    pdf.extend(header)
    offsets = [0]
    offset = len(header)
    for obj in objects:
        offsets.append(offset)
        pdf.extend(obj)
        offset += len(obj)
    xref_pos = len(pdf)
    pdf.extend(f"xref\n0 {len(objects) + 1}\n".encode())
    pdf.extend(b"0000000000 65535 f \n")
    running = len(header)
    for obj in objects:
        pdf.extend(f"{running:010d} 00000 n \n".encode())
        running += len(obj)
    pdf.extend(b"trailer<< /Size " + str(len(objects) + 1).encode() + b" /Root 1 0 R >>\nstartxref\n")
    pdf.extend(str(xref_pos).encode())
    pdf.extend(b"\n%%EOF")
    path.write_bytes(pdf)


def build_roster(rows: list[tuple[str, str, str, str]], xlsx_path: Path) -> None:
    sheet_rows = [('firstname', 'lastname', 'email', 'studentid')] + rows
    cols = ['A', 'B', 'C', 'D']
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData>'
    ]
    for r_idx, row in enumerate(sheet_rows, start=1):
        lines.append(f'<row r="{r_idx}">')
        for c_idx, value in enumerate(row):
            col = cols[c_idx]
            val = escape(str(value))
            lines.append(f'<c r="{col}{r_idx}" t="inlineStr"><is><t>{val}</t></is></c>')
        lines.append('</row>')
    lines.extend(['</sheetData>', '</worksheet>'])

    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Students" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>'''

    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

    wb_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>'''

    with zipfile.ZipFile(xlsx_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types)
        zf.writestr('_rels/.rels', rels)
        zf.writestr('xl/workbook.xml', workbook_xml)
        zf.writestr('xl/_rels/workbook.xml.rels', wb_rels)
        zf.writestr('xl/worksheets/sheet1.xml', '\n'.join(lines))


def main() -> None:
    args = parse_args()
    local_part, domain = normalize_email_base(args.gmail)

    root = Path('sample-set')
    if root.exists():
        shutil.rmtree(root)
    files_dir = root / 'feedback-files'
    files_dir.mkdir(parents=True, exist_ok=True)

    random.seed(42)
    pairs = list(product(FIRST_NAMES, LAST_NAMES))
    random.shuffle(pairs)
    selected = pairs[:60]
    records = []

    for idx, (first, last) in enumerate(selected, start=1):
        slug = '-'.join(filter(None, [slugify(first), slugify(last)])) or f'student-{idx:02d}'
        pdf_path = files_dir / f"{slug}.pdf"
        create_pdf(pdf_path, f"Feedback for {first} {last}")
        email = f"{local_part}+{slug}-{idx:02d}@{domain}"
        student_id = f"S{idx:04d}"
        records.append((first, last, email, student_id))

    roster_path = root / 'student-sampleset.xlsx'
    build_roster(records, roster_path)
    print(f"Created sample set in {root}")


if __name__ == '__main__':
    main()
