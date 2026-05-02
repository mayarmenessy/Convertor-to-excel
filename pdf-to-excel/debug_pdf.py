"""
Run this script to debug what pdfplumber extracts from your PDF.
Usage: python debug_pdf.py "path\to\your\file.pdf"
"""
import sys
import pdfplumber

pdf_path = sys.argv[1] if len(sys.argv) > 1 else input("Enter PDF path: ").strip('"')

with pdfplumber.open(pdf_path) as pdf:
    print(f"\n=== PDF has {len(pdf.pages)} pages ===\n")

    # Check first 2 pages only
    for page_num in range(min(2, len(pdf.pages))):
        page = pdf.pages[page_num]
        print(f"\n{'='*60}")
        print(f"PAGE {page_num + 1}")
        print(f"{'='*60}")

        # Method 1: extract_tables
        tables = page.extract_tables()
        print(f"\n[Method 1] extract_tables() found {len(tables)} table(s)")
        for t_idx, table in enumerate(tables):
            print(f"  Table {t_idx+1}: {len(table)} rows x {len(table[0]) if table else 0} cols")
            for row in table[:5]:  # show first 5 rows
                print(f"    {row}")

        # Method 2: extract_text
        text = page.extract_text() or ""
        lines = [l for l in text.splitlines() if l.strip()]
        print(f"\n[Method 2] extract_text() found {len(lines)} lines")
        for line in lines[:10]:
            print(f"  >> {repr(line)}")

        # Method 3: extract_words
        words = page.extract_words()
        print(f"\n[Method 3] extract_words() found {len(words)} words")
        for w in words[:15]:
            print(f"  x0={w['x0']:.1f} top={w['top']:.1f} text={repr(w['text'])}")
