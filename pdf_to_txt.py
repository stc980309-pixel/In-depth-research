#!/usr/bin/env python3
"""Convert all PDFs in a directory to plain text files using PyMuPDF."""
import sys, os, fitz

def pdf_to_txt(pdf_dir, merge=False):
    pdf_dir = os.path.abspath(pdf_dir)
    pdfs = sorted([f for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')])
    if not pdfs:
        print(f"No PDFs found in {pdf_dir}")
        return

    all_text = []
    for pdf_file in pdfs:
        pdf_path = os.path.join(pdf_dir, pdf_file)
        txt_path = os.path.splitext(pdf_path)[0] + '.txt'
        try:
            doc = fitz.open(pdf_path)
            lines = [f"=== {pdf_file} ===\n"]
            for i, page in enumerate(doc):
                text = page.get_text()
                if text.strip():
                    lines.append(f"--- Page {i+1} ---\n{text}\n")
            doc.close()
            full_text = ''.join(lines)
            if not merge:
                with open(txt_path, 'w', encoding='utf-8') as f:
                    f.write(full_text)
                print(f"  -> {os.path.basename(txt_path)} ({len(full_text)} chars)")
            else:
                all_text.append(full_text)
        except Exception as e:
            print(f"  ERROR {pdf_file}: {e}")

    if merge and all_text:
        merged_path = os.path.join(pdf_dir, 'all_papers.txt')
        with open(merged_path, 'w', encoding='utf-8') as f:
            f.write('\n\n'.join(all_text))
        print(f"  -> {merged_path}")

if __name__ == '__main__':
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument('pdf_dir', help='Directory containing PDF files')
    p.add_argument('--merge', action='store_true', help='Merge all into single all_papers.txt')
    args = p.parse_args()
    pdf_to_txt(args.pdf_dir, args.merge)
