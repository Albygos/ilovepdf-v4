import os
import datetime

# From app.py
SLUG_TO_FILE = {
    '': 'index.html',
    'merge_pdf': 'merge.html',
    'split_pdf': 'split.html',
    'compress_pdf': 'compress.html',
    'word_to_pdf': 'word-to-pdf.html',
    'pdf_to_word': 'pdf-to-word.html',
    'pdf_to_powerpoint': 'pdf-to-powerpoint.html',
    'pdf_to_excel': 'pdf-to-excel.html',
    'powerpoint_to_pdf': 'powerpoint-to-pdf.html',
    'excel_to_pdf': 'excel-to-pdf.html',
    'jpg_to_pdf': 'jpg-to-pdf.html',
    'pdf_to_jpg': 'pdf-to-jpg.html',
    'rotate_pdf': 'rotate.html',
    'protect-pdf': 'protect.html',
    'unlock_pdf': 'unlock.html',
    'pdf_add_watermark': 'watermark.html',
    'add_pdf_page_number': 'page-numbers.html',
    'organize-pdf': 'organize.html',
    'html_to_pdf': 'html-to-pdf.html',
    'extract_text': 'extract-text.html',
    'remove-pages': 'delete-pages.html',
    'pdf_to_pdfa': 'pdf-to-pdfa.html',
    'repair_pdf': 'repair-pdf.html',
    'ocr_pdf': 'ocr-pdf.html',
    'pdf_summarize': 'pdf-summarize.html',
    'translate_pdf': 'translate-pdf.html',
    'sign_pdf': 'sign-pdf.html',
    'compare-pdf': 'compare-pdf.html'
}

INDIAN_LANGS = ['hi', 'bn', 'ta', 'te', 'mr', 'gu', 'kn', 'ur', 'pa', 'ml', 'or']
BASE_URL = 'https://ilovespdfs.in'

urls = []
date_str = datetime.datetime.now().strftime('%Y-%m-%d')

for lang in INDIAN_LANGS:
    for slug in SLUG_TO_FILE.keys():
        if slug == '':
            url = f"{BASE_URL}/{lang}/"
        else:
            url = f"{BASE_URL}/{lang}/{slug}"
            
        urls.append(f"""  <url>
    <loc>{url}</loc>
    <lastmod>{date_str}</lastmod>
    <changefreq>weekly</changefreq>
    <priority>0.9</priority>
  </url>""")

xml_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
{chr(10).join(urls)}
</urlset>"""

with open('sitemap-india.xml', 'w', encoding='utf-8') as f:
    f.write(xml_content)

print("Generated sitemap-india.xml with", len(urls), "URLs.")
