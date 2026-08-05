import os
from datetime import datetime

# Read SUPPORTED_LANGS and SLUG_TO_FILE from app.py
BASE_URL = "https://ilovespdfs.in"

SUPPORTED_LANGS = {'en': 'English', 'es': 'Español', 'fr': 'Français', 'de': 'Deutsch', 'pt': 'Português', 'hi': 
'हिन्दी', 'ar': 'العربية', 'zh': '中文', 'ja': '日本語', 'ko': '한국어', 'ru': 'Русский', 'it': 'Italiano', 'nl': 'Nederlands', 'tr': 'Türkçe', 'pl': 'Polski', 'sv': 'Svenska', 'id': 'Bahasa Indonesia', 'vi': 'Tiếng Việt', 'th': 'ไทย', 'uk': 'Українська', 'el': 'Ελληνικά', 'he': 
'עברית', 'bn': 'বাংলা', 'ta': 'தமிழ்', 'te': 'తెలుగు', 'mr': 'मराठी', 
'gu': 'ગુજરાતી', 'kn': 'ಕನ್ನಡ', 'ur': 'اردو', 'sw': 'Kiswahili', 'ms': 'Bahasa Melayu', 
'fil': 'Filipino', 'cs': 'Čeština', 'hu': 'Magyar', 'da': 'Dansk', 'fi': 'Suomi', 'sk': 'Slovenčina', 'ro': 
'Română', 'bg': 'Български', 'hr': 'Hrvatski', 'sr': 'Српски', 'sl': 'Slovenščina', 'et': 'Eesti', 
'lv': 'Latviešu', 'lt': 'Lietuvių', 'sq': 'Shqip', 'mk': 'Македонски', 'ka': 'ქართული', 'hy': 
'Հայերեն', 'az': 'Azərbaycan', 'kk': 'Қазақша', 'uz': 'Oʻzbekcha', 'mn': 'Монгол', 'my': 
'မြန်မာစာ', 'km': 'ខ្មែរ', 'lo': 'ລາວ', 'si': 'සිංහල', 
'ne': 'नेपाली', 'pa': 'ਪੰਜਾਬੀ', 'am': 'አማርኛ', 'so': 'Soomaali', 'yo': 'Yorùbá', 'ig': 
'Asụsụ Igbo', 'ha': 'Hausa', 'zu': 'isiZulu', 'xh': 'isiXhosa', 'af': 'Afrikaans', 'gl': 'Galego', 'ca': 'Català', 
'eu': 'Euskara', 'cy': 'Cymraeg', 'ga': 'Gaeilge', 'is': 'Íslenska', 'mt': 'Malti', 'lb': 'Lëtzebuergesch', 'bs': 
'Bosanski', 'fa': 'فارسی', 'ps': 'پښتو', 'sd': 'سنڌي', 'su': 'Basa Sunda', 'jv': 'Basa Jawa', 'ceb': 
'Cebuano', 'haw': 'ʻOlelo Hawaiʻi', 'eo': 'Esperanto', 'la': 'Latina', 'gd': 'Gàidhlig', 'fy': 'Frysk', 'co': 
'Corsu', 'mi': 'Te Reo Māori', 'sm': 'Gagana Sāmoa', 'ny': 'Chichewa', 'ht': 'Kreyòl Ayisyen', 'mg': 'Malagasy', 
'st': 'Sesotho', 'tn': 'Setswana', 've': 'Tshivenḓa'}

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
    'add_pdf_page_number': 'page-number.html',
    'html_to_pdf': 'html-to-pdf.html',
    'extract_text': 'extract-text.html',
    'remove-pages': 'remove-pages.html',
    'pdf_to_pdfa': 'pdf-to-pdfa.html',
    'repair_pdf': 'repair.html',
    'ocr_pdf': 'ocr-pdf.html',
    'ai_summary': 'ai-summary.html',
    'translate_pdf': 'translate-pdf.html',
    'sign_pdf': 'sign-pdf.html',
    'organize-pdf': 'organize-pdf.html',
    'compare-pdf': 'compare-pdf.html'
}

sitemap_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sitemap.xml')

today = datetime.now().strftime('%Y-%m-%d')

with open(sitemap_path, 'w', encoding='utf-8') as f:
    f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
    f.write('<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n')
    
    # Generate URLs for all languages and tools
    for lang in SUPPORTED_LANGS.keys():
        lang_prefix = f"/{lang}" if lang != 'en' else ""
        for slug in SLUG_TO_FILE.keys():
            path = f"{lang_prefix}/{slug}" if slug else lang_prefix
            if not path:
                path = "/"
            if not path.startswith('/'):
                path = "/" + path
            
            # Prioritize root index and popular tools
            priority = "1.0" if slug == "" else "0.8"
            if slug in ["merge_pdf", "split_pdf", "compress_pdf"]:
                priority = "0.9"
                
            f.write(f'  <url>\n')
            f.write(f'    <loc>{BASE_URL}{path}</loc>\n')
            f.write(f'    <lastmod>{today}</lastmod>\n')
            f.write(f'    <changefreq>weekly</changefreq>\n')
            f.write(f'    <priority>{priority}</priority>\n')
            f.write(f'  </url>\n')
            
    f.write('</urlset>\n')

print(f"Generated sitemap with over {len(SUPPORTED_LANGS) * len(SLUG_TO_FILE)} URLs at {sitemap_path}")
