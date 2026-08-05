import os
from datetime import datetime
import traceback

with open('sitemap_debug.log', 'w') as log:
    try:
        log.write("Starting script\n")
        # Domain
        DOMAIN = 'https://ilovespdfs.in'

        TOOLS = [
            'merge', 'split', 'compress', 'word-to-pdf', 'powerpoint-to-pdf', 'excel-to-pdf', 
            'jpg-to-pdf', 'pdf-to-word', 'pdf-to-powerpoint', 'pdf-to-excel', 'pdf-to-jpg', 
            'html-to-pdf', 'rotate', 'protect', 'unlock', 'watermark', 'page-numbers', 
            'organize', 'delete-pages', 'convert-pdf-to-pdfa', 'repair-pdf', 'ocr-pdf', 
            'pdf-summarize', 'translate-pdf', 'sign-pdf', 'compare-pdf', 'extract-text'
        ]

        # Supported languages
        LANGUAGES = ['es', 'fr', 'de', 'pt', 'hi', 'ar', 'zh', 'ja', 'ko', 'ru', 'it', 'nl', 'tr', 'pl', 'sv', 'id', 'vi', 'th', 'uk', 'el', 'he', 'bn', 'ta', 'te', 'mr', 'gu', 'kn', 'ur', 'sw', 'ms', 'fil', 'cs', 'hu', 'da', 'fi', 'sk', 'ro', 'bg', 'hr', 'sr', 'sl', 'et', 'lv', 'lt', 'sq', 'mk', 'ka', 'hy', 'az', 'kk', 'uz', 'mn', 'my', 'km', 'lo', 'si', 'ne', 'pa', 'am', 'so', 'yo', 'ig', 'ha', 'zu', 'xh', 'af', 'gl', 'ca', 'eu', 'cy', 'ga', 'is', 'mt', 'lb', 'bs', 'fa', 'ps', 'sd', 'su', 'jv', 'ceb', 'haw', 'eo', 'la', 'gd', 'fy', 'co', 'mi', 'sm', 'ny', 'ht', 'mg', 'st', 'tn', 've']

        indian_langs_set = {'hi', 'bn', 'ta', 'te', 'mr', 'gu', 'kn', 'ur', 'ml', 'pa', 'or', 'as', 'mai', 'sat', 'ks'}

        STATIC_PAGES = ['privacy-policy', 'terms', 'about', 'contact', 'cookie-policy', 'disclaimer']

        # Exclude list for SEO articles
        exclude_htmls = set([
            'index.html', 'merge.html', 'split.html', 'compress.html', 'word-to-pdf.html', 'pdf-to-word.html', 
            'jpg-to-pdf.html', 'pdf-to-jpg.html', 'rotate.html', 'protect.html', 'unlock.html', 'watermark.html', 
            'page-numbers.html', 'organize.html', 'html-to-pdf.html', 'extract-text.html', 'delete-pages.html', 
            'compare.html', 'powerpoint-to-pdf.html', 'pdf-to-powerpoint.html', 'excel-to-pdf.html', 
            'pdf-to-excel.html', 'pdf-to-pdfa.html', 'repair-pdf.html', 'ocr-pdf.html', 'pdf-summarize.html', 
            'translate-pdf.html', 'sign-pdf.html', 'privacy-policy.html', 'terms.html', 'about.html', 
            'contact.html', 'cookie-policy.html', 'disclaimer.html', '404.html', '500.html'
        ])

        seo_articles = []
        log.write("Listing directory\n")
        files = os.listdir('.')
        log.write(f"Found {len(files)} files\n")
        for f in files:
            if f.endswith('.html') and f not in exclude_htmls:
                seo_articles.append(f.replace('.html', ''))

        urls = []

        log.write(f"Building urls (found {len(seo_articles)} articles)\n")
        # 1. Homepage
        urls.append((f"{DOMAIN}/", '1.0', 'daily'))
        for lang in LANGUAGES:
            urls.append((f"{DOMAIN}/{lang}/", '1.0', 'daily'))
            if lang in indian_langs_set:
                urls.append((f"{DOMAIN}/{lang}-in/", '1.0', 'daily'))

        # 2. Tools
        for tool in TOOLS:
            urls.append((f"{DOMAIN}/{tool}", '0.9', 'weekly'))
            for lang in LANGUAGES:
                urls.append((f"{DOMAIN}/{lang}/{tool}", '0.9', 'weekly'))
                if lang in indian_langs_set:
                    urls.append((f"{DOMAIN}/{lang}-in/{tool}", '0.9', 'weekly'))

        # 3. Static Pages
        for page in STATIC_PAGES:
            urls.append((f"{DOMAIN}/{page}", '0.5', 'monthly'))
            for lang in LANGUAGES:
                urls.append((f"{DOMAIN}/{lang}/{page}", '0.5', 'monthly'))
                if lang in indian_langs_set:
                    urls.append((f"{DOMAIN}/{lang}-in/{page}", '0.5', 'monthly'))

        # 4. SEO Articles (ONLY english, no translations)
        for article in seo_articles:
            urls.append((f"{DOMAIN}/{article}", '0.8', 'weekly'))

        log.write(f"Generating XML string for {len(urls)} urls\n")
        
        # Use list for fast joining
        xml_lines = ['<?xml version="1.0" encoding="UTF-8"?>', '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
        today = datetime.now().strftime('%Y-%m-%d')
        
        for loc, priority, changefreq in urls:
            xml_lines.append('  <url>')
            xml_lines.append(f'    <loc>{loc}</loc>')
            xml_lines.append(f'    <lastmod>{today}</lastmod>')
            xml_lines.append(f'    <changefreq>{changefreq}</changefreq>')
            xml_lines.append(f'    <priority>{priority}</priority>')
            xml_lines.append('  </url>')

        xml_lines.append('</urlset>')
        
        log.write("Writing to sitemap.xml\n")
        with open('sitemap.xml', 'w', encoding='utf-8') as f:
            f.write('\n'.join(xml_lines))

        log.write("Done\n")
    except Exception as e:
        log.write(traceback.format_exc())
