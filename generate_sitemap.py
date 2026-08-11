import os
from datetime import datetime
import traceback

with open('sitemap_debug.log', 'w') as log:
    try:
        log.write("Starting script\n")
        DOMAIN = 'https://ilovespdfs.in'

        TOOLS = [
            'merge', 'split', 'compress', 'word-to-pdf', 'powerpoint-to-pdf', 'excel-to-pdf', 
            'jpg-to-pdf', 'pdf-to-word', 'pdf-to-powerpoint', 'pdf-to-excel', 'pdf-to-jpg', 
            'html-to-pdf', 'rotate', 'protect', 'unlock', 'watermark', 'page-numbers', 
            'organize', 'delete-pages', 'convert-pdf-to-pdfa', 'repair-pdf', 'ocr-pdf', 
            'pdf-summarize', 'translate-pdf', 'sign-pdf', 'compare-pdf', 'extract-text'
        ]

        LANGUAGES = ['es', 'fr', 'de', 'pt', 'hi', 'ar', 'zh', 'ja', 'ko', 'ru', 'it', 'nl', 'tr', 'pl', 'sv', 'id', 'vi', 'th', 'uk', 'el', 'he', 'bn', 'ta', 'te', 'mr', 'gu', 'kn', 'ur', 'sw', 'ms', 'fil', 'cs', 'hu', 'da', 'fi', 'sk', 'ro', 'bg', 'hr', 'sr', 'sl', 'et', 'lv', 'lt', 'sq', 'mk', 'ka', 'hy', 'az', 'kk', 'uz', 'mn', 'my', 'km', 'lo', 'si', 'ne', 'pa', 'am', 'so', 'yo', 'ig', 'ha', 'zu', 'xh', 'af', 'gl', 'ca', 'eu', 'cy', 'ga', 'is', 'mt', 'lb', 'bs', 'fa', 'ps', 'sd', 'su', 'jv', 'ceb', 'haw', 'eo', 'la', 'gd', 'fy', 'co', 'mi', 'sm', 'ny', 'ht', 'mg', 'st', 'tn', 've']

        indian_langs_set = {'hi', 'bn', 'ta', 'te', 'mr', 'gu', 'kn', 'ur', 'ml', 'pa', 'or', 'as', 'mai', 'sat', 'ks'}

        STATIC_PAGES = ['privacy-policy', 'terms', 'about', 'contact', 'cookie-policy', 'disclaimer']

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
        if os.path.exists('seo_keywords.txt'):
            with open('seo_keywords.txt', 'r', encoding='utf-16') as f:
                for line in f:
                    slug = line.strip()
                    if slug:
                        seo_articles.append(slug)

        core_urls = []
        language_urls = []
        article_urls = []

        # 1. Homepage
        core_urls.append((f"{DOMAIN}/", '1.0', 'daily'))
        for lang in LANGUAGES:
            language_urls.append((f"{DOMAIN}/{lang}/", '1.0', 'daily'))
            if lang in indian_langs_set:
                language_urls.append((f"{DOMAIN}/{lang}-in/", '1.0', 'daily'))

        # 2. Tools
        for tool in TOOLS:
            core_urls.append((f"{DOMAIN}/{tool}", '0.9', 'weekly'))
            for lang in LANGUAGES:
                language_urls.append((f"{DOMAIN}/{lang}/{tool}", '0.9', 'weekly'))
                if lang in indian_langs_set:
                    language_urls.append((f"{DOMAIN}/{lang}-in/{tool}", '0.9', 'weekly'))

        # 3. Static Pages
        for page in STATIC_PAGES:
            core_urls.append((f"{DOMAIN}/{page}", '0.5', 'monthly'))
            for lang in LANGUAGES:
                language_urls.append((f"{DOMAIN}/{lang}/{page}", '0.5', 'monthly'))
                if lang in indian_langs_set:
                    language_urls.append((f"{DOMAIN}/{lang}-in/{page}", '0.5', 'monthly'))

        # 4. SEO Articles (Generate separate lists for each language because of 50k limit per sitemap)
        # We will dynamically generate the sitemap names and URL lists
        article_sitemaps = {}
        for lang in ['en'] + LANGUAGES:
            article_sitemaps[lang] = []
            
        for lang_in in indian_langs_set:
            article_sitemaps[f"{lang_in}-in"] = []
            
        for article in seo_articles:
            # English goes to /en/
            article_sitemaps['en'].append((f"{DOMAIN}/en/{article}", '0.8', 'weekly'))
            
            # Other languages
            for lang in LANGUAGES:
                article_sitemaps[lang].append((f"{DOMAIN}/{lang}/{article}", '0.8', 'weekly'))
                if lang in indian_langs_set:
                    article_sitemaps[f"{lang}-in"].append((f"{DOMAIN}/{lang}-in/{article}", '0.8', 'weekly'))

        today = datetime.now().strftime('%Y-%m-%d')

        def write_sitemap(filename, urls):
            xml_lines = ['<?xml version="1.0" encoding="UTF-8"?>', '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
            for loc, priority, changefreq in urls:
                xml_lines.append('  <url>')
                xml_lines.append(f'    <loc>{loc}</loc>')
                xml_lines.append(f'    <lastmod>{today}</lastmod>')
                xml_lines.append(f'    <changefreq>{changefreq}</changefreq>')
                xml_lines.append(f'    <priority>{priority}</priority>')
                xml_lines.append('  </url>')
            xml_lines.append('</urlset>')
            with open(filename, 'w', encoding='utf-8') as f:
                f.write('\n'.join(xml_lines))

        # Write the individual sitemaps
        write_sitemap('sitemap-core.xml', core_urls)
        write_sitemap('sitemap-languages.xml', language_urls)
        
        sitemap_files_to_index = ['sitemap-core.xml', 'sitemap-languages.xml']
        
        # Create a sitemaps directory to avoid cluttering the root folder
        sitemaps_dir = 'sitemaps'
        if not os.path.exists(sitemaps_dir):
            os.makedirs(sitemaps_dir)
        
        # Write article sitemaps per language into the subfolder
        for lang_key, urls in article_sitemaps.items():
            if urls:
                # E.g. "sitemaps/sitemap-articles-en.xml"
                filename = f'{sitemaps_dir}/sitemap-articles-{lang_key}.xml'
                write_sitemap(filename, urls)
                # Google index needs the full path including the folder name
                sitemap_files_to_index.append(filename)

        # Generate Sitemap Index
        index_lines = ['<?xml version="1.0" encoding="UTF-8"?>', '<sitemapindex xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
        for sitemap_name in sitemap_files_to_index:
            index_lines.append('  <sitemap>')
            index_lines.append(f'    <loc>{DOMAIN}/{sitemap_name}</loc>')
            index_lines.append(f'    <lastmod>{today}</lastmod>')
            index_lines.append('  </sitemap>')
        index_lines.append('</sitemapindex>')
        
        with open('sitemap.xml', 'w', encoding='utf-8') as f:
            f.write('\n'.join(index_lines))

        log.write("Done\n")
    except Exception as e:
        log.write(traceback.format_exc())
