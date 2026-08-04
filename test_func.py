import re

DYNAMIC_LANG_MAP = {'hi': {'desc': 'Test desc', 'title': 'Test title', 'h1': 'Test h1'}}

def apply_language_translations(html, lang):
    html = html.replace('href="./style.css', 'href="/style.css')
    
    if not lang or lang == 'en':
        return html
        
    base_lang = lang.lower().split('-')[0]
    
    t_dict = DYNAMIC_LANG_MAP.get(base_lang)
    if not t_dict:
        return html
        
    static_prefixes = ('favicon', 'apple', 'site.webmanifest', 'logo', 'ad-', 'style.css', 'api/', 'download/')
    html = re.sub(
        r'href="/([^"]*)"', 
        lambda m: f'href="/{lang}/{m.group(1)}"' if not m.group(1).startswith(static_prefixes) else m.group(0), 
        html
    )
    
    html = re.sub(
        r'window\.location\.href=[\'"]\/[\'"]', 
        f'window.location.href="/{lang}/"', 
        html,
        flags=re.IGNORECASE
    )
        
    html = re.sub(r'''<html\s+lang=["'][^"']*["']''', f'<html lang="{lang}"', html, flags=re.IGNORECASE)
    
    if 'title' in t_dict:
        html = re.sub(r'<title>.*?</title>', f'<title>{t_dict["title"]}</title>', html, flags=re.DOTALL)
    if 'desc' in t_dict:
        html = re.sub(r'''<meta\s+name=["']description["']\s+content=["'][^"']*["']''', f'<meta name="description" content="{t_dict["desc"]}">', html, flags=re.IGNORECASE)
        
    if 'h1' in t_dict:
        html = html.replace('Every tool you need to work with PDFs in one place', t_dict['h1'])
    if 'desc' in t_dict:
        html = html.replace('Every tool you need to use PDFs, at your fingertips. All are 100% FREE and easy to use! Merge, split, compress, convert, rotate, unlock and watermark PDFs with just a few clicks.', t_dict['desc'])
        
    replacements = [
        ('Merge PDF', t_dict.get('merge', 'Merge PDF')),
    ]
    
    for old_str, new_str in replacements:
        html = html.replace(old_str, new_str)
        
    # Advanced Regional SEO Schema Injection for Indian Languages
    indian_langs = ['hi', 'bn', 'ta', 'te', 'mr', 'gu', 'kn', 'ur', 'pa', 'ml', 'or']
    if base_lang in indian_langs:
        schema = f"""<script type="application/ld+json">
{{
  "@context": "https://schema.org",
  "@type": "WebSite",
  "name": "i loves pdfs",
  "url": "https://ilovespdfs.in/{lang}/",
  "inLanguage": "{lang}",
  "description": "{t_dict.get('desc', '')}"
}}
</script>
<script type="application/ld+json">
{{
  "@context": "https://schema.org",
  "@type": "Service",
  "name": "i loves pdfs",
  "areaServed": "IN",
  "description": "{t_dict.get('desc', '')}",
  "url": "https://ilovespdfs.in/{lang}/"
}}
</script>"""
        html = re.sub(r'</head>', schema + '\n</head>', html, flags=re.IGNORECASE)
        
    return html

print(apply_language_translations('<head></head>', 'hi'))
