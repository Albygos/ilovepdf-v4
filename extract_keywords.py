import os
import json

KEYWORD_FILE = 'keyword_mapping.json'

exclude_htmls = {
    'index.html', 'merge.html', 'split.html', 'compress.html', 'word-to-pdf.html', 'pdf-to-word.html', 
    'jpg-to-pdf.html', 'pdf-to-jpg.html', 'rotate.html', 'protect.html', 'unlock.html', 'watermark.html', 
    'page-numbers.html', 'organize.html', 'html-to-pdf.html', 'extract-text.html', 'delete-pages.html', 
    'compare.html', 'powerpoint-to-pdf.html', 'pdf-to-powerpoint.html', 'excel-to-pdf.html', 
    'pdf-to-excel.html', 'pdf-to-pdfa.html', 'repair-pdf.html', 'ocr-pdf.html', 'pdf-summarize.html', 
    'translate-pdf.html', 'sign-pdf.html', 'privacy-policy.html', 'terms.html', 'about.html', 
    'contact.html', 'cookie-policy.html', 'disclaimer.html', '404.html', '500.html'
}

def extract_keywords():
    mapping = {}
    files = [f.name for f in os.scandir('.') if f.is_file() and f.name.endswith('.html') and f.name not in exclude_htmls]
    
    total = len(files)
    print(f"Processing {total} files...", flush=True)
    
    for i, filename in enumerate(files):
        slug = filename[:-5]
        
        try:
            with open(filename, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(4000)
        except Exception as e:
            print(f"Error reading {filename}: {e}", flush=True)
            continue
            
        title = ""
        desc = ""
        target_tool = "index.html"
        
        t_start = content.find("<title>")
        if t_start != -1:
            t_end = content.find("</title>", t_start)
            if t_end != -1:
                title = content[t_start+7:t_end]
                
        d_start = content.find('<meta name="description" content="')
        if d_start != -1:
            d_end = content.find('">', d_start + 34)
            if d_end != -1:
                desc = content[d_start+34:d_end]
                
        c_start = content.find('<link rel="canonical" href="https://ilovespdfs.in/')
        if c_start != -1:
            c_end = content.find('.html"', c_start + 50)
            if c_end != -1:
                target_tool = content[c_start+50:c_end] + ".html"
                
        mapping[slug] = {
            'target': target_tool,
            'title': title,
            'desc': desc
        }
        
        if (i+1) % 2000 == 0:
            print(f"Processed {i+1}/{total} files...", flush=True)
            
    with open(KEYWORD_FILE, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False)
        
    print(f"Saved {len(mapping)} mappings to {KEYWORD_FILE}", flush=True)

if __name__ == '__main__':
    extract_keywords()
