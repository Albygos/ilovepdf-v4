import os
import glob
from multiprocessing import Pool

def process_file(file):
    target = "dropzone.addEventListener('click', () => fileSelector.click());"
    replacement = "dropzone.addEventListener('click', (e) => { if (e.target !== fileSelector) fileSelector.click(); });"
    try:
        with open(file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        if target in content:
            content = content.replace(target, replacement)
            with open(file, 'w', encoding='utf-8') as f:
                f.write(content)
            return 1
    except Exception:
        pass
    return 0

def fix_upload():
    print("Starting script...")
    html_files = glob.glob('*.html')
    print(f"Found {len(html_files)} HTML files")
    
    with Pool(os.cpu_count()) as p:
        results = p.map(process_file, html_files)
        
    print(f"Done! Fixed {sum(results)} files.")

if __name__ == '__main__':
    fix_upload()
