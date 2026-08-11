import os
import glob
import time

def fix_upload():
    with open('fix_upload_debug.log', 'w') as log:
        try:
            log.write("Starting script\n")
            html_files = glob.glob('*.html')
            log.write(f"Found {len(html_files)} HTML files\n")
            
            target = "dropzone.addEventListener('click', () => fileSelector.click());"
            replacement = "dropzone.addEventListener('click', (e) => { if (e.target !== fileSelector) fileSelector.click(); });"
            
            count = 0
            for i, file in enumerate(html_files):
                if i % 1000 == 0:
                    log.write(f"Processed {i} files...\n")
                    log.flush()
                try:
                    with open(file, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    if target in content:
                        content = content.replace(target, replacement)
                        with open(file, 'w', encoding='utf-8') as f:
                            f.write(content)
                        count += 1
                except Exception as e:
                    log.write(f"Error processing {file}: {e}\n")
                    
            log.write(f"Done! Fixed {count} files.\n")
        except Exception as e:
            log.write(f"FATAL ERROR: {e}\n")

if __name__ == '__main__':
    fix_upload()
