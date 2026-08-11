import re
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(BASE_DIR, 'index.html'), 'r', encoding='utf-8') as f:
    html = f.read()

keyword = '100-free-add-page-numbers-to-pdf-software'
clean_title_h1 = keyword.replace('-', ' ').title().replace('100 Free', '100% Free')
desc = "Looking for 100% Free Add Page Numbers To Pdf Software?"

html = re.sub(
    r'<h1 class="f-heading1[^>]*>.*?</h1>', 
    f'<h1 class="f-heading1everytoolyouneedtoworkwithpdfsinoneplace-1006461">{clean_title_h1}</h1>', 
    html, 
    flags=re.IGNORECASE | re.DOTALL
)

html = re.sub(
    r'<h2 class="f-heading2[^>]*>.*?</h2>', 
    f'<h2 class="f-heading2everytoolyouneedtousepdfsatyourfingertipsallare100freeandeasytousemergesplitcompressconvertrotateunlockandwatermarkpdfswithjustafewclicks-1006462">{desc}</h2>', 
    html, 
    flags=re.IGNORECASE | re.DOTALL
)

match1 = re.search(r'<h1 class="f-heading1[^>]*>.*?</h1>', html, flags=re.IGNORECASE | re.DOTALL)
print("H1 Match:", match1.group(0) if match1 else "None")

match2 = re.search(r'<h2 class="f-heading2[^>]*>.*?</h2>', html, flags=re.IGNORECASE | re.DOTALL)
print("H2 Match:", match2.group(0) if match2 else "None")
