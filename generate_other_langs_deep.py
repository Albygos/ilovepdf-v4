import json
import time
from deep_translator import GoogleTranslator

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

en_dict = {
    'title': 'i loves pdf | Free Online PDF Tools',
    'h1': 'All the PDF tools you need in one place',
    'desc': 'All PDF tools at your fingertips. 100% FREE and easy to use! Merge, split, compress, convert, rotate and unlock your PDFs with just a few clicks.',
    'all': 'All',
    'org': 'Organize PDF',
    'opt': 'Optimize PDF',
    'conv': 'Convert PDF',
    'sec': 'PDF Security',
    'intel': 'PDF Intelligence',
    'merge': 'Merge PDF',
    'split': 'Split PDF',
    'compress': 'Compress PDF',
    'w2p': 'Word to PDF',
    'p2w': 'PDF to Word',
    'p2pp': 'PDF to PowerPoint',
    'p2ex': 'PDF to Excel',
    'pp2p': 'PowerPoint to PDF',
    'ex2p': 'Excel to PDF',
    'j2p': 'JPG to PDF',
    'p2j': 'PDF to JPG',
    'rot': 'Rotate PDF',
    'prot': 'Protect PDF',
    'unl': 'Unlock PDF',
    'wm': 'Watermark',
    'pn': 'Page Numbers',
    'h2p': 'HTML to PDF',
    'txt': 'Extract Text',
    'del': 'Delete Pages',
    'pdfa': 'PDF to PDF/A',
    'rep': 'Repair PDF',
    'ocr': 'OCR PDF',
    'ai': 'AI Summary',
    'trans': 'Translate PDF',
    'sign': 'Sign PDF',
    'comp': 'Compare PDF',
    'login': 'Log In',
    'signup': 'Sign Up',
    'home': 'Home',
    'all_tools': 'All PDF Tools',
    'convert_pdf': 'Convert PDF'
}

EXISTING_LANGS = ['es', 'fr', 'de', 'pt', 'hi', 'ar', 'zh', 'ja', 'ru', 'it', 'bn', 'ta', 'te', 'mr', 'gu', 'kn', 'ur', 'pa', 'ml', 'or', 'en']
MISSING_LANGS = [lang for lang in SUPPORTED_LANGS if lang not in EXISTING_LANGS]

print(f"Translating for {len(MISSING_LANGS)} missing languages...")
all_translations = {}

for lang in MISSING_LANGS:
    print(f"Processing {lang}...")
    t_dict = {}
    
    # Deep translator expects language code format. We might need to map some if they fail.
    # GoogleTranslator automatically handles most of them.
    try:
        translator = GoogleTranslator(source='en', target=lang)
    except Exception as e:
        print(f"Skipping {lang} because: {e}")
        continue
        
    for key, value in en_dict.items():
        try:
            translated = translator.translate(value)
            translated = translated.replace('\\', '')
            t_dict[key] = translated
        except Exception as e:
            t_dict[key] = value
            
    all_translations[lang] = t_dict

with open('other_langs.py', 'w', encoding='utf-8') as f:
    f.write('OTHER_LANGS = {\n')
    for lang, t_dict in all_translations.items():
        f.write(f"    '{lang}': {{\n")
        for k, v in t_dict.items():
            v_escaped = v.replace("'", "\\'")
            f.write(f"        '{k}': '{v_escaped}',\n")
        f.write("    },\n")
    f.write('}\n')

print("Done! Translations saved to other_langs.py")
