import re

# Read indian_langs
with open('indian_langs.py', 'r', encoding='utf-8') as f:
    indian_langs_content = f.read()

with open('app.py', 'r', encoding='utf-8') as f:
    app_content = f.read()

# Insert the dictionary right before apply_language_translations
insert_pos = app_content.find('def apply_language_translations')

new_content = app_content[:insert_pos] + indian_langs_content + '\nDYNAMIC_LANG_MAP.update(INDIAN_LANGS)\n\n' + app_content[insert_pos:]

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(new_content)

print("Added INDIAN_LANGS to app.py")
