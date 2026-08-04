const fs = require('fs');

const indianLangsContent = fs.readFileSync('indian_langs.py', 'utf8');
let appContent = fs.readFileSync('app.py', 'utf8');

const insertPos = appContent.indexOf('def apply_language_translations');

const newContent = appContent.substring(0, insertPos) + 
                   indianLangsContent + 
                   '\nDYNAMIC_LANG_MAP.update(INDIAN_LANGS)\n\n' + 
                   appContent.substring(insertPos);

fs.writeFileSync('app.py', newContent, 'utf8');
console.log('Added INDIAN_LANGS to app.py');
