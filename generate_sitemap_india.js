const fs = require('fs');

const SLUG_TO_FILE = {
    '': 'index.html',
    'merge_pdf': 'merge.html',
    'split_pdf': 'split.html',
    'compress_pdf': 'compress.html',
    'word_to_pdf': 'word-to-pdf.html',
    'pdf_to_word': 'pdf-to-word.html',
    'pdf_to_powerpoint': 'pdf-to-powerpoint.html',
    'pdf_to_excel': 'pdf-to-excel.html',
    'powerpoint_to_pdf': 'powerpoint-to-pdf.html',
    'excel_to_pdf': 'excel-to-pdf.html',
    'jpg_to_pdf': 'jpg-to-pdf.html',
    'pdf_to_jpg': 'pdf-to-jpg.html',
    'rotate_pdf': 'rotate.html',
    'protect-pdf': 'protect.html',
    'unlock_pdf': 'unlock.html',
    'pdf_add_watermark': 'watermark.html',
    'add_pdf_page_number': 'page-numbers.html',
    'organize-pdf': 'organize.html',
    'html_to_pdf': 'html-to-pdf.html',
    'extract_text': 'extract-text.html',
    'remove-pages': 'delete-pages.html',
    'pdf_to_pdfa': 'pdf-to-pdfa.html',
    'repair_pdf': 'repair-pdf.html',
    'ocr_pdf': 'ocr-pdf.html',
    'pdf_summarize': 'pdf-summarize.html',
    'translate_pdf': 'translate-pdf.html',
    'sign_pdf': 'sign-pdf.html',
    'compare-pdf': 'compare-pdf.html'
};

const INDIAN_LANGS = ['hi', 'bn', 'ta', 'te', 'mr', 'gu', 'kn', 'ur', 'pa', 'ml', 'or'];
const BASE_URL = 'https://ilovespdfs.in';

let urls = [];
const dateStr = new Date().toISOString().split('T')[0];

for (const lang of INDIAN_LANGS) {
    for (const slug of Object.keys(SLUG_TO_FILE)) {
        const url = slug === '' ? `${BASE_URL}/${lang}/` : `${BASE_URL}/${lang}/${slug}`;
        urls.push(`  <url>\n    <loc>${url}</loc>\n    <lastmod>${dateStr}</lastmod>\n    <changefreq>weekly</changefreq>\n    <priority>0.9</priority>\n  </url>`);
    }
}

const xmlContent = `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n${urls.join('\n')}\n</urlset>`;

fs.writeFileSync('sitemap-india.xml', xmlContent, 'utf8');
console.log(`Generated sitemap-india.xml with ${urls.length} URLs.`);
