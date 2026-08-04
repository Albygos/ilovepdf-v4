const fs = require('fs');
const path = require('path');

const filepath = path.join(__dirname, 'app.py');
let content = fs.readFileSync(filepath, 'utf8');

// Replace request.files.get('file') with (request.files.get('file') or request.files.get('files'))
const newContent = content.replace(/request\.files\.get\('file'\)/g, "(request.files.get('file') or request.files.get('files'))");

fs.writeFileSync(filepath, newContent, 'utf8');
console.log("Replaced instances in app.py");
