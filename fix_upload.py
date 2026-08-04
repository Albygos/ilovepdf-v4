import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace request.files.get('file') with (request.files.get('file') or request.files.get('files'))
new_content = re.sub(
    r"request\.files\.get\('file'\)",
    r"(request.files.get('file') or request.files.get('files'))",
    content
)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(new_content)

print("Replaced instances in app.py")
