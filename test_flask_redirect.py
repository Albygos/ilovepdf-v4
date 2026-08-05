from flask import Flask, redirect

app = Flask(__name__)

legacy_redirects = {
    'split': 'split_pdf'
}

def serve_tool_page(slug, lang=None):
    if slug in legacy_redirects:
        target = legacy_redirects[slug]
        url = f"/{lang}/{target}" if lang else f"/{target}"
        return redirect(url, code=301)
    return f"Served {slug} for {lang}"

@app.route('/<path:slug>')
def dynamic_seo_page(slug):
    slug = slug.strip('/')
    if '/' in slug:
        parts = slug.split('/')
        return serve_tool_page(parts[1], parts[0])
    return serve_tool_page(slug, None)

@app.route('/<lang>/<slug>')
def lang_tool_route(lang, slug):
    return dynamic_seo_page(f"{lang}/{slug}")

if __name__ == '__main__':
    with app.test_client() as client:
        print("Testing /split ...")
        res = client.get('/split')
        print(f"Status: {res.status_code}, Location: {res.headers.get('Location')}")
        
        print("\nTesting /hi/split ...")
        res = client.get('/cy/html_to_pdf')
        print(f"Status: {res.status_code}, Location: {res.headers.get('Location')}")
        
        if res.status_code >= 400:
            print("Error data:", res.data)
