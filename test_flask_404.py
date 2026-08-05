from flask import Flask

app = Flask(__name__)

@app.errorhandler(Exception)
def handle_global_exception(e):
    import traceback
    traceback.print_exc()
    return '{"error":"An internal server error occurred."}', 500

def serve_tool_page(slug, lang=None):
    if slug == "html_to_pdf":
        return "Page not found.", 404
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
        res = client.get('/route_not_found')
        print(f"Status: {res.status_code}")
        print(f"Data: {res.data.decode()}")
