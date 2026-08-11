import sys
import app

def test_routes():
    test_cases = [
        "100-free-add-logo-to-pdf-high-quality",
        "100-free-add-logo-to-pdf",
        "100-free-add-page-numbers-to-pdf",
        "100-free-remove-pdf-password",
        "100-free-combine-pdf",
        "100-free-ppt-to-pdf",
        "100-free-ai-pdf-summary",
        "100-free-archive-pdf"
    ]
    
    for slug in test_cases:
        base = app.get_base_tool_for_seo_keyword(slug)
        title, desc = app.get_seo_metadata(slug)
        print(f"SLUG: {slug}")
        print(f"  -> BASE: {base}")
        print(f"  -> TITLE: {title}")
        print(f"  -> DESC: {desc}\n")

if __name__ == '__main__':
    test_routes()
