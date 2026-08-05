import sys
import os

# Mock Flask context so we can test app.py
from unittest.mock import MagicMock
sys.modules['flask'] = MagicMock()
sys.modules['pypdf'] = MagicMock()
sys.modules['fitz'] = MagicMock()
sys.modules['PIL'] = MagicMock()
sys.modules['docx'] = MagicMock()
sys.modules['reportlab'] = MagicMock()
sys.modules['reportlab.pdfgen'] = MagicMock()
sys.modules['reportlab.lib'] = MagicMock()
sys.modules['reportlab.lib.pagesizes'] = MagicMock()
sys.modules['reportlab.platypus'] = MagicMock()
sys.modules['reportlab.lib.styles'] = MagicMock()
sys.modules['pptx'] = MagicMock()
sys.modules['pptx.util'] = MagicMock()
sys.modules['openpyxl'] = MagicMock()
sys.modules['boto3'] = MagicMock()

try:
    import app
    
    # Try calling serve_tool_page
    print("Calling serve_tool_page('split', 'hi')...")
    res = app.serve_tool_page('split', 'hi')
    print("Result:", res)

    print("\nCalling serve_tool_page('split_pdf', 'hi')...")
    res2 = app.serve_tool_page('split_pdf', 'hi')
    print("Result length:", len(res2) if isinstance(res2, str) else type(res2))
    
except Exception as e:
    import traceback
    traceback.print_exc()
