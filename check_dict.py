import json
import re

from indian_langs import INDIAN_LANGS

for lang, t_dict in INDIAN_LANGS.items():
    for k, v in t_dict.items():
        if '\\' in v:
            print(f"Backslash in {lang} -> {k}: {v}")

print("Done checking.")
