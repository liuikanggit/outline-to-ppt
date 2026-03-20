#!/usr/bin/env python3
import json

with open("output/mubu_parsed_structure.json", "r", encoding="utf-8") as f:
    data = json.load(f)

def check_content(chapters):
    for ch in chapters:
        for c in ch.get('content', []):
            if c.get('type') == 'text':
                t = c.get('text')
                print(f"Code: {ch.get('code')} - Content Title: {c.get('title')}")
                print(f"  Type of text field: {type(t)}")
                if isinstance(t, list):
                    print(f"  ❌ FOUND A LIST: {t}")
        check_content(ch.get('subChapter', []))

check_content(data.get('chapters', []))
