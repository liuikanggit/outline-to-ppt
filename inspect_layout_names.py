#!/usr/bin/env python3
import os
from pptx import Presentation

pptx_path = "output/master_with_toc.pptx"
prs = Presentation(pptx_path)

print(f"Total Masters: {len(prs.slide_masters)}")
for m_idx, m in enumerate(prs.slide_masters):
    print(f"\n--- Master {m_idx+1} ---")
    for l_idx, lay in enumerate(m.slide_layouts):
        print(f"  Layout {l_idx}: Space name='{lay.name}'")
