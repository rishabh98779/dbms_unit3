# Build notes: DBMS one-day exam notes

## Layout
- `style.css` — print design system (A4, two-column grids, callout boxes, tables).
- `part00_cover.html` … `part08_final.html` — content, one file per major block.
- `pipeline.py` — concatenates the parts, strips `§Sxx` markers, writes `notes_final.html`,
  renders with headless chromium to `final_raw.pdf`, then stamps footer and page numbers.
- `stamp.py` — pymupdf pass that writes the footer, page numbers and metadata.
- `docx_build.py` — converts `notes_final.html` into the editable DOCX; figures are
  cropped from the rendered PDF (`figs/figN.png`).

## Commands
```bash
python pipeline.py     # build final_raw.pdf (21 pages)
python stamp.py        # write /workspace/project/DBMS_One_Day_Exam_Notes.pdf
python docx_build.py   # write /workspace/project/DBMS_One_Day_Exam_Notes.docx
```

## Conventions
- Section headings carry `<span class="num">` and `<span class="mk">§Sxx</span>`; the TOC
  page numbers on the cover are derived from these markers, so re-run the pipeline after
  any content edit and re-inject the numbers printed by `pass1`.
- Diagrams are inline SVG with explicit `viewBox`; check for clipping by rendering pages to
  `png/pNN.png` (110–120 dpi) after each layout change.
- Keep runaway page growth in check: the target is 12–20 pages.
