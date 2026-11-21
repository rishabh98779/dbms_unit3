# DBMS Unit III — One-Day Exam Notes

Condensed, exam-oriented revision notes for DBMS Unit III (transaction processing, concurrency
control, recovery, PL/SQL), built from 136 pages of source material and reduced to 21 A4 pages
that can realistically be revised in one day.

## Deliverables

| File | What it is |
|---|---|
| `DBMS_One_Day_Exam_Notes.pdf` | Final 21-page A4 print-ready notes |
| `DBMS_One_Day_Exam_Notes.docx` | Editable version of the same content |

## Coverage

| Pages | Section |
|---|---|
| 1 | Cover: topic map, one-day plan, TOC, source coverage |
| 2 | Transactions & ACID, states, schedules and conflicts |
| 3–4 | Serializability (conflict vs view, precedence graph), recoverability |
| 4–6 | Concurrency problems, locking, 2PL, strict 2PL, distributed 2PL |
| 6–9 | Timestamp ordering, strict TO, optimistic CC, granularity, MGL, distributed CC |
| 10–13 | Recovery: failures, logs, deferred vs immediate update, checkpoints, shadow paging, backups |
| 13–17 | PL/SQL: procedures and functions, parameter modes, triggers, packages, cursors, overloading |
| 17–18 | Query processing and optimization: four steps, cost model, selection and join algorithms, heuristics, cost-based optimization, worked example |
| 19–21 | Final revision: checklist, comparisons, notation sheet, recall list, answer skeletons, traps |

Section 12 (query processing and optimization) is supplementary: it is not in the Unit III source,
and the notes say so. Everything else is derived from the source.

Each topic ends with an **Exam Focus** box giving 2-, 5- and 10-mark guidance plus likely questions.
Points where the source is wrong or self-contradictory are flagged in "Check this point" boxes
rather than silently corrected.

## Rebuilding

Everything under `build/` regenerates the PDF and DOCX:

```bash
cd build
python pipeline.py     # HTML -> final_raw.pdf (headless chromium)
python stamp.py        # footer, page numbers, metadata -> ../DBMS_One_Day_Exam_Notes.pdf
python docx_build.py   # -> ../DBMS_One_Day_Exam_Notes.docx
```

`build/AGENTS.md` documents the file layout and the section-marker convention used to derive the
table-of-contents page numbers.

The source PDF is not redistributed here.
