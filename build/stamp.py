import pymupdf

src = 'final_raw.pdf'
out = '/workspace/project/DBMS_One_Day_Exam_Notes.pdf'
doc = pymupdf.open(src)
n = doc.page_count
gray = (0.48, 0.53, 0.58)
for i, page in enumerate(doc):
    if i == 0:
        continue
    w, h = page.rect.width, page.rect.height
    page.draw_line(pymupdf.Point(36.8, h - 34), pymupdf.Point(w - 36.8, h - 34),
                   color=(0.83, 0.87, 0.91), width=0.5)
    page.insert_text((36.8, h - 24), 'DBMS Unit III  ·  One-Day Exam Notes',
                     fontname='helv', fontsize=7, color=gray)
    label = f'Page {i+1} of {n}'
    tw = pymupdf.get_text_length(label, fontname='helv', fontsize=7)
    page.insert_text((w - 36.8 - tw, h - 24), label, fontname='helv', fontsize=7, color=gray)

doc.set_metadata({
    'title': 'DBMS Unit III — One-Day Exam Notes',
    'author': 'Condensed from Unit III source material',
    'subject': 'Transactions, Concurrency Control, Recovery, PL/SQL',
    'keywords': 'DBMS, ACID, serializability, 2PL, timestamp ordering, recovery, PL/SQL, triggers, packages',
})
doc.save(out, garbage=4, deflate=True)
print('saved', out, doc.page_count)
