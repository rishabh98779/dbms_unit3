import re, subprocess, pymupdf, json, sys

parts=['part00_cover.html','part01_transactions.html','part02_serializability.html','part03_recoverability.html',
       'part04_locking.html','part05_timestamp.html','part06_recovery.html','part07_plsql.html','part08_final.html']
html=''.join(open(p).read() for p in parts)
open('notes.html','w').write(html)

def render(src,out):
    subprocess.run(['chromium','--headless','--disable-gpu','--no-sandbox','--no-pdf-header-footer',
                    f'--print-to-pdf={out}','--virtual-time-budget=8000',
                    f'file:///workspace/project/build/{src}'],check=True,
                   stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)

def mapping(pdf):
    d=pymupdf.open(pdf); m={}
    for i,p in enumerate(d):
        for mt in re.finditer(r'§S(\d\d)', p.get_text()):
            m.setdefault(mt.group(1), i+1)
    return d.page_count, m

render('notes.html','raw.pdf')
n1,m1=mapping('raw.pdf')
print('pass1',n1,m1)

# fill TOC
cover=open('part00_cover.html').read()
for k,v in m1.items():
    num=int(k)
    cover=re.sub(r'<span data-ref="s%d">p\.\s*\d*</span>'%num, '<span data-ref="s%d">p. %d</span>'%(num,v), cover)
open('part00_cover.html','w').write(cover)

html=''.join(open(p).read() for p in parts)
open('notes.html','w').write(html)
render('notes.html','raw.pdf')
n2,m2=mapping('raw.pdf')
print('pass2',n2,m2)
assert n1==n2 and m1==m2, 'pagination moved after TOC fill'

# final without markers
open('notes_final.html','w').write(re.sub(r'<span class="mk">§S\d\d</span>','',html))
render('notes_final.html','final_raw.pdf')
n3,m3=mapping('final_raw.pdf')
print('final',n3,m3)
