import json, os, shutil, zipfile
B='/home/claude'; OUT=B+'/web/site'; F=OUT+'/fail'
shutil.rmtree(OUT, ignore_errors=True); os.makedirs(F)
D=B+'/Latihan_Claude_IAB/'
src={'a1':D+'01_Profil_Program_PKPB3.docx','a2':D+'02_Senarai_Penceramah.xlsx','a3':D+'03_Komen_Peserta.txt',
 'a4':D+'04_Jadual_Program_Draf.xlsx','a5':D+'05_Nota_Mesyuarat_JK.txt','a6':D+'06_Senarai_Peserta.xlsx',
 'a7':D+'07_Penilaian_PKPB3_mentah.xlsx','a8':D+'08_Ujian_Pra_Pasca_PKPB3.xlsx','a9':D+'09_Data_Operasi_PKPB3.xlsx',
 'prog_code':B+'/sistem_iab/Code.gs','prog_index':B+'/sistem_iab/Index.html',
 'pem_code':B+'/pemantauan_iab/Code.gs','pem_index':B+'/pemantauan_iab/Index.html','pem_wbs':B+'/pemantauan_iab/WBS_Contoh_PKPB3.csv',
 'agen':B+'/pemantauan_iab/Agen.gs','skill':B+'/pemantauan_iab/minit-mesyuarat-iab.zip','zipdata':B+'/Latihan_Claude_IAB.zip'}
# pakej kod
zk=B+'/web/Kod_Sistem_Claude_IAB.zip'
readme="""KOD SISTEM - KURSUS CLAUDE IAB (bahan latihan, data rekaan)

1_Sistem_Pengurusan_Program  : Modul Tambahan 1 (Aktiviti B.3). Tampal Code.gs dan Index.html
                               dalam Extensions > Apps Script, jalankan setupSheets, Deploy > Web app.
2_Sistem_Pemantauan_Tugasan  : Aktiviti 5.1. Tampal Code.gs dan Index.html, jalankan setupSheet,
                               import WBS_Contoh_PKPB3.csv ke helaian Tugasan, Deploy > Web app.
3_Agen_AI                    : Modul Tambahan 2 (Tahap 4). Tambah Agen.gs ke projek Sistem Pemantauan,
                               tetapkan Script Property ANTHROPIC_API_KEY, jalankan setupAgen.
4_Skill                      : Modul Tambahan 2 (Tahap 2). Muat naik minit-mesyuarat-iab.zip di Customize > Skills.

Fail HTML mesti dinamakan "Index" dalam editor Apps Script.
"""
with zipfile.ZipFile(zk,'w',zipfile.ZIP_DEFLATED) as z:
    r='Kod_Sistem_Claude_IAB/'
    z.writestr(r+'BACA_SAYA.txt',readme)
    z.write(src['prog_code'],r+'1_Sistem_Pengurusan_Program/Code.gs'); z.write(src['prog_index'],r+'1_Sistem_Pengurusan_Program/Index.html')
    z.write(src['pem_code'],r+'2_Sistem_Pemantauan_Tugasan/Code.gs'); z.write(src['pem_index'],r+'2_Sistem_Pemantauan_Tugasan/Index.html')
    z.write(src['pem_wbs'],r+'2_Sistem_Pemantauan_Tugasan/WBS_Contoh_PKPB3.csv')
    z.write(src['agen'],r+'3_Agen_AI/Agen.gs'); z.write(src['skill'],r+'4_Skill/minit-mesyuarat-iab.zip')
src['zipkod']=zk
TEXT={'a3','a5','prog_code','prog_index','pem_code','pem_index','pem_wbs','agen'}
import re, openpyxl
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
def book(path):
    wb=openpyxl.load_workbook(path); out=[]
    for ws in wb.worksheets:
        rows=[]; money=set()
        for r in ws.iter_rows():
            row=[]
            for c in r:
                v=c.value
                if isinstance(v,str) and v.startswith('='):
                    m=re.fullmatch(r'=SUM\(([A-Z]+)(\d+):\1(\d+)\)',v); assert m,v
                    col=openpyxl.utils.column_index_from_string(m.group(1))
                    tot=sum(ws.cell(row=i,column=col).value or 0 for i in range(int(m.group(2)),int(m.group(3))+1))
                    v={'f':v[1:],'v':tot}
                elif hasattr(v,'isoformat'): raise SystemExit('date cell '+c.coordinate)
                if c.number_format=='#,##0.00': money.add(c.column-1)
                row.append(v)
            while row and row[-1] is None: row.pop()
            rows.append(row)
        ncol=max(len(r) for r in rows)
        cols=[round(ws.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width or 12,1) for i in range(ncol)]
        sh={'name':ws.title,'rows':rows,'cols':cols,'hdr':len(rows)>1}
        if money: sh['money']=sorted(money)
        out.append(sh)
    return out
def docblocks(path):
    doc=Document(path); out=[]; ti=0
    for el in doc.element.body.iterchildren():
        tag=el.tag.split('}')[1]
        if tag=='p':
            p=Paragraph(el,doc)
            if not p.text.strip(): continue
            b={'text':p.text}; sn=p.style.name
            if sn.startswith('Heading'): b['style']='Heading2'
            elif sn=='List Bullet': b['style']='ListBullet'; b['text']='\u2022 '+p.text
            if p.alignment is not None and int(p.alignment)==1: b['center']=True
            if p.runs and all(r.bold for r in p.runs): b['bold']=True
            out.append(b)
        elif tag=='tbl':
            t=Table(el,doc); out.append({'tbl':[[c.text for c in r.cells] for r in t.rows],'boldCol':ti==0,'boldRow':ti==1}); ti+=1
    return out
inl={}
for k,p in src.items():
    inl[k]={'size':os.path.getsize(p)}
    if k in TEXT: inl[k]['text']=open(p,encoding='utf-8').read()
    elif p.endswith('.xlsx'): inl[k]['book']=book(p)
    elif p.endswith('.docx'): inl[k]['doc']=docblocks(p)
inl['skill']['md']=open(B+'/pemantauan_iab/skill/minit-mesyuarat-iab/SKILL.md',encoding='utf-8').read()
inl['zipkod']['readme']=readme
shutil.rmtree(F)
js=json.dumps(inl,ensure_ascii=False).replace('<','\\u003c').replace('\u2028','\\u2028').replace('\u2029','\\u2029')
html=open(B+'/web/index.src.html',encoding='utf-8').read()
assert html.count('__INLINE__')==1
open(OUT+'/index.html','w',encoding='utf-8').write(html.replace('__INLINE__',js))
print(os.path.getsize(OUT+'/index.html'), os.listdir(OUT))
