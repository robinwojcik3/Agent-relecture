import os,docx,re,csv
from pathlib import Path
fn=[f for f in os.listdir('input') if f.endswith('.docx')][0]
d=docx.Document('input/'+fn)
paras=[p.text.strip() for p in d.paragraphs]
rep={'�%':'','�?T':'É','�?"':'—','��':'','Ǹ':'e','Ǧ':'e','ǩ':'i','ǯ':'u','�?':'É'}
clean=[]
for t in paras:
    s=t
    for a,b in rep.items(): s=s.replace(a,b)
    s=re.sub(r'\\s+',' ',s)
    clean.append(s)
out=[]
anchor_comments=[]
for i,t in enumerate(clean):
    if not t: continue
    if t.isupper() or t.startswith(('PRÉAMBULE','INTRODUCTION','MÉTHODOLOGIE','CONTEXTE','SYNTHÈSE','CONCLUSION','RÉGLEMENTAIRE','RÉGLEMENTATION','BIBLIOGRAPHIE','INDEX DES FIGURES','INDEX DES TABLEAUX')):
        out.append('\n# '+t+'\n')
    else:
        out.append(t)
    import re as _re
    if len(t)>20:
        anchor=f"p{i+1}_{_re.sub(r'[^a-z0-9]+','- ',t.lower())[:20].strip('- ')}"
    else:
        anchor=f"p{i+1}_{_re.sub(r'[^a-z0-9]+','- ',t.lower())}"
    cmt=None;grav='P2';cat='redaction'
    if 'Figure' in t:
        cmt='Vérifier pagination et renvois de figures';cat='carto';grav='P2'
    if 'ZNIEFF' in t:
        cmt='Préciser type, code et date de mise à jour ZNIEFF';cat='reglementaire';grav='P2'
    if 'Natura 2000' in t:
        cmt='Indiquer codes FR et périmètres des sites Natura 2000';cat='reglementaire';grav='P2'
    if cmt:
        anchor_comments.append((anchor,cmt,grav,cat))
    out[-1]=f"<a id='{anchor}'></a>\n"+out[-1]
Path('work').mkdir(parents=True,exist_ok=True)
Path('work/rapport_revise.md').write_text('\n\n'.join(out),encoding='utf-8')
with open('work/commentaires.csv','w',encoding='utf-8',newline='') as f:
    w=csv.writer(f);w.writerow(['ancre_textuelle','commentaire','gravite','categorie']);
    for row in anchor_comments: w.writerow(row)
print('written')