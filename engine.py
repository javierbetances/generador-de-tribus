import io,re,unicodedata,math,random
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter
FLAGS=['Cardíaca','Diabetes','Respiratoria','Convulsiones','Medicamentos','Alergias','Cirugía','Psiquiátrica','Dieta especial','Restricción dieta limitada']
COL={'Rojo':'F4CCCC','Azul':'9DC3E6','Verde':'C6E0B4','Amarillo':'FFF2CC','Rosado':'F4B6C2','Naranja':'F4B183'}
def norm(x):return re.sub('[^a-z0-9]','',unicodedata.normalize('NFKD',str(x)).encode('ascii','ignore').decode().lower())
def yes(x):return str(x).strip().upper() in {'SI','SÍ','●','TRUE','1','TAG ESPECIAL'}
def number(x):
 m=re.search(r'\d+(?:[.,]\d+)?',str(x));
 try:return float(m.group().replace(',','.')) if m else None
 except:return None
def read(data,name):
 eng='xlrd' if name.lower().endswith('.xls') else 'openpyxl'
 return pd.read_excel(io.BytesIO(data),sheet_name=None,header=None,engine=eng)
def table(data,name):
 eng='xlrd' if name.lower().endswith('.xls') else 'openpyxl'
 for sn,raw in read(data,name).items():
  for i in range(min(40,len(raw))):
   vals=' '.join(str(v) for v in raw.iloc[i].dropna())
   if 'nombre' in vals.lower():
    d=pd.read_excel(io.BytesIO(data),sheet_name=sn,header=i,engine=eng);d.columns=[str(c).strip() for c in d.columns]
    return d.dropna(how='all')
 raise ValueError('No se encontró una tabla con columna de nombre.')
def findcol(df,*names):
 mp={norm(c):c for c in df.columns}
 for n in names:
  if norm(n) in mp:return mp[norm(n)]
def height(x):
 s=str(x).replace('’',"'").replace(',','.').replace('/','.');m=re.search(r'(\d)[\'\.\-](\d{1,2})',s)
 if m:return (int(m.group(1))*12+int(m.group(2)))*.0254
 v=number(s)
 if v and 1.3<=v<=2.2:return v
 if v and 400<=v<=700:
  q=str(int(v));return (int(q[0])*12+int(q[1:]))*.0254

def surnames(n):
 p=unicodedata.normalize('NFKD',str(n)).encode('ascii','ignore').decode().lower().split();return set(x for x in p[1:] if len(x)>2 and x not in {'del','los','las','de','la'})
def prepare(g,m,lmin,lmax):
 nc=findcol(g,'Nombre participante','Nombre');
 if not nc:raise ValueError('Falta la columna Nombre participante o Nombre.')
 g=g[g[nc].notna()].copy();g['_k']=g[nc].map(norm);g=g[g['_k']!=''].drop_duplicates('_k')
 mm={}
 if m is not None:
  mc=findcol(m,'Nombre','Nombre participante');
  for _,r in m[m[mc].notna()].iterrows():mm[norm(r[mc])]=r.to_dict()
 ac=findcol(g,'Edad');fc=findcol(g,'Condición física','Condicion fisica');wc=findcol(g,'Peso declarado (lbs)','Peso','Peso declarado');hc=findcol(g,'Estatura declarada','Estatura','Altura')
 out=[]
 for _,r in g.iterrows():
  d=r.to_dict();n=str(r[nc]).strip();md=mm.get(norm(n),{});w=number(r.get(wc)) if wc else None;h=height(r.get(hc)) if hc else None;a=number(r.get(ac)) if ac else None;f=number(r.get(fc)) if fc else None
  for c in FLAGS:d[c]=md.get(c,d.get(c,''))
  d.update({'Nombre participante':n,'Edad calculada':a,'Condición física normalizada':f,'Peso normalizado':w,'Estatura m':round(h,2) if h else None,'IMC operativo':round(w*.453592/h**2,1) if w and h else None,'Carga mínima %':round(lmin/w*100,1) if w else None,'Carga máxima %':round(lmax/w*100,1) if w else None,'TAG ESPECIAL':'TAG ESPECIAL' if md else '','_tag':bool(md),'_med':sum(yes(d.get(c)) for c in FLAGS),'_sur':surnames(n)})
  out.append(d)
 return pd.DataFrame(out)
def distribute(df,k):
 caps=[len(df)//k+(i<len(df)%k) for i in range(k)];gs=[[] for _ in range(k)];ids=list(df.index);random.Random(26).shuffle(ids);ids.sort(key=lambda i:(df.at[i,'_tag'],df.at[i,'_med'],df.at[i,'Carga máxima %'] or 0),reverse=True)
 for i in ids:
  best=None
  for j,g in enumerate(gs):
   if len(g)>=caps[j]:continue
   common=set().union(*(df.at[x,'_sur'] for x in g))&df.at[i,'_sur'] if g else set();tags=sum(df.at[x,'_tag'] for x in g)
   score=1000*bool(common)+30*tags+len(g)+sum(df.at[x,'_med'] for x in g)
   if best is None or score<best[0]:best=(score,j)
  gs[best[1]].append(i)
 df=df.copy();df['Tribu']=''
 for j,g in enumerate(gs,1):
  for i in g:df.at[i,'Tribu']=f'Tribu {j:02d}'
 return df
def book(df,event,k,lmin,lmax):
 wb=Workbook();wb.remove(wb.active)
 def title(ws,t,n):ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=n);c=ws.cell(1,1,t);c.font=Font(size=16,bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='17365D');c.alignment=Alignment(horizontal='center')
 def put(ws,d):
  for j,h in enumerate(d.columns,1):c=ws.cell(3,j,h);c.font=Font(bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='2F75B5');c.alignment=Alignment(wrap_text=True)
  for i,(_,r) in enumerate(d.iterrows(),4):
   for j,h in enumerate(d.columns,1):ws.cell(i,j,'' if pd.isna(r[h]) else r[h])
  ws.freeze_panes='A4';ws.auto_filter.ref=f'A3:{get_column_letter(len(d.columns))}{3+len(d)}';ws.sheet_view.showGridLines=False
  for j,h in enumerate(d.columns,1):ws.column_dimensions[get_column_letter(j)].width=min(36,max(12,len(str(h))+3))
 p=wb.create_sheet('Portada');title(p,f'{event} | {k} TRIBUS',10)
 for i,(a,b) in enumerate([('Participantes',len(df)),('TAG ESPECIAL',int(df._tag.sum())),('Mochila',f'{lmin}-{lmax} lb'),('Criterio','Equilibrio numérico, médico, físico, edad, IMC, carga proporcional y separación preventiva por apellidos.'),('Advertencia','Organización operativa; no sustituye evaluación médica individual.')],3):p.cell(i,1,a).font=Font(bold=True);p.merge_cells(start_row=i,start_column=2,end_row=i,end_column=10);p.cell(i,2,b)
 base=['Tribu','Nombre participante','TAG ESPECIAL','Edad calculada','Condición física normalizada','Peso normalizado','Estatura m','IMC operativo','Carga mínima %','Carga máxima %']+[c for c in FLAGS if c in df]
 ass=df[base].sort_values(['Tribu','Nombre participante'])
 for name,data,heading in [('Asignación de Tribus',ass,'ASIGNACIÓN POR TRIBU'),('Asignación por nombre',ass.sort_values('Nombre participante'),'ASIGNACIÓN ALFABÉTICA')]:ws=wb.create_sheet(name);title(ws,heading,len(data.columns));put(ws,data)
 ws=wb.create_sheet('Agrupación por Tribu');title(ws,'AGRUPACIÓN EXPANDIBLE',3);ws.append([]);ws.append(['Tribu','Nombre','Cantidad']);r=4
 for t,g in ass.groupby('Tribu'):
  ws.cell(r,1,t);ws.cell(r,3,len(g));ws.cell(r,1).font=Font(bold=True);r+=1
  for n in g['Nombre participante']:ws.cell(r,2,n);ws.row_dimensions[r].outlineLevel=1;r+=1
 med=df[df._tag][['Tribu','Nombre participante','TAG ESPECIAL']+[c for c in FLAGS if c in df]].sort_values(['Tribu','Nombre participante']);ws=wb.create_sheet('Indicadores Médicos');title(ws,'INDICADORES MÉDICOS',len(med.columns));put(ws,med)
 summary=df.groupby('Tribu').agg(Participantes=('Nombre participante','size'),TAG_ESPECIAL=('_tag','sum'),Edad_promedio=('Edad calculada','mean'),Condicion_promedio=('Condición física normalizada','mean'),IMC_promedio=('IMC operativo','mean'),Carga_promedio=('Carga máxima %','mean')).reset_index().round(1);ws=wb.create_sheet('Resumen de Tribus');title(ws,'RESUMEN EJECUTIVO',len(summary.columns));put(ws,summary)
 check=ass[['Tribu','Nombre participante','TAG ESPECIAL']].copy();check['Check-In']='';check['Observación']='';ws=wb.create_sheet('Check-In');title(ws,'CHECK-IN',len(check.columns));put(ws,check)
 src=df.drop(columns=['_k','_tag','_med','_sur'],errors='ignore');ws=wb.create_sheet('Fuente Original');title(ws,'TRAZABILIDAD',len(src.columns));put(ws,src)
 qc=pd.DataFrame([['Participantes',len(df),'OK'],['Asignados',(df.Tribu!='').sum(),'OK'],['Duplicados',df['Nombre participante'].map(norm).duplicated().sum(),'OK'],['Tamaño mínimo',df.groupby('Tribu').size().min(),'OK'],['Tamaño máximo',df.groupby('Tribu').size().max(),'OK']],columns=['Control','Resultado','Estado']);ws=wb.create_sheet('Control de Calidad');title(ws,'AUDITORÍA FINAL',3);put(ws,qc)
 b=io.BytesIO();wb.save(b);return b.getvalue()
def full_process(gb,gn,mb,mn,event,k,lmin,lmax):
 g=table(gb,gn);m=table(mb,mn) if mb else None;df=distribute(prepare(g,m,lmin,lmax),k);out=book(df,event,k,lmin,lmax)
 return out,{'participantes':len(df),'tribus':k,'tag':int(df._tag.sum()),'minimo':int(df.groupby('Tribu').size().min()),'maximo':int(df.groupby('Tribu').size().max())}
