import io, re, unicodedata
from dataclasses import dataclass
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

COLORS={'Rojo':'F4CCCC','Azul':'9DC3E6','Verde':'C6E0B4','Amarillo':'FFF2CC','Rosado':'F4B6C2','Naranja':'F4B183'}

def norm(s):
    s=unicodedata.normalize('NFKD',str(s)).encode('ascii','ignore').decode().lower()
    return re.sub(r'[^a-z0-9]','',s)

def read_excel_bytes(data:bytes, filename:str, header=None, sheet_name=0):
    engine='xlrd' if filename.lower().endswith('.xls') else 'openpyxl'
    return pd.read_excel(io.BytesIO(data), header=header, sheet_name=sheet_name, engine=engine)

def find_header(raw, required):
    for i in range(min(len(raw),40)):
        vals={str(v).strip() for v in raw.iloc[i].tolist() if pd.notna(v)}
        if all(any(req.lower() in v.lower() for v in vals) for req in required): return i
    raise ValueError('No fue posible identificar la fila de encabezados.')

def load_medical(data, filename):
    sheets=read_excel_bytes(data,filename,header=None,sheet_name=None)
    best=None
    for name,raw in sheets.items():
        try:
            h=find_header(raw,['Tribu','Nombre'])
            df=read_excel_bytes(data,filename,header=h,sheet_name=name)
            if 'Tribu' in df.columns and 'Nombre' in df.columns:
                best=df.dropna(subset=['Nombre']).copy(); break
        except Exception: pass
    if best is None: raise ValueError('No se encontró una hoja con columnas Tribu y Nombre.')
    return best

def load_observed(data, filename):
    raw=read_excel_bytes(data,filename,header=None)
    h=find_header(raw,['Nombre participante','Edad'])
    df=read_excel_bytes(data,filename,header=h)
    df=df.dropna(subset=['Nombre participante']).copy()
    stop=df.index[df['Nombre participante'].astype(str).str.strip().str.lower().eq('color')]
    if len(stop): df=df.loc[:stop[0]-1]
    return df

def enrich_observed(observed, medical):
    mapping={norm(n):(t,n) for n,t in zip(medical['Nombre'],medical['Tribu'])}
    out=observed.copy(); tribes=[]; matched=[]
    for n in out['Nombre participante']:
        m=mapping.get(norm(n)); tribes.append(m[0] if m else ''); matched.append(bool(m))
    out.insert(0,'Tribu',tribes)
    return out, matched

def medical_priority(row):
    def yes(col): return str(row.get(col,'')).strip().upper() in {'SÍ','SI','●','TRUE'}
    if yes('Presión / Cardíaco') or yes('Cardíaca'): return 'Rojo'
    if yes('Psiquiátrico') or yes('Psiquiátrica'): return 'Azul'
    if yes('Diabetes'): return 'Verde'
    if yes('Alergias'): return 'Amarillo'
    if yes('Otra enfermedad') or yes('Respiratoria') or yes('Convulsiones'): return 'Rosado'
    try:
        if float(row.get('Edad',0))>70: return 'Naranja'
    except: pass
    return None

def to_workbook(observed, medical, matched):
    wb=Workbook(); ws=wb.active; ws.title='Observados con Tribus'; ws.sheet_view.showGridLines=False
    cols=list(observed.columns)
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(cols))
    ws.cell(1,1,'REGISTRO DE SENDERISTAS OBSERVADOS CON TRIBUS')
    ws.cell(1,1).font=Font(size=16,bold=True,color='FFFFFF');ws.cell(1,1).fill=PatternFill('solid',fgColor='17365D');ws.cell(1,1).alignment=Alignment(horizontal='center')
    for c,h in enumerate(cols,1):
        ws.cell(3,c,h);ws.cell(3,c).font=Font(bold=True,color='FFFFFF');ws.cell(3,c).fill=PatternFill('solid',fgColor='666666');ws.cell(3,c).alignment=Alignment(wrap_text=True,horizontal='center')
    for rr,(_,row) in enumerate(observed.iterrows(),4):
        for c,h in enumerate(cols,1): ws.cell(rr,c,'' if pd.isna(row[h]) else row[h])
        p=medical_priority(row)
        if p:
            for c in range(1,len(cols)+1): ws.cell(rr,c).fill=PatternFill('solid',fgColor=COLORS[p])
        for h,label in [('Presión / Cardíaco','Rojo'),('Psiquiátrico','Azul'),('Diabetes','Verde'),('Alergias','Amarillo'),('Otra enfermedad','Rosado')]:
            if h in cols and str(row.get(h,'')).strip()=='●':
                c=cols.index(h)+1; ws.cell(rr,c).fill=PatternFill('solid',fgColor=COLORS[label]);ws.cell(rr,c).font=Font(bold=True)
    start=4+len(observed)+2
    ws.cell(start,1,'LEYENDA DE COLORES');ws.cell(start,1).font=Font(bold=True,color='FFFFFF');ws.cell(start,1).fill=PatternFill('solid',fgColor='17365D')
    legend=[('Rojo','Presión / Cardíaco'),('Azul','Psiquiátrico'),('Verde','Diabetes'),('Amarillo','Alergias'),('Rosado','Otra enfermedad'),('Naranja','Solo por peso mayor a 230 lbs y/o edad mayor a 70 años')]
    for i,(lab,desc) in enumerate(legend,start+1):
        ws.cell(i,1,lab);ws.cell(i,2,desc);ws.cell(i,1).font=Font(bold=True);ws.cell(i,1).fill=PatternFill('solid',fgColor=COLORS[lab]);ws.cell(i,2).fill=PatternFill('solid',fgColor=COLORS[lab])
    note=start+len(legend)+2
    ws.merge_cells(start_row=note,start_column=1,end_row=note,end_column=min(len(cols),10))
    ws.cell(note,1,'Prioridad del color del renglón: rojo > azul > verde > amarillo > rosado. Las columnas con ● conservan todas las categorías. Las tribus se toman del archivo de condición médica.')
    ws.cell(note,1).font=Font(bold=True);ws.cell(note,1).alignment=Alignment(wrap_text=True);ws.row_dimensions[note].height=40
    for c,h in enumerate(cols,1): ws.column_dimensions[get_column_letter(c)].width=max(13,min(38,len(str(h))+8))
    if 'Nombre participante' in cols: ws.column_dimensions[get_column_letter(cols.index('Nombre participante')+1)].width=36
    ws.freeze_panes='A4';ws.auto_filter.ref=f'A3:{get_column_letter(len(cols))}{3+len(observed)}'
    q=wb.create_sheet('Diagnóstico')
    q.append(['Indicador','Resultado']);q.append(['Observados procesados',len(observed)]);q.append(['Nombres con tribu',sum(matched)]);q.append(['Nombres sin correspondencia',len(matched)-sum(matched)]);q.append(['Casos en condición médica',len(medical)]);q.append(['Tribus detectadas',medical['Tribu'].nunique()])
    for c in q[1]: c.font=Font(bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='2F75B5')
    buf=io.BytesIO();wb.save(buf);return buf.getvalue()

def process(observed_bytes, observed_name, medical_bytes, medical_name):
    med=load_medical(medical_bytes,medical_name)
    obs=load_observed(observed_bytes,observed_name)
    enriched,matched=enrich_observed(obs,med)
    output=to_workbook(enriched,med,matched)
    return output, {'observados':len(obs),'medicos':len(med),'coincidencias':sum(matched),'sin_coincidencia':len(matched)-sum(matched),'tribus':med['Tribu'].nunique()}
