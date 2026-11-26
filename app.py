import math,streamlit as st
from engine import full_process,table
st.set_page_config(page_title='Generador de Tribus',page_icon='🥾',layout='wide')
st.markdown('<div style="padding:24px;border-radius:18px;background:linear-gradient(90deg,#17365d,#0f6b78);color:white"><h1>Generador de Tribus</h1><p>Distribución multicriterio y libro operativo automático.</p></div>',unsafe_allow_html=True)
a,b=st.columns(2)
with a:g=st.file_uploader('Listado general de senderistas',type=['xls','xlsx'])
with b:m=st.file_uploader('Condición médica / observados',type=['xls','xlsx'])
with st.expander('Configuración del evento',expanded=True):
 c1,c2,c3=st.columns(3);event=c1.text_input('Nombre del evento','Senderismo');mode=c2.radio('Definir tribus por',['Cantidad','Tamaño deseado']);value=c3.number_input('Cantidad / tamaño',2,100,14)
 c4,c5,c6=st.columns(3);days=c4.number_input('Duración (días)',1,30,3);lmin=c5.number_input('Mochila mínima (lb)',0,200,35);lmax=c6.number_input('Mochila máxima (lb)',1,200,50)
if not g:st.info('Carga el listado general. El archivo médico es opcional, pero recomendado.')
if st.button('Generar distribución completa',type='primary',disabled=not g,use_container_width=True):
 try:
  k=int(value) if mode=='Cantidad' else max(1,math.ceil(len(table(g.getvalue(),g.name))/value))
  with st.spinner('Equilibrando tribus y creando el Excel...'):out,s=full_process(g.getvalue(),g.name,m.getvalue() if m else None,m.name if m else '',event,k,int(lmin),int(lmax))
  st.session_state.out=out;st.session_state.stats=s
 except Exception as e:st.error(str(e))
if 'stats' in st.session_state:
 s=st.session_state.stats;st.success('Distribución generada.');x=st.columns(5)
 for c,(n,v) in zip(x,[('Participantes',s['participantes']),('Tribus',s['tribus']),('TAG especial',s['tag']),('Mínimo',s['minimo']),('Máximo',s['maximo'])]):c.metric(n,v)
 st.download_button('Descargar Excel final',st.session_state.out,'Generador_de_Tribus_Resultado.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary',use_container_width=True)
