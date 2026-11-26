import streamlit as st
from engine import process
st.set_page_config(page_title='Generador de Tribus',page_icon='🥾',layout='wide')
st.markdown('''<style>.stApp{background:#f5f8fc}.block-container{max-width:1100px}.hero{padding:24px;border-radius:18px;background:linear-gradient(90deg,#17365d,#0f6b78);color:white;margin-bottom:18px}.ok{padding:12px;border-radius:10px;background:#e8fff3;color:#065f46}</style>''',unsafe_allow_html=True)
st.markdown('<div class="hero"><h1>Generador de Tribus</h1><p>Carga los archivos reales, cruza participantes y genera el Excel procesado.</p></div>',unsafe_allow_html=True)
left,right=st.columns(2)
with left:
    observed=st.file_uploader('Listado de observados',type=['xls','xlsx'],help='Archivo con Nombre participante, Edad y categorías médicas.')
with right:
    medical=st.file_uploader('Condición médica con tribus',type=['xlsx','xls'],help='Debe contener las columnas Tribu y Nombre.')
with st.expander('Configuración del evento'):
    c1,c2,c3=st.columns(3)
    event=c1.text_input('Nombre del evento','Senderismo noviembre 2026')
    c2.number_input('Mochila mínima (lb)',min_value=0,value=35)
    c3.number_input('Mochila máxima (lb)',min_value=0,value=50)
ready=observed is not None and medical is not None
if not ready: st.info('Carga los dos archivos obligatorios para habilitar el motor.')
if st.button('Procesar archivos',type='primary',disabled=not ready,use_container_width=True):
    try:
        with st.spinner('Leyendo, cruzando nombres y generando el Excel...'):
            output,stats=process(observed.getvalue(),observed.name,medical.getvalue(),medical.name)
        st.session_state['output']=output;st.session_state['stats']=stats
    except Exception as e: st.error(f'No fue posible procesar los archivos: {e}')
if 'stats' in st.session_state:
    s=st.session_state['stats'];st.markdown('<div class="ok"><b>Procesamiento completado.</b></div>',unsafe_allow_html=True)
    a,b,c,d,e=st.columns(5)
    a.metric('Observados',s['observados']);b.metric('Condición médica',s['medicos']);c.metric('Coincidencias',s['coincidencias']);d.metric('Sin coincidencia',s['sin_coincidencia']);e.metric('Tribus',s['tribus'])
    st.download_button('Descargar Excel procesado',st.session_state['output'],'Senderistas_observados_con_tribus.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary',use_container_width=True)
    st.caption('El archivo conserva los nombres del listado de observados, incorpora las tribus del archivo de condición médica y aplica la leyenda de colores y múltiples condiciones.')
