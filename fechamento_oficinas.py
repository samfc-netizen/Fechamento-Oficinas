from pathlib import Path
import io, re, csv, zipfile, unicodedata
from datetime import datetime
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

st.set_page_config(page_title="Fechamento de Oficinas", page_icon="📄", layout="wide")

def norm(s):
    s='' if s is None else str(s); s=unicodedata.normalize('NFKD',s).encode('ascii','ignore').decode('ascii')
    return re.sub(r'\s+',' ',s).strip().upper()
def br_money(v):
    try:v=float(v)
    except:v=0.0
    return f"R$ {v:,.2f}".replace(',','X').replace('.',',').replace('X','.')

def parse_brl_input(v):
    if v is None:
        return 0.0
    s=str(v).strip().replace('R$','').replace(' ','')
    if not s:
        return 0.0
    if ',' in s:
        s=s.replace('.','').replace(',','.')
    try:
        return float(s)
    except:
        return 0.0

def format_boleto_field(key):
    st.session_state[key]=br_money(parse_brl_input(st.session_state.get(key,0)))
def br_num(v):
    if pd.isna(v): return 0.0
    if isinstance(v,(int,float)): return float(v)
    s=str(v).strip().replace('R$','').replace(' ','')
    if ',' in s:s=s.replace('.','').replace(',','.')
    return float(pd.to_numeric(s,errors='coerce')) if s and not pd.isna(pd.to_numeric(s,errors='coerce')) else 0.0
def clean_code(v):
    # Chave padronizada para cruzar Fechamento x Tabela 06.
    # 09921 -> 9921 | 00005 -> 5 | 9921.0 -> 9921
    if pd.isna(v):
        return ''
    s = str(v).strip()
    s = re.sub(r'\.0$', '', s)
    if re.fullmatch(r'\d+', s):
        return s.lstrip('0') or '0'
    return s.upper()

def raw(upload): upload.seek(0); return upload.read()
def decode(data):
    for e in ('utf-8-sig','cp1252','latin1'):
        try:return data.decode(e)
        except UnicodeDecodeError:pass
    return data.decode('latin1',errors='replace')
def sep(text):
    try:return csv.Sniffer().sniff('\n'.join(text.splitlines()[:60]),delimiters=';,\t|').delimiter
    except:return ';'
def load_tab(u):
    data=raw(u)
    df=pd.read_excel(io.BytesIO(data),dtype=str) if u.name.lower().endswith(('.xlsx','.xls')) else pd.read_csv(io.StringIO(decode(data)),sep=sep(decode(data)),dtype=str,engine='python')
    cols={norm(c):c for c in df.columns}; cc=next((v for k,v in cols.items() if k in ('COD.ITEM','COD ITEM','CODIGO ITEM','CODIGO')),None); pc=next((v for k,v in cols.items() if 'NOVO PR. VENDA 6' in k or 'NOVO PR VENDA 6' in k),None)
    if not cc or not pc:raise ValueError('Na Tabela 06 não encontrei Cód.Item e Novo Pr. Venda 6.')
    x=df[[cc,pc]].copy();x.columns=['COD','PRECO6'];x['COD']=x['COD'].map(clean_code);x['PRECO6']=x['PRECO6'].map(br_num);return x.drop_duplicates('COD',keep='last')
def load_fech(u):
    data=raw(u)
    if u.name.lower().endswith(('.xlsx','.xls')):
        tmp=pd.read_excel(io.BytesIO(data),header=None,dtype=str); hi=next((i for i,r in tmp.iterrows() if 'CLIENTE' in ' | '.join(norm(x) for x in r.tolist()) and 'DOC' in ' | '.join(norm(x) for x in r.tolist()) and 'COD' in ' | '.join(norm(x) for x in r.tolist())),None)
        if hi is None:raise ValueError('Cabeçalho do fechamento não encontrado.')
        df=pd.read_excel(io.BytesIO(data),header=hi,dtype=str)
    else:
        text=decode(data); lines=text.splitlines(); hi=next((i for i,l in enumerate(lines) if 'CLIENTE' in norm(l) and 'DOC' in norm(l) and 'COD' in norm(l) and 'VR.TOTAL' in norm(l)),None)
        if hi is None:raise ValueError('Cabeçalho CLIENTE / Nº DOC / CÓD / VR.TOTAL não encontrado.')
        df=pd.read_csv(io.StringIO('\n'.join(lines[hi:])),sep=';',dtype=str,engine='python')
    mp={}
    for c in df.columns:
        n=norm(c)
        if n=='CLIENTE':mp[c]='CLIENTE'
        elif 'DOC' in n and 'ORI' not in n and ('N' in n or 'NO' in n):mp[c]='DOC'
        elif n=='DATA':mp[c]='DATA'
        elif n in ('COD','CODIGO'):mp[c]='COD'
        elif n=='QTD':mp[c]='QTD'
        elif n in ('VR.TOTAL','VR TOTAL'):mp[c]='VR_TOTAL'
    need={'CLIENTE','DOC','DATA','COD','QTD','VR_TOTAL'}
    if not need.issubset(set(mp.values())):raise ValueError('Colunas ausentes: '+', '.join(sorted(need-set(mp.values()))))
    x=df[list(mp)].rename(columns=mp);x=x[x.CLIENTE.notna()&x.DOC.notna()&x.COD.notna()].copy();x['CLIENTE']=x.CLIENTE.astype(str).str.strip();x['DOC']=x.DOC.astype(str).str.strip();x['COD']=x.COD.map(clean_code);x['QTD']=x.QTD.map(br_num);x['VR_TOTAL']=x.VR_TOTAL.map(br_num);x['DATA_DT']=pd.to_datetime(x.DATA,dayfirst=True,errors='coerce');x['DATA']=x.DATA_DT.dt.strftime('%d/%m/%Y').fillna(x.DATA.astype(str));x['TIPO']=x.DOC.str.extract(r'-(FF|NF)-',expand=False).fillna('OUTRO');return x
def competencia(df):
    d=df.DATA_DT.dropna()
    if d.empty:return 'Fechamento de Oficinas'
    m=['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];return f'{m[d.iloc[0].month-1]}/{d.iloc[0].year}'
def cab(canvas,doc):
    canvas.saveState();canvas.setFont('Helvetica-Bold',9);canvas.drawString(15*mm,287*mm,'FECHAMENTO DE OFICINAS');canvas.setFont('Helvetica',8);canvas.drawRightString(195*mm,287*mm,datetime.now().strftime('%d/%m/%Y %H:%M'));canvas.setStrokeColor(colors.HexColor('#D0D5DD'));canvas.line(15*mm,284*mm,195*mm,284*mm);canvas.restoreState()
def pdf_resumo(r,comp):
    b=io.BytesIO();doc=SimpleDocTemplate(b,pagesize=A4,rightMargin=15*mm,leftMargin=15*mm,topMargin=20*mm,bottomMargin=15*mm);sty=getSampleStyleSheet();story=[Paragraph('Resumo do Fechamento',ParagraphStyle('t',parent=sty['Title'],fontName='Helvetica-Bold',fontSize=17,alignment=TA_CENTER)),Paragraph(comp,ParagraphStyle('s',parent=sty['Normal'],alignment=TA_CENTER,textColor=colors.HexColor('#667085'))),Spacer(1,7*mm)];rows=[['Oficina','Valor Boleto','Valor Real','Diferença']]+[[x['Oficina'],br_money(x['Valor Boleto']),br_money(x['Valor Real']),br_money(x['Diferença'])] for _,x in r.iterrows()]+[['TOTAL GERAL',br_money(r['Valor Boleto'].sum()),br_money(r['Valor Real'].sum()),br_money(r['Diferença'].sum())]];t=Table(rows,colWidths=[82*mm,32*mm,32*mm,32*mm],repeatRows=1);t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#101828')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),('BACKGROUND',(0,-1),(-1,-1),colors.HexColor('#F2F4F7')),('GRID',(0,0),(-1,-1),.35,colors.HexColor('#D0D5DD')),('ALIGN',(1,1),(-1,-1),'RIGHT'),('FONTSIZE',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6)]));story.append(t);doc.build(story,onFirstPage=cab,onLaterPages=cab);return b.getvalue()
def pdf_documentos(base,comp):
    docs=base[base.TIPO.isin(['FF','NF'])].groupby(['CLIENTE','TIPO','DOC','DATA'],as_index=False).VALOR_REAL.sum().sort_values(['CLIENTE','DATA','TIPO','DOC']);b=io.BytesIO();doc=SimpleDocTemplate(b,pagesize=A4,rightMargin=15*mm,leftMargin=15*mm,topMargin=20*mm,bottomMargin=15*mm);sty=getSampleStyleSheet();story=[Paragraph('Documentos do Fechamento por Oficina',ParagraphStyle('t',parent=sty['Title'],fontName='Helvetica-Bold',fontSize=17,alignment=TA_CENTER)),Paragraph(comp,ParagraphStyle('s',parent=sty['Normal'],alignment=TA_CENTER,textColor=colors.HexColor('#667085'))),Spacer(1,6*mm)]
    for oficina,g in docs.groupby('CLIENTE',sort=True):
        # Cada oficina inicia obrigatoriamente em uma nova página.
        # Se a oficina anterior terminar no meio da folha, o espaço restante fica em branco.
        if len(story) > 3:
            from reportlab.platypus import PageBreak
            story.append(PageBreak())
        story+=[Paragraph(oficina,ParagraphStyle('h',parent=sty['Heading2'],fontName='Helvetica-Bold',fontSize=11,spaceBefore=6,spaceAfter=4))];rows=[['Tipo','Documento','Data','Valor Real']]+[[x.TIPO,x.DOC,x.DATA,br_money(x.VALOR_REAL)] for x in g.itertuples()]+[['','TOTAL DA OFICINA','',br_money(g.VALOR_REAL.sum())]];t=Table(rows,colWidths=[18*mm,78*mm,34*mm,48*mm],repeatRows=1);t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#344054')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),('BACKGROUND',(0,-1),(-1,-1),colors.HexColor('#F2F4F7')),('GRID',(0,0),(-1,-1),.3,colors.HexColor('#D0D5DD')),('ALIGN',(3,1),(3,-1),'RIGHT'),('FONTSIZE',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]));story.append(t)
    doc.build(story,onFirstPage=cab,onLaterPages=cab);return b.getvalue()

st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#F6F8FB}
[data-testid="stHeader"]{background:rgba(255,255,255,.88)}
.block-container{max-width:1500px;padding-top:1.7rem;padding-bottom:3rem}
.hero{background:linear-gradient(135deg,#101828 0%,#344054 100%);border-radius:20px;padding:26px 30px;margin-bottom:22px;box-shadow:0 12px 32px rgba(16,24,40,.12)}
.hero h1{color:#fff!important;margin:0;font-size:2rem}.hero p{color:#D0D5DD;margin:7px 0 0}
div[data-testid="stMetric"]{background:#fff;border:1px solid #EAECF0;padding:16px 18px;border-radius:15px;box-shadow:0 4px 14px rgba(16,24,40,.04)}
div[data-testid="stMetricLabel"]{color:#667085} div[data-testid="stMetricValue"]{font-weight:750}
[data-testid="stFileUploader"]{background:#fff;border:1px solid #EAECF0;border-radius:15px;padding:12px}
[data-testid="stDataFrame"],[data-testid="stDataEditor"]{border:1px solid #EAECF0;border-radius:14px;overflow:hidden;background:#fff}
.stButton>button,.stDownloadButton>button{border-radius:10px;font-weight:650;min-height:42px}
div[data-testid="stAlert"]{border-radius:12px}
h2,h3{color:#101828;letter-spacing:-.02em}
.section-note{color:#667085;font-size:.9rem;margin-top:-8px;margin-bottom:12px}
.boleto-card{background:#fff;border:1px solid #EAECF0;border-radius:14px;padding:14px 16px;margin-bottom:8px}
hr{border-color:#EAECF0}
</style>
<div class="hero">
  <h1>Fechamento de Oficinas</h1>
  <p>Importe os relatórios, revise as NFs, informe os boletos e gere o fechamento.</p>
</div>
""",unsafe_allow_html=True)
st.markdown("### Manual de instruções")
st.caption("Consulte o POP completo com o passo a passo de extração no Autcom e utilização da aplicação.")
manual_path = Path(__file__).with_name("POP_Fechamento_de_Oficinas.pdf")
if manual_path.exists():
    with open(manual_path, "rb") as manual_file:
        st.download_button(
            "📘 Abrir / baixar Manual de Instruções (PDF)",
            data=manual_file.read(),
            file_name="POP_Fechamento_de_Oficinas.pdf",
            mime="application/pdf",
            use_container_width=True
        )
else:
    st.warning("Manual de instruções não encontrado. Coloque o arquivo POP_Fechamento_de_Oficinas.pdf na mesma pasta do aplicativo.")

st.divider()

a,b=st.columns(2)
with a:f1=st.file_uploader('1. Fechamento das Oficinas',type=['csv','xlsx','xls'])
with b:f2=st.file_uploader('2. Tabela 06',type=['csv','xlsx','xls'])
if not(f1 and f2):st.info('Selecione os dois arquivos para iniciar.');st.stop()
try:
    fech=load_fech(f1);tab=load_tab(f2);base=fech.merge(tab,on='COD',how='left',validate='many_to_one');miss=base[base.PRECO6.isna()][['COD']].drop_duplicates();base['VALOR_REAL']=base.QTD*base.PRECO6.fillna(0)
except Exception as e:st.error(f'Erro ao processar: {e}');st.stop()
if len(miss):st.error(f'Fechamento bloqueado: {len(miss)} código(s) realmente não encontrados na Tabela 06 após normalizar os códigos.');st.dataframe(miss,hide_index=True);st.stop()
d=base.DATA_DT.dropna();periodo=f'{d.min():%d/%m/%Y} a {d.max():%d/%m/%Y}' if len(d) else '-';c1,c2,c3,c4=st.columns(4);c1.metric('Período',periodo);c2.metric('Oficinas',base.CLIENTE.nunique());c3.metric('Documentos',base.DOC.nunique());c4.metric('Itens',len(base));st.success('Todos os códigos possuem correspondência na Tabela 06.')

st.subheader('1. NFs — manter ou retirar')
nf=base[base.TIPO.eq('NF')].groupby(['CLIENTE','DOC','DATA'],as_index=False).agg(VALOR_SISTEMA=('VR_TOTAL','sum'),VALOR_REAL=('VALOR_REAL','sum')).sort_values(['CLIENTE','DATA','DOC'])
if nf.empty:st.info('Nenhuma NF encontrada.');keep=set()
else:
    if 'nf_keep' not in st.session_state:st.session_state.nf_keep={}
    v=nf.rename(columns={'CLIENTE':'Oficina','DOC':'Documento','DATA':'Data','VALOR_SISTEMA':'Valor Sistema','VALOR_REAL':'Valor Real'});v.insert(0,'Manter',[st.session_state.nf_keep.get(x,True) for x in v.Documento]);e=st.data_editor(v,hide_index=True,use_container_width=True,disabled=['Oficina','Documento','Data','Valor Sistema','Valor Real'],column_config={'Manter':st.column_config.CheckboxColumn('Manter'),'Valor Sistema':st.column_config.NumberColumn(format='R$ %.2f'),'Valor Real':st.column_config.NumberColumn(format='R$ %.2f')});st.session_state.nf_keep={x.Documento:bool(x.Manter) for x in e.itertuples()};keep=set(e.loc[e.Manter,'Documento'])
basef=base[(base.TIPO!='NF')|base.DOC.isin(keep)].copy()

st.subheader('2. Valores dos boletos')
st.markdown('<div class="section-note">Informe o valor efetivo de cada boleto. O campo aceita 29412,62 ou R$ 29.412,62 e formata automaticamente ao confirmar.</div>',unsafe_allow_html=True)
ag=basef.groupby('CLIENTE',as_index=False).agg(Valor_Sistema=('VR_TOTAL','sum'),Valor_Real=('VALOR_REAL','sum')).sort_values('CLIENTE')

if 'boletos' not in st.session_state:
    st.session_state.boletos={}

linhas_boleto=[]
for _,row in ag.iterrows():
    oficina=str(row['CLIENTE'])
    safe=re.sub(r'[^A-Za-z0-9]+','_',oficina)
    key='boleto_'+safe
    if key not in st.session_state:
        st.session_state[key]=br_money(st.session_state.boletos.get(oficina,0.0))

    with st.container(border=True):
        c_of,c_sis,c_real,c_bol=st.columns([2.3,1,1,1.15],vertical_alignment='center')
        with c_of:
            st.markdown(f"**{oficina}**")
        with c_sis:
            st.caption('Valor Sistema')
            st.markdown(f"**{br_money(row['Valor_Sistema'])}**")
        with c_real:
            st.caption('Valor Real')
            st.markdown(f"**{br_money(row['Valor_Real'])}**")
        with c_bol:
            valor_txt=st.text_input(
                'Valor Boleto',
                key=key,
                on_change=format_boleto_field,
                args=(key,),
                label_visibility='visible'
            )

    valor_boleto=parse_brl_input(valor_txt)
    st.session_state.boletos[oficina]=valor_boleto
    linhas_boleto.append({
        'Oficina':oficina,
        'Valor Sistema':float(row['Valor_Sistema']),
        'Valor Real':float(row['Valor_Real']),
        'Valor Boleto':valor_boleto
    })

be=pd.DataFrame(linhas_boleto)
r=be[['Oficina','Valor Boleto','Valor Real']].copy()
r['Diferença']=r['Valor Boleto']-r['Valor Real']

st.subheader('3. Conferência final');st.markdown('<div class="section-note">Comparativo consolidado entre o valor informado no boleto e o valor recalculado pela Tabela 06.</div>',unsafe_allow_html=True);st.dataframe(r.style.format({'Valor Boleto':br_money,'Valor Real':br_money,'Diferença':br_money}),hide_index=True,use_container_width=True);x,y,z=st.columns(3);x.metric('Total Boletos',br_money(r['Valor Boleto'].sum()));y.metric('Total Real',br_money(r['Valor Real'].sum()));z.metric('Diferença Geral',br_money(r['Diferença'].sum()))

p1=pdf_resumo(r,competencia(base));p2=pdf_documentos(basef,competencia(base));zb=io.BytesIO()
with zipfile.ZipFile(zb,'w',zipfile.ZIP_DEFLATED) as zz:zz.writestr('Resumo_Fechamento_Oficinas.pdf',p1);zz.writestr('Documentos_Fechamento_Oficinas.pdf',p2)
st.subheader('4. Relatórios');q1,q2,q3=st.columns(3)
with q1:st.download_button('Baixar PDF — Resumo',p1,'Resumo_Fechamento_Oficinas.pdf','application/pdf',use_container_width=True)
with q2:st.download_button('Baixar PDF — Documentos',p2,'Documentos_Fechamento_Oficinas.pdf','application/pdf',use_container_width=True)
with q3:st.download_button('Baixar pacote completo (.zip)',zb.getvalue(),'Fechamento_Oficinas.zip','application/zip',use_container_width=True)
with st.expander('Auditoria'):st.write(f"FFs incluídas: **{basef.loc[basef.TIPO.eq('FF'),'DOC'].nunique()}** | NFs incluídas: **{basef.loc[basef.TIPO.eq('NF'),'DOC'].nunique()}** de **{base.loc[base.TIPO.eq('NF'),'DOC'].nunique()}**");st.caption('O PDF de Documentos inclui todas as FFs e somente as NFs marcadas como Manter. Valor Real = QTD × Novo Pr. Venda 6.')
