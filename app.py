import io
import re
import hashlib
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
import streamlit as st

st.set_page_config(page_title='SANA | Contabilidad bancaria', page_icon='📒', layout='wide')
st.title('📒 SANA STUDIO BY PAUNANI S.A.S.')
st.caption('Contabilidad bancaria · Libro diario · Mayor · Balance de comprobación | USD')

DEFAULT_ACCOUNTS = [
 ('1.1.01','Bancos - cuenta corriente','Activo'),
 ('1.1.02','Inventarios - mercaderías','Activo'),
 ('1.1.03','Anticipos a proveedores','Activo'),
 ('1.1.04','Cuentas por cobrar a socias','Activo'),
 ('1.1.05','IVA crédito tributario','Activo'),
 ('2.1.01','Préstamos por pagar - terceros','Pasivo'),
 ('2.1.02','Préstamos por pagar - socias','Pasivo'),
 ('2.1.03','Tarjeta de crédito por pagar','Pasivo'),
 ('2.1.04','Cuentas por pagar a proveedores','Pasivo'),
 ('3.1.01','Capital social','Patrimonio'),
 ('3.1.02','Aportes de socias por formalizar','Patrimonio'),
 ('4.1.01','Ingresos por ventas','Ingreso'),
 ('5.1.01','Costo de ventas','Costo'),
 ('5.2.01','Gastos de viaje','Gasto'),
 ('5.2.02','Comisiones bancarias','Gasto'),
 ('5.2.03','Gastos administrativos','Gasto'),
 ('9.9.01','Pendiente de identificación (deudor)','Transitoria'),
 ('9.9.02','Pendiente de identificación (acreedor)','Transitoria'),
]
ACCOUNTS = pd.DataFrame(DEFAULT_ACCOUNTS,columns=['Código','Cuenta','Tipo'])
ACCOUNT_LABELS = {str(r['Código']):f"{r['Código']} · {r['Cuenta']}" for _,r in ACCOUNTS.iterrows()}
BANK='1.1.01'


def parse_amount(value):
 if pd.isna(value): return 0.0
 s=str(value).strip().replace('$','').replace(' ','')
 if not s: return 0.0
 if ',' in s and '.' in s:
  s=s.replace(',','') if s.rfind('.')>s.rfind(',') else s.replace('.','').replace(',','.')
 elif ',' in s:
  s=s.replace(',','.')
 try: return float(Decimal(s).quantize(Decimal('0.01'),rounding=ROUND_HALF_UP))
 except Exception: return float('nan')


def read_bank(file):
 raw=pd.read_excel(file, sheet_name=0, header=None, dtype=str)
 header_idx=None
 for idx in range(min(len(raw),50)):
  normalized=[str(v).strip().upper() for v in raw.iloc[idx].tolist()]
  if 'FECHA' in normalized and 'REFERENCIA' in normalized and ('VALOR' in normalized or 'MONTO' in normalized):
   header_idx=idx;break
 if header_idx is None: raise ValueError('No se localizó el encabezado FECHA / REFERENCIA / VALOR.')
 headers=[str(x).strip().upper() if pd.notna(x) else f'COL_{i}' for i,x in enumerate(raw.iloc[header_idx])]
 df=raw.iloc[header_idx+1:].copy()
 df.columns=headers
 df=df.dropna(how='all')
 def col(name,default=''):
  return df[name] if name in df.columns else pd.Series([default]*len(df),index=df.index)
 date= pd.to_datetime(col('FECHA'),errors='coerce',format='%m/%d/%Y %I:%M:%S %p')
 bad=date.isna()
 if bad.any(): date.loc[bad]=pd.to_datetime(col('FECHA')[bad],errors='coerce',dayfirst=True)
 out=pd.DataFrame({
  'Fecha':date,
  'Referencia':col('REFERENCIA').fillna('').astype(str).str.strip(),
  'Descripción banco':col('DESCRIPCION').fillna('').astype(str).str.strip(),
  'Signo':col('+/-').fillna('').astype(str).str.strip(),
  'Valor':col('VALOR').apply(parse_amount),
  'Saldo banco':col('SALDO CONTABLE').apply(parse_amount),
  'Tipo informado':col('TIPO').fillna('').astype(str).str.strip(),
  'Archivo':file.name,
  'Fila origen':df.index+1,
 })
 out=out[out['Fecha'].notna()].copy()
 if out.empty: raise ValueError('No se encontraron movimientos con fechas válidas.')
 if out['Valor'].isna().any(): raise ValueError('Hay importes no numéricos.')
 if not out['Signo'].isin(['+','-']).all(): raise ValueError('Existen signos distintos de + y -.')
 out['Valor']=out['Valor'].abs().round(2)
 if (out['Valor']<=0).any(): raise ValueError('Existen valores cero; revíselos antes de contabilizar.')
 # Se incluye saldo en la huella para distinguir débitos sucesivos con igual referencia.
 out['ID']=out.apply(lambda r:hashlib.sha256('|'.join([str(r['Fecha']),r['Referencia'],r['Descripción banco'],r['Signo'],f"{r['Valor']:.2f}",f"{r['Saldo banco']:.2f}"]).encode()).hexdigest()[:18],axis=1)
 return out


def guess(row):
 typ=str(row['Tipo informado']).lower().strip()
 desc=str(row['Descripción banco']).lower()
 incoming=row['Signo']=='+'
 account=('9.9.02' if incoming else '9.9.01')
 category='Por identificar'; note='Revisar sustento y naturaleza de la transacción.'
 if ('capital' in typ or 'aporte' in typ) and incoming:
  account='3.1.02'; category='Aporte socias';note='Verificar acta y registro societario; no confundir con capital inscrito.'
 elif ('prestamo' in typ or 'préstamo' in typ or 'credito' in typ or 'crédito' in typ):
  account='2.1.02' if 'socia' in typ else '2.1.01'
  category='Préstamo recibido' if incoming else 'Pago de préstamo'
  note='Verificar acreedor y distinguir capital de intereses; este asiento solo refleja capital provisional.'
 elif ('comision' in typ or 'comisión' in typ or 'tarifa' in desc):
  if not incoming: account='5.2.02';category='Comisión bancaria';note='Verificar comprobante y tratamiento tributario.'
 elif ('viaje' in typ):
  account='9.9.02' if incoming else '5.2.01'
  category='Reintegro por viaje' if incoming else 'Gasto de viaje'
  note='Verificar beneficiario, factura y si corresponde reintegro, préstamo o ingreso.'
 elif ('tarjeta' in typ or 'tarjeta de credito' in desc or 'tarjeta de crédito' in desc) and not incoming:
  account='2.1.03';category='Pago tarjeta de crédito';note='Comprobar conciliación con estado de tarjeta y separar cargos financieros.'
 return pd.Series([category,account,note])


def as_excel(sheets):
 b=io.BytesIO()
 with pd.ExcelWriter(b,engine='xlsxwriter',datetime_format='dd/mm/yyyy hh:mm',date_format='dd/mm/yyyy') as writer:
  wb=writer.book
  money=wb.add_format({'num_format':'#,##0.00;[Red](#,##0.00)'})
  header=wb.add_format({'bold':True,'bg_color':'#19365B','font_color':'white','border':0})
  for title,frame in sheets.items():
   clean=frame.copy()
   for col in clean.columns:
    if pd.api.types.is_datetime64_any_dtype(clean[col]):
     clean[col]=clean[col].dt.tz_localize(None) if getattr(clean[col].dt,'tz',None) else clean[col]
   clean.to_excel(writer,sheet_name=title[:31],index=False)
   ws=writer.sheets[title[:31]]
   ws.freeze_panes(1,0)
   ws.autofilter(0,0,len(clean),max(0,len(clean.columns)-1))
   for j,c in enumerate(clean.columns):
    vals=[len(str(x)) for x in clean[c].head(100).dropna().values]
    width=min(max(len(str(c))+3,max(vals,default=0)+2,14),48)
    ws.set_column(j,j,width,money if c in ('Debe','Haber','Saldo','Valor','Saldo banco') else None)
    ws.write(0,j,str(c),header)
 b.seek(0);return b.getvalue()

st.info('Los asientos son propuestas a partir del banco: comprueba documentos de respaldo, IVA, retenciones y clasificación antes del cierre contable.')
files=st.file_uploader('1. Carga uno o varios extractos bancarios Excel',type=['xlsx','xls'],accept_multiple_files=True)
if not files:
 st.write('Formato reconocido: FECHA · REFERENCIA · DESCRIPCION · +/- · VALOR · SALDO CONTABLE · tipo (opcional).')
 st.stop()
frames=[]; errors=[]
for file in files:
 try: frames.append(read_bank(file))
 except Exception as exc: errors.append(f'{file.name}: {exc}')
for error in errors: st.error(error)
if not frames: st.stop()
combined=pd.concat(frames,ignore_index=True)
duplicated=combined.duplicated('ID',keep='first')
st.caption(f'{len(files)} archivo(s) recibidos · {int(duplicated.sum())} movimiento(s) repetido(s) descartados entre extractos. Si importas cuentas bancarias distintas con movimientos coincidentes, revísalas por separado.')
mov=combined.loc[~duplicated].copy().sort_values(['Fecha','ID']).reset_index(drop=True)
mov[['Categoría','Contracuenta','Observación']]=mov.apply(guess,axis=1)
mov['Estado']=mov['Contracuenta'].apply(lambda v:'REVISAR' if str(v).startswith('9.') else 'PROPUESTO')

st.subheader('2. Configuración')
with st.expander('Ver plan de cuentas inicial',expanded=False):
 st.dataframe(ACCOUNTS,use_container_width=True,hide_index=True)
 st.caption('Plan referencial; adapta sus códigos y cuentas al catálogo oficial de SANA antes de contabilizar.')
with st.expander('Reglas y criterio contable',expanded=False):
 st.write('Cada abono bancario: Debe Bancos, Haber contracuenta. Cada débito: Debe contracuenta, Haber Bancos. Se propone una contracuenta inicial usando la columna tipo y la descripción. Las cuentas transitorias 9.9.xx representan movimientos no clasificados y no deben considerarse una contabilidad cerrada.')

min_date=mov['Fecha'].dt.date.min();max_date=mov['Fecha'].dt.date.max()
mode=st.radio('Período de consulta', ['Todo','Diario','Mensual','Anual','Rango'],horizontal=True)
period=mov.copy()
if mode=='Diario':
 chosen=st.date_input('Día',value=max_date)
 period=period[period['Fecha'].dt.date==chosen]
elif mode=='Mensual':
 options=sorted(mov['Fecha'].dt.strftime('%Y-%m').unique().tolist())
 chosen=st.selectbox('Mes',options,index=len(options)-1)
 period=period[period['Fecha'].dt.strftime('%Y-%m')==chosen]
elif mode=='Anual':
 options=sorted(mov['Fecha'].dt.year.unique().tolist())
 chosen=st.selectbox('Año',options,index=len(options)-1)
 period=period[period['Fecha'].dt.year==chosen]
elif mode=='Rango':
 start,end=st.date_input('Desde / hasta',value=(min_date,max_date))
 period=period[period['Fecha'].dt.date.between(start,end)]

st.subheader('3. Revisión y clasificación')
st.caption('Edita la contracuenta, categoría y observación. Para transacciones de préstamos, IVA, retenciones o compras con varios componentes puede ser necesario un asiento adicional fuera de esta app.')
show=['Fecha','Referencia','Descripción banco','Tipo informado','Signo','Valor','Categoría','Contracuenta','Estado','Observación','Archivo','ID']
editor=st.data_editor(period[show].reset_index(drop=True),hide_index=True,use_container_width=True,num_rows='fixed',key='transactions_editor',disabled=['Fecha','Referencia','Descripción banco','Tipo informado','Signo','Valor','Archivo','ID'],column_config={
 'Contracuenta':st.column_config.SelectboxColumn('Contracuenta',options=list(ACCOUNT_LABELS.keys()),required=True),
 'Valor':st.column_config.NumberColumn('Valor USD',format='%.2f'),
 'Observación':st.column_config.TextColumn('Observación',width='large')
})
editor['Estado']=editor['Contracuenta'].apply(lambda x:'REVISAR' if str(x).startswith('9.') else 'PROPUESTO')
needs=editor['Estado'].eq('REVISAR').sum()
col1,col2,col3=st.columns(3)
col1.metric('Movimientos',len(editor))
col2.metric('Abonos USD',f"{editor.loc[editor['Signo']=='+','Valor'].sum():,.2f}")
col3.metric('Débitos USD',f"{editor.loc[editor['Signo']=='-','Valor'].sum():,.2f}")
if needs: st.warning(f'{needs} movimiento(s) con cuentas transitorias por identificar. Se exportarán como borrador, no como asientos definitivos.')
if editor.empty: st.warning('No hay movimientos en el período seleccionado.');st.stop()

labels={str(r['Código']):r['Cuenta'] for _,r in ACCOUNTS.iterrows()}
rows=[]
for _,r in editor.sort_values(['Fecha','ID']).iterrows():
 date=r['Fecha'];base={'Fecha':date,'Asiento':r['ID'],'Referencia':r['Referencia'],'Detalle':str(r['Categoría'])+' | '+str(r['Descripción banco']),'Archivo origen':r['Archivo'],'Estado':r['Estado']}
 val=round(float(r['Valor']),2)
 contra=str(r['Contracuenta'])
 for idx,(code,debit,credit) in enumerate(([(BANK,val,0.0),(contra,0.0,val)] if r['Signo']=='+' else [(contra,val,0.0),(BANK,0.0,val)]),start=1):
  rows.append({**base,'Línea':idx,'Código':code,'Cuenta':labels.get(code,'Cuenta no identificada'),'Debe':debit,'Haber':credit})
journal=pd.DataFrame(rows)
accounts_df=journal.groupby(['Código','Cuenta'],as_index=False)[['Debe','Haber']].sum()
accounts_df['Saldo deudor']=(accounts_df['Debe']-accounts_df['Haber']).clip(lower=0).round(2)
accounts_df['Saldo acreedor']=(accounts_df['Haber']-accounts_df['Debe']).clip(lower=0).round(2)
major=journal.sort_values(['Código','Fecha','Asiento','Línea']).copy()
major['Variación']=major['Debe']-major['Haber']
major['Saldo']=major.groupby('Código')['Variación'].cumsum().round(2)
major=major.drop(columns=['Variación'])
difference=round(journal['Debe'].sum()-journal['Haber'].sum(),2)
st.subheader('4. Reportes y exportación')
if abs(difference)<0.005:st.success(f'Libro diario cuadrado: debe = haber = ${journal["Debe"].sum():,.2f}')
else:st.error(f'Desbalance: ${difference:,.2f}')
tabs=st.tabs(['Libro diario','Libro mayor','Balance de comprobación','Movimientos base'])
with tabs[0]:st.dataframe(journal,use_container_width=True,hide_index=True)
with tabs[1]:st.dataframe(major,use_container_width=True,hide_index=True)
with tabs[2]:st.dataframe(accounts_df,use_container_width=True,hide_index=True)
with tabs[3]:st.dataframe(editor,use_container_width=True,hide_index=True)
meta=pd.DataFrame([{'Empresa':'SANA STUDIO BY PAUNANI S.A.S.','Período':mode,'Generado':datetime.now().strftime('%Y-%m-%d %H:%M'),'Movimientos':len(editor),'Cuentas pendientes':int(needs),'Observación':'BORRADOR. Conciliar y sustentar documentos; no sustituye contabilidad completa.'}])
books={'Resumen':meta,'Libro diario':journal,'Libro mayor':major,'Balance comprobacion':accounts_df,'Movimientos bancarios':editor,'Plan de cuentas':ACCOUNTS}
out=as_excel(books)
st.download_button('⬇️ Descargar contabilidad bancaria (Excel)',data=out,file_name='SANA_contabilidad_bancaria.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
st.caption('Los cambios del editor se incluyen en el Excel descargado. La aplicación no guarda datos permanentemente en servidor ni modifica el archivo bancario original.')
