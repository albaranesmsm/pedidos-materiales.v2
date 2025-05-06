import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.workbook.protection import WorkbookProtection
# --- DATOS FIJOS ---
COMPRADOR = "612539"
# --- RELACIÓN ARTÍCULOS Y PROVEEDORES ---
proveedores = {
   "1600043": "13161", "1600050": "13161", "1600051": "13161", "1600052": "13161",
   "1600053": "13161", "1600054": "13161", "1600055": "13161", "1600911": "13161",
   "1600921": "13161", "1601104": "13161", "1601161": "13161", "1601271": "13161",
   "1601306": "13161", "0400153": "10381", "0400176": "10381", "0400177": "10381",
   "0400232": "10381", "0400543": "10381", "0400548": "10381", "0400699": "10381",
   "1601001": "10381"
}
# --- RELACIÓN ARTÍCULOS Y OB ---
ob_values = {
   "1600043": "14001536", "1600050": "14001536", "1600051": "14001536", "1600052": "14001536",
   "1600053": "14001536", "1600054": "14001536", "1600055": "14001536", "1600911": "14001536",
   "1600921": "14001536", "1601104": "14001536", "1601161": "14001536", "1601271": "14001536",
   "1601306": "14001536", "0400153": "31005151", "0400176": "31005151", "0400177": "31005151",
   "0400232": "31005151", "0400543": "31005151", "0400548": "31005151", "0400699": "31005151",
   "1601001": "31005151"
}
# --- LISTADO DE ARTÍCULOS ---
articulos = [
   {"Nº artículo": "1600043", "Descripción": "TUB DESAG PVC//PVC"},
   {"Nº artículo": "1600050", "Descripción": "PYTHON A Inund 1 P"},
   {"Nº artículo": "1600051", "Descripción": "PYTHON A2 Inund 2 P"},
   {"Nº artículo": "1600052", "Descripción": "PYTHON L Contac 1P"},
   {"Nº artículo": "1600053", "Descripción": "PYTHON L2 Contac 2P"},
   {"Nº artículo": "1600054", "Descripción": "PYTHON L3 Contac 3P"},
   {"Nº artículo": "1600055", "Descripción": "TUB AGUA REF"},
   {"Nº artículo": "1600911", "Descripción": "PYTHON COL"},
   {"Nº artículo": "1600921", "Descripción": "LUPULUS 6,35 1P eventos"},
   {"Nº artículo": "1601104", "Descripción": "TUB ARM Riego arm"},
   {"Nº artículo": "1601161", "Descripción": "TUB GAS LDP"},
   {"Nº artículo": "1601271", "Descripción": "KIT ANTI-COND"},
   {"Nº artículo": "1601306", "Descripción": "FLEXLAYER 1P Vermut"},
   {"Nº artículo": "0400153", "Descripción": "DINFEX Antialgas"},
   {"Nº artículo": "0400176", "Descripción": "COMPACT 200 Limp inst."},
   {"Nº artículo": "0400177", "Descripción": "TOPFOAM Limp máq"},
   {"Nº artículo": "0400232", "Descripción": "PLUS ESPEC Limp inst."},
   {"Nº artículo": "0400543", "Descripción": "ALUTRAT Limp inst."},
   {"Nº artículo": "0400548", "Descripción": "ULTRASON Liquido"},
   {"Nº artículo": "0400699", "Descripción": "DIVOSAN TC86 SDC"},
   {"Nº artículo": "1601001", "Descripción": "GLICOL Anticong"}
]
# --- RESTRICCIONES POR ARTÍCULO ---
restricciones = {
   "1600043": {"multiplo": 25, "max": 1500}, "1600050": {"multiplo": 25, "max": 2000},
   "1600051": {"multiplo": 25, "max": 500}, "1600052": {"multiplo": 25, "max": 500},
   "1600053": {"multiplo": 25, "max": 500}, "1600054": {"multiplo": 25, "max": 500},
   "1600055": {"multiplo": 25, "max": 1500}, "1600911": {"multiplo": 25, "max": 1000},
   "1600921": {"multiplo": 25, "max": 6000}, "1601104": {"multiplo": 25, "max": 50},
   "1601161": {"multiplo": 25, "max": 5000}, "1601271": {"multiplo": 25, "max": 300},
   "1601306": {"multiplo": 25, "max": 300}, "0400153": {"multiplo": 10, "max": 300},
   "0400176": {"multiplo": 10, "max": 80}, "0400177": {"multiplo": 20, "max": 80},
   "0400232": {"multiplo": 600, "max": 1800}, "0400543": {"multiplo": 20, "max": 300},
   "0400548": {"multiplo": 20, "max": 50}, "0400699": {"multiplo": 24, "max": 240},
   "1601001": {"multiplo": 25, "max": 600}
}
# --- INTERFAZ ---
st.title("Pedido de tuberías y químicos")
# Acceder a las direcciones desde los secrets
direcciones = st.secrets["direcciones"]
# Solicitar al usuario el código de dirección de entrega
dir_entrega = st.text_input("Código de Dirección de Entrega (4 cifras empezando por 8):", max_chars=4)
# Validar el código de dirección de entrega
if not dir_entrega or not (dir_entrega.isdigit() and len(dir_entrega) == 4 and dir_entrega.startswith("8")):
   st.error("Debe introducir el código de almacén de envío")
   st.stop()
elif dir_entrega in direcciones:
   st.success(f"Dirección seleccionada: {direcciones[dir_entrega]}")
else:
   st.warning("Almacén no reconocido, contacte con OOVV antes de hacer el pedido.")
st.subheader("Selecciona las cantidades:")
pedido = []
errores = []
for articulo in articulos:
   codigo = str(articulo["Nº artículo"])
   descripcion = articulo["Descripción"]
   proveedor = proveedores.get(codigo)
   ob = ob_values.get(codigo)
   if not proveedor or not ob:
       errores.append(f"Artículo {codigo} sin proveedor u OB.")
       continue
   maximo = restricciones.get(codigo, {}).get("max", 1000)
   multiplo = restricciones.get(codigo, {}).get("multiplo", 1)
   cantidad = st.number_input(
       f"{descripcion} — Ref: {codigo} (Múltiplo: {multiplo}, Máx: {maximo})",
       min_value=0, max_value=maximo, step=multiplo, value=0,
   )
   if cantidad > 0:
       pedido.append({
           "Fecha solicitud": datetime.date.today(),
           "OB": ob,
           "Comprador": COMPRADOR,
           "LM aux": "00004014",
           "Cód Prov": proveedor,
           "Proveedor": "",
           "Suc/planta": 8040,
           "Dir entr": dir_entrega,
           "Nº artículo": codigo,
           "Descripción": descripcion,
           "Autorizar cant": cantidad,
       })
def crear_excel_protegido(df):
   wb = Workbook()
   ws = wb.active
   ws.append(df.columns.tolist())
   for _, row in df.iterrows():
       ws.append(row.tolist())
   wb.security = WorkbookProtection(workbookPassword="NESTARES_24", lockStructure=True)
   excel_stream = BytesIO()
   wb.save(excel_stream)
   excel_stream.seek(0)
   return excel_stream
# --- BOTÓN PARA GENERAR PEDIDO Y MOSTRAR INSTRUCCIONES ---
mostrar_instrucciones = False
if st.button("Generar Pedido"):
   if pedido:
       df = pd.DataFrame(pedido)
       excel_bytes = crear_excel_protegido(df)
       st.success("¡Pedido generado correctamente!")
       st.download_button("Descargar Pedido", data=excel_bytes, file_name="pedido_materiales.xlsx")
       mostrar_instrucciones = True
   else:
       st.warning("No se ha seleccionado ningún artículo.")
if mostrar_instrucciones:
   st.markdown("""
<div style='border: 2px solid #4CAF50; padding: 20px; border-radius: 10px; background-color: #f6fff6'>
<h4>¡Enhorabuena! Ya has generado tu pedido</h4>
<p>Ahora tienes que enviar el fichero generado por mail siguiendo estas instrucciones:</p>
<ol>
<li>HAZ CLICK EN <b>Descargar Pedido</b></li>
<li>Abre tu correo electrónico y crea un nuevo correo.</li>
<li><b>Destinatario:</b> robot1@mahou-sanmiguel.com</li>
<li><b>Asunto:</b>OAs pedidos materiales operaciones de venta</li>
<li>Adjunta el fichero que acabas de descargar.</li>
<li>Haz clic en <b>Enviar</b>.</li>
</ol>
<p>Una vez enviado, tendrás que esperar en torno a unos 30 minutos. Recibirás un correo automático con el resultado, es decir, con las OAs generadas. <b>IMPORTANTE, Asegura que el destinatario y el asunto es tal el que pone en las instrucciones si no no generará nada</b></p>
</div>
   """, unsafe_allow_html=True)
