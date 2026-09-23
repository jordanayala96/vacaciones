import streamlit as st

from procesador import crear_excel, leer_archivo, transformar


st.set_page_config(page_title="Marcaciones por día", page_icon="🕒", layout="wide")
st.title("Seguimiento diario de marcaciones")
st.caption("Convierte el reporte vertical del biométrico en una fila por funcionario y fecha.")

archivo = st.file_uploader("Carga el reporte de marcaciones", type=["xls", "xlsx", "csv", "tsv", "txt"])

if archivo is None:
    st.info("Admite el reporte XLS de Marcaciones Integrales; también archivos Excel, CSV y TSV con columnas Empleado y Fecha Hora.")
    st.stop()

try:
    datos, hoja = leer_archivo(archivo, archivo.name)
    matriz, descartados = transformar(datos)
except Exception as exc:
    st.error(f"No pude procesar el archivo: {exc}")
    st.stop()

st.success(f"Reporte leído: {len(datos):,} filas de la hoja «{hoja}».")
a, b, c, d = st.columns(4)
a.metric("Funcionarios", matriz["Empleado"].nunique())
b.metric("Días registrados", len(matriz))
c.metric("Marcaciones", int(matriz["N.º marcaciones"].sum()))
d.metric("Días para revisar", int(matriz["N.º marcaciones"].ne(4).sum()))

if len(descartados):
    st.warning(f"Se excluyeron {len(descartados)} filas sin empleado o sin fecha/hora válida. Puedes verlas en el Excel descargado.")

funcionarios = sorted(matriz["Empleado"].unique().tolist())
seleccion = st.multiselect("Filtrar por funcionario (solo vista previa)", funcionarios)
vista = matriz[matriz["Empleado"].isin(seleccion)] if seleccion else matriz
solo_revisar = st.checkbox("Mostrar únicamente días con una cantidad distinta de cuatro")
if solo_revisar:
    vista = vista[vista["N.º marcaciones"].ne(4)]

st.dataframe(vista, use_container_width=True, hide_index=True,
             column_config={"Fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY")})
st.caption("Las horas se ordenan de menor a mayor. Se conserva cada registro, incluso si dos marcaciones tienen la misma hora. Con cuatro marcaciones, las posiciones suelen corresponder a entrada, inicio de almuerzo, fin de almuerzo y salida; confirma cualquier excepción con el registro original.")

st.download_button(
    "Descargar matriz completa en Excel",
    data=crear_excel(matriz, descartados),
    file_name="seguimiento_marcaciones_diarias.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
