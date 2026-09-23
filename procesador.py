"""Lectura y organización de reportes de marcaciones integrales."""

from io import BytesIO
import re
import unicodedata

import pandas as pd


def _normalizar(texto):
    texto = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in texto if not unicodedata.combining(c)).strip().lower()


def leer_archivo(archivo, nombre):
    """Devuelve (tabla, hoja) con las dos columnas necesarias.

    El encabezado del XLS de INEC viene después de cinco filas descriptivas.
    También se aceptan archivos con encabezados en la primera fila.
    """
    extension = nombre.rsplit(".", 1)[-1].lower()
    if extension in ("xls", "xlsx"):
        motor = "xlrd" if extension == "xls" else "openpyxl"
        hojas = pd.read_excel(archivo, sheet_name=None, header=None, engine=motor)
    elif extension in ("csv", "tsv", "txt"):
        datos = archivo.read() if hasattr(archivo, "read") else archivo
        if isinstance(datos, str):
            datos = datos.encode("utf-8")
        ultimo_error = None
        for codificacion in ("utf-8-sig", "cp1252"):
            try:
                tabla = pd.read_csv(BytesIO(datos), header=None, sep=None,
                                    engine="python", encoding=codificacion,
                                    dtype=str, on_bad_lines="skip")
                break
            except (UnicodeDecodeError, pd.errors.ParserError) as exc:
                ultimo_error = exc
        else:
            raise ValueError(f"No se pudo leer el archivo de texto: {ultimo_error}")
        hojas = {"Datos": tabla}
    else:
        raise ValueError("Formato no admitido. Carga un archivo .xls, .xlsx, .csv o .tsv.")

    for hoja, tabla in hojas.items():
        for fila in range(min(len(tabla), 40)):
            nombres = [_normalizar(x) for x in tabla.iloc[fila]]
            empleado = next((i for i, x in enumerate(nombres)
                             if x in ("empleado", "funcionario", "nombre", "servidor")), None)
            fecha = next((i for i, x in enumerate(nombres)
                          if x in ("fecha hora", "fecha y hora", "fechahora", "marcacion", "fecha de marcacion")), None)
            if empleado is not None and fecha is not None:
                salida = tabla.iloc[fila + 1:, [empleado, fecha]].copy()
                salida.columns = ["Empleado", "Fecha Hora"]
                return salida.reset_index(drop=True), hoja
    raise ValueError("No encontré las columnas 'Empleado' y 'Fecha Hora' en las primeras 40 filas de ninguna hoja.")


def _convertir_fecha(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return pd.NaT
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return pd.to_datetime(valor, unit="D", origin="1899-12-30", errors="coerce")
    texto = str(valor).strip()
    # Excel puede exportar el número de serie como texto.
    if re.fullmatch(r"\d{5}(?:[.,]\d+)?", texto):
        return pd.to_datetime(float(texto.replace(",", ".")), unit="D",
                              origin="1899-12-30", errors="coerce")
    return pd.to_datetime(valor, dayfirst=True, errors="coerce")


def transformar(datos):
    """Una fila por empleado y día, sin eliminar marcaciones repetidas."""
    datos = datos.copy()
    datos["Empleado"] = datos["Empleado"].astype("string").str.strip()
    fechas = datos["Fecha Hora"].map(_convertir_fecha)
    validos = datos["Empleado"].notna() & datos["Empleado"].ne("") & fechas.notna()
    descartados = datos.loc[~validos].copy()
    if not validos.any():
        raise ValueError("No hay marcaciones válidas. Revisa las columnas de empleado y fecha/hora.")

    datos = datos.loc[validos, ["Empleado"]].copy()
    datos["Instante"] = pd.to_datetime(fechas.loc[validos])
    datos["Fecha"] = datos["Instante"].dt.normalize()
    # mergesort mantiene el orden original si hay dos marcas idénticas.
    datos = datos.sort_values(["Empleado", "Fecha", "Instante"], kind="mergesort")
    datos["Orden"] = datos.groupby(["Empleado", "Fecha"]).cumcount() + 1
    datos["Hora"] = datos["Instante"].dt.strftime("%H:%M:%S")
    matriz = datos.pivot(index=["Empleado", "Fecha"], columns="Orden", values="Hora")
    matriz.columns = [f"Marcación {numero}" for numero in matriz.columns]
    matriz = matriz.reset_index()
    cantidades = datos.groupby(["Empleado", "Fecha"]).size().rename("N.º marcaciones").reset_index()
    matriz = matriz.merge(cantidades, on=["Empleado", "Fecha"], validate="one_to_one")
    cols = ["Empleado", "Fecha", "Día", "N.º marcaciones", "Estado"]
    dias = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
    matriz["Día"] = matriz["Fecha"].dt.dayofweek.map(lambda x: dias[x])
    matriz["Estado"] = matriz["N.º marcaciones"].map(
        lambda n: "4 marcaciones" if n == 4 else f"Revisar: {n} marcación" + ("" if n == 1 else "es"))
    matriz = matriz[cols + [c for c in matriz if c.startswith("Marcación ")]]
    matriz = matriz.sort_values(["Empleado", "Fecha"], kind="mergesort").reset_index(drop=True)
    return matriz, descartados


def crear_excel(matriz, descartados):
    """Crea la descarga para Excel con horas de origen y filas descartadas."""
    salida = BytesIO()
    with pd.ExcelWriter(salida, engine="xlsxwriter", datetime_format="dd/mm/yyyy",
                        date_format="dd/mm/yyyy") as escritor:
        matriz.to_excel(escritor, sheet_name="Matriz diaria", index=False)
        libro = escritor.book
        hoja = escritor.sheets["Matriz diaria"]
        encabezado = libro.add_format({"bold": True, "bg_color": "#143B55", "font_color": "white",
                                       "border": 0, "valign": "vcenter"})
        fecha = libro.add_format({"num_format": "dd/mm/yyyy"})
        alerta = libro.add_format({"bg_color": "#FFF0D9", "font_color": "#8C4A00"})
        for i, nombre in enumerate(matriz.columns):
            hoja.write(0, i, nombre, encabezado)
            hoja.set_column(i, i, 19 if nombre.startswith("Marcación") else 22)
        hoja.set_column(0, 0, 32)
        hoja.set_column(1, 1, 16, fecha)
        hoja.set_column(4, 4, 24)
        hoja.freeze_panes(1, 2)
        hoja.autofilter(0, 0, len(matriz), len(matriz.columns) - 1)
        if len(matriz):
            hoja.conditional_format(1, 4, len(matriz), 4,
                                    {"type": "text", "criteria": "containing",
                                     "value": "Revisar", "format": alerta})
        if len(descartados):
            descartados.to_excel(escritor, sheet_name="Filas sin procesar", index=False)
            escritor.sheets["Filas sin procesar"].set_column(0, 1, 29)
    return salida.getvalue()
