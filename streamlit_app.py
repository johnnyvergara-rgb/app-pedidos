# app_pedidos_aggrid.py

import streamlit as st
import pandas as pd
import sqlite3
import os
import datetime
import getpass
import time
import subprocess
import re
import shutil
import glob
import fractions
import numpy as np
import base64

# Imports opcionales que solo funcionan en Windows (para versión local)
try:
    import pyodbc  # Solo disponible en entorno Windows con ODBC
    PYODBC_AVAILABLE = True
except ImportError:
    PYODBC_AVAILABLE = False

try:
    import win32com.client
    import pythoncom
    from win32com.client import Dispatch
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    win32com = None
    pythoncom = None
    Dispatch = None

from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode

# === Conexión a SQLite ===
db_path = "pedidos.sqlite"
conn = sqlite3.connect(db_path)

# === Preparar columnas: renombrar y eliminar duplicadas ===
def preparar_columnas_pedidos(conn):
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(Pedidos);")
    columnas = [col[1] for col in cursor.fetchall()]

    cambios = {}
    if "Fec#Puesta" in columnas and "Fec.Puesta" not in columnas:
        cambios["Fec#Puesta"] = "Fec.Puesta"
    if "Pos#OFA" in columnas and "Pos.OFA" not in columnas:
        cambios["Pos#OFA"] = "Pos.OFA"

    if not cambios:
        return

    selects = [f'"{k}" AS "{v}"' for k, v in cambios.items()]
    columnas_preservar = [f'"{col}"' for col in columnas if col not in cambios]
    consulta = f"""
    BEGIN TRANSACTION;
    CREATE TABLE Pedidos_new AS
    SELECT {", ".join(columnas_preservar + selects)} FROM Pedidos;
    DROP TABLE Pedidos;
    ALTER TABLE Pedidos_new RENAME TO Pedidos;
    COMMIT;
    """
    cursor.executescript(consulta)
    conn.commit()
    st.success("✅ Columnas renombradas correctamente.")

def renombrar_columna_fec_puesta(conn):
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(Pedidos);")
    columnas = [col[1] for col in cursor.fetchall()]

    if "Fec#Puesta" in columnas and "Fec.Puesta" not in columnas:
        conn.executescript("""
        BEGIN TRANSACTION;
        CREATE TABLE Pedidos_new AS
        SELECT *, "Fec#Puesta" AS "Fec.Puesta" FROM Pedidos;
        DROP TABLE Pedidos;
        ALTER TABLE Pedidos_new RENAME TO Pedidos;
        COMMIT;
        """)
        st.success("✅ Columna 'Fec#Puesta' renombrada a 'Fec.Puesta' correctamente.")

renombrar_columna_fec_puesta(conn)

# === Configuración de la página Streamlit ===
st.set_page_config(page_title="📋 Sistema SysPro V1", layout="wide")
st.sidebar.title("📚 Navegación")
if "pagina" not in st.session_state:
    st.session_state["pagina"] = "📋 Principal"

pagina = st.sidebar.radio("Ir a:", ["📋 Principal", "📊 Ver Tablas", "📊 Análisis de OFAs"])

# === Utilidades varias ===

def formatear_fecha(df, columna):
    if columna in df.columns:
        df[columna] = pd.to_datetime(df[columna], errors="coerce").dt.date
    return df

def convertir_a_numero(df, columnas):
    for col in columnas:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df

def agregar_filtros_sidebar(df):
    st.sidebar.markdown("### 🔍 Filtros")
    filtros = {}
    for col in df.columns:
        if df[col].dtype == "object":
            valores = sorted(df[col].dropna().unique().tolist())
            seleccion = st.sidebar.multiselect(f"Filtrar {col}", ["(Todos)"] + valores, default="(Todos)")
            filtros[col] = None if "(Todos)" in seleccion else seleccion
        elif np.issubdtype(df[col].dtype, np.number):
            min_val = float(df[col].min())
            max_val = float(df[col].max())
            rango = st.sidebar.slider(f"Rango de {col}", min_val, max_val, (min_val, max_val))
            filtros[col] = rango
        elif np.issubdtype(df[col].dtype, "datetime64[ns]"):
            min_val = df[col].min().date()
            max_val = df[col].max().date()
            rango = st.sidebar.date_input(f"Rango de {col}", (min_val, max_val))
            filtros[col] = None if len(rango) != 2 else rango
    return filtros

def aplicar_filtros(df, filtros):
    for col, val in filtros.items():
        if val is None:
            continue
        if isinstance(val, list):
            df = df[df[col].isin(val)]
        elif isinstance(val, tuple) and len(val) == 2:
            if np.issubdtype(df[col].dtype, np.number):
                df = df[(df[col] >= val[0]) & (df[col] <= val[1])]
            elif np.issubdtype(df[col].dtype, "datetime64[ns]"):
                df = df[(df[col] >= pd.to_datetime(val[0])) & (df[col] <= pd.to_datetime(val[1]))]
    return df

def mostrar_tabla_aggrid(df, editable=False, key="tabla"):
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_pagination(paginationAutoPageSize=True)
    gb.configure_side_bar()
    gb.configure_default_column(editable=editable, filter=True, sortable=True, resizable=True)
    grid_options = gb.build()
    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        enable_enterprise_modules=False,
        fit_columns_on_grid_load=True,
        theme="streamlit",
        key=key,
    )
    return grid_response["data"]

# === Funciones específicas de negocio ===
# (aquí siguen todas tus funciones originales: cargar_manual_listmarc,
# obtener_ofas_nuevas_desde_outlook, login_sap_via_sendkeys, ejecutar_sap, etc.
# No las modifico para no romper la lógica, solo las dejamos tal cual.)

# ... (tus funciones de negocio existentes van aquí; no las tocamos) ...


# === Cargar Stock blanks (VERSIÓN COMPATIBLE NUBE) ===
def cargar_stock_blanks(uploaded_file):
    """Carga el stock de blanks desde un archivo XLS/XLSX subido por el usuario
    y lo normaliza antes de guardarlo en la tabla StockBlanks de SQLite.
    """
    try:
        import fractions as _fractions

        if uploaded_file is None:
            st.error("⚠️ Debes subir un archivo de stock (XLS o XLSX).")
            return

        # Leer archivo directamente desde el uploader
        df = pd.read_excel(uploaded_file)
        st.write("📦 Columnas de stock detectadas:", df.columns.tolist())

        # Eliminar columnas tipo CAMP
        columnas_filtradas = [col for col in df.columns if not col.upper().startswith("CAMP")]
        df = df[columnas_filtradas]

        # Normalizar columnas numéricas
        def convertir_a_coma(valor):
            try:
                valor = str(valor).strip().replace(",", ".")
                if "_" in valor:
                    entero, fraccion = valor.split("_")
                    valor = float(entero) + float(_fractions.Fraction(fraccion))
                elif " " in valor and "/" in valor:
                    entero, fraccion = valor.split()
                    valor = float(entero) + float(_fractions.Fraction(fraccion))
                elif "/" in valor:
                    valor = float(_fractions.Fraction(valor))
                else:
                    valor = float(valor)
                return str(int(valor)) if valor.is_integer() else str(round(valor, 4)).replace(".", ",")
            except Exception:
                return valor

        for col in ["ESP_CUB", "ANC_CUB", "LAR_CUB"]:
            if col in df.columns:
                df[col] = df[col].apply(convertir_a_coma)

        # Normalizar calidad
        def transformar_calidad(valor):
            valor = str(valor).strip().upper()
            if valor == "CLEAR_GB":
                return "CLEAR"
            elif valor in ["MCM_PECA", "MCM"]:
                return "MCM"
            elif valor in ["CLEAR_GB_PECA", "CLEAR_GB_TA", "(EN BLANCO)", "", "NONE"]:
                return "USA"
            else:
                return "MCR"

        if "CALIDAD" in df.columns:
            df["CALIDAD"] = df["CALIDAD"].apply(transformar_calidad)

        df.to_sql("StockBlanks", conn, if_exists="replace", index=False)
        st.success(f"✅ Stock cargado y normalizado con {len(df)} registros.")

    except Exception as e:
        st.error(f"❌ Error al cargar y normalizar stock: {e}")

# === BOTONES DE ACCIÓN EN FILA ===
st.markdown("### 🚀 Acciones principales")

# Usamos 4 columnas de tamaño igual, y un pequeño gap visual
cols = st.columns([1, 1, 1, 1], gap="small")

# Columna 0: OFAs + SAP (solo disponible en versión escritorio)
with cols[0]:
    if not WIN32_AVAILABLE:
        st.info("📥 Extracción de OFAs y ejecución de SAP solo está disponible en la versión de escritorio.")
    else:
        if st.button("📥 Extraer OFAs nuevas + Ejecutar SAP", use_container_width=True):
            try:
                nuevas = obtener_ofas_nuevas_desde_outlook(conn)
                if nuevas:
                    st.toast(f"✅ {len(nuevas)} OFAs nuevas extraídas y guardadas.")
                    time.sleep(2.2)

                    login_sap_via_sendkeys()
                    st.toast("✅ Login SAP completado.")
                    time.sleep(2.2)

                    ejecutar_sap()
                    st.toast("✅ SAP ejecutado correctamente.")
            except Exception as e:
                st.toast(f"❌ Error extrayendo OFAs: {e}", icon="❌")

# Columna 1: Cargar ListMarc.XLS (solo escritorio)
with cols[1]:
    if not WIN32_AVAILABLE:
        st.info("📂 Carga de ListMarc.XLS solo disponible en la versión de escritorio.")
    else:
        if st.button("📂 Cargar ListMarc.XLS", use_container_width=True):
            try:
                cargar_manual_listmarc()
                st.toast("✅ ListMarc.XLS cargado correctamente.")
            except Exception as e:
                st.toast(f"❌ Error cargando ListMarc.XLS: {e}", icon="❌")

# Columna 2: Cargar Stock (funciona en la nube)
with cols[2]:
    archivo_stock = st.file_uploader(
        "📤 Subir archivo de stock (XLS/XLSX)",
        type=["xls", "xlsx"],
        key="uploader_stock"
    )

    if st.button("📦 Cargar Stock", use_container_width=True):
        if archivo_stock is None:
            st.toast("⚠️ Primero debes subir un archivo de stock.", icon="⚠️")
        else:
            try:
                cargar_stock_blanks(archivo_stock)
                st.toast("✅ Stock cargado correctamente.")
            except Exception as e:
                st.toast(f"❌ Error cargando stock: {e}", icon="❌")

# Columna 3: Importar desde Access (solo escritorio con pyodbc)
with cols[3]:
    if not PYODBC_AVAILABLE:
        st.info("🗃 Importar pedidos desde Access requiere pyodbc y solo está disponible en la versión de escritorio.")
    else:
        if st.button("🔄 Importar Pedidos desde Access", use_container_width=True, key="import_access"):
            try:
                ruta_access = r"C:\Users\Johnny.Vergara\OneDrive - ARAUCO\Escritorio\Plan_SisRem13_01_2025_15_00.accdb"
                access_conn_str = (
                    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
                    f"DBQ={ruta_access};"
                )

                with pyodbc.connect(access_conn_str) as access_conn:
                    df_import = pd.read_sql("SELECT * FROM Pedidos", access_conn)

                df_import.to_sql("Pedidos", conn, if_exists="replace", index=False)
                st.success("✅ Tabla 'Pedidos' importada exitosamente desde Access.")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error importando desde Access: {e}")

# Botón adicional para exportar tablas (funciona en ambos entornos)
with cols[3]:
    if st.button("📊 Exportar tablas a Excel", use_container_width=True, key="exportar_excel"):
        try:
            pedidos_path = os.path.join(os.getcwd(), "Pedidos_export.xlsx")
            stock_path = os.path.join(os.getcwd(), "StockBlanks_export.xlsx")

            with pd.ExcelWriter(pedidos_path, engine="openpyxl") as writer:
                pd.read_sql("SELECT * FROM Pedidos", conn).to_excel(writer, sheet_name="Pedidos", index=False)

            with pd.ExcelWriter(stock_path, engine="openpyxl") as writer:
                pd.read_sql("SELECT * FROM StockBlanks", conn).to_excel(writer, sheet_name="Stock", index=False)

            st.toast("✅ Tablas exportadas exitosamente a Excel")
        except Exception as e:
            st.toast(f"❌ Error al exportar: {e}", icon="❌")

# === (Resto de tu código original: tablas, análisis, historial, etc.) ===
# Mantén aquí todas las secciones que no dependen de SAP/Access/Excel COM.
# Este archivo ya es ejecutable en Streamlit Cloud siempre que tu requirements.txt
# NO incluya pyodbc, win32com ni pythoncom (los importamos opcionalmente).

