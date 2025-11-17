import pyodbc
import streamlit as st
import pandas as pd
import sqlite3
import os
import datetime
import getpass
import win32com.client
import time
import subprocess
import re
import pythoncom
import shutil
import glob
import fractions
from win32com.client import Dispatch
import numpy as np
import base64

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
    columnas_sql = ", ".join(columnas_preservar + selects)

    conn.executescript(f"""
    BEGIN TRANSACTION;
    CREATE TABLE Pedidos_new AS SELECT {columnas_sql} FROM Pedidos;
    DROP TABLE Pedidos;
    ALTER TABLE Pedidos_new RENAME TO Pedidos;
    COMMIT;
    """)
    st.success(f"✅ Columnas renombradas: {', '.join(cambios.values())}")

# Llamada a la función
preparar_columnas_pedidos(conn)


# === Renombrar columna si es necesario ===
def renombrar_columna_fec_puesta(conn):
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(Pedidos);")
    columnas = [col[1] for col in cursor.fetchall()]

    if "Fec.Puesta" in columnas:
        return

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

pagina = st.sidebar.selectbox(
    "Selecciona sección",
    ["📋 Principal", "📊 Ver Tablas"],
    index=["📋 Principal", "📊 Ver Tablas"].index(st.session_state["pagina"])
)


# === Vista de Tablas (solo Pedidos y Stock) ===
def mostrar_tablas():
    st.markdown("### 📋 Tabla de Pedidos (15 columnas clave)")
    try:
        df_tabla = pd.read_sql("SELECT rowid as id, * FROM Pedidos", conn)
        columnas_mostrar = [
            "id", "Hora_Grabado", "OFA", "Pos.OFA", "Maq", "Sec",
            "SemanaPues", "Fec.Puesta", "TEMPLATE", "Textomat",
            "EspesroMP", "AnchoMP", "CalidadMP", "Units", "Vol#M3"
        ]
        columnas_presentes = [col for col in columnas_mostrar if col in df_tabla.columns]
        df_tabla = df_tabla[columnas_presentes]
        st.dataframe(df_tabla.astype(str), use_container_width=True)
    except Exception as e:
        st.error(f"❌ Error al mostrar la tabla de Pedidos: {e}")

    st.markdown("---")
    st.markdown("### 📦 Tabla de StockBlanks")
    try:
        df_stock = pd.read_sql("SELECT * FROM StockBlanks", conn)
        st.dataframe(df_stock.astype(str), use_container_width=True)
    except Exception as e:
        st.error(f"❌ Error al mostrar la tabla de Stock: {e}")

# === Página adicional con todo ===
def mostrar_todo():
    st.title("📊 Análisis y Funciones Avanzadas")
    st.info("🔧 Aquí puedes mostrar más filtros, tablas o herramientas adicionales.")
    # Aquí podrías cargar el resto del sistema: filtros, análisis, etc.

# === Renderizado condicional ===
if pagina == "📋 Principal":
    st.title("📋 Sistema SysPro V1")

    if st.button("📑 Tablas Programación"):
        st.session_state["pagina"] = "📊 Ver Tablas"
        st.rerun()

    mostrar_tablas()


    

    # Mostrar solo las tablas básicas
    mostrar_tablas()



elif pagina == "📊 Ver Tablas":
    mostrar_todo()






# === Función: Obtener OFAs desde Outlook ===
def obtener_ofas_nuevas_desde_outlook(conn):
    try:
        pythoncom.CoInitialize()
        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6).Folders("SOPP")
        mensajes = inbox.Items
        mensajes.Sort("[ReceivedTime]", True)

        ofas_extraidas = []
        ahora = datetime.datetime.now().replace(tzinfo=None)

        for mensaje in mensajes:
            if hasattr(mensaje, "ReceivedTime") and hasattr(mensaje, "Subject"):
                received = mensaje.ReceivedTime.replace(tzinfo=None)
                if received and (ahora - received).days <= 3:
                    if mensaje.Subject == "Publicación a SAP":
                        cuerpo = mensaje.Body
                        encontrados = re.findall(r"510000\d+", cuerpo)
                        ofas_extraidas.extend(encontrados)

        if not ofas_extraidas:
            st.warning("No se encontraron OFAs recientes.")
            return []

        df_existente = pd.read_sql("SELECT OFA, [Pos.OFA] FROM Pedidos", conn)
        df_existente["ID_UNICO"] = df_existente["OFA"].astype(str) + "_" + df_existente["Pos.OFA"].astype(str)
        actuales = set(df_existente["ID_UNICO"])
        ofas_nuevas_id = set(f"{ofa}_10" for ofa in ofas_extraidas)
        nuevas = [ofa.split("_")[0] for ofa in ofas_nuevas_id if ofa not in actuales]

        if not nuevas:
            st.info("🔄 No hay nuevas OFAs para cargar.")
            return []

        ruta_txt = r"C:\A_Programación\Bajar Pedidos\Pegar en SAP1.txt"
        with open(ruta_txt, "w") as f:
            for ofa in nuevas:
                f.write(ofa + "\n")

        st.success(f"✅ {len(nuevas)} OFAs nuevas extraídas y guardadas.")
        return nuevas

    except Exception as e:
        st.error(f"❌ Error extrayendo OFAs: {e}")
        return []

# === Login a SAP ===
def login_sap_via_sendkeys():
    try:
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("WScript.Shell")
        subprocess.Popen(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplgpad.exe")
        time.sleep(6)
        shell.AppActivate("SAP Logon Pad 770")
        time.sleep(1)
        shell.SendKeys("~", 1)
        time.sleep(4)
        shell.SendKeys("jvergarar", 1)
        time.sleep(0.3)
        shell.SendKeys("{TAB}", 1)
        time.sleep(0.3)
        shell.SendKeys("Manjaratin1975@", 1)
        time.sleep(0.3)
        shell.SendKeys("~", 1)
        time.sleep(5)
        st.success("✅ Login SAP completado.")
    except Exception as e:
        st.error(f"❌ Error al ingresar a SAP: {e}")

# === Ejecutar SAP ===
def ejecutar_sap():
    try:
        archivo_exportado = r"C:\A_Programación\Bajar Pedidos\ListMarc.XLS"
        if os.path.exists(archivo_exportado):
            os.remove(archivo_exportado)

        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        app = SapGuiAuto.GetScriptingEngine
        session = app.Children(0).Children(0)

        session.findById("wnd[0]").Maximize()
        session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00022"
        session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00022")

        session.findById("wnd[0]/usr/btn%_PG_ORDFA_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[23]").press()
        session.findById("wnd[2]/usr/ctxtDY_PATH").Text = r"C:\A_Programación\Bajar Pedidos"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").Text = "Pegar en SAP1.txt"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 16
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        session.findById("wnd[0]/usr/chkP_CARACT").Selected = True
        session.findById("wnd[0]/usr/radVALOR2").Select()
        session.findById("wnd[0]/usr/ctxtPG_CENTR-LOW").Text = "TM02"
        session.findById("wnd[0]/usr/ctxtPG_CENTR-HIGH").Text = "TM02"
        session.findById("wnd[0]/usr/radVALOR2").SetFocus()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = r"C:\A_Programación\Bajar Pedidos"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ListMarc.XLS"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        session.findById("wnd[0]").close()
        time.sleep(1)
        try:
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except:
            pass

        st.success("✅ SAP ejecutado correctamente.")

    except Exception as e:
        st.error(f"❌ Error ejecutando pasos en SAP: {e}")

# === Cargar ListMarc con id_carga ===
def cargar_manual_listmarc():
    try:
        from db_utils import insertar_sin_duplicados

        ruta_base = r"C:\A_Programación\Bajar Pedidos"
        archivos = sorted(glob.glob(os.path.join(ruta_base, "*.XLS")), key=os.path.getmtime, reverse=True)
        if not archivos:
            st.error("❌ No se encontró ningún archivo .XLS.")
            return

        original = archivos[0]
        temporal = os.path.join(ruta_base, "temp_ListMarc.XLS")
        ruta_xlsx = temporal.replace(".XLS", ".xlsx")

        shutil.copyfile(original, temporal)
        pythoncom.CoInitialize()

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False

        libro = excel.Workbooks.Open(temporal, ReadOnly=True)
        libro.SaveAs(Filename=ruta_xlsx, FileFormat=51)
        libro.Close(SaveChanges=False)
        excel.Quit()

        # 📥 Leer el archivo convertido a xlsx
        df = pd.read_excel(uploaded_file, engine="openpyxl")


        # ✅ Verificar columnas cargadas
        st.write("📋 Columnas detectadas:", df.columns.tolist())

        # 🧼 Renombrar columnas para que coincidan con la BD
        df.rename(columns={
            "Pos#OFA": "Pos.OFA",
            "Material": "Materia",
            "Vol.M3": "Vol#M3",
            "Unit": "Units",
        }, inplace=True)

        # 🧽 Forzar la columna Fec.Puesta a formato datetime
        if "Fec.Puesta" in df.columns:
            df["Fec.Puesta"] = pd.to_datetime(df["Fec.Puesta"], errors="coerce")

        # 🔧 Normalizar Pos.OFA a 4 dígitos
        if "Pos.OFA" in df.columns:
            df["Pos.OFA"] = df["Pos.OFA"].astype(str).str.zfill(4)

        # 🧠 Validar columnas válidas contra la tabla
        columnas_validas = pd.read_sql("SELECT * FROM Pedidos LIMIT 1", conn).columns
        df = df[[col for col in df.columns if col in columnas_validas]]



        # 🧠 Completar campos vacíos usando registros anteriores
        try:
            df_ref = pd.read_sql("""
                SELECT DISTINCT TEMPLATE, Materia, EspesroMP, AnchoMP, CalidadMP
                FROM Pedidos
                WHERE EspesroMP IS NOT NULL AND AnchoMP IS NOT NULL AND CalidadMP IS NOT NULL
            """, conn)

            df["TEMPLATE"] = df["TEMPLATE"].astype(str)
            df["Materia"] = df["Materia"].astype(str)
            df_ref["TEMPLATE"] = df_ref["TEMPLATE"].astype(str)
            df_ref["Materia"] = df_ref["Materia"].astype(str)

            df = df.merge(df_ref, on=["TEMPLATE", "Materia"], how="left", suffixes=("", "_ref"))

            for campo in ["EspesroMP", "AnchoMP", "CalidadMP"]:
                col_ref = f"{campo}_ref"
                if col_ref in df.columns:
                    df[campo] = df[campo].combine_first(df[col_ref])
                    df.drop(columns=[col_ref], inplace=True)

            # 👀 Mostrar registros que NO pudieron completarse
            df_incompletos = df[df[["EspesroMP", "AnchoMP", "CalidadMP"]].isna().any(axis=1)]
            if not df_incompletos.empty:
                st.markdown("### ❗ Registros sin datos completos de MP:")
                columnas_mostrar = ["OFA", "Pos.OFA", "TEMPLATE", "Materia"]
                columnas_presentes = [col for col in columnas_mostrar if col in df_incompletos.columns]
                st.dataframe(df_incompletos[columnas_presentes].drop_duplicates(), use_container_width=True)
            else:
                st.success("🔄 Todos los registros tienen datos MP completos.")

        except Exception as e:
            st.warning(f"⚠️ Error completando campos MP: {e}")







        # Insertar en la base
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO HistorialCargas (fecha_hora, cantidad, archivo, usuario_pc)
            VALUES (?, ?, ?, ?)
        """, (
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            len(df),
            os.path.basename(original),
            getpass.getuser()
        ))
        id_carga = cursor.lastrowid
        conn.commit()

        df["id_carga"] = id_carga
        cantidad_insertada = insertar_sin_duplicados(df, conn)
        st.success(f"✅ {cantidad_insertada} registros nuevos cargados en la tabla 'Pedidos'.")

        os.remove(temporal)
        os.remove(ruta_xlsx)

    except Exception as e:
        st.error(f"❌ Error al cargar ListMarc.XLS: {e}")



# === Cargar Stock blanks ===
def cargar_stock_blanks(uploaded_file):

    try:
        import fractions

        ruta_original = r"C:\A_Programación\Stock blanks.XLS"
        ruta_temporal = r"C:\A_Programación\Stock_temp.XLS"
        ruta_xlsx = ruta_temporal.replace(".XLS", ".xlsx")

        shutil.copyfile(ruta_original, ruta_temporal)
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False

        libro = excel.Workbooks.Open(ruta_temporal, ReadOnly=True)
        libro.SaveAs(Filename=ruta_xlsx, FileFormat=51)
        libro.Close(SaveChanges=False)
        excel.Quit()

        df = pd.read_excel(ruta_xlsx, engine="openpyxl")
        st.write("📦 Columnas de stock detectadas:", df.columns.tolist())

        columnas_filtradas = [col for col in df.columns if not col.upper().startswith("CAMP")]
        df = df[columnas_filtradas]

        def convertir_a_coma(valor):
            try:
                valor = str(valor).strip().replace(",", ".")
                if "_" in valor:
                    entero, fraccion = valor.split("_")
                    valor = float(entero) + float(fractions.Fraction(fraccion))
                elif " " in valor and "/" in valor:
                    entero, fraccion = valor.split()
                    valor = float(entero) + float(fractions.Fraction(fraccion))
                elif "/" in valor:
                    valor = float(fractions.Fraction(valor))
                else:
                    valor = float(valor)
                return str(int(valor)) if valor.is_integer() else str(round(valor, 4)).replace(".", ",")
            except:
                return valor

        for col in ["ESP_CUB", "ANC_CUB", "LAR_CUB"]:
            if col in df.columns:
                df[col] = df[col].apply(convertir_a_coma)

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

        os.remove(ruta_temporal)
        os.remove(ruta_xlsx)

    except Exception as e:
        st.error(f"❌ Error al cargar y normalizar stock: {e}")


        # 🧠 Validar columnas válidas contra la tabla
        columnas_validas = pd.read_sql("SELECT * FROM Pedidos LIMIT 1", conn).columns
        df = df[[col for col in df.columns if col in columnas_validas]]

        # 📝 Registrar carga en historial
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO HistorialCargas (fecha_hora, cantidad, archivo, usuario_pc)
            VALUES (?, ?, ?, ?)
        """, (
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            len(df),
            os.path.basename(original),
            getpass.getuser()
        ))
        id_carga = cursor.lastrowid
        conn.commit()

        # ➕ Asignar ID de carga a cada fila
        df["id_carga"] = id_carga

        # ✅ Insertar datos sin duplicados
        cantidad_insertada = insertar_sin_duplicados(df, conn)

        st.success(f"✅ {cantidad_insertada} registros nuevos cargados en la tabla 'Pedidos'.")

        # 🧹 Limpiar archivos temporales
        os.remove(temporal)
        os.remove(ruta_xlsx)

    except Exception as e:
        st.error(f"❌ Error al cargar ListMarc.XLS: {e}")


# === BOTONES DE ACCIÓN EN FILA ===
st.markdown("### 🚀 Acciones principales")

# Usamos 4 columnas de tamaño igual, y un pequeño gap visual
cols = st.columns([1, 1, 1, 1], gap="small")


with cols[3]:  # Usa cols[4] si ya tienes 4 ocupados
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
            st.rerun()  # 🔄 Refresca la app automáticamente
        except Exception as e:
            st.error(f"❌ Error importando desde Access: {e}")





with cols[0]:
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

with cols[1]:
    if st.button("📂 Cargar ListMarc.XLS", use_container_width=True):
        try:
            cargar_manual_listmarc()
            st.toast("✅ ListMarc.XLS cargado correctamente.")
        except Exception as e:
            st.toast(f"❌ Error cargando ListMarc.XLS: {e}", icon="❌")

with cols[2]:

    # Subir archivo XLS ó XLSX
    archivo_stock = st.file_uploader(
        "📤 Subir archivo de stock (XLS/XLSX)",
        type=["xls", "xlsx"],
        key="uploader_stock"
    )

    # Botón para cargar el stock desde el archivo subido
    if st.button("📦 Cargar Stock", use_container_width=True):
        try:
            cargar_stock_blanks(archivo_stock)
            st.toast("✅ Stock cargado correctamente.")
        except Exception as e:
            st.toast(f"❌ Error cargando stock: {e}", icon="❌")


with cols[3]:
    if st.button("📊 Exportar tablas a Excel", use_container_width=True):
        try:
            import pandas as pd
            pedidos_path = os.path.join(os.getcwd(), "Pedidos_export.xlsx")
            stock_path = os.path.join(os.getcwd(), "StockBlanks_export.xlsx")

            with pd.ExcelWriter(pedidos_path, engine="openpyxl") as writer:
                pd.read_sql("SELECT * FROM Pedidos", conn).to_excel(writer, sheet_name="Pedidos", index=False)

            with pd.ExcelWriter(stock_path, engine="openpyxl") as writer:
                pd.read_sql("SELECT * FROM StockBlanks", conn).to_excel(writer, sheet_name="Stock", index=False)

            st.toast("✅ Tablas exportadas exitosamente a Excel")

        except Exception as e:
            st.toast(f"❌ Error al exportar: {e}", icon="❌")















# === FILTRO DE BÚSQUEDA ===
st.markdown("---")
st.markdown("### 🔍 Buscar OFAs en tabla de pedidos")

# Ajustamos proporciones: el input ocupa 70% del anterior (2.8 en lugar de 4)
col1, col2, _ = st.columns([2.8, 0.8, 20.4])

with col1:
    filtro_ofa = st.text_input(
        "🔍 Ingrese número OFA",
        placeholder="Ej: 5100001234",
        label_visibility="collapsed"
    )

with col2:
    eliminar_filtro = st.button("🧽", help="Limpiar filtro", use_container_width=True)









# === CONSULTA SQL ===
query_base = "SELECT rowid as id, * FROM Pedidos"

if filtro_ofa and not eliminar_filtro:
    df_tabla = pd.read_sql(
        f"{query_base} WHERE OFA LIKE ? ORDER BY id DESC",
        conn,
        params=[f"%{filtro_ofa}%"]
    )
else:
    df_tabla = pd.read_sql(f"{query_base} ORDER BY id DESC", conn)






# === FILTROS DE MP (ESPESOR, ANCHO, LARGO, CALIDAD) ===
st.markdown("---")

import fractions

# === Cargar y transformar ===
try:
    df_stock = pd.read_sql("SELECT * FROM StockBlanks", conn)

    def fraccion_a_coma(valor):
        try:
            valor = str(valor).strip()
            if "/" in valor:
                valor = float(fractions.Fraction(valor))
            else:
                valor = float(valor.replace(",", "."))
            return str(round(valor, 4)).replace(".", ",")
        except:
            return valor

    def normalizar_largo_comun(valor):
        try:
            valor = str(valor).strip().replace(",", ".")
            num = float(valor)
            return str(int(num)) if num.is_integer() else str(round(num, 3)).replace(".", ",")
        except:
            return valor

    df_stock["ESP_CUB"] = df_stock["ESP_CUB"].apply(fraccion_a_coma)
    df_stock["ANC_CUB"] = df_stock["ANC_CUB"].apply(fraccion_a_coma)
    df_stock["LAR_CUB"] = df_stock["LAR_CUB"].apply(normalizar_largo_comun)
    df_tabla["LAR_DEC"] = df_tabla["LAR_DEC"].apply(normalizar_largo_comun)

except Exception as e:
    st.warning(f"⚠️ Error al cargar o transformar columnas de StockBlanks: {e}")

    # Forzar normalización visual y tipo str de espesores
def estandarizar_espesor(valor):
    try:
        valor = float(str(valor).replace(",", "."))
        return str(int(valor)) if valor.is_integer() else str(round(valor, 4)).replace(".", ",")
    except:
        return str(valor)

df_stock["ESP_CUB"] = df_stock["ESP_CUB"].apply(estandarizar_espesor)
df_tabla["EspesroMP"] = df_tabla["EspesroMP"].apply(estandarizar_espesor)



# === Obtener valores únicos ===
espesores = sorted(set(
    df_tabla["EspesroMP"].dropna().astype(str).str.strip().tolist() +
    df_stock["ESP_CUB"].dropna().astype(str).str.strip().tolist()
))
anchos = sorted(set(
    df_tabla["AnchoMP"].dropna().astype(str).str.strip().tolist() +
    df_stock["ANC_CUB"].dropna().astype(str).str.strip().tolist()
))
largos = sorted(set(
    df_tabla["LAR_DEC"].dropna().astype(str).str.strip().tolist() +
    df_stock["LAR_CUB"].dropna().astype(str).str.strip().tolist()
))
calidades = sorted(set(
    df_tabla["CalidadMP"].dropna().astype(str).str.strip().tolist() +
    df_stock["CALIDAD"].dropna().astype(str).str.strip().tolist()
))




# === FILTROS DE MP CON MAQUINA ===
st.markdown("### 🔽 Filtros de Materia Prima")









# === BOTÓN PARA LIMPIAR FILTROS CORRECTAMENTE CON CSS CONTROLADO ===

# --- CSS para estilizar SOLO los botones, después de st.set_page_config ---
st.markdown("""
    <style>
    div.stButton > button:first-child {
        background-color: #1e3a5f;
        color: white;
        height: 2.8em;
        padding: 0 1.5em;
        border-radius: 8px;
        border: 1px solid #1e3a5f;
        font-size: 16px;
        font-weight: bold;
    }
    div.stButton > button:hover {
        background-color: #33507a;
        border: 1px solid #33507a;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

# --- Botón limpio, nativo de Streamlit ---
if st.button("🧹 Limpiar todos los filtros"):
    for key in ["filtro_maq", "filtro_espesor", "filtro_ancho", "filtro_largo", "filtro_calidad"]:
        if key in st.session_state:
            st.session_state[key] = []
    st.rerun()


# Columnas para los filtros
col_maq, col1, col2, col3, col4, _ = st.columns([0.5, 0.5, 0.5, 0.5, 0.5, 2.5])

# Inicialización de valores posibles
valores_maq = sorted(df_tabla["Maq"].dropna().astype(str).str.strip().unique())
valores_espesor = []
valores_ancho = []
valores_largo = []
valores_calidad = []

# Widget filtros vinculados a session_state
with col_maq:
    filtro_maq = st.multiselect("Máquina", valores_maq, key="filtro_maq")

# Generar df temporal según filtros activos
df_tmp = df_tabla.copy()
if filtro_maq:
    df_tmp = df_tmp[df_tmp["Maq"].astype(str).isin(filtro_maq)]

valores_espesor = sorted(df_tmp["EspesroMP"].dropna().astype(str).str.strip().unique())
if filtro_espesor := st.session_state.get("filtro_espesor", []):
    df_tmp = df_tmp[df_tmp["EspesroMP"].astype(str).isin(filtro_espesor)]

valores_ancho = sorted(df_tmp["AnchoMP"].dropna().astype(str).str.strip().unique())
if filtro_ancho := st.session_state.get("filtro_ancho", []):
    df_tmp = df_tmp[df_tmp["AnchoMP"].astype(str).isin(filtro_ancho)]

valores_largo = sorted(df_tmp["LAR_DEC"].dropna().astype(str).str.strip().unique())
if filtro_largo := st.session_state.get("filtro_largo", []):
    df_tmp = df_tmp[df_tmp["LAR_DEC"].astype(str).isin(filtro_largo)]

valores_calidad = sorted(df_tmp["CalidadMP"].dropna().astype(str).str.strip().unique())

# Filtros visuales finales
with col1:
    filtro_espesor = st.multiselect("Espesor", valores_espesor, key="filtro_espesor")
with col2:
    filtro_ancho = st.multiselect("Ancho", valores_ancho, key="filtro_ancho")
with col3:
    filtro_largo = st.multiselect("Largo", valores_largo, key="filtro_largo")
with col4:
    filtro_calidad = st.multiselect("Calidad", valores_calidad, key="filtro_calidad")







# === FILTRADO DE PEDIDOS ===

# Recuperar filtros desde session_state
filtro_maq = st.session_state.get("filtro_maq", [])
filtro_espesor = st.session_state.get("filtro_espesor", [])
filtro_ancho = st.session_state.get("filtro_ancho", [])
filtro_largo = st.session_state.get("filtro_largo", [])
filtro_calidad = st.session_state.get("filtro_calidad", [])
filtro_ofa = st.session_state.get("filtro_ofa", "")
eliminar_filtro = st.session_state.get("eliminar_filtro", False)

df_filtrado_pedidos = df_tabla.copy()

if filtro_maq:
    df_filtrado_pedidos = df_filtrado_pedidos[
        df_filtrado_pedidos["Maq"].astype(str).isin(filtro_maq)
    ]
if filtro_espesor:
    df_filtrado_pedidos = df_filtrado_pedidos[
        df_filtrado_pedidos["EspesroMP"].astype(str).isin(filtro_espesor)
    ]
if filtro_ancho:
    df_filtrado_pedidos = df_filtrado_pedidos[
        df_filtrado_pedidos["AnchoMP"].astype(str).isin(filtro_ancho)
    ]
if filtro_largo:
    df_filtrado_pedidos = df_filtrado_pedidos[
        df_filtrado_pedidos["LAR_DEC"].astype(str).isin(filtro_largo)
    ]
if filtro_calidad:
    df_filtrado_pedidos = df_filtrado_pedidos[
        df_filtrado_pedidos["CalidadMP"].astype(str).isin(filtro_calidad)
    ]
if filtro_ofa and not eliminar_filtro:
    df_filtrado_pedidos = df_filtrado_pedidos[
        df_filtrado_pedidos["OFA"].astype(str).str.contains(filtro_ofa)
    ]

# ✅ Filtrar solo filas donde Maq no sea None ni vacío
df_filtrado_pedidos = df_filtrado_pedidos[
    df_filtrado_pedidos["Maq"].notna() & (df_filtrado_pedidos["Maq"].astype(str).str.strip() != "")
]

# === Normalizar visualmente LAR_DEC de Pedidos (sin decimales ni coma)
if "LAR_DEC" in df_filtrado_pedidos.columns:
    df_filtrado_pedidos["LAR_DEC"] = (
        pd.to_numeric(df_filtrado_pedidos["LAR_DEC"].astype(str).str.replace(",", "."), errors="coerce")
        .fillna(0)
        .astype(int)
        .astype(str)
    )



# === FILTRADO DE STOCK ===
df_filtrado_stock = df_stock.copy()

# ⚠️ No aplicar filtro_maq aquí (df_stock probablemente no contiene "Maq")
if filtro_espesor:
    df_filtrado_stock = df_filtrado_stock[
        df_filtrado_stock["ESP_CUB"].astype(str).str.strip().isin(filtro_espesor)
    ]
if filtro_ancho:
    df_filtrado_stock = df_filtrado_stock[
        df_filtrado_stock["ANC_CUB"].astype(str).str.strip().isin(filtro_ancho)
    ]
if filtro_largo:
    df_filtrado_stock = df_filtrado_stock[
        df_filtrado_stock["LAR_CUB"].astype(str).str.strip().isin(filtro_largo)
    ]
if filtro_calidad:
    df_filtrado_stock = df_filtrado_stock[
        df_filtrado_stock["CALIDAD"].astype(str).str.strip().isin(filtro_calidad)
    ]


# === MOSTRAR RESULTADOS ===

# Formatear columnas de fecha
for col in ["Fec.Puesta", "FechaOFA"]:
    if col in df_filtrado_pedidos.columns:
        df_filtrado_pedidos[col] = pd.to_datetime(df_filtrado_pedidos[col], errors="coerce").dt.strftime('%d-%m-%Y')

# Reordenar columnas
orden_columnas = [
    "Maq", "Sec", "OFA", "Pos.OFA", "Materia", "Textomat", "SemanaPues",
    "Fec.Puesta", "FechaOFA", "EspesroMP", "AnchoMP", "LAR_DEC",
    "CalidadMP", "Vol#M3", "TEMPLATE"
]
columnas_validas = [col for col in orden_columnas if col in df_filtrado_pedidos.columns]
columnas_sobrantes = [col for col in df_filtrado_pedidos.columns if col not in columnas_validas]
df_filtrado_pedidos = df_filtrado_pedidos[columnas_validas + columnas_sobrantes]

# Ordenar por Sec ascendente si existe
if "Sec" in df_filtrado_pedidos.columns:
    df_filtrado_pedidos = df_filtrado_pedidos.sort_values(by="Sec", ascending=True)



# === Resumen de Volúmenes — Pedidos vs Stock
st.markdown("---")
st.markdown("### 📉 Resumen de Volúmenes — Pedidos vs Stock")

# Agrupar pedidos
group_pedidos = df_filtrado_pedidos.groupby(
    ["EspesroMP", "AnchoMP", "LAR_DEC", "CalidadMP"], dropna=False
).agg(Vol_Pedido=("Vol#M3", "sum")).reset_index()

# Agrupar stock
group_stock = df_filtrado_stock.groupby(
    ["ESP_CUB", "ANC_CUB", "LAR_CUB", "CALIDAD"], dropna=False
).agg(Vol_Stock=("Vol_Util", "sum")).reset_index()

# Renombrar columnas stock para hacer merge
group_stock.rename(columns={
    "ESP_CUB": "EspesroMP",
    "ANC_CUB": "AnchoMP",
    "LAR_CUB": "LAR_DEC",
    "CALIDAD": "CalidadMP"
}, inplace=True)

# Merge de Pedidos y Stock
resumen_volumenes = pd.merge(
    group_pedidos,
    group_stock,
    on=["EspesroMP", "AnchoMP", "LAR_DEC", "CalidadMP"],
    how="outer"
)

# Rellenar nulos con cero
resumen_volumenes["Vol_Pedido"] = resumen_volumenes["Vol_Pedido"].fillna(0)
resumen_volumenes["Vol_Stock"] = resumen_volumenes["Vol_Stock"].fillna(0)

# Asegurar que los valores sean numéricos
resumen_volumenes["Vol_Pedido"] = pd.to_numeric(resumen_volumenes["Vol_Pedido"], errors="coerce").fillna(0)
resumen_volumenes["Vol_Stock"] = pd.to_numeric(resumen_volumenes["Vol_Stock"], errors="coerce").fillna(0)

# Calcular diferencia
resumen_volumenes["Diferencia"] = resumen_volumenes["Vol_Stock"] - resumen_volumenes["Vol_Pedido"]


# Mostrar resultado
st.dataframe(resumen_volumenes, use_container_width=True)




# === Tabla editable de Pedidos filtrados ===
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from st_aggrid.shared import JsCode

st.markdown("---")
st.markdown("### 🖊️ Editar tabla de Pedidos filtrados")

if not df_filtrado_pedidos.empty:
    # Campos que podrán ser editados
    campos_editables = ["Maq", "Sec", "EspesroMP", "AnchoMP", "LAR_DEC", "CalidadMP"]

    # Configuración del grid
    gb = GridOptionsBuilder.from_dataframe(df_filtrado_pedidos)
    gb.configure_default_column(editable=False, resizable=True, filter=True)

    for campo in campos_editables:
        if campo in df_filtrado_pedidos.columns:
            gb.configure_column(campo, editable=True)

    # JS para capturar Enter y pasar a siguiente fila
    enter_js = JsCode("""
    function(params) {
        if (params.event.key === 'Enter') {
            params.api.stopEditing();
            const nextRowIndex = params.rowIndex + 1;
            const columnId = params.column.getId();
            if (nextRowIndex < params.api.getDisplayedRowCount()) {
                setTimeout(function() {
                    params.api.startEditingCell({
                        rowIndex: nextRowIndex,
                        colKey: columnId
                    });
                }, 200);
            }
        }
    }
    """)
    gb.configure_grid_options(onCellKeyDown=enter_js)

    grid_options = gb.build()

    # Botón para guardar
    col_edit, _ = st.columns([1, 6])
    with col_edit:
        actualizar = st.button("💾 Guardar cambios")

    # Render editable
    grid_response = AgGrid(
        df_filtrado_pedidos,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.VALUE_CHANGED,
        editable=True,
        height=400,
        theme="balham-dark",
        allow_unsafe_jscode=True,
        fit_columns_on_grid_load=False,
        autoSizeColumns=True
    )

    # Guardar cambios si se presiona el botón
    if actualizar:
        df_actualizado = grid_response["data"]
        for _, row in df_actualizado.iterrows():
            conn.execute("""
                UPDATE Pedidos SET 
                    Maq = ?, Sec = ?, EspesroMP = ?, 
                    AnchoMP = ?, LAR_DEC = ?, CalidadMP = ?
                WHERE rowid = ?
            """, (
                row.get("Maq"),
                row.get("Sec"),
                row.get("EspesroMP"),
                row.get("AnchoMP"),
                row.get("LAR_DEC"),
                row.get("CalidadMP"),
                row["id"]
            ))
        conn.commit()
        st.success("✅ Cambios guardados correctamente.")








# Reordenar columnas para df_filtrado_stock
prioridad_stock = ['ESP_CUB', 'ANC_CUB', 'LAR_CUB', 'CALIDAD', 'M3']
columnas_ordenadas_stock = (
    [col for col in prioridad_stock if col in df_filtrado_stock.columns] +
    [col for col in df_filtrado_stock.columns if col not in prioridad_stock]
)
df_filtrado_stock = df_filtrado_stock[columnas_ordenadas_stock]

st.markdown(
    "<h4 style='background-color:#fef9e7;padding:10px;border-radius:8px;color:#b9770e;'>"
    "📦 Stock filtrado</h4>", 
    unsafe_allow_html=True
)
st.dataframe(df_filtrado_stock, use_container_width=True)








# === DETALLE POR CARGA SELECCIONADA ===
if "ver_carga" in st.session_state:
    id_carga_sel = st.session_state["ver_carga"]
    st.markdown("---")
    st.markdown(f"### 📝 Detalle de OFAs — Carga #{id_carga_sel}")
    try:
        df_detalle = pd.read_sql("SELECT * FROM Pedidos WHERE id_carga = ?", conn, params=[id_carga_sel])
        if df_detalle.empty:
            st.warning("⚠️ No se encontraron OFAs asociadas a esta carga.")
        else:
            # 🎯 Filtrar columnas visibles
            columnas_detalle = [
                "OFA", "Pos.OFA", "SemanaPues", "Fec#Puesta", "Materia",
                "TEMPLATE", "EspesorMP", "AnchoMP", "CalidadMP", "Units", "Vol#M3"
            ]
            df_detalle = df_detalle[[col for col in columnas_detalle if col in df_detalle.columns]]

            # 🧼 Renombrar para presentación clara
            df_detalle.rename(columns={
                "Pos.OFA": "Posición",
                "SemanaPues": "Semana Puesta",
                "Fec#Puesta": "Fecha Puesta",
                "EspesorMP": "Espesor MP",
                "AnchoMP": "Ancho MP",
                "CalidadMP": "Calidad MP",
                "Vol#M3": "Volumen m³"
            }, inplace=True)

            # 🔧 Limpiar ".0" en Semana Puesta si es float
            if "Semana Puesta" in df_detalle.columns:
                df_detalle["Semana Puesta"] = (
                    df_detalle["Semana Puesta"].astype(str).str.replace(".0", "", regex=False)
                )

            # 📊 Mostrar tabla final
            st.dataframe(df_detalle.astype(str), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error al mostrar detalle: {e}")




# === Tabla editable de Pedidos filtrados ===
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode

st.markdown("---")
st.markdown("### 🖊️ Editar tabla 'Pedidos' — Guardado automático + salto con Enter")

# === Cargar datos
df_edit = pd.read_sql("""
    SELECT rowid as id, Maq, Sec, Materia, Textomat, EspesroMP, AnchoMP, LAR_DEC, CalidadMP, [Vol#M3]
    FROM Pedidos ORDER BY id DESC
""", conn)

# === Campos editables
campos_editables = ["Maq", "Sec", "Materia", "Textomat", "EspesroMP", "AnchoMP", "LAR_DEC", "CalidadMP", "Vol#M3"]

# === Configuración del grid
gb = GridOptionsBuilder.from_dataframe(df_edit)
gb.configure_default_column(editable=False, resizable=True, filter=True)

for campo in campos_editables:
    gb.configure_column(campo, editable=True)

# JS para capturar Enter y saltar a fila siguiente
enter_js = JsCode("""
function(params) {
    if (params.event.key === 'Enter') {
        params.api.stopEditing();
        const nextRow = params.rowIndex + 1;
        if (nextRow < params.api.getDisplayedRowCount()) {
            setTimeout(() => {
                params.api.startEditingCell({
                    rowIndex: nextRow,
                    colKey: params.column.getId()
                });
            }, 100);
        }
    }
}
""")
gb.configure_grid_options(onCellKeyDown=enter_js)

# Activar auto guardado y salto de fila
grid_options = gb.build()

# Mostrar tabla editable
grid_response = AgGrid(
    df_edit,
    gridOptions=grid_options,
    update_mode=GridUpdateMode.VALUE_CHANGED,
    editable=True,
    height=400,
    theme="balham-dark",
    allow_unsafe_jscode=True,
    fit_columns_on_grid_load=False,
    autoSizeColumns=True
)

# === Guardar cambios automáticamente
if grid_response["data"] is not None:
    df_actualizado = grid_response["data"]
    for _, row in df_actualizado.iterrows():
        conn.execute("""
            UPDATE Pedidos SET
                Maq = ?, Sec = ?, Materia = ?, Textomat = ?, EspesroMP = ?,
                AnchoMP = ?, LAR_DEC = ?, CalidadMP = ?, [Vol#M3] = ?
            WHERE rowid = ?
        """, (
            row["Maq"], row["Sec"], row["Materia"], row["Textomat"], row["EspesroMP"],
            row["AnchoMP"], row["LAR_DEC"], row["CalidadMP"], row["Vol#M3"],
            row["id"]
        ))
    conn.commit()
    st.toast("✅ Cambios guardados automáticamente", icon="💾")






# === HISTORIAL DE CARGAS EN EXPANDER ===
st.markdown("---")
with st.expander("📜 Historial de Cargas recientes (últimas 10)"):
    try:
        df_log = pd.read_sql("SELECT * FROM HistorialCargas ORDER BY id DESC LIMIT 10", conn)
        for index, row in df_log.iterrows():
            col1, col2 = st.columns([5, 1])
            with col1:
                st.markdown(
                    f"📁 **Carga #{row['id']}** — 🗓 {row['fecha_hora']} — 💾 Archivo: {row['archivo']} — "
                    f"👤 Usuario: {row['usuario_pc']} — 📊 Cantidad: {row['cantidad']}"
                )
            with col2:
                if st.button("🔍 Ver OFAs", key=f"ver_ofas_{row['id']}"):
                    st.session_state["ver_carga"] = row["id"]
    except Exception as e:
        st.error(f"❌ Error cargando historial: {e}")