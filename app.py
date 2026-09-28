import streamlit as st
import pandas as pd
import re
import os
import base64
import unicodedata
from datetime import datetime, timedelta

st.set_page_config(page_title="Precios en la calle", layout="wide")

# ============================================================
# CONFIGURACIÓN
# ============================================================
CARPETA_APP = os.path.dirname(os.path.abspath(__file__))

ARCHIVO_MES_ACTUAL = os.path.join(CARPETA_APP, "mes actual.xlsx")
ARCHIVO_MES_ANTERIOR = os.path.join(CARPETA_APP, "mes anterior.xlsx")

HOJA = "Hoja2"
FILA_ENCABEZADOS = 3  # fila 4 del Excel = índice 3 (0-indexado)

COL_FECHA = "Fecha Comprobante"
COL_CLIENTE = "Cliente"
COL_RAZON = "Razon Social"
COL_DIVISION = "Descripcion DIVISION"
COL_CODIGO = "Codigo de Articulo"
COL_DESC = "Descripcion de Articulo"
COL_UXB = "Unidades por Bulto"
COL_DESC_PCT = "% desc"
COL_BTOS = "Btos"
COL_PTR = "Suma de PTR"
COL_PTR_UNIT = "PTR Unit"
COL_PTC = "PTC"
COL_MES = "mes"

COLUMNAS_ESPERADAS = [
    COL_FECHA, COL_CLIENTE, COL_RAZON, COL_DIVISION, COL_CODIGO,
    COL_DESC, COL_UXB, COL_DESC_PCT, COL_BTOS, COL_PTR,
    COL_PTR_UNIT, COL_PTC, COL_MES,
]

# ============================================================
# LOGO EN ESQUINA SUPERIOR DERECHA
# ============================================================
def agregar_logo_esquina(ruta_imagen, ancho=180):
    with open(ruta_imagen, "rb") as f:
        datos = base64.b64encode(f.read()).decode()

    st.markdown(
        f"""
        <style>
        .logo-esquina {{
            position: fixed;
            top: 12px;
            right: 25px;
            z-index: 9999;
        }}
        .logo-esquina img {{
            width: {ancho}px;
        }}
        </style>
        <div class="logo-esquina">
            <img src="data:image/png;base64,{datos}">
        </div>
        """,
        unsafe_allow_html=True,
    )


ruta_logo = os.path.join(CARPETA_APP, "logo.png")
if os.path.exists(ruta_logo):
    agregar_logo_esquina(ruta_logo)
else:
    st.warning("⚠️ No se encontró el archivo 'logo.png' en la carpeta de la app.")


# ============================================================
# UTILIDADES
# ============================================================
def normalizar_texto(s):
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"\s+", " ", s)
    return s


def emparejar_columnas(df, columnas_esperadas):
    mapa_normalizado = {normalizar_texto(c): c for c in df.columns}
    renombres = {}
    faltantes = []

    for esperada in columnas_esperadas:
        clave = normalizar_texto(esperada)
        if clave in mapa_normalizado:
            renombres[mapa_normalizado[clave]] = esperada
        else:
            faltantes.append(esperada)

    df = df.rename(columns=renombres)
    return df, faltantes


def limpiar_encabezados(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df


def extraer_mes(valor):
    if valor is None:
        return None
    valor = str(valor).strip()
    if valor.startswith("{"):
        match = re.search(r'"result"\s*:\s*(-?\d+\.?\d*)', valor)
        return int(float(match.group(1))) if match else None
    try:
        return int(float(valor))
    except ValueError:
        return None


def serial_a_fecha(valor):
    """Convierte un valor de Excel (serial o texto) a un objeto date,
    interpretando el texto en formato día-mes-año (D-M-AAAA)."""
    if valor is None or str(valor).strip() == "" or str(valor).strip().lower() == "nan":
        return None

    valor_str = str(valor).strip()

    # Caso: número serial de Excel
    try:
        num = float(valor_str)
        return (datetime(1899, 12, 30) + timedelta(days=num)).date()
    except ValueError:
        pass

    # Caso: texto con fecha (día-mes-año)
    try:
        fecha = pd.to_datetime(valor_str, dayfirst=True, errors="coerce")
        return fecha.date() if pd.notna(fecha) else None
    except Exception:
        return None


def formato_fecha_ar(valor):
    """Formatea una fecha (date) como texto DD/MM/AAAA."""
    if valor is None or pd.isna(valor):
        return ""
    return valor.strftime("%d/%m/%Y")


def limpiar_numero(valor):
    if valor is None or pd.isna(valor):
        return None
    s = str(valor).strip()
    if s == "" or s.lower() == "nan":
        return None

    s = re.sub(r"[^\d.,\-]", "", s)

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")

    return s


def formato_pesos_ar(valor):
    """Formatea un número como moneda argentina: $ 1.234,5"""
    if pd.isna(valor):
        return ""
    texto = f"{valor:,.1f}"  # formato US: 1,234.5
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"$ {texto}"


def formato_porcentaje(valor):
    """% con 2 decimales. El valor ya viene como número plano (ej: 15.5)."""
    if pd.isna(valor):
        return ""
    return f"{valor:.2f}%"


def formato_bultos(valor):
    """Bultos con 2 decimales."""
    if pd.isna(valor):
        return ""
    return f"{valor:.2f}"


def formato_mes(valor):
    """Mes sin decimales."""
    if pd.isna(valor):
        return ""
    return f"{valor:.0f}"


# ============================================================
# CARGA DE DATOS
# ============================================================
@st.cache_data
def cargar_datos():
    archivos = [ARCHIVO_MES_ACTUAL, ARCHIVO_MES_ANTERIOR]
    lista_df = []
    archivos_encontrados = []

    for archivo in archivos:
        if not os.path.exists(archivo):
            st.warning(f"⚠️ No se encontró el archivo: {os.path.basename(archivo)}")
            continue
        try:
            df_temp = pd.read_excel(
                archivo, sheet_name=HOJA, dtype=str, header=FILA_ENCABEZADOS
            )
            df_temp = limpiar_encabezados(df_temp)
            df_temp["__archivo_origen"] = os.path.basename(archivo)
            lista_df.append(df_temp)
            archivos_encontrados.append(archivo)
        except ValueError as e:
            st.warning(f"⚠️ No se pudo leer '{HOJA}' en {os.path.basename(archivo)}: {e}")

    if not lista_df:
        st.error(
            "⚠️ No se pudo cargar ningún archivo. Verificá que existan "
            "'mes actual.xlsx' y 'mes anterior.xlsx' en la carpeta de la app."
        )
        st.stop()

    df = pd.concat(lista_df, ignore_index=True)

    df, faltantes = emparejar_columnas(df, COLUMNAS_ESPERADAS)

    if faltantes:
        st.error(f"⚠️ Faltan columnas en el archivo: {faltantes}")
        st.write("Columnas encontradas en el Excel:", list(df.columns))
        st.stop()

    df[COL_MES] = df[COL_MES].apply(extraer_mes)
    df[COL_FECHA] = df[COL_FECHA].apply(serial_a_fecha)

    columnas_clave = [COL_MES, COL_FECHA, COL_CLIENTE, COL_CODIGO, COL_PTR]
    df = df.drop_duplicates(subset=columnas_clave, keep="last")

    dicc_div = (
        df.dropna(subset=[COL_CODIGO, COL_DIVISION])
        .groupby(COL_CODIGO)[COL_DIVISION]
        .agg(lambda x: x.mode().iloc[0] if not x.mode().empty else None)
        .to_dict()
    )
    df[COL_DIVISION] = df[COL_DIVISION].fillna(df[COL_CODIGO].map(dicc_div))

    dicc_cod = (
        df.dropna(subset=[COL_DESC, COL_CODIGO])
        .groupby(COL_DESC)[COL_CODIGO]
        .agg(lambda x: x.mode().iloc[0] if not x.mode().empty else None)
        .to_dict()
    )
    df[COL_CODIGO] = df[COL_CODIGO].fillna(df[COL_DESC].map(dicc_cod))

    columnas_numericas = [COL_UXB, COL_DESC_PCT, COL_BTOS, COL_PTR, COL_PTR_UNIT, COL_PTC]
    for col in columnas_numericas:
        df[col] = df[col].apply(limpiar_numero)
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df, archivos_encontrados


df, archivos_cargados = cargar_datos()

# ============================================================
# AVISOS DE CALIDAD DE DATOS
# ============================================================
with st.sidebar.expander("🔎 Calidad de datos", expanded=False):
    nulos_fecha = df[COL_FECHA].isna().sum()
    if nulos_fecha > 0:
        st.warning(f"{nulos_fecha} filas sin fecha válida")

    for col in [COL_PTR, COL_PTC]:
        nulos = df[col].isna().sum()
        if nulos > 0:
            st.warning(f"{nulos} valores no numéricos en '{col}'")

    if nulos_fecha == 0 and all(df[c].isna().sum() == 0 for c in [COL_PTR, COL_PTC]):
        st.success("Sin inconsistencias detectadas")

# ============================================================
# SIDEBAR - FILTROS (en cascada: Mes → Cliente → División → Producto)
# ============================================================
st.sidebar.title("🔍 Filtros")

with st.sidebar.expander("📁 Archivos cargados"):
    for a in archivos_cargados:
        st.caption(os.path.basename(a))

# --- 1. Mes ---
meses_disponibles = sorted(df[COL_MES].dropna().unique())
mes_sel = st.sidebar.multiselect("Mes", meses_disponibles, default=meses_disponibles)

df_base_mes = df[df[COL_MES].isin(mes_sel)]

# --- 2. Cliente ---
clientes_disponibles = sorted(df_base_mes[COL_RAZON].dropna().unique())
cliente_sel = st.sidebar.multiselect("Cliente", clientes_disponibles)

df_base_cliente = df_base_mes
if cliente_sel:
    df_base_cliente = df_base_cliente[df_base_cliente[COL_RAZON].isin(cliente_sel)]

if not cliente_sel:
    st.sidebar.caption("💡 Seleccioná un cliente para acotar división y producto")

# --- 3. División (según cliente seleccionado) ---
divisiones_disponibles = sorted(df_base_cliente[COL_DIVISION].dropna().unique())
division_sel = st.sidebar.multiselect(
    "División", divisiones_disponibles, default=divisiones_disponibles
)

df_base_division = df_base_cliente[df_base_cliente[COL_DIVISION].isin(division_sel)]

# --- 4. Producto (según cliente + división seleccionados) ---
productos_disponibles = sorted(df_base_division[COL_DESC].dropna().unique())
producto_sel = st.sidebar.multiselect(
    "Producto (según cliente y división seleccionados)", productos_disponibles
)

# --- Filtro final ---
df_filtrado = df_base_division
if producto_sel:
    df_filtrado = df_filtrado[df_filtrado[COL_DESC].isin(producto_sel)]

# ============================================================
# CONTENIDO PRINCIPAL
# ============================================================
st.title("📊 Precios en la Calle")

# ============================================================
# CUADRO COMPARATIVO: División (filas) x Mes (columnas) - Suma de Btos
# ============================================================
st.subheader("📊 Comparativo de Bultos por División y Mes")

if df_filtrado.empty:
    st.info("No hay datos para los filtros seleccionados.")
else:
    tabla_comparativa = pd.pivot_table(
        df_filtrado,
        index=COL_DIVISION,
        columns=COL_MES,
        values=COL_BTOS,
        aggfunc="sum",
        fill_value=0,
    )

    tabla_comparativa = tabla_comparativa.reindex(
        sorted(tabla_comparativa.columns), axis=1
    )

    tabla_comparativa["Total"] = tabla_comparativa.sum(axis=1)
    tabla_comparativa.loc["Total"] = tabla_comparativa.sum(axis=0)

    # column_config para minimizar el ancho de cada columna numérica
    config_pivot = {
        str(col): st.column_config.Column(width="small")
        for col in tabla_comparativa.columns
    }

    st.dataframe(
        tabla_comparativa.style.format("{:,.2f}"),
        column_config=config_pivot,
    )

st.divider()

# ============================================================
# TABLA DE DETALLE
# ============================================================
st.subheader("📋 Detalle de registros")

df_ordenado = df_filtrado.sort_values(by=COL_FECHA, ascending=False)

# Columnas visibles en la tabla
# (se ocultan "Descripcion DIVISION", "Suma de PTR" y "Unidades por Bulto";
#  el índice también se oculta al mostrar la tabla)
columnas_visibles = [
    COL_MES, COL_FECHA, COL_CLIENTE, COL_RAZON,
    COL_CODIGO, COL_DESC, COL_DESC_PCT, COL_BTOS,
    COL_PTR_UNIT, COL_PTC,
]

df_vista = df_ordenado[columnas_visibles].copy()

# --- Formateamos como texto directamente (sin Styler) ---
# Esto evita el límite de celdas de Pandas Styler y es más rápido
# en tablas grandes.
df_mostrar = df_vista.copy()
df_mostrar[COL_MES] = df_mostrar[COL_MES].apply(formato_mes)
df_mostrar[COL_FECHA] = df_mostrar[COL_FECHA].apply(formato_fecha_ar)
df_mostrar[COL_DESC_PCT] = df_mostrar[COL_DESC_PCT].apply(formato_porcentaje)
df_mostrar[COL_BTOS] = df_mostrar[COL_BTOS].apply(formato_bultos)
df_mostrar[COL_PTR_UNIT] = df_mostrar[COL_PTR_UNIT].apply(formato_pesos_ar)
df_mostrar[COL_PTC] = df_mostrar[COL_PTC].apply(formato_pesos_ar)

# column_config para minimizar el ancho de cada columna
config_detalle = {
    col: st.column_config.Column(width="small")
    for col in df_mostrar.columns
}

st.dataframe(
    df_mostrar,
    height=500,
    column_config=config_detalle,
    hide_index=True,
)

# Descarga (mantiene todas las columnas originales, incluidas "Suma de PTR" y
# "Unidades por Bulto", sin formato de moneda ni ocultamiento)
csv = df_ordenado.to_csv(index=False).encode("utf-8")
st.download_button(
    "⬇️ Descargar tabla filtrada (CSV)",
    csv,
    "datos_filtrados.csv",
    "text/csv",
)