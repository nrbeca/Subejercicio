"""
app.py — Generador de Cuadros de Subejercicio (DGAPSGB)
==========================================================
Streamlit app: sube el MAP, ajusta los filtros, revisa el resumen general
y descarga el Excel con un cuadro por cada UR+Pp.
"""

import datetime

import pandas as pd
import streamlit as st

import cuadros_lib as lib

st.set_page_config(page_title="Cuadros de Subejercicio", layout="wide")

BURGUNDY = "#9F2241"
CAFE = "#BC955C"
CREAM = "#F6F2EB"

st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {CREAM}; }}
    h1, h2, h3 {{ color: {BURGUNDY}; }}
    .stButton>button, .stDownloadButton>button {{
        background-color: {BURGUNDY}; color: white; border: none;
    }}
    .stButton>button:hover, .stDownloadButton>button:hover {{
        background-color: {CAFE}; color: white;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Generador de Cuadros de Subejercicio — DGAPSGB")
st.caption("Sube el MAP crudo (CSV o XLSX), ajusta los filtros y descarga el Excel con un cuadro por UR+Pp.")

# ── Sidebar: filtros ─────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Archivo")
    archivo = st.file_uploader("MAP (CSV o XLSX)", type=["csv", "xlsx"])

    st.header("Filtros de selección")
    disp_minimo = st.number_input(
        "MDP mínimo de disponible por capítulo",
        min_value=0.0, value=1.0, step=0.1,
        help="Un capítulo se incluye solo si su disponible (Modificado − Ejercido) es mayor o igual a este monto, en millones de pesos."
    )
    capitulo_max = st.number_input(
        "Capítulo máximo a considerar",
        min_value=1000, max_value=9000, value=4999, step=1,
        help="Los capítulos por arriba de este valor se excluyen (por default 4999, para dejar fuera Inversión 5000/6000/7000, que se reporta aparte)."
    )

    st.header("Lista manual (opcional)")
    st.caption("Si la llenas, solo se generan estas combinaciones UR,Pp. Un renglón por combinación, formato UR,Pp — ej. 923,S318")
    lista_manual_txt = st.text_area("UR,Pp por renglón", value="", height=100, label_visibility="collapsed")

    st.header("Periodo")
    modo_fecha = st.radio("Fecha de corte", ["Detectar del nombre del archivo", "Especificar manualmente"], index=0)
    corte_manual = None
    if modo_fecha == "Especificar manualmente":
        corte_manual = st.date_input("Último día del periodo", value=datetime.date.today())

def _parse_lista_manual(txt):
    combos = []
    for line in txt.strip().splitlines():
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in line.split(",")]
        if len(parts) != 2:
            continue
        ur_raw, pp = parts
        ur = int(ur_raw) if ur_raw.lstrip("-").isdigit() else ur_raw
        combos.append((ur, pp))
    return combos

lista_manual = _parse_lista_manual(lista_manual_txt)

if archivo is None:
    st.info("Sube el archivo MAP para comenzar.")
    st.stop()

# ── Determinar corte ──────────────────────────────────────────────────────────
if modo_fecha == "Detectar del nombre del archivo":
    corte = lib.parse_fecha(archivo.name)
    if corte is None:
        st.warning("No se pudo detectar la fecha de corte del nombre del archivo. Usa 'Especificar manualmente' en la barra lateral.")
        st.stop()
else:
    corte = corte_manual

cm, anio, periodo_label, fecha_corte_str, fecha_rep_str = lib.periodo_info(corte)

st.markdown(f"**Corte:** {fecha_corte_str}  |  **Periodo:** {periodo_label}  |  **Fecha de reporte:** {fecha_rep_str}")

# ── Procesar ──────────────────────────────────────────────────────────────────
with st.spinner("Leyendo y calculando…"):
    try:
        df = lib.leer_y_calcular(archivo, archivo.name, cm)
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}")
        st.stop()
    grp = lib.agregar(df)
    caps_sel = lib.seleccionar(grp, disp_minimo, capitulo_max, lista_manual)

if caps_sel.empty:
    st.warning("No hay combinaciones que cumplan los filtros actuales.")
    st.stop()

resumen = lib.resumen_general(caps_sel)

# ── Resumen general ────────────────────────────────────────────────────────────
st.subheader("Resumen general")

n_combos = len(resumen)
total_mod = resumen["MOD_MDP"].sum()
total_disp = resumen["DISP_MDP"].sum()

c1, c2, c3 = st.columns(3)
c1.metric("Combinaciones UR + Pp", n_combos)
c2.metric("Modificado total (MDP)", f"{total_mod:,.1f}")
c3.metric("Disponible total (MDP)", f"{total_disp:,.1f}")

resumen_fmt = resumen.copy()
resumen_fmt["MOD_MDP"] = resumen_fmt["MOD_MDP"].round(1)
resumen_fmt["EJE_MDP"] = resumen_fmt["EJE_MDP"].round(1)
resumen_fmt["DISP_MDP"] = resumen_fmt["DISP_MDP"].round(1)
resumen_fmt["PCT"] = resumen_fmt["PCT"].round(1)
resumen_fmt = resumen_fmt.rename(columns={
    "UR": "UR", "UR_Nombre": "Unidad Responsable", "PP": "Pp", "PP_Nombre": "Programa presupuestario",
    "MOD_MDP": "Modificado (MDP)", "EJE_MDP": "Ejercido (MDP)", "DISP_MDP": "Disponible (MDP)", "PCT": "% disp/mod",
})

st.dataframe(
    resumen_fmt,
    use_container_width=True,
    hide_index=True,
    column_config={
        "% disp/mod": st.column_config.NumberColumn(format="%.1f%%"),
        "Modificado (MDP)": st.column_config.NumberColumn(format="%.1f"),
        "Ejercido (MDP)": st.column_config.NumberColumn(format="%.1f"),
        "Disponible (MDP)": st.column_config.NumberColumn(format="%.1f"),
    },
)

# Detalle por capítulo, opcional
with st.expander("Ver detalle por capítulo"):
    detalle = caps_sel.copy()
    detalle["UR_Nombre"] = detalle["UR2"].apply(lambda u: lib.UR_NOMBRES.get(u, str(u)))
    detalle["PP_Nombre"] = detalle["PP"].apply(lambda p: lib.PP_NOMBRES.get(p, p))
    detalle["CAP_Nombre"] = detalle["CAPITULO"].apply(lambda c: lib.CAP_NOMBRES.get(c, str(c)))
    detalle["PCT"] = detalle.apply(lambda r: (r["DISP"] / r["MOD"] * 100) if r["MOD"] else 0, axis=1)
    detalle = detalle[["UR2", "UR_Nombre", "PP", "PP_Nombre", "CAPITULO", "CAP_Nombre", "MOD", "EJE", "DISP", "PCT"]]
    detalle = detalle.sort_values(["PP", "UR2", "CAPITULO"]).round(1)
    st.dataframe(detalle, use_container_width=True, hide_index=True)

# ── Descarga ──────────────────────────────────────────────────────────────────
st.subheader("Descargar")

with st.spinner("Generando el Excel…"):
    xlsx_bytes, n_hojas = lib.generar_workbook(caps_sel, anio, periodo_label, fecha_corte_str, fecha_rep_str)

periodo_safe = periodo_label.replace(" ", "_").replace("-", "_")
nombre_salida = f"Cuadros_Subejercicio_{periodo_safe}_{anio}.xlsx"

st.download_button(
    label=f"Descargar Excel ({n_hojas} hojas)",
    data=xlsx_bytes,
    file_name=nombre_salida,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
