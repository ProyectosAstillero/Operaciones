import os

import pandas as pd
import streamlit as st


BASE_DIR = os.path.dirname(os.path.dirname(__file__))
SHEET_NAME = "DATA"
CURRENCY_SYMBOL = "$"

st.set_page_config(page_title="Astillero-Operaciones | Presupuesto", layout="wide")


@st.cache_data(show_spinner=False)
def _read_excel_sheet(excel_path: str, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(excel_path, sheet_name=sheet_name)


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).replace("\r", " ").replace("\n", " ").strip() for c in df.columns]
    return df


def _find_col(cols: list[str], candidates: list[str]) -> str | None:
    cols_norm = {c.lower().strip(): c for c in cols}
    for cand in candidates:
        key = cand.lower().strip()
        if key in cols_norm:
            return cols_norm[key]
    for cand in candidates:
        key = cand.lower().strip()
        for c in cols:
            if key in c.lower():
                return c
    return None


def _monthly_summary(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = _normalize_columns(df)

    month_names_es = {
        1: "Enero",
        2: "Febrero",
        3: "Marzo",
        4: "Abril",
        5: "Mayo",
        6: "Junio",
        7: "Julio",
        8: "Agosto",
        9: "Septiembre",
        10: "Octubre",
        11: "Noviembre",
        12: "Diciembre",
    }

    mes_col = _find_col(df.columns.tolist(), ["Mes"])
    real_col = _find_col(df.columns.tolist(), ["Cst.reales", "Cst.reales ", "Cst reales", "Cst. reales"])
    plan_col = _find_col(df.columns.tolist(), ["Cst.plan", "Cst.plan ", "Cst plan", "Cst. plan"])

    if mes_col is None:
        raise ValueError("No se encontró la columna 'Mes'.")
    if real_col is None:
        raise ValueError("No se encontró la columna de costo real (ej: 'Cst.reales').")
    if plan_col is None:
        raise ValueError("No se encontró la columna de costo planificado (ej: 'Cst.plan').")

    tmp = df[[mes_col, real_col, plan_col]].copy()
    tmp["_mes"] = pd.to_datetime(tmp[mes_col], errors="coerce", dayfirst=True)
    tmp["_real"] = pd.to_numeric(tmp[real_col], errors="coerce")
    tmp["_plan"] = pd.to_numeric(tmp[plan_col], errors="coerce")
    tmp = tmp.dropna(subset=["_mes"])
    tmp["_mes"] = tmp["_mes"].dt.to_period("M").dt.to_timestamp("M")

    mens = (
        tmp.groupby("_mes", as_index=False)[["_plan", "_real"]]
        .sum()
        .rename(columns={"_mes": "Mes_dt", "_plan": f"Plan ({CURRENCY_SYMBOL})", "_real": f"Real ({CURRENCY_SYMBOL})"})
        .sort_values("Mes_dt")
    )

    mens["Mes"] = mens["Mes_dt"].dt.month.map(month_names_es) + " " + mens["Mes_dt"].dt.year.astype(str)

    acum = mens.copy()
    acum[f"Plan Acum ({CURRENCY_SYMBOL})"] = acum[f"Plan ({CURRENCY_SYMBOL})"].cumsum()
    acum[f"Real Acum ({CURRENCY_SYMBOL})"] = acum[f"Real ({CURRENCY_SYMBOL})"].cumsum()
    return mens, acum


@st.cache_data(show_spinner=False)
def _load_monthly_for_file(excel_filename: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    excel_path = os.path.join(BASE_DIR, excel_filename)
    df = _read_excel_sheet(excel_path, SHEET_NAME)
    return _monthly_summary(df)


def _render_excel_tab(title: str, excel_filename: str, d1, d2, mode: str) -> None:
    excel_path = os.path.join(BASE_DIR, excel_filename)

    st.subheader(title)

    if not os.path.exists(excel_path):
        st.error(f"No se encontró el archivo: {excel_path}")
        return

    try:
        mensual, acumulado = _load_monthly_for_file(excel_filename)
    except Exception as e:
        st.error(f"Error cargando la hoja '{SHEET_NAME}' o calculando resumen: {e}")
        return

    if mensual.empty:
        st.warning("No hay datos válidos en la hoja DATA para construir el resumen.")
        return

    mensual_f = mensual[(mensual["Mes_dt"].dt.date >= d1) & (mensual["Mes_dt"].dt.date <= d2)].copy()
    acumulado_f = acumulado[(acumulado["Mes_dt"].dt.date >= d1) & (acumulado["Mes_dt"].dt.date <= d2)].copy()

    plan_col = f"Plan ({CURRENCY_SYMBOL})"
    real_col = f"Real ({CURRENCY_SYMBOL})"
    plan_ac_col = f"Plan Acum ({CURRENCY_SYMBOL})"
    real_ac_col = f"Real Acum ({CURRENCY_SYMBOL})"

    if mode == "Mensual":
        st.dataframe(mensual_f[["Mes", plan_col, real_col]], width="stretch", hide_index=True)
        chart_df = mensual_f.set_index("Mes_dt")[[plan_col, real_col]]
        st.line_chart(chart_df)
    else:
        show = acumulado_f[["Mes", plan_ac_col, real_ac_col]].copy()
        st.dataframe(show, width="stretch", hide_index=True)
        chart_df = acumulado_f.set_index("Mes_dt")[[plan_ac_col, real_ac_col]]
        st.line_chart(chart_df)

    with st.expander("Ver DATA (tabla completa)"):
        try:
            df_full = _read_excel_sheet(excel_path, SHEET_NAME)
        except Exception as e:
            st.error(f"Error cargando DATA: {e}")
            return
        st.dataframe(df_full, width="stretch", hide_index=True)


st.title("Presupuesto")

files = [
    ("Jefatura Astillero", "Jefatura Astillero.xlsx"),
    ("Maniobras", "Maniobras.xlsx"),
    ("Varadero", "Varadero.xlsx"),
]

min_dt = None
max_dt = None
for _, fn in files:
    try:
        mensual_tmp, _ = _load_monthly_for_file(fn)
    except Exception:
        continue
    if mensual_tmp.empty:
        continue
    mn = mensual_tmp["Mes_dt"].min()
    mx = mensual_tmp["Mes_dt"].max()
    min_dt = mn if min_dt is None else min(min_dt, mn)
    max_dt = mx if max_dt is None else max(max_dt, mx)

with st.sidebar:
    st.header("Filtros")
    if min_dt is None or max_dt is None:
        st.warning("No se pudieron determinar fechas (revisar hoja DATA en los excels).")
        d1 = d2 = None
    else:
        d1, d2 = st.date_input(
            "Rango de meses",
            value=(min_dt.date(), max_dt.date()),
            min_value=min_dt.date(),
            max_value=max_dt.date(),
        )
    mode = st.radio("Vista", options=["Mensual", "Acumulado"], horizontal=False)
    st.caption(f"Moneda: {CURRENCY_SYMBOL} (USD)")

tab1, tab2, tab3 = st.tabs([t for t, _ in files])

if d1 is None or d2 is None:
    st.stop()

with tab1:
    _render_excel_tab(files[0][0], files[0][1], d1, d2, mode)

with tab2:
    _render_excel_tab(files[1][0], files[1][1], d1, d2, mode)

with tab3:
    _render_excel_tab(files[2][0], files[2][1], d1, d2, mode)
