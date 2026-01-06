
import os

import pandas as pd
import streamlit as st

EXCEL_DEFAULT_PATH = os.path.join(os.path.dirname(__file__), "bd faturacion.xlsx")
USD_TO_PEN = 3.4


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).replace("\r", " ").replace("\n", " ").strip() for c in df.columns]
    return df


@st.cache_data(show_spinner=False)
def load_data(excel_path: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name=0)
    df = _normalize_columns(df)

    if "Unnamed: 18" in df.columns:
        df = df.drop(columns=["Unnamed: 18"])

    # Column mapping (handles minor variations)
    fecha_col = None
    for c in df.columns:
        if "Fecha" in c or "fecha" in c:
            fecha_col = c
            break

    if fecha_col is None:
        raise ValueError("No se encontró una columna de fecha (ej: 'Fecha de creación de liquidación').")

    if "Precio" not in df.columns:
        raise ValueError("No se encontró la columna 'Precio'.")

    contratista_col = None
    for cand in ["Nombre Acreedor", "Acreedor", "Contratista", "Proveedor"]:
        if cand in df.columns:
            contratista_col = cand
            break
    if contratista_col is None:
        raise ValueError("No se encontró columna de contratista (ej: 'Nombre Acreedor').")

    df["_fecha"] = pd.to_datetime(df[fecha_col], errors="coerce", dayfirst=True)
    df["_precio"] = pd.to_numeric(df["Precio"], errors="coerce")
    df["_contratista"] = df[contratista_col].astype(str).str.strip()
    df["_moneda"] = df["Mon/"] if "Mon/" in df.columns else "(sin moneda)"
    moneda_upper = df["_moneda"].astype(str).str.upper().str.strip()
    df["_precio_pen"] = df["_precio"].where(moneda_upper != "USD", df["_precio"] * USD_TO_PEN)
    if "Especialidad" in df.columns:
        df["_especialidad"] = df["Especialidad"].astype(str).str.strip()
    else:
        df["_especialidad"] = "(sin especialidad)"

    # Month key
    df["_mes"] = df["_fecha"].dt.to_period("M").astype(str)
    return df


def main() -> None:
    st.set_page_config(page_title="Astillero-Operaciones", layout="wide")

    st.title("Facturación contratista")

    with st.sidebar:
        st.header("Filtros")

    excel_path = EXCEL_DEFAULT_PATH
    if not os.path.exists(excel_path):
        st.error(f"No se encontró el archivo Excel en: {excel_path}")
        st.stop()

    try:
        df = load_data(excel_path)
    except Exception as e:
        st.error(f"Error cargando el archivo: {e}")
        st.stop()

    # Filter to meaningful rows
    base = df.dropna(subset=["_fecha", "_precio_pen"]).copy()
    base = base[base["_contratista"].str.lower() != "nan"]

    with st.sidebar:
        min_date = base["_fecha"].min().date()
        max_date = base["_fecha"].max().date()
        date_range = st.date_input("Rango de fechas", value=(min_date, max_date), min_value=min_date, max_value=max_date)
        if isinstance(date_range, tuple) and len(date_range) == 2:
            d1, d2 = date_range
        else:
            d1, d2 = min_date, max_date

        especialidades = sorted(base["_especialidad"].dropna().astype(str).unique().tolist())
        especialidad_sel = st.multiselect("Especialidad", options=especialidades, default=especialidades)

        contratistas = sorted(base["_contratista"].dropna().astype(str).unique().tolist())

    mask = (
        (base["_fecha"].dt.date >= d1)
        & (base["_fecha"].dt.date <= d2)
        & (base["_especialidad"].astype(str).isin([str(x) for x in especialidad_sel]))
    )
    data = base.loc[mask].copy()

    total = float(data["_precio_pen"].sum())
    n_docs = int(len(data))
    n_contratistas = int(data["_contratista"].nunique())

    c1, c2, c3 = st.columns(3)
    c1.metric("Total (PEN)", f"{total:,.2f}")
    c2.metric("Registros", f"{n_docs:,}")
    c3.metric("Contratistas", f"{n_contratistas:,}")

    st.divider()

    st.subheader("Acumulado por contratista")
    acc = (
        data.groupby(["_contratista"], as_index=False)["_precio_pen"]
        .sum()
        .rename(columns={"_contratista": "Contratista", "_precio_pen": "Total (PEN)"})
        .sort_values("Total (PEN)", ascending=False)
    )
    acc_display = acc.copy()
    acc_display["Total (PEN)"] = acc_display["Total (PEN)"].map(lambda x: f"S/ {x:,.2f}")
    st.dataframe(acc_display, use_container_width=True, hide_index=True)
    st.bar_chart(acc.set_index("Contratista")["Total (PEN)"].head(25))

    with st.expander("Ver datos filtrados"):
        st.dataframe(
            data[[
                "_fecha",
                "_mes",
                "_contratista",
                "_especialidad",
                "_precio",
                "_precio_pen",
            ]].rename(
                columns={
                    "_fecha": "Fecha",
                    "_mes": "Mes",
                    "_contratista": "Contratista",
                    "_especialidad": "Especialidad",
                    "_precio": "Precio (Original)",
                    "_precio_pen": "Precio (PEN)",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )


if __name__ == "__main__":
    main()
