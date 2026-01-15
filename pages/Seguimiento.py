import os

import pandas as pd
import streamlit as st


BASE_DIR = os.path.dirname(os.path.dirname(__file__))
EXCEL_FILENAME = "Seg. salidas no conformes 2025.xlsx"
SHEET_NAME = "DATA2025"

st.set_page_config(page_title="Seguimiento", layout="wide")


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).replace("\r", " ").replace("\n", " ").strip() for c in df.columns]
    return df


def _truthy(v) -> bool:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    s = str(v).strip().lower()
    if s in {"", "0", "false", "no", "nan", "none"}:
        return False
    return True


@st.cache_data(show_spinner=False)
def load_data(excel_path: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name=SHEET_NAME)
    df = _normalize_columns(df)
    return df


def main() -> None:
    st.title("Seguimiento - Salidas no conformes")

    excel_path = os.path.join(BASE_DIR, EXCEL_FILENAME)

    if not os.path.exists(excel_path):
        st.error(f"No se encontró el archivo: {excel_path}")
        st.stop()

    try:
        df = load_data(excel_path)
    except PermissionError:
        st.error(
            "No puedo leer el Excel porque está bloqueado (PermissionError).\n\n"
            "Cierra el archivo si lo tienes abierto en Excel y vuelve a cargar la página. "
            "Si sigue pasando, haz una copia del archivo con otro nombre y usa esa copia."
        )
        st.stop()
    except Exception as e:
        st.error(f"Error cargando el Excel: {e}")
        st.stop()

    # Detect columns
    supervisor_col = None
    if "Supervisor de Proyectos" in df.columns:
        supervisor_col = "Supervisor de Proyectos"
    elif "Supervisor" in df.columns:
        supervisor_col = "Supervisor"
    else:
        for c in df.columns:
            if "supervisor" in c.lower():
                supervisor_col = c
                break

    ac_col = "AC" if "AC" in df.columns else None
    re_col = "RE" if "RE" in df.columns else None

    # Build status column
    df2 = df.copy()
    if ac_col is not None or re_col is not None:
        def _row_status(row):
            if ac_col is not None and _truthy(row.get(ac_col)):
                return "AC"
            if re_col is not None and _truthy(row.get(re_col)):
                return "RE"
            return ""

        df2["_estado"] = df2.apply(_row_status, axis=1)
    else:
        df2["_estado"] = ""

    with st.sidebar:
        st.header("Filtros")

        if supervisor_col is None:
            st.warning("No se encontró la columna 'Supervisor'.")
            supervisor_sel = []
        else:
            supervisors = sorted(df2[supervisor_col].dropna().astype(str).str.strip().unique().tolist())
            supervisor_sel = st.multiselect("Supervisor de Proyectos", options=supervisors, default=supervisors)

        estados = [e for e in ["AC", "RE"] if (e in df2["_estado"].unique())]
        if not estados:
            estados = ["AC", "RE"]
        estado_sel = st.multiselect("Estado (AC/RE)", options=estados, default=estados)

    mask = pd.Series(True, index=df2.index)
    if supervisor_col is not None and supervisor_sel:
        mask &= df2[supervisor_col].astype(str).str.strip().isin([str(x).strip() for x in supervisor_sel])
    if estado_sel:
        mask &= df2["_estado"].isin(estado_sel)

    filtered = df2.loc[mask].copy()

    # KPIs
    ac_count = int((filtered["_estado"] == "AC").sum())
    re_count = int((filtered["_estado"] == "RE").sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("Registros", f"{len(filtered):,}")
    c2.metric("AC", f"{ac_count:,}")
    c3.metric("RE", f"{re_count:,}")

    display_df = filtered.drop(columns=["_estado"], errors="ignore")
    if supervisor_col is not None:
        display_df = display_df.drop(columns=[supervisor_col], errors="ignore")
    display_df = display_df.drop(columns=["AC", "RE"], errors="ignore")
    cols = list(display_df.columns)
    cols_to_hide = []
    if len(cols) >= 2:
        cols_to_hide.append(cols[1])
    if len(cols) >= 1:
        cols_to_hide.append(cols[-1])
    cols_to_hide = list(dict.fromkeys(cols_to_hide))
    display_df = display_df.drop(columns=cols_to_hide, errors="ignore")

    st.dataframe(display_df, width="stretch", hide_index=True)


if __name__ == "__main__":
    main()
