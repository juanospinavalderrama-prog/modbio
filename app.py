"""
Dashboard Finca Las Margaritas — Simulación bioeconómica Monte Carlo.

Aplicación Streamlit para Juan Pablo Ospina. Lee el Excel del modelo
bioeconómico de la finca, permite ajustar las distribuciones de probabilidad
de cada parámetro, ejecuta una simulación Monte Carlo de 5,000 iteraciones
y produce KPIs, histograma interactivo y un reporte PDF descargable para
toma de decisiones.

Ejecutar con:
    streamlit run app.py --server.port 8501 --server.address 0.0.0.0
"""
from __future__ import annotations

import io
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from pdf_report import build_pdf_report
from simulation import (
    ALL_DISTS,
    DIST_CONSTANTE,
    DIST_NORMAL,
    DIST_TRIANGULAR,
    DIST_UNIFORME,
    ParamSpec,
    compute_kpis,
    default_param_specs,
    load_specs_from_excel,
    run_simulation,
)

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"
DEFAULT_XLSX_NAMES = [
    "Finca-Las-Margaritas-Modelo-Bioeconomico.xlsx",
]
PRIMARY = "#01696F"
PRIMARY_DARK = "#0C4E54"
ACCENT = "#20808D"
NEUTRAL_BG = "#F7F6F2"
NEUTRAL_BORDER = "#D4D1CA"
TEXT_DARK = "#28251D"
TEXT_MUTED = "#7A7974"
SUCCESS = "#437A22"
ERROR = "#A12C7B"

st.set_page_config(
    page_title="Finca Las Margaritas · Monte Carlo",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Estilo personalizado (paleta verde/teal, sobria)
# ---------------------------------------------------------------------------
st.markdown(
    f"""
<style>
:root {{
  --primary: {PRIMARY};
  --primary-dark: {PRIMARY_DARK};
  --accent: {ACCENT};
  --bg: {NEUTRAL_BG};
  --border: {NEUTRAL_BORDER};
  --text: {TEXT_DARK};
  --text-muted: {TEXT_MUTED};
}}
.block-container {{ padding-top: 1.6rem; padding-bottom: 2rem; max-width: 1400px; }}
h1, h2, h3, h4 {{ color: var(--primary-dark); letter-spacing: -0.01em; }}
section[data-testid="stSidebar"] {{ background: #FBFAF6; border-right: 1px solid var(--border); }}
.kpi-card {{
    background: white; border: 1px solid var(--border); border-radius: 10px;
    padding: 14px 16px; height: 100%;
}}
.kpi-label {{ color: var(--text-muted); font-size: 0.78rem; text-transform: uppercase;
    letter-spacing: 0.04em; font-weight: 600; }}
.kpi-value {{ color: var(--primary-dark); font-size: 1.45rem; font-weight: 700;
    line-height: 1.15; margin-top: 6px; font-variant-numeric: tabular-nums; }}
.kpi-sub {{ color: var(--text-muted); font-size: 0.78rem; margin-top: 4px; }}
.header-band {{
    background: linear-gradient(135deg, {PRIMARY_DARK} 0%, {PRIMARY} 70%, {ACCENT} 100%);
    color: white; padding: 22px 26px; border-radius: 12px; margin-bottom: 18px;
}}
.header-band h1 {{ color: white; margin: 0 0 4px 0; font-size: 1.6rem; }}
.header-band p {{ color: rgba(255,255,255,0.86); margin: 0; font-size: 0.92rem; }}
div[data-testid="stMetricValue"] {{ color: var(--primary-dark); font-variant-numeric: tabular-nums; }}
.stButton > button {{
    background: var(--primary); color: white; border: none; border-radius: 8px;
    padding: 0.55rem 1.1rem; font-weight: 600;
}}
.stButton > button:hover {{ background: var(--primary-dark); color: white; }}
.stDownloadButton > button {{
    background: white; color: var(--primary-dark);
    border: 1px solid var(--primary); border-radius: 8px; font-weight: 600;
}}
.stDownloadButton > button:hover {{ background: var(--bg); }}
[data-testid="stExpander"] {{ border: 1px solid var(--border) !important; border-radius: 8px; }}
.small-muted {{ color: var(--text-muted); font-size: 0.85rem; }}
</style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Helpers de formato
# ---------------------------------------------------------------------------

def fmt_cop(x: float, short: bool = True) -> str:
    if x is None or not np.isfinite(x):
        return "N/D"
    sign = "-" if x < 0 else ""
    x = abs(x)
    if short:
        if x >= 1e9:
            return f"{sign}${x/1e9:,.2f} mil M"
        if x >= 1e6:
            return f"{sign}${x/1e6:,.2f} M"
        if x >= 1e3:
            return f"{sign}${x/1e3:,.0f} K"
        return f"{sign}${x:,.0f}"
    return f"{sign}${x:,.0f}"


def fmt_num(x: float, decimals: int = 2) -> str:
    if x is None or not np.isfinite(x):
        return "N/D"
    return f"{x:,.{decimals}f}"


# ---------------------------------------------------------------------------
# Carga del Excel
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def load_excel_metadata(xlsx_bytes_or_path) -> Dict[str, pd.DataFrame]:
    """Devuelve un dict con cada hoja relevante como DataFrame y los warnings."""
    target_sheets = ["Parámetros base", "Variables salida", "Resumen Monte Carlo", "Curva "]
    out: Dict[str, pd.DataFrame] = {}
    missing: List[str] = []
    try:
        xl = pd.ExcelFile(xlsx_bytes_or_path)
        names = xl.sheet_names
        for s in target_sheets:
            if s in names:
                try:
                    out[s] = xl.parse(s, header=None)
                except Exception as e:  # noqa: BLE001
                    out[s] = pd.DataFrame()
                    missing.append(f"No se pudo leer '{s}': {e}")
            else:
                missing.append(f"Falta hoja '{s}'.")
    except Exception as e:  # noqa: BLE001
        missing.append(f"Error al abrir el Excel: {e}")
    out["_missing"] = pd.DataFrame({"warning": missing})  # type: ignore[assignment]
    return out


def get_default_xlsx_path() -> Optional[Path]:
    """Busca el Excel por defecto incluido dentro del proyecto."""
    for name in DEFAULT_XLSX_NAMES:
        p = DATA_DIR / name
        if p.exists():
            return p
    # Fallback al workspace
    fallback = Path("/home/user/workspace") / DEFAULT_XLSX_NAMES[0]
    if fallback.exists():
        return fallback
    return None


# ---------------------------------------------------------------------------
# Estado de sesión
# ---------------------------------------------------------------------------
if "specs" not in st.session_state:
    st.session_state.specs = None
if "specs_overridden" not in st.session_state:
    st.session_state.specs_overridden = {}
if "sim_result" not in st.session_state:
    st.session_state.sim_result = None
if "xlsx_source_label" not in st.session_state:
    st.session_state.xlsx_source_label = ""


# ---------------------------------------------------------------------------
# Encabezado
# ---------------------------------------------------------------------------
st.markdown(
    """
<div class="header-band">
  <h1>🌿 Dashboard Bioeconómico · Finca Las Margaritas</h1>
  <p>Simulación Monte Carlo del margen anual del hato bajo incertidumbre productiva,
     reproductiva y de precios. Adaptación del modelo de Vargas-Leitón &amp; Cuevas-Abrego (2023).</p>
</div>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Sidebar: carga de datos + configuración general
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 📂 Datos de entrada")
    upload = st.file_uploader(
        "Cargar Excel del modelo bioeconómico",
        type=["xlsx"],
        help="Si no cargas un archivo, se usa el modelo base de la Finca Las Margaritas.",
    )

    if upload is not None:
        xlsx_bytes = upload.read()
        xlsx_source = io.BytesIO(xlsx_bytes)
        st.session_state.xlsx_source_label = f"Subido: {upload.name}"
        # Guardar a temporal para leer también por hoja
        tmp_path = APP_DIR / ".uploaded.xlsx"
        tmp_path.write_bytes(xlsx_bytes)
        xlsx_path_for_specs = str(tmp_path)
    else:
        default_path = get_default_xlsx_path()
        if default_path is None:
            st.error("No se encontró el Excel por defecto. Por favor sube uno.")
            st.stop()
        xlsx_path_for_specs = str(default_path)
        st.session_state.xlsx_source_label = f"Por defecto: {default_path.name}"

    st.caption(st.session_state.xlsx_source_label)

    sheets = load_excel_metadata(xlsx_path_for_specs)
    missing = sheets.get("_missing", pd.DataFrame())
    if not missing.empty:
        for w in missing["warning"].tolist():
            st.warning(w)

    # Cargar specs desde Excel
    if st.session_state.specs is None or upload is not None:
        st.session_state.specs = load_specs_from_excel(xlsx_path_for_specs)
        st.session_state.specs_overridden = {}

    st.markdown("---")
    st.markdown("### ⚙️ Configuración de simulación")

    n_iter = st.number_input(
        "Número de iteraciones",
        min_value=100, max_value=50_000, value=5000, step=500,
        help="El artículo de Vargas-Leitón usa 5,000 iteraciones por hipercubo latino. Aquí usamos Monte Carlo simple."
    )

    use_seed = st.checkbox("Fijar semilla aleatoria (reproducible)", value=True)
    seed_val: Optional[int] = None
    if use_seed:
        seed_val = int(st.number_input("Semilla", min_value=0, max_value=2**31 - 1, value=42))

    st.markdown("---")
    st.markdown("### 🌱 Constantes operativas")
    st.caption("Ajusta supuestos de la finca que no son aleatorios por defecto.")
    spec_pesonac = st.session_state.specs.get("PesoNac")
    spec_conv = st.session_state.specs.get("ConvAlim")
    spec_rm = st.session_state.specs.get("FactorRM")
    if spec_pesonac:
        new_v = st.number_input("Peso al nacimiento (kg)", value=float(spec_pesonac.base), step=1.0)
        spec_pesonac.base = new_v; spec_pesonac.mean = new_v; spec_pesonac.constant_value = new_v
    if spec_conv:
        new_v = st.number_input("Conversión leche/concentrado (kg/kg)", value=float(spec_conv.base), step=0.1)
        spec_conv.base = new_v; spec_conv.mean = new_v; spec_conv.constant_value = new_v
    if spec_rm:
        new_v = st.number_input("Factor de reducción por mastitis", value=float(spec_rm.base),
                                step=0.01, min_value=0.0, max_value=1.0, format="%.3f")
        spec_rm.base = new_v; spec_rm.mean = new_v; spec_rm.constant_value = new_v


# ---------------------------------------------------------------------------
# Tabs principales
# ---------------------------------------------------------------------------
tab_params, tab_sim, tab_outputs, tab_excel, tab_about = st.tabs([
    "1 · Parámetros y distribuciones",
    "2 · Resultados Monte Carlo",
    "3 · Variables de salida",
    "4 · Hojas del Excel",
    "ℹ️ Metodología",
])


# ---------------------------------------------------------------------------
# Tab 1: Parámetros y distribuciones
# ---------------------------------------------------------------------------
with tab_params:
    st.markdown("### Ajuste de distribuciones de probabilidad")
    st.markdown(
        '<p class="small-muted">Para cada parámetro de entrada elige la distribución '
        '(Normal, Triangular, Uniforme o Constante) y sus argumentos. Los valores se '
        'pre-llenan desde el Excel y siguen la convención del modelo bioeconómico de la finca.</p>',
        unsafe_allow_html=True,
    )

    # Agrupar por bloque
    specs: Dict[str, ParamSpec] = st.session_state.specs
    groups: Dict[str, List[ParamSpec]] = {}
    for code, s in specs.items():
        # Saltar los meta-parámetros del sidebar
        if code in {"PesoNac", "ConvAlim", "FactorRM"}:
            continue
        groups.setdefault(s.group or "Otros", []).append(s)

    group_order = ["Hato", "Reemplazo", "Crecimiento", "Reproducción",
                   "Reproducción novillas", "Producción", "Composición",
                   "Carne", "Salud", "Económicos", "Alimentación", "Otros"]
    sorted_groups = sorted(groups.items(),
                           key=lambda kv: group_order.index(kv[0]) if kv[0] in group_order else 99)

    for group_name, params_in_group in sorted_groups:
        with st.expander(f"▾  {group_name}  ({len(params_in_group)} parámetros)",
                         expanded=(group_name in {"Hato", "Producción", "Económicos"})):
            for sp in params_in_group:
                cols = st.columns([2.5, 1.4, 3.0])
                with cols[0]:
                    st.markdown(f"**{sp.label}**  \n"
                                f"<span class='small-muted'>código `{sp.code}` · {sp.unit or '·'} · "
                                f"base {fmt_num(sp.base, 4)}</span>",
                                unsafe_allow_html=True)
                with cols[1]:
                    new_dist = st.selectbox(
                        "Distribución",
                        ALL_DISTS,
                        index=ALL_DISTS.index(sp.distribution) if sp.distribution in ALL_DISTS else 0,
                        key=f"dist_{sp.code}",
                        label_visibility="collapsed",
                    )
                    sp.distribution = new_dist
                with cols[2]:
                    if new_dist == DIST_NORMAL:
                        c1, c2 = st.columns(2)
                        sp.mean = c1.number_input("Media", value=float(sp.mean),
                                                  key=f"mean_{sp.code}",
                                                  format="%.6f")
                        sp.sd = c2.number_input("Desv. estándar (σ)", value=float(max(sp.sd, 0.0)),
                                                key=f"sd_{sp.code}", min_value=0.0,
                                                format="%.6f")
                    elif new_dist == DIST_TRIANGULAR:
                        c1, c2, c3 = st.columns(3)
                        sp.tri_min = c1.number_input("Mínimo", value=float(sp.tri_min),
                                                    key=f"tmin_{sp.code}", format="%.6f")
                        sp.tri_mode = c2.number_input("Moda", value=float(sp.tri_mode),
                                                     key=f"tmod_{sp.code}", format="%.6f")
                        sp.tri_max = c3.number_input("Máximo", value=float(sp.tri_max),
                                                    key=f"tmax_{sp.code}", format="%.6f")
                    elif new_dist == DIST_UNIFORME:
                        c1, c2 = st.columns(2)
                        sp.uni_min = c1.number_input("Mínimo", value=float(sp.uni_min),
                                                    key=f"umin_{sp.code}", format="%.6f")
                        sp.uni_max = c2.number_input("Máximo", value=float(sp.uni_max),
                                                    key=f"umax_{sp.code}", format="%.6f")
                    else:  # Constante
                        sp.constant_value = st.number_input(
                            "Valor", value=float(sp.constant_value),
                            key=f"const_{sp.code}", format="%.6f",
                        )
                        sp.mean = sp.constant_value
                        sp.base = sp.constant_value

    st.markdown("---")
    rcol1, rcol2, rcol3 = st.columns([1.4, 1.0, 1.0])
    with rcol1:
        run_sim = st.button(
            "▶  Ejecutar simulación Monte Carlo",
            type="primary",
            use_container_width=True,
        )
    with rcol2:
        if st.button("↺  Restaurar valores del Excel", use_container_width=True):
            st.session_state.specs = load_specs_from_excel(xlsx_path_for_specs)
            st.rerun()
    with rcol3:
        all_const = st.button("⏸  Tratar todos como constantes", use_container_width=True)
        if all_const:
            for code, s in st.session_state.specs.items():
                s.distribution = DIST_CONSTANTE
            st.rerun()

    if run_sim:
        with st.spinner(f"Ejecutando {int(n_iter):,} iteraciones..."):
            res = run_simulation(
                st.session_state.specs,
                n_iter=int(n_iter),
                seed=seed_val,
            )
        st.session_state.sim_result = res
        st.success(f"Simulación completada: {int(n_iter):,} iteraciones.")
        for w in res.warnings:
            st.warning(w)


# ---------------------------------------------------------------------------
# Tab 2: Resultados Monte Carlo
# ---------------------------------------------------------------------------
with tab_sim:
    res = st.session_state.sim_result
    if res is None:
        st.info("Configura tus distribuciones en la pestaña anterior y pulsa "
                "**Ejecutar simulación Monte Carlo** para ver los resultados.")
    else:
        margin = res.outputs["Margen bruto COP"].values
        kpis = compute_kpis(margin)

        st.markdown("### Margen bruto anual del hato (COP)")

        # KPIs en cards
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(f"""<div class="kpi-card">
              <div class="kpi-label">Media</div>
              <div class="kpi-value">{fmt_cop(kpis['media'])}</div>
              <div class="kpi-sub">Esperanza del margen</div></div>""", unsafe_allow_html=True)
        with k2:
            st.markdown(f"""<div class="kpi-card">
              <div class="kpi-label">Mediana</div>
              <div class="kpi-value">{fmt_cop(kpis['mediana'])}</div>
              <div class="kpi-sub">Valor central</div></div>""", unsafe_allow_html=True)
        with k3:
            st.markdown(f"""<div class="kpi-card">
              <div class="kpi-label">Desv. estándar</div>
              <div class="kpi-value">{fmt_cop(kpis['sd'])}</div>
              <div class="kpi-sub">CV = {fmt_num(kpis['cv'], 3)}</div></div>""", unsafe_allow_html=True)
        with k4:
            color = ERROR if kpis["prob_neg"] > 0.10 else (PRIMARY_DARK if kpis["prob_neg"] > 0.01 else SUCCESS)
            st.markdown(f"""<div class="kpi-card">
              <div class="kpi-label">Probabilidad de pérdida</div>
              <div class="kpi-value" style="color:{color}">{kpis['prob_neg']*100:.1f}%</div>
              <div class="kpi-sub">P(margen &lt; 0)</div></div>""", unsafe_allow_html=True)

        st.markdown("")
        k5, k6, k7, k8 = st.columns(4)
        with k5:
            st.markdown(f"""<div class="kpi-card">
              <div class="kpi-label">Percentil 2.5</div>
              <div class="kpi-value">{fmt_cop(kpis['p2.5'])}</div>
              <div class="kpi-sub">Escenario pesimista</div></div>""", unsafe_allow_html=True)
        with k6:
            st.markdown(f"""<div class="kpi-card">
              <div class="kpi-label">Percentil 97.5</div>
              <div class="kpi-value">{fmt_cop(kpis['p97.5'])}</div>
              <div class="kpi-sub">Escenario optimista</div></div>""", unsafe_allow_html=True)
        with k7:
            ic_amplitud = kpis["p97.5"] - kpis["p2.5"]
            st.markdown(f"""<div class="kpi-card">
              <div class="kpi-label">Amplitud IC 95%</div>
              <div class="kpi-value">{fmt_cop(ic_amplitud)}</div>
              <div class="kpi-sub">p97.5 − p2.5</div></div>""", unsafe_allow_html=True)
        with k8:
            st.markdown(f"""<div class="kpi-card">
              <div class="kpi-label">Iteraciones válidas</div>
              <div class="kpi-value">{int(kpis['n']):,}</div>
              <div class="kpi-sub">de {res.n_iter:,} totales</div></div>""", unsafe_allow_html=True)

        st.markdown("")

        # Histograma plotly
        st.markdown("#### Distribución del margen bruto")
        valid = margin[np.isfinite(margin)]
        valid_m = valid / 1e6
        fig = go.Figure()
        fig.add_trace(go.Histogram(
            x=valid_m, nbinsx=60,
            marker_color=ACCENT, marker_line_color="white", marker_line_width=1,
            name="Iteraciones",
            hovertemplate="Margen: $%{x:,.1f} M<br>Frecuencia: %{y}<extra></extra>",
        ))
        fig.add_vline(x=kpis["media"]/1e6, line_width=2, line_color=PRIMARY_DARK,
                      annotation_text=f"Media ${kpis['media']/1e6:,.1f}M",
                      annotation_position="top right",
                      annotation_font_color=PRIMARY_DARK)
        fig.add_vline(x=kpis["p2.5"]/1e6, line_width=1.5, line_dash="dash",
                      line_color="#A84B2F",
                      annotation_text=f"p2.5",
                      annotation_position="top left",
                      annotation_font_color="#A84B2F")
        fig.add_vline(x=kpis["p97.5"]/1e6, line_width=1.5, line_dash="dash",
                      line_color="#A84B2F",
                      annotation_text=f"p97.5",
                      annotation_position="top right",
                      annotation_font_color="#A84B2F")
        fig.add_vline(x=0, line_width=1, line_dash="dot", line_color=TEXT_MUTED)
        fig.update_layout(
            height=460,
            margin=dict(l=20, r=20, t=30, b=40),
            paper_bgcolor="white", plot_bgcolor="white",
            xaxis=dict(title="Margen bruto (millones de COP)", gridcolor="#EFEEE9", zerolinecolor="#CFCDC6"),
            yaxis=dict(title="Frecuencia (iteraciones)", gridcolor="#EFEEE9", zerolinecolor="#CFCDC6"),
            font=dict(family="Helvetica, Arial, sans-serif", color=TEXT_DARK, size=12),
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True)

        # Intervalo de confianza explícito
        st.markdown(
            f'<p class="small-muted">Intervalo de confianza del 95% '
            f'(percentiles 2.5–97.5): <b>{fmt_cop(kpis["p2.5"])}</b> '
            f'a <b>{fmt_cop(kpis["p97.5"])}</b>. '
            f'Iteraciones excluidas por dominio inválido: '
            f'{res.n_iter - int(kpis["n"]):,}.</p>',
            unsafe_allow_html=True,
        )

        # Descargas
        st.markdown("---")
        st.markdown("#### Exportar resultados")
        dcol1, dcol2, dcol3 = st.columns([1.2, 1.2, 1.2])

        # CSV de salidas
        with dcol1:
            csv_out = res.outputs.to_csv(index_label="iter").encode("utf-8")
            st.download_button(
                "⬇  CSV — Variables de salida",
                data=csv_out,
                file_name=f"margaritas_outputs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with dcol2:
            csv_in = res.inputs.to_csv(index_label="iter").encode("utf-8")
            st.download_button(
                "⬇  CSV — Entradas muestreadas",
                data=csv_in,
                file_name=f"margaritas_inputs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with dcol3:
            # PDF
            specs_used = []
            for code, sp in st.session_state.specs.items():
                if sp.distribution == DIST_NORMAL:
                    params = f"μ={sp.mean:.6g}, σ={sp.sd:.6g}"
                elif sp.distribution == DIST_TRIANGULAR:
                    params = f"min={sp.tri_min:.6g}, moda={sp.tri_mode:.6g}, max={sp.tri_max:.6g}"
                elif sp.distribution == DIST_UNIFORME:
                    params = f"min={sp.uni_min:.6g}, max={sp.uni_max:.6g}"
                else:
                    params = f"= {sp.constant_value:.6g}"
                specs_used.append({
                    "code": code, "label": sp.label,
                    "dist": sp.distribution, "params": params,
                })
            pdf_bytes = build_pdf_report(res, specs_used)
            st.download_button(
                "⬇  PDF — Resumen estadístico",
                data=pdf_bytes,
                file_name=f"margaritas_resumen_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )


# ---------------------------------------------------------------------------
# Tab 3: Variables de salida
# ---------------------------------------------------------------------------
with tab_outputs:
    res = st.session_state.sim_result
    if res is None:
        st.info("Aún no hay resultados de simulación.")
    else:
        st.markdown("### Resumen estadístico por variable de salida")
        st.markdown(
            '<p class="small-muted">Cada fila corresponde a una variable del modelo bioeconómico. '
            'Las cifras económicas (IT, CT, Margen bruto) están en COP. Las cifras "kg vaca" o "kg hato" '
            'están en kilogramos. Las demográficas son número de cabezas.</p>',
            unsafe_allow_html=True,
        )
        summary = res.summary_outputs()
        # Mostrar tabla con formato
        display = summary.copy()
        for col in ["Media", "DE", "Mediana", "p2.5", "p97.5", "Min", "Max"]:
            display[col] = display[col].apply(lambda x: f"{x:,.2f}" if np.isfinite(x) else "N/D")
        st.dataframe(display, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("### Explorar una variable individual")
        var_options = list(res.outputs.columns)
        chosen = st.selectbox("Variable", var_options,
                              index=var_options.index("Margen bruto COP"))
        x = res.outputs[chosen].values
        valid = x[np.isfinite(x)]
        if valid.size == 0:
            st.warning("No hay valores válidos para esta variable.")
        else:
            colA, colB = st.columns([2, 1])
            with colA:
                fig2 = go.Figure()
                fig2.add_trace(go.Histogram(
                    x=valid, nbinsx=50,
                    marker_color=ACCENT, marker_line_color="white", marker_line_width=1,
                ))
                fig2.add_vline(x=np.mean(valid), line_color=PRIMARY_DARK, line_width=2,
                               annotation_text="Media", annotation_position="top right",
                               annotation_font_color=PRIMARY_DARK)
                fig2.update_layout(
                    height=380, margin=dict(l=20, r=20, t=30, b=40),
                    paper_bgcolor="white", plot_bgcolor="white",
                    xaxis=dict(title=chosen, gridcolor="#EFEEE9"),
                    yaxis=dict(title="Frecuencia", gridcolor="#EFEEE9"),
                    font=dict(family="Helvetica, Arial, sans-serif", color=TEXT_DARK),
                    showlegend=False,
                )
                st.plotly_chart(fig2, use_container_width=True)
            with colB:
                stats = {
                    "Media": np.mean(valid),
                    "Mediana": np.median(valid),
                    "Desv. estándar": np.std(valid, ddof=1) if valid.size > 1 else 0.0,
                    "Mínimo": np.min(valid),
                    "Máximo": np.max(valid),
                    "p2.5": np.percentile(valid, 2.5),
                    "p25": np.percentile(valid, 25),
                    "p75": np.percentile(valid, 75),
                    "p97.5": np.percentile(valid, 97.5),
                    "n válidos": valid.size,
                }
                df_stats = pd.DataFrame(
                    [{"Estadístico": k, "Valor": (f"{v:,.4f}" if isinstance(v, float) else v)}
                     for k, v in stats.items()]
                )
                st.dataframe(df_stats, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# Tab 4: Hojas del Excel
# ---------------------------------------------------------------------------
with tab_excel:
    st.markdown("### Hojas del modelo bioeconómico")
    st.markdown(
        '<p class="small-muted">Vista previa de las hojas relevantes del Excel cargado. '
        'Útil para auditar parámetros de origen y referencias del modelo.</p>',
        unsafe_allow_html=True,
    )
    sheets = load_excel_metadata(xlsx_path_for_specs)
    available_sheets = [k for k in sheets.keys()
                        if k != "_missing" and not sheets[k].empty]
    if not available_sheets:
        st.warning("No se pudo leer ninguna hoja relevante.")
    else:
        sheet_pick = st.selectbox("Hoja", available_sheets)
        df = sheets[sheet_pick]
        # Limpiar para visualización
        df_view = df.copy()
        st.dataframe(df_view, use_container_width=True, height=520)


# ---------------------------------------------------------------------------
# Tab 5: Acerca de / Metodología
# ---------------------------------------------------------------------------
with tab_about:
    st.markdown("### Metodología")
    st.markdown(
        """
**Modelo bioeconómico** — adaptación simplificada y trazable del marco propuesto por
Vargas-Leitón &amp; Cuevas-Abrego (2023, *Stochastic model to estimate economic values of
production and functional traits in dairy cattle*). El modelo simula el ciclo
productivo-reproductivo del hato de la finca durante un año calendario y agrega los
resultados a nivel de hato.

**Ecuaciones principales** (siguen la hoja `Variables salida` del Excel):
- **Curva de lactancia (Wood):** `y(t) = a · t^b · exp(−c·t)`, t = 1…DLAC.
- **Edad al primer servicio (Gompertz inverso):**
  `EIA = (ln(ln(PIA/Ag)/−1) − Bg) / (−Cg)`, en meses.
- **Edad al primer parto:** `EPA = EIA + (IAC + DP)/30.4`.
- **Vacas lactantes:** `VL = VA·(1 − TDIv/100) − VA·TMv/100`.
- **Hembras de reemplazo:** `HR = VL · 0.5 · (1 − TDIr/100) · 24 / EPA`.
- **Reducción por mastitis:** `RM = PLn · (IMC/100) · factor_reducción_mastitis`.
- **Ingresos totales:** `IT = (PGR+PPR)·VL·$KGP + PLM·VL·$LM + PCAR·$CAR + NV·$NOV`.
- **Costos totales:** `CT = $CON·(CON + CON/3) + VL·$VAC + HR·$REM + $AD`.
- **Margen bruto:** `IT − CT`.

**Simulación Monte Carlo.** Por defecto se ejecutan **5,000 iteraciones**. Las variables de
entrada se muestrean independientemente desde sus distribuciones (normal, triangular o
uniforme). Las constantes se mantienen fijas. Iteraciones con dominio inválido
(p. ej. ln(PIA/Ag) cuando PIA ≥ Ag) se marcan como `NaN` y se excluyen del cálculo de KPIs.

**Diferencias frente al artículo.**
- Usamos muestreo Monte Carlo simple en vez de Hipercubo Latino. El número de iteraciones
  (5,000) sigue la recomendación del artículo y produce KPIs estables.
- No se implementa la regresión de valores económicos parciales (rasgo a rasgo); el
  dashboard se centra en el riesgo del margen bruto agregado, que era el objetivo del usuario.
- Se utiliza el valor de `c (Wood)` provisto por la hoja `Curva` del Excel (≈0.00486),
  no el valor inconsistente publicado en la hoja `Parámetros base` (≈0.034).

**Cobertura de variables.** La simulación calcula 25 variables de salida:
VL, HR, VDI, VDV, NE, NV, TR, EIA, GDP, EPA, VP, PLn, PP, DS, RM, PLsm,
PGR, PPR, PLM, PSOL, PCAR, CON, IT, CT y Margen bruto COP.

**Manejo de errores.** Variables faltantes, dispersiones no numéricas, valores negativos
imposibles y dominios logarítmicos no válidos se detectan en `simulation.py` y producen
advertencias visibles en la UI.

**Reproducibilidad.** Si fijas la semilla en la barra lateral, dos ejecuciones con la
misma configuración producen exactamente los mismos resultados.
"""
    )
    st.markdown(
        '<p class="small-muted">Construido para Juan Pablo Ospina · Finca Las Margaritas · '
        f'{datetime.now().strftime("%Y")}</p>',
        unsafe_allow_html=True,
    )
