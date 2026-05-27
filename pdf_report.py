"""
Generación de PDF de resumen estadístico de la simulación Monte Carlo.

Usa ReportLab + matplotlib para producir un PDF profesional, en español,
listo para toma de decisiones en finca.
"""
from __future__ import annotations

import io
from datetime import datetime
from typing import Dict, List, Optional

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from simulation import SimulationResult, compute_kpis


# Paleta sobria verde/teal
PRIMARY = colors.HexColor("#01696F")   # Hydra teal
PRIMARY_DARK = colors.HexColor("#0C4E54")
ACCENT = colors.HexColor("#20808D")    # Chart teal
NEUTRAL_BG = colors.HexColor("#F7F6F2")
NEUTRAL_BORDER = colors.HexColor("#D4D1CA")
TEXT_DARK = colors.HexColor("#28251D")
TEXT_MUTED = colors.HexColor("#7A7974")
SUCCESS = colors.HexColor("#437A22")
ERROR = colors.HexColor("#A12C7B")


def _fmt_cop(x: float) -> str:
    if not np.isfinite(x):
        return "N/D"
    sign = "-" if x < 0 else ""
    x = abs(x)
    if x >= 1e9:
        return f"{sign}${x/1e9:,.2f} mil M COP"
    if x >= 1e6:
        return f"{sign}${x/1e6:,.2f} M COP"
    if x >= 1e3:
        return f"{sign}${x/1e3:,.0f} K COP"
    return f"{sign}${x:,.0f} COP"


def _fmt_num(x: float, decimals: int = 2) -> str:
    if not np.isfinite(x):
        return "N/D"
    return f"{x:,.{decimals}f}"


def _make_histogram(margin: np.ndarray, kpis: Dict[str, float]) -> bytes:
    """Genera un histograma del Margen bruto en COP como PNG en memoria."""
    fig, ax = plt.subplots(figsize=(8, 4.2), dpi=150)
    x = margin[np.isfinite(margin)] / 1e6  # millones de COP
    ax.hist(x, bins=50, color="#20808D", edgecolor="white", alpha=0.92)

    mean_m = kpis["media"] / 1e6
    p_lo = kpis["p2.5"] / 1e6
    p_hi = kpis["p97.5"] / 1e6

    ax.axvline(mean_m, color="#0C4E54", linestyle="-",
               linewidth=1.6, label=f"Media: ${mean_m:,.1f} M")
    ax.axvline(p_lo, color="#A84B2F", linestyle="--", linewidth=1.2,
               label=f"p2.5: ${p_lo:,.1f} M")
    ax.axvline(p_hi, color="#A84B2F", linestyle="--", linewidth=1.2,
               label=f"p97.5: ${p_hi:,.1f} M")
    ax.axvline(0, color="#7A7974", linestyle=":", linewidth=1.0, alpha=0.7)

    ax.set_xlabel("Margen bruto (millones de COP)", fontsize=10)
    ax.set_ylabel("Frecuencia (iteraciones)", fontsize=10)
    ax.set_title("Distribución del margen bruto anual del hato",
                 fontsize=12, fontweight="bold", loc="left", color="#28251D")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(True, axis="y", linestyle="-", alpha=0.25)
    ax.legend(loc="upper right", frameon=False, fontsize=9)
    plt.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def _styles():
    base = getSampleStyleSheet()
    styles = {
        "title": ParagraphStyle(
            "title", parent=base["Title"], fontName="Helvetica-Bold",
            fontSize=20, leading=24, textColor=PRIMARY_DARK,
            spaceAfter=4, alignment=0,
        ),
        "subtitle": ParagraphStyle(
            "subtitle", parent=base["Normal"], fontName="Helvetica",
            fontSize=11, leading=14, textColor=TEXT_MUTED, spaceAfter=14,
        ),
        "h2": ParagraphStyle(
            "h2", parent=base["Heading2"], fontName="Helvetica-Bold",
            fontSize=13, leading=16, textColor=PRIMARY_DARK,
            spaceBefore=10, spaceAfter=6,
        ),
        "body": ParagraphStyle(
            "body", parent=base["BodyText"], fontName="Helvetica",
            fontSize=10, leading=14, textColor=TEXT_DARK, spaceAfter=4,
        ),
        "small": ParagraphStyle(
            "small", parent=base["BodyText"], fontName="Helvetica",
            fontSize=8.5, leading=11, textColor=TEXT_MUTED,
        ),
        "kpi_label": ParagraphStyle(
            "kpi_label", parent=base["Normal"], fontName="Helvetica",
            fontSize=8.5, leading=10, textColor=TEXT_MUTED,
            alignment=1,
        ),
        "kpi_value": ParagraphStyle(
            "kpi_value", parent=base["Normal"], fontName="Helvetica-Bold",
            fontSize=12, leading=14, textColor=PRIMARY_DARK, alignment=1,
        ),
    }
    return styles


def _kpi_table(kpis: Dict[str, float], styles) -> Table:
    rows = [
        ["Media", "Mediana", "Desv. estándar", "Coef. variación",
         "p2.5", "p97.5", "P(margen<0)", "Iteraciones"],
        [
            _fmt_cop(kpis["media"]).replace(" COP", ""),
            _fmt_cop(kpis["mediana"]).replace(" COP", ""),
            _fmt_cop(kpis["sd"]).replace(" COP", ""),
            _fmt_num(kpis["cv"], 3),
            _fmt_cop(kpis["p2.5"]).replace(" COP", ""),
            _fmt_cop(kpis["p97.5"]).replace(" COP", ""),
            f"{kpis['prob_neg']*100:.1f}%",
            f"{int(kpis['n']):,}",
        ],
    ]
    # Anchos diferenciados: las primeras columnas necesitan más espacio para cifras COP
    col_widths = [2.55 * cm, 2.55 * cm, 2.55 * cm, 1.85 * cm,
                  2.4 * cm, 2.4 * cm, 1.85 * cm, 1.85 * cm]
    t = Table(rows, colWidths=col_widths, hAlign="LEFT")
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica"),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8.0),
        ("FONTSIZE", (0, 1), (-1, 1), 9.5),
        ("TEXTCOLOR", (0, 0), (-1, 0), TEXT_MUTED),
        ("TEXTCOLOR", (0, 1), (-1, 1), PRIMARY_DARK),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, 0), NEUTRAL_BG),
        ("BOX", (0, 0), (-1, -1), 0.4, NEUTRAL_BORDER),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, NEUTRAL_BORDER),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    return t


def _distributions_table(specs_used: List[Dict[str, str]]) -> Table:
    header = ["Código", "Variable", "Distribución", "Parámetros"]
    data = [header] + [[r["code"], r["label"], r["dist"], r["params"]] for r in specs_used]
    t = Table(data, colWidths=[2.0 * cm, 6.0 * cm, 2.6 * cm, 6.5 * cm], hAlign="LEFT", repeatRows=1)
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 0), (-1, 0), PRIMARY),
        ("TEXTCOLOR", (0, 1), (-1, -1), TEXT_DARK),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, NEUTRAL_BG]),
        ("BOX", (0, 0), (-1, -1), 0.4, NEUTRAL_BORDER),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, NEUTRAL_BORDER),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


def _outputs_summary_table(summary_df: pd.DataFrame, styles) -> Table:
    # Para variables económicas, usamos formato corto en millones de COP
    economic_vars = {"IT", "CT", "Margen bruto COP", "PSOL"}

    # Etiquetas más cortas para la columna Variable (evita desbordes en el PDF)
    label_map = {
        "Margen bruto COP": "Margen bruto",
    }

    # Estilo de párrafo para la columna Variable: permite ajuste de línea si hace falta
    var_style = ParagraphStyle(
        "var_cell", parent=styles["body"], fontName="Helvetica",
        fontSize=8.0, leading=10, textColor=TEXT_DARK, alignment=0,
    )

    def fmt(val: float, var: str) -> str:
        if not np.isfinite(val):
            return "N/D"
        if var in economic_vars:
            # Mostrar en millones para mantener legibilidad
            return f"{val/1e6:,.1f}M"
        if abs(val) >= 10000:
            return f"{val:,.0f}"
        return f"{val:,.2f}"

    header = ["Variable", "Media", "DE", "Mediana", "p2.5", "p97.5", "Min", "Max"]
    rows = [header]
    for _, r in summary_df.iterrows():
        var = str(r["Variable"])
        display = label_map.get(var, var)
        # Envolver la etiqueta en un Paragraph permite wrap automático
        rows.append([
            Paragraph(display, var_style),
            fmt(r["Media"], var),
            fmt(r["DE"], var),
            fmt(r["Mediana"], var),
            fmt(r["p2.5"], var),
            fmt(r["p97.5"], var),
            fmt(r["Min"], var),
            fmt(r["Max"], var),
        ])
    # Variable más ancha (3.0 cm) para evitar overflow de etiquetas largas;
    # el resto se ajusta para que la tabla siga cabiendo dentro de A4 con
    # márgenes laterales de 1.6 cm (ancho útil ≈ 17.78 cm).
    t = Table(
        rows,
        colWidths=[3.0 * cm, 2.05 * cm, 2.05 * cm, 2.15 * cm,
                   2.05 * cm, 2.15 * cm, 2.0 * cm, 2.0 * cm],
        hAlign="LEFT",
        repeatRows=1,
    )
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (1, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.0),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 0), (-1, 0), PRIMARY),
        ("TEXTCOLOR", (1, 1), (-1, -1), TEXT_DARK),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, NEUTRAL_BG]),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 0), (0, 0), "CENTER"),
        ("ALIGN", (1, 0), (-1, 0), "CENTER"),
        ("ALIGN", (0, 1), (0, -1), "LEFT"),
        ("LEFTPADDING", (0, 1), (0, -1), 6),
        ("BOX", (0, 0), (-1, -1), 0.4, NEUTRAL_BORDER),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, NEUTRAL_BORDER),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


def _interpretation(kpis: Dict[str, float]) -> str:
    """Genera una breve interpretación textual del riesgo financiero."""
    mean = kpis["media"]
    p_lo = kpis["p2.5"]
    p_hi = kpis["p97.5"]
    cv = kpis["cv"]
    prob_neg = kpis["prob_neg"]

    nivel_riesgo = "bajo"
    if prob_neg > 0.25:
        nivel_riesgo = "alto"
    elif prob_neg > 0.05:
        nivel_riesgo = "moderado"

    if cv < 0.15:
        var_text = "muy estable"
    elif cv < 0.30:
        var_text = "moderadamente variable"
    elif cv < 0.60:
        var_text = "variable"
    else:
        var_text = "altamente variable"

    txt = (
        f"La simulación Monte Carlo proyecta un margen bruto anual esperado de "
        f"<b>{_fmt_cop(mean)}</b> para el hato de la Finca Las Margaritas. El intervalo de "
        f"confianza del 95% se sitúa entre <b>{_fmt_cop(p_lo)}</b> y <b>{_fmt_cop(p_hi)}</b>, "
        f"con un coeficiente de variación de <b>{cv:.2f}</b>, lo que indica que el resultado "
        f"es <b>{var_text}</b>. "
        f"La probabilidad de obtener un margen negativo es del <b>{prob_neg*100:.1f}%</b>, "
        f"lo que clasifica el riesgo financiero del escenario simulado como <b>{nivel_riesgo}</b>. "
        f"Se recomienda revisar las distribuciones de las variables con mayor sensibilidad "
        f"(producción de leche, precios de grasa y proteína, costos de concentrado y reemplazo) "
        f"antes de decisiones de inversión o desafío productivo."
    )
    return txt


def build_pdf_report(
    sim: SimulationResult,
    specs_used: List[Dict[str, str]],
    farm_name: str = "Finca Las Margaritas",
    author: str = "Perplexity Computer",
    extra_notes: Optional[str] = None,
) -> bytes:
    """Construye el PDF y devuelve los bytes."""
    buf = io.BytesIO()
    title = f"Resumen Monte Carlo — {farm_name}"
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=1.6 * cm, rightMargin=1.6 * cm,
        topMargin=1.6 * cm, bottomMargin=1.6 * cm,
        title=title,
        author=author,
        subject="Simulación bioeconómica del hato",
        creator=author,
    )
    styles = _styles()
    story: List = []

    # Cabecera
    story.append(Paragraph(title, styles["title"]))
    story.append(Paragraph(
        f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')} · "
        f"Iteraciones: {sim.n_iter:,} · "
        f"Semilla: {sim.seed if sim.seed is not None else 'aleatoria'}",
        styles["subtitle"],
    ))

    # Resumen ejecutivo / KPIs
    story.append(Paragraph("KPIs del margen bruto anual del hato (cifras en COP)", styles["h2"]))
    margin = sim.outputs["Margen bruto COP"].values
    kpis = compute_kpis(margin)
    story.append(_kpi_table(kpis, styles))
    story.append(Spacer(1, 0.4 * cm))

    # Histograma
    img_bytes = _make_histogram(margin, kpis)
    img = Image(io.BytesIO(img_bytes), width=17 * cm, height=9 * cm)
    story.append(img)
    story.append(Spacer(1, 0.2 * cm))

    # Interpretación de riesgo
    story.append(Paragraph("Interpretación para toma de decisiones", styles["h2"]))
    story.append(Paragraph(_interpretation(kpis), styles["body"]))
    if extra_notes:
        story.append(Spacer(1, 0.2 * cm))
        story.append(Paragraph(extra_notes, styles["small"]))

    story.append(PageBreak())

    # Distribuciones usadas
    story.append(Paragraph("Supuestos y distribuciones de probabilidad", styles["h2"]))
    story.append(Paragraph(
        "Cada parámetro de entrada se muestreó mediante el método Monte Carlo. "
        "Las variables marcadas como 'Constante' se mantuvieron fijas en su valor base.",
        styles["body"],
    ))
    story.append(Spacer(1, 0.2 * cm))
    story.append(_distributions_table(specs_used))

    story.append(PageBreak())

    # Resumen estadístico por variable de salida
    story.append(Paragraph("Resumen estadístico por variable de salida", styles["h2"]))
    story.append(Paragraph(
        "Estadísticos descriptivos de las 25 variables del modelo bioeconómico simuladas "
        "sobre la totalidad de iteraciones. Las cifras económicas (IT, CT, PSOL, Margen bruto) "
        "se reportan en <b>millones de COP</b> para facilitar la lectura; las demás variables están en "
        "sus unidades originales (cabezas, kg, %, meses, días).",
        styles["body"],
    ))
    story.append(Spacer(1, 0.2 * cm))
    story.append(_outputs_summary_table(sim.summary_outputs(), styles))

    story.append(Spacer(1, 0.6 * cm))
    nota_method = (
        "Nota metodológica: el modelo es una adaptación simplificada del modelo estocástico "
        "de Vargas-Leitón &amp; Cuevas-Abrego (Stochastic model to estimate economic values of "
        "production and functional traits in dairy cattle). Se utilizan distribuciones de probabilidad "
        "para representar la incertidumbre de los rasgos productivos y económicos; la simulación "
        "se ejecuta con muestreo Monte Carlo (no estratificado por hipercubo latino). "
        "Las fórmulas estructurales para edad al primer servicio (Gompertz inverso), curva de lactancia "
        "(Wood), reducción por mastitis y agregados de ingresos y costos siguen la hoja "
        "&quot;Variables salida&quot; del modelo bioeconómico de la finca."
    )
    story.append(Paragraph(nota_method, styles["small"]))

    if sim.warnings:
        story.append(Spacer(1, 0.4 * cm))
        story.append(Paragraph("Advertencias de ejecución", styles["h2"]))
        for w in sim.warnings:
            story.append(Paragraph(f"• {w}", styles["small"]))

    doc.build(story)
    buf.seek(0)
    return buf.read()
