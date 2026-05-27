"""
Motor de simulación bioeconómica estocástica para la Finca Las Margaritas.

Implementa un modelo simplificado tipo Vargas-Leitón & Cuevas-Abrego (2023):
- Muestreo Monte Carlo de variables productivas, reproductivas, sanitarias y económicas.
- Curva de lactancia de Wood: y(t) = a * t^b * exp(-c * t).
- Edad al primer servicio derivada de Gompertz inverso.
- Cálculo de ingresos, costos y margen bruto del hato en un año calendario.

Las fórmulas siguen la hoja "Variables salida" del Excel del usuario.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# Modelos de distribución
# ---------------------------------------------------------------------------

DIST_NORMAL = "Normal"
DIST_TRIANGULAR = "Triangular"
DIST_UNIFORME = "Uniforme"
DIST_CONSTANTE = "Constante"

ALL_DISTS = [DIST_NORMAL, DIST_TRIANGULAR, DIST_UNIFORME, DIST_CONSTANTE]


@dataclass
class ParamSpec:
    """Especificación de un parámetro de entrada."""
    code: str                    # Código corto (e.g. "VA", "Ag", "$KGP")
    label: str                   # Nombre descriptivo
    unit: str                    # Unidad
    distribution: str            # Normal/Triangular/Uniforme/Constante
    # Parámetros de la distribución
    mean: float = 0.0            # Para Normal
    sd: float = 0.0              # Para Normal
    tri_min: float = 0.0         # Para Triangular
    tri_mode: float = 0.0
    tri_max: float = 0.0
    uni_min: float = 0.0         # Para Uniforme
    uni_max: float = 0.0
    constant_value: float = 0.0  # Para Constante
    # Metadatos heredados del Excel
    base: float = 0.0
    dispersion: Any = None
    group: str = ""              # Bloque (Crecimiento, Hato, etc.)
    allow_negative: bool = False
    min_floor: Optional[float] = None  # piso de validez (p.ej. 0 para %, 1 para días)


def sample_parameter(spec: ParamSpec, n: int, rng: np.random.Generator) -> np.ndarray:
    """Genera n muestras de la distribución del parámetro."""
    if spec.distribution == DIST_CONSTANTE:
        return np.full(n, float(spec.constant_value))

    if spec.distribution == DIST_NORMAL:
        sd = max(float(spec.sd), 0.0)
        samples = rng.normal(loc=float(spec.mean), scale=sd, size=n)

    elif spec.distribution == DIST_TRIANGULAR:
        lo = float(spec.tri_min)
        mo = float(spec.tri_mode)
        hi = float(spec.tri_max)
        if not (lo <= mo <= hi) or hi <= lo:
            # Fallback: degeneramos a constante en la moda
            return np.full(n, mo)
        samples = rng.triangular(lo, mo, hi, size=n)

    elif spec.distribution == DIST_UNIFORME:
        lo = float(spec.uni_min)
        hi = float(spec.uni_max)
        if hi <= lo:
            return np.full(n, lo)
        samples = rng.uniform(lo, hi, size=n)

    else:
        raise ValueError(f"Distribución no soportada: {spec.distribution}")

    # Validación: piso de no-negatividad
    if not spec.allow_negative:
        samples = np.maximum(samples, 0.0)
    if spec.min_floor is not None:
        samples = np.maximum(samples, spec.min_floor)
    return samples


# ---------------------------------------------------------------------------
# Especificaciones por defecto derivadas del Excel "Finca Las Margaritas"
# ---------------------------------------------------------------------------

def _spec_from_row(code: str, label: str, unit: str, base: float, dispersion: Any,
                   distribution_hint: str, group: str,
                   allow_negative: bool = False,
                   min_floor: Optional[float] = 0.0) -> ParamSpec:
    """Construye un ParamSpec a partir de filas tipo 'Parámetros base'."""
    try:
        disp = float(dispersion)
    except (TypeError, ValueError):
        disp = 0.0

    # Si en el Excel viene Constante/Costante o Discreta/Extremo/Exponencial,
    # lo tratamos por defecto como Normal con SD razonable. El usuario podrá cambiar
    # la distribución desde la UI.
    hint = (distribution_hint or "").strip().lower()
    if hint in {"costante", "constante"} or disp == 0.0:
        dist = DIST_CONSTANTE
    else:
        dist = DIST_NORMAL

    spec = ParamSpec(
        code=code,
        label=label,
        unit=unit,
        distribution=dist,
        mean=float(base),
        sd=disp,
        tri_min=float(base) - 2.0 * disp,
        tri_mode=float(base),
        tri_max=float(base) + 2.0 * disp,
        uni_min=float(base) - 2.0 * disp,
        uni_max=float(base) + 2.0 * disp,
        constant_value=float(base),
        base=float(base),
        dispersion=dispersion,
        group=group,
        allow_negative=allow_negative,
        min_floor=min_floor,
    )
    return spec


def default_param_specs() -> Dict[str, ParamSpec]:
    """Catálogo por defecto de parámetros (los valores serán refrescados
    desde el Excel cuando esté disponible)."""
    specs: Dict[str, ParamSpec] = {}

    def add(code, label, unit, base, disp, hint, group, **kwargs):
        specs[code] = _spec_from_row(code, label, unit, base, disp, hint, group, **kwargs)

    # Novillas de reemplazo / Crecimiento
    add("Ag", "a (Gompertz)", "parámetro", 487.4, 9.6367, "Normal", "Crecimiento")
    add("Bg", "b (Gompertz)", "parámetro", 1.10007, 0.04818, "Normal", "Crecimiento")
    add("Cg", "c (Gompertz)", "parámetro", 0.11009, 0.004818, "Normal", "Crecimiento")
    add("PIA", "Peso a 1er servicio", "kg", 270, 9.651, "Normal", "Crecimiento")
    add("IAC", "Intervalo 1er servicio-concepción", "d", 21, 5, "Normal", "Reproducción novillas")
    add("TDIr", "Desecho involuntario (reemplazo)", "%", 12, 2.5, "Normal", "Reemplazo")
    add("TDVr", "Desecho voluntario (reemplazo)", "%", 50, 10, "Normal", "Reemplazo")

    # Hato en producción
    add("VA", "Vacas adultas", "n", 68, 0, "Constante", "Hato", min_floor=1)
    add("TDIv", "Desecho involuntario adulto", "%", 18, 2, "Normal", "Reemplazo")
    add("TMv", "Mortalidad adulta", "%", 2, 0.5, "Normal", "Reemplazo")
    add("DA", "Días abiertos", "d", 138, 24.80, "Normal", "Reproducción", min_floor=1)
    add("DP", "Duración de preñez", "d", 282, 2.995, "Normal", "Reproducción", min_floor=200)
    add("RC", "Rendimiento en canal", "%", 48, 1, "Normal", "Carne")

    # Curva de Wood (producción de leche)
    add("a", "a (Wood)", "parámetro", 18.84999, 4.4668, "Normal", "Producción", min_floor=0.1)
    add("b", "b (Wood)", "parámetro", 0.132, 0.005, "Normal", "Producción", min_floor=0.01)
    # Nota metodológica: la hoja "Parámetros base" del Excel reporta c=0.034 que es
    # inconsistente con la curva de lactancia simulada (PLn≈4663). El valor real
    # de c en la hoja "Curva" es ~0.00486. Usamos este último como base por defecto.
    add("c", "c (Wood)", "parámetro", 0.00486, 0.000764, "Normal", "Producción", min_floor=0.001)
    add("DLAC", "Longitud de lactancia", "d", 305, 39.72, "Normal", "Producción", min_floor=60)

    # Composición de leche
    add("%GR", "Grasa en leche", "%", 4.3203, 0.6450, "Normal", "Composición")
    add("%PR", "Proteína en leche", "%", 3.5004, 0.27798, "Normal", "Composición")
    add("%LM", "Lactosa + minerales", "%", 5.45, 0, "Constante", "Composición")

    # Salud
    add("IMC", "Incidencia mastitis clínica", "%", 15, 3, "Normal", "Salud")
    add("FactorRM", "Factor reducción por mastitis", "·", 0.095, 0, "Constante", "Salud")

    # Económicos
    add("$KGP", "Precio grasa/proteína", "COP/kg", 42300, 95.44, "Normal", "Económicos")
    add("$LM", "Precio lactosa+minerales", "COP/kg", 15000, 73, "Normal", "Económicos")
    add("$CAR", "Precio carne en canal", "COP/kg", 20000, 73, "Normal", "Económicos")
    add("$NOV", "Precio novilla preñada", "COP/cab", 3_700_000, 180_000, "Normal", "Económicos")
    add("$CON", "Costo concentrado", "COP/kg", 2025, 73, "Normal", "Económicos")
    add("$REM", "Costo novilla de reemplazo", "COP/cab", 2_270_000, 220_000, "Normal", "Económicos")
    add("$VAC", "Costo fijo por vaca adulta", "COP", 1_500_000, 110_000, "Normal", "Económicos")
    add("$AD", "Costos por administración", "COP", 18_000_000, 0, "Constante", "Económicos")

    # Parámetros operativos adicionales (constantes pero ajustables)
    add("PesoNac", "Peso al nacimiento", "kg", 32, 0, "Constante", "Crecimiento", min_floor=10)
    add("ConvAlim", "Conversión kg leche/kg concentrado", "·", 3.0, 0, "Constante", "Alimentación", min_floor=0.5)

    return specs


def load_specs_from_excel(xlsx_path: str) -> Dict[str, ParamSpec]:
    """Carga y refresca los parámetros desde la hoja 'Parámetros base' del Excel."""
    specs = default_param_specs()
    try:
        df = pd.read_excel(xlsx_path, sheet_name="Parámetros base", header=2)
    except Exception:
        return specs

    df.columns = [str(c).strip() for c in df.columns]
    expected = {"Parámetros", "Unidades", "Código", "Distribución", "Base", "Dispersión"}
    if not expected.issubset(set(df.columns)):
        return specs

    df = df.dropna(subset=["Código"])
    for _, row in df.iterrows():
        code = str(row["Código"]).strip()
        if not code or code.lower() in {"nan"}:
            continue
        base = row["Base"]
        disp = row["Dispersión"]
        try:
            base_f = float(base)
        except (TypeError, ValueError):
            continue
        try:
            disp_f = float(disp)
        except (TypeError, ValueError):
            disp_f = 0.0
        label = str(row["Parámetros"]).strip()
        unit = "" if pd.isna(row["Unidades"]) else str(row["Unidades"]).strip()
        dist_hint = "" if pd.isna(row["Distribución"]) else str(row["Distribución"]).strip()

        if code in specs:
            # Override metodológico: c (Wood) en "Parámetros base" es inconsistente
            # con la curva real (≈0.034 vs 0.00486 de la hoja "Curva"). Mantenemos
            # el valor por defecto correcto y solo refrescamos metadatos.
            if code == "c":
                s = specs[code]
                s.label = label or s.label
                s.unit = unit or s.unit
                s.dispersion = disp
                continue
            s = specs[code]
            s.label = label or s.label
            s.unit = unit or s.unit
            s.base = base_f
            s.dispersion = disp
            s.mean = base_f
            s.sd = disp_f
            s.tri_min = base_f - 2 * disp_f
            s.tri_mode = base_f
            s.tri_max = base_f + 2 * disp_f
            s.uni_min = base_f - 2 * disp_f
            s.uni_max = base_f + 2 * disp_f
            s.constant_value = base_f
            # Reajustar distribución: si dispersión es 0 -> Constante
            hint_low = dist_hint.lower()
            if hint_low in {"costante", "constante"} or disp_f == 0.0:
                s.distribution = DIST_CONSTANTE
            else:
                s.distribution = DIST_NORMAL

    # Intentar refrescar c (Wood) desde la hoja "Curva" si está disponible
    try:
        curva_df = pd.read_excel(xlsx_path, sheet_name="Curva ", header=None)
        for _, r in curva_df.iterrows():
            label = str(r.iloc[0]).strip() if pd.notna(r.iloc[0]) else ""
            if label.lower().startswith("c ="):
                val = r.iloc[1]
                try:
                    c_val = float(val)
                    s = specs["c"]
                    s.base = c_val
                    s.mean = c_val
                    s.constant_value = c_val
                    s.tri_min = c_val - 2 * s.sd
                    s.tri_mode = c_val
                    s.tri_max = c_val + 2 * s.sd
                    s.uni_min = c_val - 2 * s.sd
                    s.uni_max = c_val + 2 * s.sd
                except (TypeError, ValueError):
                    pass
                break
    except Exception:
        pass

    return specs


# ---------------------------------------------------------------------------
# Modelo bioeconómico
# ---------------------------------------------------------------------------

@dataclass
class SimulationResult:
    n_iter: int
    seed: Optional[int]
    inputs: pd.DataFrame                  # muestras de variables de entrada
    outputs: pd.DataFrame                 # variables de salida calculadas
    warnings: List[str] = field(default_factory=list)

    @property
    def margin(self) -> np.ndarray:
        return self.outputs["Margen bruto COP"].values

    def summary_outputs(self) -> pd.DataFrame:
        """Tabla resumen por variable de salida."""
        df = self.outputs.copy()
        rows = []
        for col in df.columns:
            x = pd.to_numeric(df[col], errors="coerce").dropna().values
            if x.size == 0:
                rows.append({"Variable": col, "Media": np.nan, "DE": np.nan,
                             "Mediana": np.nan, "p2.5": np.nan, "p97.5": np.nan,
                             "Min": np.nan, "Max": np.nan, "n": 0})
                continue
            rows.append({
                "Variable": col,
                "Media": float(np.mean(x)),
                "DE": float(np.std(x, ddof=1)) if x.size > 1 else 0.0,
                "Mediana": float(np.median(x)),
                "p2.5": float(np.percentile(x, 2.5)),
                "p97.5": float(np.percentile(x, 97.5)),
                "Min": float(np.min(x)),
                "Max": float(np.max(x)),
                "n": int(x.size),
            })
        return pd.DataFrame(rows)


def _wood_curve_total(a: np.ndarray, b: np.ndarray, c: np.ndarray,
                      dlac: np.ndarray) -> np.ndarray:
    """Calcula la suma de la curva de Wood para t=1..DLAC para cada iteración.

    Vectorizado: usa el DLAC máximo como longitud común y enmascara t > DLAC_i.
    Si los parámetros producen valores no finitos, se devuelve NaN para esa iteración.
    """
    n = a.shape[0]
    dlac_int = np.clip(np.rint(dlac).astype(int), 1, 1000)
    t_max = int(dlac_int.max())
    t = np.arange(1, t_max + 1, dtype=float).reshape(1, t_max)  # (1, T)
    a_col = a.reshape(n, 1)
    b_col = b.reshape(n, 1)
    c_col = c.reshape(n, 1)
    with np.errstate(over="ignore", invalid="ignore", divide="ignore"):
        y = a_col * np.power(t, b_col) * np.exp(-c_col * t)
    mask = t <= dlac_int.reshape(n, 1)
    y = np.where(mask, y, 0.0)
    total = np.nansum(np.where(np.isfinite(y), y, 0.0), axis=1)
    return total


def _wood_peak(a: np.ndarray, b: np.ndarray, c: np.ndarray) -> np.ndarray:
    """Producción al pico = a*(b/c)^b * exp(-b). Devuelve NaN si c<=0."""
    with np.errstate(invalid="ignore", divide="ignore"):
        ratio = np.where(c > 0, b / c, np.nan)
        peak = a * np.power(ratio, b) * np.exp(-b)
    return peak


def _eia_gompertz(pia: np.ndarray, ag: np.ndarray, bg: np.ndarray,
                  cg: np.ndarray) -> np.ndarray:
    """Edad al primer servicio: (ln(ln(PIA/Ag)/-1) - Bg) / (-Cg).

    Dominios: PIA/Ag < 1 y ln(PIA/Ag) < 0 para que ln(...)/-1 > 0.
    Si falla el dominio, devuelve NaN para excluir esa iteración.
    """
    with np.errstate(invalid="ignore", divide="ignore"):
        ratio = np.where(ag > 0, pia / ag, np.nan)
        log_ratio = np.log(np.where(ratio > 0, ratio, np.nan))
        inner = -log_ratio  # debe ser > 0 (ratio < 1)
        log_inner = np.log(np.where(inner > 0, inner, np.nan))
        eia = (log_inner - bg) / np.where(cg != 0, -cg, np.nan)
    return eia


def run_simulation(specs: Dict[str, ParamSpec],
                   n_iter: int = 5000,
                   seed: Optional[int] = None) -> SimulationResult:
    """Ejecuta la simulación Monte Carlo.

    Calcula todas las variables de salida solicitadas:
    VL, HR, VDI, VDV, NE, NV, TR, EIA, GDP, EPA, VP, PLn, PP, DS,
    RM, PLsm, PGR, PPR, PLM, PSOL, PCAR, CON, IT, CT, Margen bruto COP.
    """
    rng = np.random.default_rng(seed)
    warnings: List[str] = []

    # Muestreo de entradas
    samples: Dict[str, np.ndarray] = {}
    for code, spec in specs.items():
        try:
            samples[code] = sample_parameter(spec, n_iter, rng)
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"No se pudo muestrear {code}: {exc}. Se usa la base como constante.")
            samples[code] = np.full(n_iter, spec.base)

    def s(code: str, default: float = 0.0) -> np.ndarray:
        if code in samples:
            return samples[code]
        warnings.append(f"Falta parámetro {code}; se usa valor por defecto {default}.")
        return np.full(n_iter, default)

    # --- Demografía del hato ---
    VA = s("VA")
    TDIv = s("TDIv")
    TMv = s("TMv")
    TDIr = s("TDIr")
    TDVr = s("TDVr")

    VL = VA * (1 - TDIv / 100.0) - VA * TMv / 100.0
    VDI = VA * ((TDIv + TMv) / 100.0)

    # --- Reproducción y crecimiento de novillas ---
    PIA = s("PIA")
    Ag = s("Ag")
    Bg = s("Bg")
    Cg = s("Cg")
    IAC = s("IAC")
    DA = s("DA")
    DP = s("DP")
    PesoNac = s("PesoNac", 32.0)

    EIA = _eia_gompertz(PIA, Ag, Bg, Cg)
    EPA = EIA + (IAC + DP) / 30.4
    with np.errstate(invalid="ignore", divide="ignore"):
        HR = np.where(EPA > 0, VL * 0.5 * (1 - TDIr / 100.0) * 24.0 / EPA, np.nan)
        TR = np.where(VA > 0, HR / VA * 100.0, np.nan)
        VP = np.where(TR > 0, 100.0 / TR * 12.0, np.nan)
        GDP = np.where(EIA > 0, (PIA - PesoNac) / (EIA * 30.4), np.nan)

    NE = np.maximum(HR - VDI, 0.0)
    NV = NE * TDVr / 100.0
    VDV = np.maximum(NE - NV, 0.0)

    # --- Producción de leche (curva de Wood) ---
    a = s("a")
    b = s("b")
    c = s("c")
    DLAC = s("DLAC")
    wood_total = _wood_curve_total(a, b, c, DLAC)
    PP = _wood_peak(a, b, c)

    with np.errstate(invalid="ignore", divide="ignore"):
        denom = DA + DP
        PLn = np.where(denom > 0, wood_total / denom * 365.0, np.nan)

    DS = (DP + DA) - DLAC

    # --- Reducción por mastitis ---
    IMC = s("IMC")
    FactorRM = s("FactorRM", 0.095)
    RM = PLn * (IMC / 100.0) * FactorRM

    PLsm = PLn - RM

    # --- Composición ---
    GR = s("%GR")
    PR = s("%PR")
    LM = s("%LM")
    PGR = PLsm * GR / 100.0
    PPR = PLsm * PR / 100.0
    PLM = PLsm * LM / 100.0

    PSOL = (PGR + PPR + PLM) * VL

    # --- Carne ---
    RC = s("RC")
    PCAR = (Ag * RC / 100.0) * (NE - NV)
    PCAR = np.maximum(PCAR, 0.0)

    # --- Consumo de concentrado ---
    ConvAlim = s("ConvAlim", 3.0)
    with np.errstate(invalid="ignore", divide="ignore"):
        leche_anual_vaca = np.where(denom > 0, wood_total / denom * 365.0, np.nan)
        CON_vaca = np.where(ConvAlim > 0, leche_anual_vaca / ConvAlim, np.nan)
        CON = CON_vaca * VL  # kg hato/año

    # --- Economía ---
    KGP = s("$KGP")
    LMprice = s("$LM")
    CAR = s("$CAR")
    NOV = s("$NOV")
    CONprice = s("$CON")
    REM = s("$REM")
    VACprice = s("$VAC")
    AD = s("$AD")

    IT = (PGR + PPR) * VL * KGP + PLM * VL * LMprice + PCAR * CAR + NV * NOV
    CT = CONprice * (CON + CON / 3.0) + VL * VACprice + HR * REM + AD
    MB = IT - CT

    outputs = pd.DataFrame({
        "VL": VL,
        "HR": HR,
        "VDI": VDI,
        "VDV": VDV,
        "NE": NE,
        "NV": NV,
        "TR": TR,
        "EIA": EIA,
        "GDP": GDP,
        "EPA": EPA,
        "VP": VP,
        "PLn": PLn,
        "PP": PP,
        "DS": DS,
        "RM": RM,
        "PLsm": PLsm,
        "PGR": PGR,
        "PPR": PPR,
        "PLM": PLM,
        "PSOL": PSOL,
        "PCAR": PCAR,
        "CON": CON,
        "IT": IT,
        "CT": CT,
        "Margen bruto COP": MB,
    })

    # Diagnóstico de NaN
    nan_iter = outputs["Margen bruto COP"].isna().sum()
    if nan_iter > 0:
        warnings.append(
            f"Se descartaron {nan_iter} iteraciones con dominio inválido "
            f"(EIA logarítmico, denominadores cero, etc.)."
        )

    inputs_df = pd.DataFrame({k: v for k, v in samples.items()})

    return SimulationResult(
        n_iter=n_iter,
        seed=seed,
        inputs=inputs_df,
        outputs=outputs,
        warnings=warnings,
    )


# ---------------------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------------------

def compute_kpis(margin: np.ndarray) -> Dict[str, float]:
    x = pd.to_numeric(pd.Series(margin), errors="coerce").dropna().values
    if x.size == 0:
        return {k: float("nan") for k in
                ["media", "mediana", "sd", "p2.5", "p97.5", "prob_neg", "cv", "n"]}
    mean = float(np.mean(x))
    sd = float(np.std(x, ddof=1)) if x.size > 1 else 0.0
    return {
        "media": mean,
        "mediana": float(np.median(x)),
        "sd": sd,
        "p2.5": float(np.percentile(x, 2.5)),
        "p97.5": float(np.percentile(x, 97.5)),
        "prob_neg": float(np.mean(x < 0)),
        "cv": float(sd / mean) if mean != 0 else float("nan"),
        "n": int(x.size),
    }
