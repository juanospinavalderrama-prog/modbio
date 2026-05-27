# 🌿 Dashboard Bioeconómico — Finca Las Margaritas

Aplicación Streamlit para Juan Pablo Ospina que ejecuta una simulación Monte
Carlo del margen bruto anual del hato de la Finca Las Margaritas, basada en el
modelo bioeconómico de la finca y en la metodología propuesta por
Vargas-Leitón &amp; Cuevas-Abrego (*Stochastic model to estimate economic
values of production and functional traits in dairy cattle*, 2023).

---

## ✨ Funcionalidad

- 📂 **Carga del Excel del modelo** (`Finca-Las-Margaritas-Modelo-Bioeconomico.xlsx`)
  o uso del archivo por defecto incluido en `data/`.
- 🎛️ **Configuración interactiva** de la distribución de probabilidad de cada
  parámetro: Normal (μ, σ), Triangular (mín, moda, máx), Uniforme (mín, máx)
  o Constante.
- 🎲 **Simulación Monte Carlo de 5,000 iteraciones por defecto** (configurable),
  con semilla opcional para reproducibilidad.
- 📊 **KPIs financieros**: media, mediana, desviación estándar, coeficiente de
  variación, percentiles 2.5 y 97.5, intervalo de confianza al 95% y
  probabilidad de margen negativo.
- 📈 **Histograma interactivo** del margen bruto con línea de media e intervalo
  95% (Plotly).
- 📋 **Tabla resumen** por cada una de las 25 variables de salida del modelo.
- ⬇️ **Exportación** de resultados en CSV (entradas + salidas) y un PDF con
  resumen estadístico listo para toma de decisiones. En el PDF, las cifras
  económicas (IT, CT, PSOL, Margen bruto) se reportan en **millones de COP**
  para mantener la legibilidad y evitar desbordes de celda; la nota al pie de
  la tabla lo explica al lector.

## 🏗 Arquitectura

```
las_margaritas_streamlit/
├── app.py                # Aplicación Streamlit (UI + orquestación)
├── simulation.py         # Motor de simulación Monte Carlo + parámetros
├── pdf_report.py         # Construcción del PDF de resumen (ReportLab)
├── requirements.txt
├── README.md
└── data/
    └── Finca-Las-Margaritas-Modelo-Bioeconomico.xlsx
```

## 🚀 Instalación

```bash
cd las_margaritas_streamlit
python -m venv .venv && source .venv/bin/activate     # opcional
pip install -r requirements.txt
```

## ▶️ Ejecución

```bash
streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```

Luego abre tu navegador en <http://localhost:8501>.

## 📖 Flujo de uso

1. **Carga el Excel** desde la barra lateral (o usa el archivo por defecto).
2. **Revisa y ajusta** las distribuciones en la pestaña *1 · Parámetros y
   distribuciones*. Las distribuciones se pre-llenan con los valores base y
   las dispersiones de la hoja `Parámetros base` del Excel.
3. **Configura** el número de iteraciones (5,000 por defecto) y opcionalmente
   fija una semilla en la barra lateral.
4. Pulsa **Ejecutar simulación Monte Carlo**. Los resultados aparecen en la
   pestaña *2 · Resultados Monte Carlo* con KPIs e histograma interactivo.
5. Explora cada variable individual en *3 · Variables de salida*.
6. **Descarga** los CSV de entradas y salidas y el **PDF de resumen**
   estadístico para llevarlo a la mesa de decisiones.

## 🧪 Notas metodológicas

### Modelo simulado

Cada iteración Monte Carlo:

1. Muestrea cada parámetro de entrada desde su distribución.
2. Calcula la edad al primer servicio mediante el **Gompertz inverso**.
3. Genera la **curva de lactancia de Wood** (`y(t) = a·t^b·exp(−c·t)`) y suma
   sobre `t = 1…DLAC` para obtener la producción anual por vaca.
4. Aplica la **reducción por mastitis clínica** y calcula sólidos por
   componente (grasa, proteína, lactosa+minerales).
5. Calcula ingresos por leche, carne y novillas vendidas; costos por
   concentrado, vacas adultas, reemplazo y administración.
6. Devuelve el **margen bruto = IT − CT** en COP.

### Diferencias frente al artículo de Vargas-Leitón

- **Muestreo Monte Carlo simple** en vez de Hipercubo Latino. El número de
  iteraciones (5,000) sigue la recomendación del artículo y produce KPIs
  estables a este nivel.
- No se calculan los **valores económicos parciales** (Δmargen por unidad de
  cambio en cada rasgo) mediante regresión: el foco está en el riesgo
  agregado del margen bruto, según solicitó el usuario.
- Se utiliza el valor de **`c` de Wood ≈ 0.00486** (hoja `Curva`), no el
  valor inconsistente reportado en `Parámetros base` (≈0.034). Esta
  corrección permite reproducir el PLn ≈ 4,663 kg/vaca-año reportado en la
  hoja `Variables salida`.

### Validación

Con los parámetros base del Excel, la simulación reproduce los siguientes
valores promedio de la hoja `Variables salida` (tolerancia estocástica):

| Variable | Excel | Simulado (n=5,000) |
|----------|------:|-------------------:|
| VL       | 54.4  | 54.4 |
| HR       | 23.2  | 23.2 |
| EIA      | 14.8  | 14.9 |
| EPA      | 24.7  | 24.8 |
| PLn      | 4,663 | 4,683 |
| PP       | 25.5  | 25.5 |
| IT       | 1,131 M COP | ~1,074 M COP |
| CT       | 419 M COP   | ~381 M COP |
| Margen   | 712 M COP   | ~693 M COP |

Las diferencias residuales en valores económicos provienen del término
no-lineal `Wood·VL·precio` cuando se introduce ruido estocástico.

### Manejo de errores

- Iteraciones con dominio inválido para el Gompertz inverso
  (p. ej. `PIA ≥ Ag`) se marcan como `NaN` y se excluyen de los KPIs.
- Dispersiones no numéricas o negativas se corrigen a cero con advertencia.
- Distribuciones triangulares con `min > moda > max` mal definidas se
  degeneran al valor moda.
- Valores negativos imposibles (números de animales, producción) se acotan
  a cero.

## 📝 Atribuciones

- Modelo bioeconómico de la Finca Las Margaritas: Juan Pablo Ospina.
- Marco metodológico: B. Vargas-Leitón &amp; J. Cuevas-Abrego, *Stochastic
  model to estimate economic values of production and functional traits in
  dairy cattle*, 2023.
- Implementación: Perplexity Computer.
