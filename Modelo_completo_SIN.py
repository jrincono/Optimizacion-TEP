# -*- coding: utf-8 -*-
"""
=============================================================================
TEP — COSTO DIFERIDO POR COMUNIDADES ENERGÉTICAS BTM (Behind-the-Meter)
Backbone Equivalente STN Colombia 500 kV — MILP DC (Pyomo + Gurobi)
=============================================================================

CONCEPTO CENTRAL:
    Las comunidades energéticas BTM actúan como un RECURSO VIRTUAL DE
    EXPANSIÓN al reducir la demanda neta del sistema. Esta reducción
    disminuye o elimina la necesidad de construir nuevas líneas, difiriendo
    o evitando inversiones en transmisión.

    Formulación de reducción de demanda:
        D_nuevo[i] = D_base[i] × (1 − α[i])

    Donde α[i] es la penetración efectiva de comunidades energéticas en el
    nodo i, que puede ser:
        a) Uniforme:   α[i] = α_total  (igual para todos los nodos)
        b) Inteligente (CEI): α[i] = α_total × (CEI[i] / CEI_mean)
           — prioriza nodos con mayor potencial solar, mayor estrés de red
             y mayor necesidad de refuerzo

INDICADORES NODALES (adaptados al SIN Colombia):
    SPI — Solar Potential Index: potencial fotovoltaico de cada área
    NSI — Network Stress Index:  déficit generación/demanda normalizado
    GRI — Grid Reinforcement Index: carga relativa de corredores adyacentes
    CEI — Community Energy Index: índice compuesto (0.4·SPI + 0.3·NSI + 0.3·GRI)

ESCENARIOS ANALIZADOS:
    Escenario base futuro 2039 (alta demanda, congestión severa):
      Oriental: 4 200 MW | Caribe: 3 800 MW | Suroccidental: 3 600 MW

    Penetraciones BTM evaluadas: α = 0%, 5%, 10%, 15%, 20%, 25%, 30%, 35%
    Estrategias:
      • Reducción uniforme
      • Reducción inteligente basada en índice CEI

SALIDAS:
    - Inversión óptima por escenario
    - Líneas construidas / evitadas
    - Costo evitado y diferido vs. caso base
    - Ranking CEI de nodos
    - Heatmaps de cargabilidad de corredores
    - Curvas de ahorro acumulado
    - Comparación uniforme vs. CEI inteligente

FUENTES:
    - DIgSILENT PowerFactory 2025 | Proyecto Uniandes GCM
    - UPME PIEG 2025-2039 | UPME PET 2022-2036
    - XM (2024) informes operación anual
    - Atlas de Radiación Solar Colombia — IDEAM 2014

Tesis Ingeniería Eléctrica — Uniandes 2026
=============================================================================
"""

import pyomo.environ as pyo
from pyomo.opt import SolverFactory
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.colors as mcolors
from matplotlib.gridspec import GridSpec
import numpy as np
import pandas as pd
import time
import warnings
warnings.filterwarnings('ignore')

# =============================================================================
# CONFIGURACIÓN GLOBAL
# =============================================================================

GUARDAR_GRAFICAS     = True

# ── Configuración de exportación de resultados ───────────────────────────────
import os
OUTPUT_DIR    = "resultados_tesis"
DIR_FIGURAS   = os.path.join(OUTPUT_DIR, "figuras")
DIR_TABLAS    = os.path.join(OUTPUT_DIR, "tablas")
DIR_REPORTES  = os.path.join(OUTPUT_DIR, "reportes")
for _d in (OUTPUT_DIR, DIR_FIGURAS, DIR_TABLAS, DIR_REPORTES):
    os.makedirs(_d, exist_ok=True)

# Archivo Excel maestro consolidado (un solo libro con varias hojas)
EXCEL_MASTER  = os.path.join(DIR_REPORTES, "resultados_tesis_consolidado.xlsx")

def fpath_fig(name): return os.path.join(DIR_FIGURAS, name)
def fpath_tbl(name): return os.path.join(DIR_TABLAS,  name)
def fpath_rep(name): return os.path.join(DIR_REPORTES, name)

# Acumulador de hojas para el Excel maestro (nombre_hoja -> DataFrame)
EXCEL_SHEETS  = {}
def add_to_excel(sheet_name, df):
    """Acumula un DataFrame para escribirlo al Excel maestro al final."""
    EXCEL_SHEETS[sheet_name[:31]] = df  # Excel limita nombres de hoja a 31 chars

CORRER_MULTIPERIODO  = True    # poner False para saltarse los modelos multiperiodo

# Penetraciones BTM a evaluar (fracción de reducción de demanda total)
ALPHAS = [0.00, 0.05, 0.10, 0.15, 0.20, 0.25, 0.30]
ALPHA_LABELS = {a: f"α={int(a*100):02d}%" for a in ALPHAS}

# Tasa de descuento y horizonte para análisis multiperiodo
TASA_DESCUENTO  = 0.10
NT              = 16
ANO_INI         = 2024
PERIODOS        = list(range(1, NT + 1))

# =============================================================================
# 1. DATOS DEL SISTEMA (SIN Colombia — Backbone 500 kV)
# =============================================================================

Sbase = 100  # MVA

NODOS  = [1, 2, 3, 4, 5, 6]
NOMBRE = {
    1: 'Caribe',
    2: 'Nordeste',
    3: 'Oriental',
    4: 'Antioquia',
    5: 'Suroccidental',
    6: 'La Guajira',
}

# ── Demanda base 2024 [MW] — DIgSILENT PowerFactory 2025 ─────────────────────
D_2024 = {1: 2829.8, 2: 1178.9, 3: 2840.1, 4: 1442.6, 5: 2598.6, 6: 0.0} # Caso Base
# D_2024 = {1: 2829.8*2, 2: 1178.9*3, 3: 2840.1*2, 4: 1442.6*3, 5: 2598.6*3, 6: 0.0*3}
total_D_2024 = sum(D_2024.values())

print('D_2024',total_D_2024)
# ── Demanda futura 2039 — Escenario ALTO (UPME PIEG 2025-2039) ───────────────
# Oriental: +48% (electrificación transporte, data centers, Bogotá)
# Caribe:   +34% (industria petroquímica, electrificación costa)
# Suroccd:  +38% (Cali, minería verde, agroindustria)
# La Guajira: 0 (nodo exportador neto eólico)
# D_2039 = {1: 3800.0, 2: 1600.0, 3: 4200.0, 4: 2100.0, 5: 3600.0, 6: 0.0} # Caso Base
D_2039 = {1: 3800.0*2, 2: 1600.0*2, 3: 4200.0*2, 4: 2100.0*2, 5: 3600.0*2, 6: 0.0*2}
total_D_2039 = sum(D_2039.values())

print('D_2039',total_D_2039)

# ── Generación máxima 2039 [MW] ───────────────────────────────────────────────
# Antioquia: Ituango completo (2 400 MW hidroeléctrica) + expansión
# La Guajira: fase 2 eólica habilitada (~4 000 MW)
# G_2039 = {1: 2000.0, 2: 1400.0, 3: 2200.0, 4: 6000.0, 5: 2000.0, 6: 4000.0} # Caso Base
G_2039 = {1: 2000.0, 2: 1400.0, 3: 2200.0, 4: 6000.0*3, 5: 2000.0, 6: 4000.0*2}
total_G_2039 = sum(G_2039.values())

print('G_2039',total_G_2039)

# ── Generación 2024 para interpolación multiperiodo ──────────────────────────
G_2024 = {1: 3419.8, 2: 976.3, 3: 1452.0, 4: 3291.1, 5: 1368.3, 6: 1926.0} # Caso Base
# G_2024 = {1: 3419.8*4, 2: 976.3, 3: 1452.0, 4: 3291.1*5, 5: 1368.3, 6: 1926.0}
total_G_2024 = sum(G_2024.values())

print('G_2024',total_G_2024)

# ── Corredores ────────────────────────────────────────────────────────────────
CORREDORES = [
    '1,2', '1,3', '1,4', '1,5', '1,6',
    '2,3', '2,4', '2,5', '2,6',
    '3,4', '3,5', '3,6',
    '4,5', '4,6', '5,6',
]

DESC = {
    '1,2': 'La Loma–Sogamoso (Caribe↔Nordeste)',
    '1,3': 'Cerromatoso–Chinú (Caribe↔Oriental)',
    '1,4': 'San Carlos–Cerromatoso (Caribe↔Antioquia)',
    '1,5': 'Virginia–San Marcos (Caribe↔Suroccidental)',
    '1,6': 'Cuestecitas–Colectora (Caribe↔La Guajira)',
    '2,3': 'Sogamoso–Norte (Nordeste↔Oriental)',
    '2,4': 'Porce III–Sogamoso (Nordeste↔Antioquia)',
    '2,5': 'Nuevo corredor (Nordeste↔Suroccidental)',
    '2,6': 'La Loma–Colectora (Nordeste↔La Guajira)',
    '3,4': 'Nuevo Bogotá–Medellín (Oriental↔Antioquia)',
    '3,5': 'Tequendama–San Marcos (Oriental↔Suroccidental)',
    '3,6': 'Nuevo (Oriental↔La Guajira)',
    '4,5': 'Heliconia–La Virginia RSO (Antioquia↔Suroccidental)',
    '4,6': 'HVDC La Guajira–Interior (3000 MW)',
    '5,6': 'Nuevo (Suroccidental↔La Guajira)',
}

N0 = {
    '1,2': 2, '1,3': 1, '1,4': 2, '1,5': 1, '1,6': 1,
    '2,3': 1, '2,4': 1, '2,5': 0, '2,6': 1,
    '3,4': 0, '3,5': 1, '3,6': 0,
    '4,5': 1, '4,6': 0, '5,6': 0,
}

FMAX = {
    '1,2': 900,  '1,3': 900,  '1,4': 900,  '1,5': 900,  '1,6': 900,
    '2,3': 900,  '2,4': 900,  '2,5': 900,  '2,6': 900,
    '3,4': 900,  '3,5': 900,  '3,6': 900,
    '4,5': 1200, '4,6': 3000, '5,6': 900,
}

B_SIN = {
    '1,2': 111.9, '1,3': 241.0, '1,4': 88.7,  '1,5': 237.5, '1,6': 250.0,
    '2,3': 110.0, '2,4': 131.1, '2,5': 83.3,  '2,6': 166.7,
    '3,4': 111.1, '3,5': 123.3, '3,6': 125.0,
    '4,5': 171.5, '4,6': 200.0, '5,6': 90.0,
}

COSTO = {
    '1,2': 290,  '1,3': 140,  '1,4': 360,  '1,5': 155,  '1,6': 140,
    '2,3': 300,  '2,4': 250,  '2,5': 460,  '2,6': 280,
    '3,4': 340,  '3,5': 270,  '3,6': 380,
    '4,5': 215,  '4,6': 1950, '5,6': 380,
}

NCAND = 5
NREF  = 4     # Antioquia — mayor exportador neto
BIG_M = max(FMAX.values()) * 2.5 / Sbase

# ── Posiciones geográficas ─────────────────────────────────────────────────
POS = {
    1: (0.65, 0.82),   # Caribe
    2: (0.58, 0.55),   # Nordeste
    3: (0.52, 0.38),   # Oriental
    4: (0.32, 0.50),   # Antioquia
    5: (0.28, 0.28),   # Suroccidental
    6: (0.80, 0.95),   # La Guajira
}

COL_NODOS = {
    1: '#1E88E5', 2: '#FF8F00', 3: '#7B1FA2',
    4: '#2E7D32', 5: '#C62828', 6: '#F9A825',
}

# =============================================================================
# 2. INDICADORES NODALES — CEI (adaptados al SIN Colombia)
# =============================================================================
"""
Los indicadores CEI permiten identificar DÓNDE priorizar las comunidades
energéticas BTM para maximizar el beneficio en transmisión.

SPI — Solar Potential Index:
    Basado en irradiancia horizontal global promedio por región
    (IDEAM, Atlas de Radiación Solar Colombia 2014).
    Mayor SPI → mayor potencial FV → mayor autoabastecimiento.

NSI — Network Stress Index:
    NSI[i] = (D[i] - G[i]) / D[i]
    Mide la dependencia neta de importación de cada nodo.
    NSI > 0: nodo importador (alta dependencia de red) → prioridad alta
    NSI < 0: nodo exportador (gen. local excede demanda) → prioridad baja

GRI — Grid Reinforcement Index:
    Suma ponderada de la cargabilidad de los corredores adyacentes al nodo.
    Mayor GRI → más corredores saturados en el entorno → mayor presión.

CEI — Community Energy Index:
    CEI[i] = 0.4·SPI[i] + 0.3·NSI[i] + 0.3·GRI[i]
    Índice compuesto que prioriza nodos para instalación de comunidades BTM.
"""

# Irradiancia horizontal global promedio [kWh/m²/día] por área SIN
# Fuente: IDEAM Atlas de Radiación Solar Colombia (2014)
GHI_SIN = {
    1: 5.5,   # Caribe: costa caribeña, alta irradiancia
    2: 4.8,   # Nordeste: Santander, montañas, moderada
    3: 4.6,   # Oriental: sabana de Bogotá, nubosidad frecuente
    4: 4.3,   # Antioquia: Medellín, valles, baja-moderada
    5: 4.7,   # Suroccidental: Cali/Cauca, moderada
    6: 6.2,   # La Guajira: zona más soleada del país (hotspot solar-eólico)
}

def calcular_indices_CEI(demanda, gmax, flujos_aprox=None):
    """
    Calcula SPI, NSI, GRI y CEI para cada nodo del SIN Colombia.

    Parámetros:
        demanda     : dict {nodo: MW demanda}
        gmax        : dict {nodo: MW generación máxima}
        flujos_aprox: dict {corredor: MW flujo} (si None, se estiman)

    Retorna:
        DataFrame con columnas: Node, Nombre, Demand, Generation,
                                  GHI, SPI, NSI, GRI, CEI
    """
    # ── SPI: Solar Potential Index ────────────────────────────────────────────
    ghi_max = max(GHI_SIN.values())
    SPI = {i: GHI_SIN[i] / ghi_max for i in NODOS}

    # ── NSI: Network Stress Index ─────────────────────────────────────────────
    # NSI[i] = max(0, (D[i]-G[i])/D[i]) — solo nodos importadores son relevantes
    NSI_raw = {}
    for i in NODOS:
        if demanda[i] > 0:
            nsi = (demanda[i] - gmax[i]) / demanda[i]
        else:
            nsi = 0.0
        NSI_raw[i] = max(0.0, nsi)  # negativos → 0 (exportadores)

    nsi_max = max(NSI_raw.values()) if max(NSI_raw.values()) > 0 else 1.0
    NSI = {i: NSI_raw[i] / nsi_max for i in NODOS}

    # ── GRI: Grid Reinforcement Index ─────────────────────────────────────────
    # Si no hay flujos reales, estimar carga proporcional al déficit
    if flujos_aprox is None:
        # Estimación: flujo = déficit del nodo más déficit del vecino / 2
        flujos_aprox = {}
        for j in CORREDORES:
            a, b = map(int, j.split(','))
            def_a = max(0, demanda[a] - gmax[a])
            def_b = max(0, demanda[b] - gmax[b])
            flujos_aprox[j] = (def_a + def_b) / 2.0

    GRI_raw = {i: 0.0 for i in NODOS}
    for j in CORREDORES:
        a, b = map(int, j.split(','))
        n_tot = max(N0[j], 1)
        cap   = FMAX[j] * n_tot
        ratio = abs(flujos_aprox[j]) / cap if cap > 0 else 0.0
        ratio = min(ratio, 1.0)  # cap en 100% de carga
        GRI_raw[a] += ratio
        GRI_raw[b] += ratio

    gri_max = max(GRI_raw.values()) if max(GRI_raw.values()) > 0 else 1.0
    GRI = {i: GRI_raw[i] / gri_max for i in NODOS}

    # ── CEI: Community Energy Index ───────────────────────────────────────────
    w1, w2, w3 = 0.4, 0.3, 0.3
    CEI = {i: w1 * SPI[i] + w2 * NSI[i] + w3 * GRI[i] for i in NODOS}

    # ── DataFrame de resultados ───────────────────────────────────────────────
    df = pd.DataFrame({
        'Node':       NODOS,
        'Nombre':     [NOMBRE[i] for i in NODOS],
        'Demand_MW':  [demanda[i] for i in NODOS],
        'Gen_MW':     [gmax[i] for i in NODOS],
        'GHI':        [GHI_SIN[i] for i in NODOS],
        'SPI':        [SPI[i] for i in NODOS],
        'NSI':        [NSI[i] for i in NODOS],
        'GRI':        [GRI[i] for i in NODOS],
        'CEI':        [CEI[i] for i in NODOS],
    })

    return df

def calcular_reduccion_inteligente(alpha_total, df_cei):
    """
    Calcula la reducción por nodo basada en el índice CEI (reducción inteligente).

    La idea es asignar mayor reducción a nodos con mayor CEI, manteniendo
    el promedio igual a alpha_total.

    Formulación:
        CEI_norm[i] = CEI[i] / CEI_mean  (ratio respecto a la media)
        alpha_nodo[i] = alpha_total × CEI_norm[i]
        (se clampea para que alpha_nodo[i] ∈ [0, 0.40])

    Parámetro:
        alpha_total : penetración promedio del sistema
        df_cei      : DataFrame con columna 'CEI' indexado por 'Node'

    Retorna:
        dict {nodo: alpha_efectivo}
    """
    cei_vals   = {int(row['Node']): row['CEI'] for _, row in df_cei.iterrows()}
    # Solo nodos con demanda > 0 reciben reducción
    dem_vals   = {int(row['Node']): row['Demand_MW'] for _, row in df_cei.iterrows()}

    cei_mean   = np.mean([cei_vals[i] for i in NODOS if dem_vals[i] > 0])
    if cei_mean == 0:
        return {i: alpha_total for i in NODOS}

    alpha_cei = {}
    for i in NODOS:
        if dem_vals[i] > 0:
            ratio = cei_vals[i] / cei_mean
            alpha_cei[i] = min(alpha_total * ratio, 0.40)  # cap máx 40%
        else:
            alpha_cei[i] = 0.0

    return alpha_cei

def aplicar_reduccion(demanda_base, alpha_dict):
    """
    Aplica reducción de demanda neta por comunidades energéticas BTM.

    D_nuevo[i] = D_base[i] × (1 − alpha[i])

    Esta formulación captura el efecto behind-the-meter: la generación
    distribuida reduce la demanda neta que el sistema de transmisión
    debe satisfacer, sin modificar la topología o capacidad de la red.
    """
    return {i: demanda_base[i] * (1.0 - alpha_dict[i]) for i in NODOS}

# =============================================================================
# 3. CONSTRUCCIÓN DEL MODELO MILP DC
# =============================================================================

def get_S(i, j_str):
    """Incidencia nodo–corredor: +1 si i es origen, -1 si destino, 0 si ajeno."""
    a, b = map(int, j_str.split(','))
    if i == a: return  1
    if i == b: return -1
    return 0

def construir_modelo(demanda, gmax):
    """
    Modelo MILP DC de expansión de transmisión (TEP) — formulación estática.

    Variables de decisión:
        y[j,k]   : binaria — instalar circuito k del corredor j (1=sí, 0=no)
        f[j,k]   : flujo de potencia [p.u.] por circuito k, corredor j
        g[i]     : generación del nodo i [p.u.]
        theta[i] : ángulo de tensión del nodo i [rad]

    Función objetivo:
        min Σ_j Σ_k COSTO[j] · y[j,k]   (solo circuitos NUEVOS: k > N0[j])

    Restricciones principales:
        (1) Balance nodal DC
        (2) Flujo DC con linealización Big-M
        (3) Límites térmicos |f[j,k]| ≤ FMAX[j] · y[j,k]
        (4) Generación máxima g[i] ≤ gmax[i]
        (5) Líneas existentes fijas: y[j,k]=1 para k ≤ N0[j]
        (6) Referencia angular: theta[NREF] = 0
    """
    model = pyo.ConcreteModel()

    model.I = pyo.Set(initialize=NODOS)
    model.J = pyo.Set(initialize=CORREDORES)
    model.K = pyo.RangeSet(1, NCAND)

    model.S    = pyo.Param(model.I, model.J,
                           initialize={(i,j): get_S(i,j) for i in NODOS for j in CORREDORES})
    model.d    = pyo.Param(model.I, initialize={i: demanda[i]/Sbase for i in NODOS})
    model.gmax = pyo.Param(model.I, initialize={i: gmax[i]/Sbase for i in NODOS})
    model.n0   = pyo.Param(model.J, initialize=N0)
    model.fmax = pyo.Param(model.J, initialize={j: FMAX[j]/Sbase for j in CORREDORES})
    model.B    = pyo.Param(model.J, initialize=B_SIN)
    model.c    = pyo.Param(model.J, initialize=COSTO)

    model.y     = pyo.Var(model.J, model.K, domain=pyo.Binary)
    model.f     = pyo.Var(model.J, model.K, domain=pyo.Reals)
    model.g     = pyo.Var(model.I, domain=pyo.NonNegativeReals)
    model.theta = pyo.Var(model.I, domain=pyo.Reals)

    # Referencia angular
    model.slack = pyo.Constraint(expr=model.theta[NREF] == 0)

    # Función objetivo: minimizar inversión en circuitos nuevos
    def obj_rule(m):
        return sum(m.c[j] * m.y[j,k]
                   for j in m.J for k in m.K
                   if k > pyo.value(m.n0[j]))
    model.obj = pyo.Objective(rule=obj_rule, sense=pyo.minimize)

    # (1) Balance nodal: Σ S[i,j]·f[j,k] + g[i] = d[i]  ∀i
    def balance_nodal(m, i):
        return (sum(m.S[i,j]*m.f[j,k] for j in m.J for k in m.K)
                + m.g[i] == m.d[i])
    model.balance_nodal = pyo.Constraint(model.I, rule=balance_nodal)

    # (2) Flujo DC — Big-M: f[j,k] - B[j]·(θ_a-θ_b) ∈ [-M(1-y), M(1-y)]
    def flujo_pos(m, j, k):
        a, b = map(int, j.split(','))
        return m.f[j,k] - m.B[j]*(m.theta[a]-m.theta[b]) <=  BIG_M*(1-m.y[j,k])
    def flujo_neg(m, j, k):
        a, b = map(int, j.split(','))
        return m.f[j,k] - m.B[j]*(m.theta[a]-m.theta[b]) >= -BIG_M*(1-m.y[j,k])
    model.flujo_pos = pyo.Constraint(model.J, model.K, rule=flujo_pos)
    model.flujo_neg = pyo.Constraint(model.J, model.K, rule=flujo_neg)

    # (3) Límites térmicos: |f[j,k]| ≤ fmax[j]·y[j,k]
    def lim_pos(m, j, k): return  m.f[j,k] <=  m.fmax[j]*m.y[j,k]
    def lim_neg(m, j, k): return -m.f[j,k] <=  m.fmax[j]*m.y[j,k]
    model.lim_pos = pyo.Constraint(model.J, model.K, rule=lim_pos)
    model.lim_neg = pyo.Constraint(model.J, model.K, rule=lim_neg)

    # (4) Generación máxima
    def gen_max(m, i): return m.g[i] <= m.gmax[i]
    model.gen_max = pyo.Constraint(model.I, rule=gen_max)

    # (5) Líneas existentes fijas
    def lineas_exist(m, j, k):
        if k <= pyo.value(m.n0[j]):
            return m.y[j,k] == 1
        return pyo.Constraint.Skip
    model.lineas_exist = pyo.Constraint(model.J, model.K, rule=lineas_exist)

    return model

# =============================================================================
# 4. CONSTRUCCIÓN DEL MODELO MILP DC MULTIPERIODO
# =============================================================================

def demanda_t_btm(i, t, alpha_dict_ini=None, alpha_dict_fin=None):
    """
    Demanda del nodo i en el periodo t, con penetración BTM interpolada.

    La penetración BTM crece linealmente desde 0 en 2024 hasta alpha_fin en 2039,
    modelando la adopción progresiva de comunidades energéticas.

    D[i,t] = D_base[i,t] × (1 − α[i,t])
    α[i,t] = alpha_ini[i] + (alpha_fin[i]-alpha_ini[i]) × (t-1)/(NT-1)
    """
    d_base = D_2024[i] + (D_2039[i] - D_2024[i]) * (t - 1) / (NT - 1)
    if alpha_dict_ini is None:
        return d_base
    a_ini = alpha_dict_ini.get(i, 0.0)
    a_fin = alpha_dict_fin.get(i, 0.0)
    alpha_t = a_ini + (a_fin - a_ini) * (t - 1) / (NT - 1)
    return d_base * (1.0 - alpha_t)

def gmax_t(i, t):
    """Generación máxima del nodo i en el periodo t — interpolación lineal."""
    return G_2024[i] + (G_2039[i] - G_2024[i]) * (t - 1) / (NT - 1)

def construir_modelo_multiperiodo(alpha_fin_dict=None):
    """
    Modelo MILP DC multiperiodo 2024-2039.

    Las comunidades energéticas BTM se modelan como una reducción lineal
    de la demanda que crece desde 0% en 2024 hasta alpha_fin en 2039,
    simulando la adopción progresiva del recurso BTM.

    Variables adicionales respecto al modelo estático:
        y[j,k,t] : binaria — construir circuito k del corredor j en el año t
        x[j,k,t] : binaria — estado acumulado (circuito activo en t)
    """
    if alpha_fin_dict is None:
        alpha_fin_dict = {i: 0.0 for i in NODOS}
    alpha_ini_dict = {i: 0.0 for i in NODOS}

    mdl = pyo.ConcreteModel()
    mdl.I = pyo.Set(initialize=NODOS)
    mdl.J = pyo.Set(initialize=CORREDORES)
    mdl.K = pyo.RangeSet(1, NCAND)
    mdl.T = pyo.RangeSet(1, NT)

    mdl.S    = pyo.Param(mdl.I, mdl.J,
                         initialize={(i,j): get_S(i,j) for i in NODOS for j in CORREDORES})
    mdl.d    = pyo.Param(mdl.I, mdl.T,
                         initialize={(i,t): demanda_t_btm(i, t, alpha_ini_dict, alpha_fin_dict)/Sbase
                                     for i in NODOS for t in PERIODOS})
    mdl.gmax = pyo.Param(mdl.I, mdl.T,
                         initialize={(i,t): gmax_t(i,t)/Sbase for i in NODOS for t in PERIODOS})
    mdl.n0   = pyo.Param(mdl.J, initialize=N0)
    mdl.fmax = pyo.Param(mdl.J, initialize={j: FMAX[j]/Sbase for j in CORREDORES})
    mdl.B    = pyo.Param(mdl.J, initialize=B_SIN)
    mdl.c    = pyo.Param(mdl.J, initialize=COSTO)
    # Factor de descuento VPN
    desc = {t: 1.0/(1.0+TASA_DESCUENTO)**(t-1) for t in PERIODOS}
    mdl.desc = pyo.Param(mdl.T, initialize=desc)

    mdl.y     = pyo.Var(mdl.J, mdl.K, mdl.T, domain=pyo.Binary)
    mdl.x     = pyo.Var(mdl.J, mdl.K, mdl.T, domain=pyo.Binary)
    mdl.f     = pyo.Var(mdl.J, mdl.K, mdl.T, domain=pyo.Reals)
    mdl.g     = pyo.Var(mdl.I, mdl.T, domain=pyo.NonNegativeReals)
    mdl.theta = pyo.Var(mdl.I, mdl.T, domain=pyo.Reals)

    def slack_rule(m, t): return m.theta[NREF, t] == 0
    mdl.slack = pyo.Constraint(mdl.T, rule=slack_rule)

    # Función objetivo: minimizar VPN de inversiones en circuitos nuevos
    def obj_rule(m):
        return sum(m.c[j]*m.y[j,k,t]*m.desc[t]
                   for j in m.J for k in m.K for t in m.T
                   if k > pyo.value(m.n0[j]))
    mdl.obj = pyo.Objective(rule=obj_rule, sense=pyo.minimize)

    # Balance nodal por periodo
    def balance_nodal(m, i, t):
        return (sum(m.S[i,j]*m.f[j,k,t] for j in m.J for k in m.K)
                + m.g[i,t] == m.d[i,t])
    mdl.balance_nodal = pyo.Constraint(mdl.I, mdl.T, rule=balance_nodal)

    # Flujo DC Big-M
    def flujo_pos(m, j, k, t):
        a, b = map(int, j.split(','))
        return m.f[j,k,t] - m.B[j]*(m.theta[a,t]-m.theta[b,t]) <=  BIG_M*(1-m.x[j,k,t])
    def flujo_neg(m, j, k, t):
        a, b = map(int, j.split(','))
        return m.f[j,k,t] - m.B[j]*(m.theta[a,t]-m.theta[b,t]) >= -BIG_M*(1-m.x[j,k,t])
    mdl.flujo_pos = pyo.Constraint(mdl.J, mdl.K, mdl.T, rule=flujo_pos)
    mdl.flujo_neg = pyo.Constraint(mdl.J, mdl.K, mdl.T, rule=flujo_neg)

    # Límites térmicos
    def lim_pos(m, j, k, t): return  m.f[j,k,t] <=  m.fmax[j]*m.x[j,k,t]
    def lim_neg(m, j, k, t): return -m.f[j,k,t] <=  m.fmax[j]*m.x[j,k,t]
    mdl.lim_pos = pyo.Constraint(mdl.J, mdl.K, mdl.T, rule=lim_pos)
    mdl.lim_neg = pyo.Constraint(mdl.J, mdl.K, mdl.T, rule=lim_neg)

    # Generación máxima
    def gen_max(m, i, t): return m.g[i,t] <= m.gmax[i,t]
    mdl.gen_max = pyo.Constraint(mdl.I, mdl.T, rule=gen_max)

    # Dinámica de expansión: x[j,k,t] = x[j,k,t-1] + y[j,k,t]
    def dinamica(m, j, k, t):
        base = 1 if k <= pyo.value(m.n0[j]) else 0
        if t == 1:
            return m.x[j,k,t] == base + m.y[j,k,t]
        return m.x[j,k,t] == m.x[j,k,t-1] + m.y[j,k,t]
    mdl.dinamica = pyo.Constraint(mdl.J, mdl.K, mdl.T, rule=dinamica)

    # Cada circuito se construye máximo una vez
    def una_vez(m, j, k):
        return sum(m.y[j,k,t] for t in m.T) <= 1
    mdl.una_vez = pyo.Constraint(mdl.J, mdl.K, rule=una_vez)

    # Máximo NCAND circuitos nuevos por corredor
    def max_circ(m, j):
        return sum(m.y[j,k,t] for k in m.K for t in m.T
                   if k > pyo.value(m.n0[j])) <= NCAND
    mdl.max_circ = pyo.Constraint(mdl.J, rule=max_circ)

    # x ≤ 1 (refuerzo binario)
    def x_bin(m, j, k, t): return m.x[j,k,t] <= 1
    mdl.x_bin = pyo.Constraint(mdl.J, mdl.K, mdl.T, rule=x_bin)

    return mdl

# =============================================================================
# 5. RESOLUCIÓN DE MODELOS
# =============================================================================

def resolver_estatico(demanda, gmax, presupuesto_MUSD, titulo, verbose=True):
    """
    Construye y resuelve el modelo TEP estático.
    Retorna dict con resultados o None si no hay solución.
    """
    if verbose:
        print(f"\n  ── Resolviendo: {titulo} ──")

    model = construir_modelo(demanda, gmax)

    model.presupuesto = pyo.Constraint(
        expr=sum(model.c[j]*model.y[j,k]
                 for j in model.J for k in model.K
                 if k > pyo.value(model.n0[j])) <= presupuesto_MUSD
    )

    solver = SolverFactory('gurobi')
    solver.options['MIPGap']    = 0.003
    solver.options['TimeLimit'] = 300
    solver.options['OutputFlag'] = 0

    t0 = time.time()
    result = solver.solve(model, tee=False)
    dt = time.time() - t0

    status = result.solver.termination_condition
    cond_ok = (pyo.TerminationCondition.optimal, pyo.TerminationCondition.feasible)
    if status not in cond_ok:
        if verbose:
            print(f"    ❌ Sin solución ({status}) — {dt:.1f}s")
        return None

    if verbose:
        print(f"    ✅ {status} — {dt:.1f}s")

    nuevas_lineas, inversion, flujos = {}, 0.0, {}
    generacion = {i: pyo.value(model.g[i])*Sbase for i in NODOS}
    angulos    = {i: np.degrees(pyo.value(model.theta[i])) for i in NODOS}

    for j in CORREDORES:
        flujos[j] = sum(pyo.value(model.f[j,k])*Sbase for k in range(1, NCAND+1))
        nc = sum(1 for k in range(1, NCAND+1)
                 if k > N0[j] and pyo.value(model.y[j,k]) > 0.5)
        if nc > 0:
            nuevas_lineas[j] = nc
            inversion += COSTO[j]*nc

    # Cargabilidad de corredores
    cargabilidad = {}
    for j in CORREDORES:
        n_tot = N0[j] + nuevas_lineas.get(j, 0)
        cap   = FMAX[j]*max(n_tot, 1)
        cargabilidad[j] = abs(flujos[j])/cap if cap > 0 else 0.0

    return {
        'titulo':        titulo,
        'demanda':       demanda,
        'gmax':          gmax,
        'nuevas_lineas': nuevas_lineas,
        'flujos':        flujos,
        'generacion':    generacion,
        'angulos':       angulos,
        'inversion':     inversion,
        'cargabilidad':  cargabilidad,
        'dem_total':     sum(demanda.values()),
    }

def resolver_multiperiodo_btm(alpha_fin_dict, titulo):
    """
    Resuelve el modelo multiperiodo con penetración BTM dada.
    Retorna dict con resultados por periodo o None.
    """
    print(f"\n  ── Resolviendo multiperiodo: {titulo} ──")

    mdl = construir_modelo_multiperiodo(alpha_fin_dict)

    solver = SolverFactory('gurobi')
    solver.options['MIPGap']     = 0.005
    solver.options['TimeLimit']  = 700
    solver.options['OutputFlag'] = 0

    t0 = time.time()
    result = solver.solve(mdl, tee=False)
    dt = time.time() - t0

    status = result.solver.termination_condition
    cond_ok = (pyo.TerminationCondition.optimal, pyo.TerminationCondition.feasible)
    if status not in cond_ok:
        print(f"    ❌ Sin solución ({status}) — {dt:.1f}s")
        return None
    print(f"    ✅ {status} — {dt:.1f}s")

    expansion_por_t  = {}
    inversion_por_t  = {}
    demanda_por_t    = {}
    flujos_por_t     = {}

    alpha_ini = {i: 0.0 for i in NODOS}
    for t in PERIODOS:
        exp_t  = {}
        inv_t  = 0.0
        fl_t   = {}
        for j in CORREDORES:
            nc = sum(1 for k in range(1, NCAND+1)
                     if k > N0[j] and pyo.value(mdl.y[j,k,t]) > 0.5)
            if nc > 0:
                exp_t[j] = nc
                inv_t   += COSTO[j]*nc
            fl_t[j] = sum(pyo.value(mdl.f[j,k,t])*Sbase for k in range(1, NCAND+1))
        expansion_por_t[t] = exp_t
        inversion_por_t[t] = inv_t
        demanda_por_t[t]   = sum(demanda_t_btm(i, t, alpha_ini, alpha_fin_dict) for i in NODOS)
        flujos_por_t[t]    = fl_t

    return {
        'titulo':          titulo,
        'alpha_fin':       alpha_fin_dict,
        'expansion_por_t': expansion_por_t,
        'inversion_por_t': inversion_por_t,
        'demanda_por_t':   demanda_por_t,
        'flujos_por_t':    flujos_por_t,
        'inversion_total': sum(inversion_por_t.values()),
    }

# =============================================================================
# 6. ANÁLISIS BTM — EJECUTAR TODOS LOS ESCENARIOS
# =============================================================================

def ejecutar_analisis_btm(df_cei, presupuesto=8000):
    """
    Ejecuta el TEP estático para todos los escenarios BTM:
      - Escenario base (α=0%) con la demanda alta 2039
      - Reducción uniforme: α = 5%, 10%, 15%, 20%
      - Reducción inteligente (CEI): α = 5%, 10%, 15%, 20%

    Las comunidades energéticas reducen D_neta:
        D_nuevo[i] = D_base[i] × (1 − α[i])

    Retorna dos dicts:
        resultados_uniforme[alpha] : resultado TEP con reducción uniforme
        resultados_cei[alpha]      : resultado TEP con reducción CEI inteligente
    """
    resultados_uniforme = {}
    resultados_cei      = {}

    print("\n" + "="*65)
    print("  ANÁLISIS BTM — ESCENARIOS DE REDUCCIÓN DE DEMANDA")
    print("="*65)

    for alpha in ALPHAS:
        # ── Escenario uniforme ────────────────────────────────────────────────
        alpha_unif = {i: alpha for i in NODOS}
        dem_unif   = aplicar_reduccion(D_2039, alpha_unif)
        titulo_u   = f"Uniforme {ALPHA_LABELS[alpha]}"
        r = resolver_estatico(dem_unif, G_2039, presupuesto, titulo_u, verbose=True)
        resultados_uniforme[alpha] = r

        # ── Escenario CEI inteligente ─────────────────────────────────────────
        alpha_cei_dict = calcular_reduccion_inteligente(alpha, df_cei)
        dem_cei        = aplicar_reduccion(D_2039, alpha_cei_dict)
        titulo_c       = f"CEI-Inteligente {ALPHA_LABELS[alpha]}"
        r = resolver_estatico(dem_cei, G_2039, presupuesto, titulo_c, verbose=True)
        resultados_cei[alpha] = r

    return resultados_uniforme, resultados_cei

# =============================================================================
# 7. CÁLCULO DE COSTOS EVITADOS Y DIFERIDOS
# =============================================================================

def calcular_ahorros(resultados_uniforme, resultados_cei):
    """
    Calcula el costo evitado, diferido y ahorro total respecto al caso base
    (α=0%, sin comunidades energéticas).

    Definiciones:
        Costo_evitado[α]  = Inv_base - Inv[α]
            → Líneas que NO se construyen porque la red ya no las necesita

        Costo_diferido[α] = max(0, Inv_base - Inv[α]) cuando la reducción
            de demanda posterga la necesidad de expansión en el tiempo

        Ahorro_total[α]   = Inv_base - Inv[α]
            → Diferencia absoluta en inversión respecto al caso base

    En el modelo estático, el costo evitado coincide con el ahorro total.
    El diferido se cuantifica explícitamente en el modelo multiperiodo.
    """
    base_inv_u = resultados_uniforme[0.0]['inversion'] if resultados_uniforme[0.0] else 0
    base_inv_c = resultados_cei[0.0]['inversion']      if resultados_cei[0.0]      else 0

    rows = []
    for alpha in ALPHAS:
        inv_u = resultados_uniforme[alpha]['inversion'] if resultados_uniforme[alpha] else None
        inv_c = resultados_cei[alpha]['inversion']      if resultados_cei[alpha]      else None
        rows.append({
            'Alpha':        alpha,
            'Alfa_pct':     f"{int(alpha*100)}%",
            'Inv_Unif':     inv_u,
            'Inv_CEI':      inv_c,
            'Ahorro_Unif':  base_inv_u - inv_u if inv_u is not None else None,
            'Ahorro_CEI':   base_inv_u - inv_c if inv_c is not None else None,
            'Dem_Unif':     resultados_uniforme[alpha]['dem_total'] if resultados_uniforme[alpha] else None,
            'Dem_CEI':      resultados_cei[alpha]['dem_total']      if resultados_cei[alpha]      else None,
        })

    df_ahorro = pd.DataFrame(rows)

    print("\n" + "="*65)
    print("  TABLA RESUMEN — COSTOS EVITADOS POR BTM")
    print("="*65)
    print(f"  {'α':>5} | {'Inv. Unif (MUSD)':>16} | {'Ahorro Unif':>12} | "
          f"{'Inv. CEI (MUSD)':>15} | {'Ahorro CEI':>11}")
    print(f"  {'─'*5}─┼─{'─'*16}─┼─{'─'*12}─┼─{'─'*15}─┼─{'─'*11}")
    for _, row in df_ahorro.iterrows():
        print(f"  {row['Alfa_pct']:>5} | {row['Inv_Unif'] or 0:>16,.0f} | "
              f"{row['Ahorro_Unif'] or 0:>12,.0f} | "
              f"{row['Inv_CEI'] or 0:>15,.0f} | {row['Ahorro_CEI'] or 0:>11,.0f}")
    print()

    return df_ahorro

# =============================================================================
# 8. GRÁFICAS ACADÉMICAS
# =============================================================================

# ── Paleta de colores para los niveles de alpha ───────────────────────────────
ALPHA_COLORS_U = {
    0.00: '#1565c0',
    0.05: '#2196F3',
    0.10: '#4CAF50',
    0.15: '#FF9800',
    0.20: '#F44336',
    0.25: '#E91E63',
    0.30: '#9C27B0',
    0.35: '#6A1B9A',
}
ALPHA_COLORS_C = {
    0.00: '#1565c0',
    0.05: '#8BC34A',
    0.10: '#CDDC39',
    0.15: '#FF5722',
    0.20: '#9C27B0',
    0.25: '#E91E63',
    0.30: '#9C27B0',
    0.35: '#6A1B9A',
}

def figura_01_ranking_CEI(df_cei):
    """
    FIG 1: Ranking CEI de nodos — bar chart horizontal con descomposición
    de los tres subíndices (SPI, NSI, GRI).
    Muestra qué áreas del SIN son prioritarias para comunidades energéticas BTM.
    """
    df_s = df_cei.sort_values('CEI', ascending=True).reset_index(drop=True)

    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    fig.suptitle(
        "FIG 1 — Ranking CEI: Priorización de Nodos para Comunidades Energéticas BTM\n"
        "SIN Colombia 500 kV — Escenario Alta Demanda 2039  |  Uniandes 2026",
        fontsize=11, fontweight='bold'
    )

    ax = axes[0]
    y_pos = range(len(df_s))
    width = 0.25
    ax.barh(y_pos, df_s['SPI'], width, label='SPI (Solar)', color='#FFC107', alpha=0.85)
    ax.barh([y+width for y in y_pos], df_s['NSI'], width, label='NSI (Estrés red)', color='#F44336', alpha=0.85)
    ax.barh([y+2*width for y in y_pos], df_s['GRI'], width, label='GRI (Refuerzo)', color='#1565c0', alpha=0.85)
    ax.set_yticks([y+width for y in y_pos])
    ax.set_yticklabels([f"N{int(r['Node'])} — {r['Nombre']}" for _, r in df_s.iterrows()], fontsize=10)
    ax.set_xlabel("Índice normalizado [0–1]")
    ax.set_title("Subíndices por área (mayor = mayor prioridad BTM)")
    ax.legend(fontsize=9)
    ax.grid(True, alpha=0.3, axis='x')
    ax.set_xlim(0, 1.15)
    for _, row in df_s.iterrows():
        yy = list(df_s['Node']).index(row['Node'])
        ax.text(row['SPI']+0.01, yy, f"{row['SPI']:.2f}", va='center', fontsize=8, color='#333')

    ax2 = axes[1]
    df_s2 = df_cei.sort_values('CEI', ascending=False).reset_index(drop=True)
    bars  = ax2.barh(range(len(df_s2)),
                     df_s2['CEI'],
                     color=[COL_NODOS[int(r['Node'])] for _, r in df_s2.iterrows()],
                     alpha=0.88)
    ax2.set_yticks(range(len(df_s2)))
    ax2.set_yticklabels([f"N{int(r['Node'])} — {r['Nombre']}" for _, r in df_s2.iterrows()], fontsize=10)
    ax2.set_xlabel("CEI (Community Energy Index)")
    ax2.set_title("Ranking CEI — Mayor CEI = Prioridad más alta")
    ax2.grid(True, alpha=0.3, axis='x')
    ax2.set_xlim(0, 1.05)
    for bar, (_, row) in zip(bars, df_s2.iterrows()):
        ax2.text(bar.get_width()+0.01, bar.get_y()+bar.get_height()/2,
                 f"CEI={row['CEI']:.3f}", va='center', fontsize=9, fontweight='bold')

    plt.tight_layout()
    fname = "BTM_Fig01_Ranking_CEI.png"
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig(fname), dpi=160, bbox_inches='tight')
        print(f"  → Guardada: {fname}")
    plt.show(); plt.close()

def figura_02_demanda_vs_inversion(df_ahorro, resultados_uniforme, resultados_cei):
    """
    FIG 2: Demanda total del sistema vs. inversión óptima en transmisión.
    Muestra la relación entre reducción de demanda y necesidad de expansión.
    """
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    fig.suptitle(
        "FIG 2 — Demanda Total vs. Inversión Óptima en Transmisión\n"
        "Efecto de las Comunidades Energéticas BTM como Recurso Virtual  |  Uniandes 2026",
        fontsize=11, fontweight='bold'
    )

    ax = axes[0]
    # Línea de tendencia — uniforme
    dems_u = [resultados_uniforme[a]['dem_total'] for a in ALPHAS if resultados_uniforme[a]]
    invs_u = [resultados_uniforme[a]['inversion']  for a in ALPHAS if resultados_uniforme[a]]
    dems_c = [resultados_cei[a]['dem_total']      for a in ALPHAS if resultados_cei[a]]
    invs_c = [resultados_cei[a]['inversion']       for a in ALPHAS if resultados_cei[a]]

    ax.plot(dems_u, invs_u, 'o-', color='#1565c0', lw=2.5, ms=9, label='Reducción Uniforme')
    ax.plot(dems_c, invs_c, 's--', color='#F44336', lw=2.5, ms=9, label='Reducción CEI Inteligente')
    for a, du, iu, dc, ic in zip(ALPHAS, dems_u, invs_u, dems_c, invs_c):
        ax.annotate(f"α={int(a*100)}%", (du, iu),
                    textcoords="offset points", xytext=(6, 6), fontsize=8.5, color='#1565c0')
        ax.annotate(f"α={int(a*100)}%", (dc, ic),
                    textcoords="offset points", xytext=(6, -14), fontsize=8.5, color='#F44336')

    ax.set_xlabel("Demanda total del sistema (MW)", fontsize=10)
    ax.set_ylabel("Inversión óptima (MUSD)", fontsize=10)
    ax.set_title("Curva Demanda–Inversión por estrategia BTM")
    ax.legend(fontsize=9)
    ax.grid(True, alpha=0.35)

    # Región de ahorro sombreada
    ax2 = axes[1]
    alpha_pcts  = [int(a*100) for a in ALPHAS]
    inv_base_u  = df_ahorro[df_ahorro['Alpha']==0.0]['Inv_Unif'].values[0] or 0
    ahorros_u   = df_ahorro['Ahorro_Unif'].fillna(0).tolist()
    ahorros_c   = df_ahorro['Ahorro_CEI'].fillna(0).tolist()

    x = np.arange(len(alpha_pcts))
    width = 0.35
    bars1 = ax2.bar(x - width/2, ahorros_u, width, label='Uniforme',       color='#1565c0', alpha=0.85)
    bars2 = ax2.bar(x + width/2, ahorros_c, width, label='CEI Inteligente', color='#F44336', alpha=0.85)
    ax2.set_xticks(x)
    ax2.set_xticklabels([f"α={p}%" for p in alpha_pcts], fontsize=10)
    ax2.set_ylabel("Ahorro en inversión (MUSD)", fontsize=10)
    ax2.set_title("Ahorro respecto al caso base (α=0%)")
    ax2.legend(fontsize=9)
    ax2.grid(True, alpha=0.3, axis='y')
    for bar in bars1:
        if bar.get_height() > 0:
            ax2.text(bar.get_x()+bar.get_width()/2, bar.get_height()+5,
                     f"{bar.get_height():,.0f}", ha='center', va='bottom', fontsize=8, color='#1565c0')
    for bar in bars2:
        if bar.get_height() > 0:
            ax2.text(bar.get_x()+bar.get_width()/2, bar.get_height()+5,
                     f"{bar.get_height():,.0f}", ha='center', va='bottom', fontsize=8, color='#F44336')

    # Línea de inversión base
    ax2.axhline(0, color='gray', lw=0.8, ls='--')
    ax2.text(len(ALPHAS)-0.5, inv_base_u*0.01,
             f"Base: {inv_base_u:,.0f} MUSD", fontsize=8, color='gray')

    plt.tight_layout()
    fname = "BTM_Fig02_Demanda_vs_Inversion.png"
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig(fname), dpi=160, bbox_inches='tight')
        print(f"  → Guardada: {fname}")
    plt.show(); plt.close()

def figura_03_penetracion_vs_congestion(resultados_uniforme, resultados_cei):
    """
    FIG 3: Penetración BTM vs. congestión del sistema.
    Muestra cómo la reducción de demanda alivia la cargabilidad de los
    corredores más críticos (Oriental, Caribe, Suroccidental).
    """
    # Corredores críticos para el escenario 2039
    corredores_criticos = ['2,3', '3,4', '3,5', '1,4', '4,6']

    fig, axes = plt.subplots(1, 2, figsize=(14, 7))
    fig.suptitle(
        "FIG 3 — Penetración BTM vs. Congestión en Corredores Críticos\n"
        "La reducción de demanda neta descongestiona la red de transmisión  |  Uniandes 2026",
        fontsize=11, fontweight='bold'
    )

    for ax, (res_dict, label, title) in zip(
        axes,
        [(resultados_uniforme, 'Uniforme',      'Estrategia Uniforme'),
         (resultados_cei,      'CEI Inteligente','Estrategia CEI Inteligente')]
    ):
        for j in corredores_criticos:
            cargs = []
            for alpha in ALPHAS:
                r = res_dict[alpha]
                if r:
                    n_tot = N0[j] + r['nuevas_lineas'].get(j, 0)
                    cap   = FMAX[j]*max(n_tot, 1)
                    carg  = abs(r['flujos'].get(j, 0))/cap * 100
                    cargs.append(carg)
                else:
                    cargs.append(np.nan)
            ax.plot([int(a*100) for a in ALPHAS], cargs,
                    'o-', lw=2, ms=7,
                    label=f"[{j}] {DESC[j][:25]}...")

        ax.axhline(95, color='red',    lw=1.5, ls='--', alpha=0.7, label='Límite crítico (95%)')
        ax.axhline(80, color='orange', lw=1.2, ls='--', alpha=0.7, label='Alta carga (80%)')
        ax.set_xlabel("Penetración BTM α (%)", fontsize=10)
        ax.set_ylabel("Cargabilidad del corredor (%)", fontsize=10)
        ax.set_title(f"Estrategia: {title}", fontsize=10)
        ax.legend(fontsize=7.5, loc='upper right')
        ax.grid(True, alpha=0.3)
        ax.set_xlim(-1, 21)
        ax.set_ylim(0, 115)

    plt.tight_layout()
    fname = "BTM_Fig03_Penetracion_vs_Congestion.png"
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig(fname), dpi=160, bbox_inches='tight')
        print(f"  → Guardada: {fname}")
    plt.show(); plt.close()

def figura_04_lineas_por_escenario(resultados_uniforme, resultados_cei):
    """
    FIG 4: Comparación de líneas construidas por escenario BTM.
    Visualiza qué líneas dejan de construirse conforme aumenta la penetración BTM.
    El verde = evitada, el rojo = aún necesaria.
    """
    fig, axes = plt.subplots(2, len(ALPHAS), figsize=(5*len(ALPHAS), 10))
    fig.suptitle(
        "FIG 4 — Líneas Construidas por Escenario BTM\n"
        "Verde oscuro = evitada por BTM  |  Rojo = aún necesaria  |  Uniandes 2026",
        fontsize=11, fontweight='bold'
    )

    base_u = set(resultados_uniforme[0.0]['nuevas_lineas'].keys()) if resultados_uniforme[0.0] else set()
    base_c = set(resultados_cei[0.0]['nuevas_lineas'].keys())      if resultados_cei[0.0]      else set()

    for col, alpha in enumerate(ALPHAS):
        for row, (res_dict, base_lineas, label) in enumerate(
            [(resultados_uniforme, base_u, 'Uniforme'),
             (resultados_cei,      base_c, 'CEI')]
        ):
            ax = axes[row][col]
            r  = res_dict[alpha]
            if r is None:
                ax.text(0.5, 0.5, 'Sin solución', ha='center', va='center', transform=ax.transAxes)
                ax.axis('off')
                continue

            lineas_actual  = set(r['nuevas_lineas'].keys())
            lineas_evitadas = base_lineas - lineas_actual
            lineas_nuevas   = lineas_actual

            # Barra horizontal: verde = evitada, rojo = construida, azul = base
            all_lineas = sorted(base_lineas | lineas_actual)
            estados    = []
            colores_b  = []
            for j in all_lineas:
                if j in lineas_evitadas:
                    estados.append(f"[{j}] EVITADA ✓")
                    colores_b.append('#2E7D32')
                elif j in lineas_nuevas:
                    estados.append(f"[{j}] {r['nuevas_lineas'][j]}×")
                    colores_b.append('#C62828')
                else:
                    estados.append(f"[{j}]")
                    colores_b.append('#757575')

            valores = [1]*len(all_lineas)
            ax.barh(range(len(all_lineas)), valores, color=colores_b, alpha=0.85)
            ax.set_yticks(range(len(all_lineas)))
            ax.set_yticklabels(estados, fontsize=7)
            ax.set_xticks([])
            ax.set_title(f"{label}\nα={int(alpha*100)}%\n"
                         f"Inv: {r['inversion']:,.0f}M$\n"
                         f"Evitadas: {len(lineas_evitadas)}",
                         fontsize=8, fontweight='bold')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['bottom'].set_visible(False)

    # Leyenda global
    leg_items = [
        mpatches.Patch(color='#C62828', label='Línea construida (necesaria)'),
        mpatches.Patch(color='#2E7D32', label='Línea EVITADA por BTM'),
    ]
    fig.legend(handles=leg_items, loc='lower center', ncol=2, fontsize=10,
               bbox_to_anchor=(0.5, 0.01))
    plt.tight_layout(rect=[0, 0.04, 1, 0.95])
    fname = "BTM_Fig04_Lineas_por_Escenario.png"
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig(fname), dpi=160, bbox_inches='tight')
        print(f"  → Guardada: {fname}")
    plt.show(); plt.close()

def figura_05_heatmap_cargabilidad(resultados_uniforme, resultados_cei):
    """
    FIG 5: Heatmap de cargabilidad de corredores por escenario BTM.
    Filas = corredores, Columnas = nivel de penetración α.
    El gradiente de color muestra cómo la cargabilidad disminuye con BTM.
    """
    fig, axes = plt.subplots(1, 2, figsize=(16, 8))
    fig.suptitle(
        "FIG 5 — Heatmap de Cargabilidad de Corredores STN 500 kV\n"
        "Impacto de la penetración BTM sobre la utilización de la red  |  Uniandes 2026",
        fontsize=11, fontweight='bold'
    )

    # Solo corredores con circuitos activos (N0 > 0 o nueva en algún escenario)
    corredores_activos = [j for j in CORREDORES
                         if N0[j] > 0 or any(resultados_uniforme[a] and j in resultados_uniforme[a]['nuevas_lineas']
                                             for a in ALPHAS)]

    for ax, (res_dict, title) in zip(
        axes,
        [(resultados_uniforme, 'Reducción Uniforme'),
         (resultados_cei,      'Reducción CEI Inteligente')]
    ):
        mat = np.zeros((len(corredores_activos), len(ALPHAS)))
        for ci, j in enumerate(corredores_activos):
            for ai, alpha in enumerate(ALPHAS):
                r = res_dict[alpha]
                if r:
                    mat[ci, ai] = r['cargabilidad'].get(j, 0) * 100

        im = ax.imshow(mat, aspect='auto', cmap='RdYlGn_r', vmin=0, vmax=100)
        ax.set_xticks(range(len(ALPHAS)))
        ax.set_xticklabels([f"α={int(a*100)}%" for a in ALPHAS], fontsize=9)
        ax.set_yticks(range(len(corredores_activos)))
        ax.set_yticklabels([f"[{j}] {DESC[j][:28]}" for j in corredores_activos], fontsize=7.5)
        ax.set_title(f"Estrategia: {title}", fontsize=10)
        ax.set_xlabel("Penetración BTM (α)", fontsize=9)

        # Anotar valores
        for ci in range(len(corredores_activos)):
            for ai in range(len(ALPHAS)):
                v = mat[ci, ai]
                color_txt = 'white' if v > 70 else 'black'
                ax.text(ai, ci, f"{v:.0f}%",
                        ha='center', va='center', fontsize=7.5,
                        color=color_txt, fontweight='bold')

        plt.colorbar(im, ax=ax, label='Cargabilidad (%)', shrink=0.85)

        # Líneas de umbral (95% y 80%)
        # No se pueden añadir directamente al imshow, pero se marcan con borde
        for ci in range(len(corredores_activos)):
            for ai in range(len(ALPHAS)):
                v = mat[ci, ai]
                if v > 95:
                    ax.add_patch(plt.Rectangle((ai-0.5, ci-0.5), 1, 1,
                                               fill=False, edgecolor='red', lw=2.5))
                elif v > 80:
                    ax.add_patch(plt.Rectangle((ai-0.5, ci-0.5), 1, 1,
                                               fill=False, edgecolor='orange', lw=1.5))

    plt.tight_layout()
    fname = "BTM_Fig05_Heatmap_Cargabilidad.png"
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig(fname), dpi=160, bbox_inches='tight')
        print(f"  → Guardada: {fname}")
    plt.show(); plt.close()

def figura_06_ahorro_acumulado(df_ahorro):
    """
    FIG 6: Curvas de ahorro acumulado en inversión de transmisión
    vs. penetración BTM. Muestra el valor económico de las comunidades
    energéticas como sustituto de inversión en líneas de transmisión.
    """
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    fig.suptitle(
        "FIG 6 — Ahorro en Inversión de Transmisión por Penetración BTM\n"
        "Las comunidades energéticas como recurso virtual de expansión  |  Uniandes 2026",
        fontsize=11, fontweight='bold'
    )

    alpha_pcts = [int(a*100) for a in ALPHAS]
    ahorros_u  = df_ahorro['Ahorro_Unif'].fillna(0).tolist()
    ahorros_c  = df_ahorro['Ahorro_CEI'].fillna(0).tolist()

    ax = axes[0]
    ax.fill_between(alpha_pcts, 0, ahorros_u, alpha=0.25, color='#1565c0')
    ax.fill_between(alpha_pcts, 0, ahorros_c, alpha=0.25, color='#F44336')
    ax.plot(alpha_pcts, ahorros_u, 'o-', color='#1565c0', lw=2.5, ms=9, label='Uniforme')
    ax.plot(alpha_pcts, ahorros_c, 's-', color='#F44336', lw=2.5, ms=9, label='CEI Inteligente')
    for i, (p, au, ac) in enumerate(zip(alpha_pcts, ahorros_u, ahorros_c)):
        if au > 0:
            ax.annotate(f"{au:,.0f}M$", (p, au), textcoords="offset points",
                        xytext=(0, 10), ha='center', fontsize=8.5, color='#1565c0')
        if ac > 0:
            ax.annotate(f"{ac:,.0f}M$", (p, ac), textcoords="offset points",
                        xytext=(0, -18), ha='center', fontsize=8.5, color='#F44336')
    ax.set_xlabel("Penetración BTM α (%)", fontsize=10)
    ax.set_ylabel("Ahorro en inversión (MUSD)", fontsize=10)
    ax.set_title("Curva de ahorro acumulado vs. penetración")
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.35)
    ax.set_ylim(bottom=0)

    # Panel 2: % de ahorro relativo
    ax2 = axes[1]
    base_inv = df_ahorro[df_ahorro['Alpha']==0.0]['Inv_Unif'].values[0] or 1
    pct_u = [au/base_inv*100 for au in ahorros_u]
    pct_c = [ac/base_inv*100 for ac in ahorros_c]
    ax2.plot(alpha_pcts, pct_u, 'o-', color='#1565c0', lw=2.5, ms=9, label='Uniforme')
    ax2.plot(alpha_pcts, pct_c, 's-', color='#F44336', lw=2.5, ms=9, label='CEI Inteligente')
    ax2.fill_between(alpha_pcts, pct_u, pct_c,
                     where=[c >= u for c, u in zip(pct_c, pct_u)],
                     alpha=0.2, color='#F44336',
                     label='Ganancia adicional CEI vs. Uniforme')
    ax2.set_xlabel("Penetración BTM α (%)", fontsize=10)
    ax2.set_ylabel("Ahorro relativo respecto a caso base (%)", fontsize=10)
    ax2.set_title("Efectividad relativa de la reducción BTM")
    ax2.legend(fontsize=9)
    ax2.grid(True, alpha=0.35)
    ax2.set_ylim(0, max(max(pct_u), max(pct_c))*1.25)
    for p, pu, pc in zip(alpha_pcts, pct_u, pct_c):
        ax2.annotate(f"{pu:.1f}%", (p, pu), textcoords="offset points",
                     xytext=(0, 8), ha='center', fontsize=8, color='#1565c0')

    plt.tight_layout()
    fname = "BTM_Fig06_Ahorro_Acumulado.png"
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig(fname), dpi=160, bbox_inches='tight')
        print(f"  → Guardada: {fname}")
    plt.show(); plt.close()

def figura_07_comparacion_uniforme_vs_cei(resultados_uniforme, resultados_cei, df_ahorro):
    """
    FIG 7: Comparación directa entre estrategia uniforme e inteligente (CEI).
    Muestra la ganancia adicional de asignar la reducción de demanda
    de forma prioritaria según el índice CEI.
    """
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    fig.suptitle(
        "FIG 7 — Comparación Estrategia Uniforme vs. CEI Inteligente\n"
        "La asignación óptima de recursos BTM maximiza el ahorro en transmisión  |  Uniandes 2026",
        fontsize=11, fontweight='bold'
    )

    alpha_pcts = [int(a*100) for a in ALPHAS]

    # ── Panel (a): Inversión total por estrategia ─────────────────────────────
    ax = axes[0][0]
    inv_u = df_ahorro['Inv_Unif'].fillna(0).tolist()
    inv_c = df_ahorro['Inv_CEI'].fillna(0).tolist()
    x = np.arange(len(ALPHAS))
    w = 0.35
    b1 = ax.bar(x-w/2, inv_u, w, color='#1565c0', alpha=0.85, label='Uniforme')
    b2 = ax.bar(x+w/2, inv_c, w, color='#F44336', alpha=0.85, label='CEI Inteligente')
    ax.set_xticks(x)
    ax.set_xticklabels([f"α={p}%" for p in alpha_pcts], fontsize=9)
    ax.set_ylabel("Inversión (MUSD)")
    ax.set_title("(a) Inversión óptima por escenario")
    ax.legend(fontsize=9)
    ax.grid(True, alpha=0.3, axis='y')
    for bar in list(b1) + list(b2):
        if bar.get_height() > 0:
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+10,
                    f"{bar.get_height():,.0f}", ha='center', va='bottom', fontsize=7.5)

    # ── Panel (b): Demanda total reducida ─────────────────────────────────────
    ax = axes[0][1]
    dem_u = df_ahorro['Dem_Unif'].fillna(0).tolist()
    dem_c = df_ahorro['Dem_CEI'].fillna(0).tolist()
    ax.plot(alpha_pcts, dem_u, 'o-', color='#1565c0', lw=2, ms=8, label='Uniforme')
    ax.plot(alpha_pcts, dem_c, 's--', color='#F44336', lw=2, ms=8, label='CEI Inteligente')
    ax.set_xlabel("Penetración BTM α (%)")
    ax.set_ylabel("Demanda total (MW)")
    ax.set_title("(b) Demanda neta del sistema por escenario")
    ax.legend(fontsize=9)
    ax.grid(True, alpha=0.35)

    # ── Panel (c): Número de líneas evitadas ─────────────────────────────────
    ax = axes[1][0]
    base_lineas_u = set(resultados_uniforme[0.0]['nuevas_lineas'].keys()) if resultados_uniforme[0.0] else set()
    base_lineas_c = set(resultados_cei[0.0]['nuevas_lineas'].keys())      if resultados_cei[0.0]      else set()
    n_evitadas_u  = [len(base_lineas_u - set(resultados_uniforme[a]['nuevas_lineas'].keys()))
                     if resultados_uniforme[a] else 0 for a in ALPHAS]
    n_evitadas_c  = [len(base_lineas_c - set(resultados_cei[a]['nuevas_lineas'].keys()))
                     if resultados_cei[a] else 0 for a in ALPHAS]
    ax.step(alpha_pcts, n_evitadas_u, where='post', color='#1565c0', lw=2.5, label='Uniforme')
    ax.step(alpha_pcts, n_evitadas_c, where='post', color='#F44336', lw=2.5, ls='--', label='CEI Inteligente')
    ax.fill_between(alpha_pcts, n_evitadas_u, step='post', alpha=0.2, color='#1565c0')
    ax.fill_between(alpha_pcts, n_evitadas_c, step='post', alpha=0.2, color='#F44336')
    ax.set_xlabel("Penetración BTM α (%)")
    ax.set_ylabel("Corredores evitados (vs. α=0%)")
    ax.set_title("(c) Corredores que dejan de construirse")
    ax.legend(fontsize=9)
    ax.grid(True, alpha=0.3)
    ax.yaxis.set_major_locator(plt.MaxNLocator(integer=True))

    # ── Panel (d): Ventaja CEI sobre Uniforme (MUSD adicionales ahorrados) ───
    ax = axes[1][1]
    ganancia = [max(0, (df_ahorro[df_ahorro['Alpha']==a]['Ahorro_CEI'].values[0] or 0)
                    - (df_ahorro[df_ahorro['Alpha']==a]['Ahorro_Unif'].values[0] or 0))
                for a in ALPHAS]
    colores_g = ['#2E7D32' if g > 0 else '#757575' for g in ganancia]
    bars = ax.bar(alpha_pcts, ganancia, width=3.5, color=colores_g, alpha=0.85)
    ax.set_xlabel("Penetración BTM α (%)")
    ax.set_ylabel("MUSD adicionales ahorrados")
    ax.set_title("(d) Ganancia de usar CEI vs. reducción uniforme")
    ax.grid(True, alpha=0.3, axis='y')
    for bar, g in zip(bars, ganancia):
        if g > 0:
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+2,
                    f"+{g:,.0f}M$", ha='center', va='bottom', fontsize=9, fontweight='bold')
    ax.text(0.98, 0.97, "CEI inteligente siempre\n≥ reducción uniforme",
            transform=ax.transAxes, ha='right', va='top',
            fontsize=9, color='#2E7D32',
            bbox=dict(boxstyle='round', facecolor='#E8F5E9', alpha=0.8))

    plt.tight_layout()
    fname = "BTM_Fig07_Uniforme_vs_CEI.png"
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig(fname), dpi=160, bbox_inches='tight')
        print(f"  → Guardada: {fname}")
    plt.show(); plt.close()

def figura_08_red_comparativa(resultados_uniforme, alpha_ref=0.10):
    """
    FIG 8: Mapa de la red — comparación caso base vs. α=10% uniforme.
    Visualiza qué líneas dejan de construirse con BTM y qué corredores
    se descongestionen.
    """
    try:
        import networkx as nx
    except ImportError:
        print("  (pip install networkx para activar las gráficas de red)")
        return

    r_base = resultados_uniforme[0.0]
    r_btm  = resultados_uniforme[alpha_ref]
    if r_base is None or r_btm is None:
        return

    lineas_evitadas = set(r_base['nuevas_lineas'].keys()) - set(r_btm['nuevas_lineas'].keys())

    fig, axes = plt.subplots(1, 2, figsize=(20, 9))
    fig.suptitle(
        f"FIG 8 — Red STN 500 kV: Caso Base (α=0%) vs. BTM (α={int(alpha_ref*100)}% Uniforme)\n"
        f"Líneas verdes = EVITADAS por comunidades energéticas BTM  |  Uniandes 2026",
        fontsize=10, fontweight='bold'
    )
    titulos = [f'Caso Base α=0%\nInv: {r_base["inversion"]:,.0f} MUSD',
               f'Con BTM α={int(alpha_ref*100)}%\nInv: {r_btm["inversion"]:,.0f} MUSD\n'
               f'Ahorro: {r_base["inversion"]-r_btm["inversion"]:+,.0f} MUSD']

    for idx, (ax, r, titulo) in enumerate(zip(axes, [r_base, r_btm], titulos)):
        G = nx.DiGraph()
        G.add_nodes_from(NODOS)
        ax.axis('off')
        ax.set_title(titulo, fontsize=9, fontweight='bold')

        nx.draw_networkx_nodes(G, POS, ax=ax,
                               node_size=3000,
                               node_color=[COL_NODOS[n] for n in NODOS],
                               alpha=0.90)
        lbl = {n: f"N{n}\n{NOMBRE[n]}\nD={r['demanda'][n]:.0f}MW\nG={r['generacion'].get(n,0):.0f}MW"
               for n in NODOS}
        nx.draw_networkx_labels(G, POS, lbl, ax=ax, font_size=6.5)

        for j in CORREDORES:
            n_new = r['nuevas_lineas'].get(j, 0)
            n_tot = N0[j] + n_new
            if n_tot == 0:
                continue
            fl   = r['flujos'].get(j, 0)
            cap  = FMAX[j]*n_tot
            carg = abs(fl)/cap if cap > 0 else 0
            ec   = '#d32f2f' if carg > 0.95 else ('#f57c00' if carg > 0.80 else '#1565c0')
            w    = 1.5 + carg*4
            a, b = map(int, j.split(','))
            nx.draw_networkx_edges(G, POS, edgelist=[(a,b)], ax=ax,
                                   width=w, edge_color=ec, alpha=0.85,
                                   arrows=True, arrowsize=14,
                                   connectionstyle='arc3,rad=0.08')
            mx = (POS[a][0]+POS[b][0])/2
            my = (POS[a][1]+POS[b][1])/2 + 0.03
            ax.text(mx, my, f"{fl:.0f}MW\n{carg*100:.0f}%",
                    fontsize=5.5, ha='center', color=ec)

        # Nuevas líneas construidas
        for j, nc in r['nuevas_lineas'].items():
            if N0[j] == 0:
                a, b = map(int, j.split(','))
                nx.draw_networkx_edges(G, POS, edgelist=[(a,b)], ax=ax,
                                       width=4.5, edge_color='#00C853',
                                       alpha=0.95, arrows=True, arrowsize=20,
                                       style='dashed',
                                       connectionstyle='arc3,rad=-0.18')
                mx = (POS[a][0]+POS[b][0])/2
                my = (POS[a][1]+POS[b][1])/2 - 0.06
                ax.text(mx, my, f"NUEVA {nc}×\n{COSTO[j]*nc:.0f}M$",
                        fontsize=6, ha='center', color='#00C853',
                        fontweight='bold',
                        bbox=dict(boxstyle='round,pad=0.2', facecolor='#B2DFDB', alpha=0.9))

        # Marcar líneas evitadas (solo en panel BTM)
        if idx == 1:
            for j in lineas_evitadas:
                a, b = map(int, j.split(','))
                nx.draw_networkx_edges(G, POS, edgelist=[(a,b)], ax=ax,
                                       width=5, edge_color='#76FF03',
                                       alpha=0.70, arrows=False,
                                       style='dotted',
                                       connectionstyle='arc3,rad=0.25')
                mx = (POS[a][0]+POS[b][0])/2
                my = (POS[a][1]+POS[b][1])/2 + 0.08
                ax.text(mx, my, f"[BTM]\nEVITADA\n{COSTO[j]:.0f}M$",
                        fontsize=6.5, ha='center', color='#33691E',
                        fontweight='bold',
                        bbox=dict(boxstyle='round,pad=0.3', facecolor='#DCEDC8', alpha=0.95))

    leyenda = [
        mpatches.Patch(color='#1565c0', label='Línea OK (<80%)'),
        mpatches.Patch(color='#f57c00', label='Alta carga (80–95%)'),
        mpatches.Patch(color='#d32f2f', label='Congestionada (>95%)'),
        mpatches.Patch(color='#00C853', label='Nueva línea construida'),
        mpatches.Patch(color='#76FF03', label='Línea EVITADA por BTM'),
    ]
    fig.legend(handles=leyenda, loc='lower center', ncol=5, fontsize=9,
               bbox_to_anchor=(0.5, 0.01))
    plt.tight_layout(rect=[0, 0.05, 1, 0.95])
    fname = f"BTM_Fig08_Red_Base_vs_BTM{int(alpha_ref*100)}.png"
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig(fname), dpi=180, bbox_inches='tight')
        print(f"  → Guardada: {fname}")
    plt.show(); plt.close()

def figura_09_sensibilidad_corredores(resultados_uniforme):
    """
    FIG 9: Sensibilidad de cada corredor a la reducción BTM.
    Identifica qué corredores son más sensibles (su cargabilidad cambia más
    por unidad de penetración BTM). Los más sensibles son los que más se
    benefician de las comunidades energéticas.
    """
    # Calcular variación de cargabilidad entre α=0% y α=20%
    r0  = resultados_uniforme[0.00]
    r20 = resultados_uniforme[0.20]
    if r0 is None or r20 is None:
        return

    corr_activos = [j for j in CORREDORES if N0[j] > 0
                    or j in r0['nuevas_lineas'] or j in r20['nuevas_lineas']]

    delta_carg = {}
    carg_base  = {}
    for j in corr_activos:
        c0  = r0['cargabilidad'].get(j, 0) * 100
        c20 = r20['cargabilidad'].get(j, 0) * 100
        delta_carg[j] = c0 - c20    # cuánto bajó la cargabilidad (positivo = mejora)
        carg_base[j]  = c0

    # Ordenar por sensibilidad (delta mayor = más sensible a BTM)
    corr_sorted = sorted(corr_activos, key=lambda j: delta_carg[j], reverse=True)

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 7))
    fig.suptitle(
        "FIG 9 — Sensibilidad de Corredores a la Reducción BTM\n"
        "Δ Cargabilidad entre α=0% y α=20% — Corredores más sensibles = mayor beneficio BTM  |  Uniandes 2026",
        fontsize=11, fontweight='bold'
    )

    # Panel (a): Delta cargabilidad
    colores_d = ['#2E7D32' if delta_carg[j] > 0 else '#C62828' for j in corr_sorted]
    bars = ax1.barh(range(len(corr_sorted)),
                    [delta_carg[j] for j in corr_sorted],
                    color=colores_d, alpha=0.85)
    ax1.set_yticks(range(len(corr_sorted)))
    ax1.set_yticklabels([f"[{j}] {DESC[j][:28]}" for j in corr_sorted], fontsize=7.5)
    ax1.set_xlabel("Δ Cargabilidad (%, α=0%→20%)")
    ax1.set_title("(a) Mejora de cargabilidad por BTM")
    ax1.axvline(0, color='gray', lw=0.8)
    ax1.grid(True, alpha=0.3, axis='x')
    for bar, j in zip(bars, corr_sorted):
        v = delta_carg[j]
        x_pos = v + 0.5 if v >= 0 else v - 0.5
        ax1.text(x_pos, bar.get_y()+bar.get_height()/2,
                 f"{v:+.1f}%", va='center', fontsize=8,
                 color='#2E7D32' if v > 0 else '#C62828')

    # Panel (b): Spider de cargabilidad por escenario (corredores top 8)
    top_corr = corr_sorted[:min(8, len(corr_sorted))]
    for alpha in ALPHAS:
        r = resultados_uniforme[alpha]
        if r is None:
            continue
        cargs = [r['cargabilidad'].get(j, 0)*100 for j in top_corr]
        ax2.plot(range(len(top_corr)), cargs,
                 'o-', lw=1.8, ms=6,
                 color=ALPHA_COLORS_U[alpha],
                 label=ALPHA_LABELS[alpha])

    ax2.axhline(95, color='red',    lw=1.5, ls='--', alpha=0.7, label='Límite crítico (95%)')
    ax2.axhline(80, color='orange', lw=1.2, ls='--', alpha=0.5)
    ax2.set_xticks(range(len(top_corr)))
    ax2.set_xticklabels([f"[{j}]" for j in top_corr], fontsize=9, rotation=30, ha='right')
    ax2.set_ylabel("Cargabilidad (%)")
    ax2.set_title("(b) Evolución de cargabilidad — top corredores")
    ax2.legend(fontsize=8.5, loc='upper right')
    ax2.grid(True, alpha=0.3)
    ax2.set_ylim(0, 115)

    plt.tight_layout()
    fname = "BTM_Fig09_Sensibilidad_Corredores.png"
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig(fname), dpi=160, bbox_inches='tight')
        print(f"  → Guardada: {fname}")
    plt.show(); plt.close()

def figura_10_generacion_despachada(resultados_uniforme, resultados_cei):
    """
    FIG 10: Generación despachada por nodo y por escenario BTM.
    La reducción de demanda modifica el despacho óptimo del sistema.
    """
    fig, axes = plt.subplots(1, 2, figsize=(14, 7))
    fig.suptitle(
        "FIG 10 — Generación Despachada por Nodo y Escenario BTM\n"
        "La reducción de demanda modifica el despacho óptimo del sistema  |  Uniandes 2026",
        fontsize=11, fontweight='bold'
    )

    for ax, (res_dict, title) in zip(
        axes,
        [(resultados_uniforme, 'Reducción Uniforme'),
         (resultados_cei,      'Reducción CEI Inteligente')]
    ):
        x = np.arange(len(NODOS))
        width = 0.14
        for ai, alpha in enumerate(ALPHAS):
            r = res_dict[alpha]
            if r is None:
                continue
            gens = [r['generacion'].get(i, 0) for i in NODOS]
            offset = (ai - len(ALPHAS)/2) * width
            ax.bar(x + offset, gens, width,
                   label=ALPHA_LABELS[alpha],
                   color=list(ALPHA_COLORS_U.values())[ai], alpha=0.85)

        ax.set_xticks(x)
        ax.set_xticklabels([f"N{i}\n{NOMBRE[i]}" for i in NODOS], fontsize=9)
        ax.set_ylabel("Generación despachada (MW)")
        ax.set_title(f"Estrategia: {title}")
        ax.legend(fontsize=8.5, loc='upper right')
        ax.grid(True, alpha=0.3, axis='y')

        # Líneas de capacidad máxima
        for i, xi in enumerate(x):
            ax.plot([xi-0.35, xi+0.35], [G_2039[i+1], G_2039[i+1]],
                    'k--', lw=1.0, alpha=0.4)

    plt.tight_layout()
    fname = "BTM_Fig10_Generacion_Despachada.png"
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig(fname), dpi=160, bbox_inches='tight')
        print(f"  → Guardada: {fname}")
    plt.show(); plt.close()

def figura_11_multiperiodo_deferral(resultados_multi):
    """
    FIG 11: Análisis multiperiodo — costo diferido por BTM en el tiempo.
    Muestra cómo la expansión se posterga a años más lejanos cuando
    el sistema cuenta con comunidades energéticas BTM.
    """
    if len(resultados_multi) < 2:
        return

    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    fig.suptitle(
        "FIG 11 — Análisis Multiperiodo: Diferimiento de Inversión por BTM\n"
        "Las comunidades energéticas postergan la expansión al futuro  |  Uniandes 2026",
        fontsize=11, fontweight='bold'
    )

    anos = [ANO_INI + t - 1 for t in PERIODOS]
    colores_m = ['#1565c0', '#F44336', '#2E7D32', '#FF9800']

    # ── Panel (a): Inversión anual por escenario ──────────────────────────────
    ax = axes[0][0]
    for rm, col in zip(resultados_multi, colores_m):
        invs = [rm['inversion_por_t'][t] for t in PERIODOS]
        ax.bar(anos, invs, alpha=0.60, color=col,
               label=rm['titulo'], width=0.6,
               bottom=[sum(r['inversion_por_t'][tt] for r in resultados_multi[:i] if r != rm
                           for ttt in PERIODOS if ttt == t) for t in PERIODOS]
               if False else None)  # barras individuales, no apiladas
    # Redibujamos limpio
    ax.cla()
    x_base = np.arange(len(PERIODOS))
    w = 0.22
    for ai, (rm, col) in enumerate(zip(resultados_multi, colores_m)):
        invs = [rm['inversion_por_t'][t] for t in PERIODOS]
        offset = (ai - len(resultados_multi)/2) * w
        ax.bar(x_base + offset, invs, w, color=col, alpha=0.80, label=rm['titulo'])
    ax.set_xticks(x_base[::2])
    ax.set_xticklabels([anos[i] for i in range(0, len(anos), 2)], rotation=45, fontsize=8)
    ax.set_ylabel("Inversión anual (MUSD)")
    ax.set_title("(a) Inversión anual por escenario BTM")
    ax.legend(fontsize=7.5, loc='upper left')
    ax.grid(True, alpha=0.3, axis='y')

    # ── Panel (b): Inversión acumulada ────────────────────────────────────────
    ax = axes[0][1]
    for rm, col in zip(resultados_multi, colores_m):
        invs   = [rm['inversion_por_t'][t] for t in PERIODOS]
        inv_ac = np.cumsum(invs)
        ax.plot(anos, inv_ac, 'o-', color=col, lw=2.5, ms=6, label=rm['titulo'])
        ax.fill_between(anos, inv_ac, alpha=0.10, color=col)
    ax.set_xlabel("Año")
    ax.set_ylabel("Inversión acumulada (MUSD)")
    ax.set_title("(b) Curvas de inversión acumulada (VPN)")
    ax.legend(fontsize=7.5)
    ax.grid(True, alpha=0.35)

    # ── Panel (c): Demanda total por escenario ────────────────────────────────
    ax = axes[1][0]
    for rm, col in zip(resultados_multi, colores_m):
        dems = [rm['demanda_por_t'][t] for t in PERIODOS]
        ax.plot(anos, dems, 'o-', color=col, lw=2, ms=5, label=rm['titulo'])
    ax.set_xlabel("Año")
    ax.set_ylabel("Demanda total (MW)")
    ax.set_title("(c) Evolución de la demanda neta con BTM")
    ax.legend(fontsize=7.5)
    ax.grid(True, alpha=0.35)
    ax.set_xlim(ANO_INI-0.5, ANO_INI+NT-0.5)

    # ── Panel (d): Ahorro de VPN acumulado vs. caso base ─────────────────────
    ax = axes[1][1]
    base_res = resultados_multi[0]
    base_ac  = np.cumsum([base_res['inversion_por_t'][t] for t in PERIODOS])
    for rm, col in zip(resultados_multi[1:], colores_m[1:]):
        inv_ac  = np.cumsum([rm['inversion_por_t'][t] for t in PERIODOS])
        ahorro  = base_ac - inv_ac
        ax.plot(anos, ahorro, 'o-', color=col, lw=2.5, ms=6,
                label=f"Ahorro BTM vs. base\n({rm['titulo']})")
        ax.fill_between(anos, ahorro, 0, where=[a > 0 for a in ahorro],
                        alpha=0.15, color=col)
        ax.fill_between(anos, ahorro, 0, where=[a < 0 for a in ahorro],
                        alpha=0.10, color='red')
    ax.axhline(0, color='gray', lw=1.0, ls='--')
    ax.set_xlabel("Año")
    ax.set_ylabel("Ahorro acumulado vs. α=0% (MUSD)")
    ax.set_title("(d) Costo diferido / evitado en el tiempo")
    ax.legend(fontsize=7.5)
    ax.grid(True, alpha=0.35)

    plt.tight_layout()
    fname = "BTM_Fig11_Multiperiodo_Diferido.png"
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig(fname), dpi=160, bbox_inches='tight')
        print(f"  → Guardada: {fname}")
    plt.show(); plt.close()

def figura_12_resumen_ejecutivo(df_ahorro, df_cei, resultados_uniforme):
    """
    FIG 12: Resumen ejecutivo de todos los resultados del análisis BTM.
    Infografía académica integrada con los indicadores clave del modelo.
    """
    fig = plt.figure(figsize=(18, 12))
    gs  = GridSpec(3, 3, figure=fig, hspace=0.45, wspace=0.35)
    fig.suptitle(
        "FIG 12 — Resumen Ejecutivo: TEP con Comunidades Energéticas BTM — SIN Colombia 500 kV\n"
        "Inversión evitada, diferida y costo comparado  |  Uniandes 2026",
        fontsize=12, fontweight='bold'
    )

    # ── Métricas clave ────────────────────────────────────────────────────────
    base_inv  = df_ahorro[df_ahorro['Alpha']==0.0]['Inv_Unif'].values[0] or 0
    max_ahorro_u = df_ahorro['Ahorro_Unif'].max() or 0
    max_ahorro_c = df_ahorro['Ahorro_CEI'].max()  or 0
    alpha_umbral = ALPHAS[next((i for i, a in enumerate(
                                df_ahorro['Ahorro_Unif'].tolist())
                                if a is not None and a > 0), -1)]

    # ── Subgráfico 1: Resumen indicadores en celdas de texto ─────────────────
    ax0 = fig.add_subplot(gs[0, :])
    ax0.axis('off')
    metricas = [
        ("Inversión base\n(α=0%)",       f"{base_inv:,.0f} MUSD",   '#C62828'),
        ("Ahorro máximo\n(uniforme)",     f"{max_ahorro_u:,.0f} MUSD", '#1565c0'),
        ("Ahorro máximo\n(CEI)",          f"{max_ahorro_c:,.0f} MUSD", '#2E7D32'),
        ("Ganancia CEI\nvs. Uniforme",    f"{max(0, max_ahorro_c-max_ahorro_u):,.0f} MUSD", '#FF9800'),
        ("Sistema  SIN\n2039 escenario",  "Alta demanda\nCONGESTIÓN",   '#7B1FA2'),
        ("Nodo más prioritario\n(CEI)",   df_cei.sort_values('CEI',ascending=False).iloc[0]['Nombre'],
                                          '#F9A825'),
    ]
    for idx, (label, value, color) in enumerate(metricas):
        x0 = 0.02 + idx * 0.165
        ax0.add_patch(plt.Rectangle((x0, 0.05), 0.155, 0.90,
                                    transform=ax0.transAxes,
                                    facecolor=color, alpha=0.12,
                                    edgecolor=color, lw=2))
        ax0.text(x0+0.078, 0.68, label,  transform=ax0.transAxes,
                 ha='center', va='center', fontsize=9, color='#333',
                 fontweight='bold')
        ax0.text(x0+0.078, 0.30, value,  transform=ax0.transAxes,
                 ha='center', va='center', fontsize=11, color=color,
                 fontweight='bold')

    # ── Subgráfico 2: Inversión por escenario ─────────────────────────────────
    ax1 = fig.add_subplot(gs[1, 0])
    inv_u = df_ahorro['Inv_Unif'].fillna(0).tolist()
    inv_c = df_ahorro['Inv_CEI'].fillna(0).tolist()
    x = np.arange(len(ALPHAS))
    ax1.bar(x-0.2, inv_u, 0.4, color='#1565c0', alpha=0.85, label='Uniforme')
    ax1.bar(x+0.2, inv_c, 0.4, color='#F44336', alpha=0.85, label='CEI')
    ax1.set_xticks(x)
    ax1.set_xticklabels([f"α={int(a*100)}%" for a in ALPHAS], fontsize=8)
    ax1.set_ylabel("MUSD", fontsize=8)
    ax1.set_title("Inversión óptima", fontsize=9, fontweight='bold')
    ax1.legend(fontsize=7.5)
    ax1.grid(True, alpha=0.3, axis='y')

    # ── Subgráfico 3: Curva ahorro ─────────────────────────────────────────────
    ax2 = fig.add_subplot(gs[1, 1])
    alpha_pcts = [int(a*100) for a in ALPHAS]
    ax2.fill_between(alpha_pcts, 0, df_ahorro['Ahorro_Unif'].fillna(0).tolist(),
                     alpha=0.3, color='#1565c0')
    ax2.fill_between(alpha_pcts, 0, df_ahorro['Ahorro_CEI'].fillna(0).tolist(),
                     alpha=0.3, color='#F44336')
    ax2.plot(alpha_pcts, df_ahorro['Ahorro_Unif'].fillna(0).tolist(),
             'o-', color='#1565c0', lw=2, label='Uniforme')
    ax2.plot(alpha_pcts, df_ahorro['Ahorro_CEI'].fillna(0).tolist(),
             's-', color='#F44336', lw=2, label='CEI')
    ax2.set_xlabel("α (%)", fontsize=8)
    ax2.set_ylabel("MUSD ahorrados", fontsize=8)
    ax2.set_title("Curva de ahorro BTM", fontsize=9, fontweight='bold')
    ax2.legend(fontsize=7.5)
    ax2.grid(True, alpha=0.3)

    # ── Subgráfico 4: Ranking CEI ──────────────────────────────────────────────
    ax3 = fig.add_subplot(gs[1, 2])
    df_s = df_cei.sort_values('CEI', ascending=True)
    bars = ax3.barh(range(len(df_s)), df_s['CEI'].tolist(),
                    color=[COL_NODOS[int(r['Node'])] for _, r in df_s.iterrows()],
                    alpha=0.88)
    ax3.set_yticks(range(len(df_s)))
    ax3.set_yticklabels([f"N{int(r['Node'])} {r['Nombre']}" for _, r in df_s.iterrows()], fontsize=8)
    ax3.set_xlabel("CEI", fontsize=8)
    ax3.set_title("Ranking priorización CEI", fontsize=9, fontweight='bold')
    ax3.grid(True, alpha=0.3, axis='x')

    # ── Subgráfico 5: Heatmap cargabilidad compacto ────────────────────────────
    ax4 = fig.add_subplot(gs[2, :2])
    corredores_act = [j for j in CORREDORES if N0[j] > 0]
    mat = np.zeros((len(corredores_act), len(ALPHAS)))
    for ci, j in enumerate(corredores_act):
        for ai, alpha in enumerate(ALPHAS):
            r = resultados_uniforme[alpha]
            if r:
                mat[ci, ai] = r['cargabilidad'].get(j, 0)*100
    im = ax4.imshow(mat, aspect='auto', cmap='RdYlGn_r', vmin=0, vmax=100)
    ax4.set_xticks(range(len(ALPHAS)))
    ax4.set_xticklabels([f"α={int(a*100)}%" for a in ALPHAS], fontsize=8)
    ax4.set_yticks(range(len(corredores_act)))
    ax4.set_yticklabels([f"[{j}] {DESC[j][:22]}" for j in corredores_act], fontsize=7)
    ax4.set_title("Cargabilidad de corredores existentes vs. BTM (Uniforme)", fontsize=9, fontweight='bold')
    for ci in range(len(corredores_act)):
        for ai in range(len(ALPHAS)):
            v = mat[ci, ai]
            ax4.text(ai, ci, f"{v:.0f}", ha='center', va='center',
                     fontsize=7, color='white' if v > 65 else 'black')
    plt.colorbar(im, ax=ax4, label='%', shrink=0.8, orientation='horizontal',
                 pad=0.15)

    # # ── Subgráfico 6: Nota metodológica ───────────────────────────────────────
    # ax5 = fig.add_subplot(gs[2, 2])
    # ax5.axis('off')
    # nota = (
    #     "METODOLOGÍA\n\n"
    #     "Modelo: TEP MILP DC — Pyomo + Gurobi\n"
    #     "Sistema: SIN Colombia backbone 500 kV\n"
    #     "         6 nodos regionales, 15 corredores\n\n"
    #     "BTM Demand Reduction:\n"
    #     "  D_nuevo = D_base × (1 − α)\n\n"
    #     "Estrategias:\n"
    #     "  • Uniforme: α_i = α_total\n"
    #     "  • CEI: α_i = α × (CEI_i / CEI_mean)\n\n"
    #     "Indicadores nodales:\n"
    #     "  CEI = 0.4·SPI + 0.3·NSI + 0.3·GRI\n"
    #     "  SPI: potencial solar (IDEAM)\n"
    #     "  NSI: estrés red (déficit/demanda)\n"
    #     "  GRI: carga corredores adyacentes\n\n"
    #     "Fuentes: UPME PIEG 2025-39 | PET 2022-36\n"
    #     "         DIgSILENT PF 2025 | XM 2024"
    # )
    # ax5.text(0.05, 0.98, nota,
    #          transform=ax5.transAxes, ha='left', va='top',
    #          fontsize=8, fontfamily='monospace',
    #          bbox=dict(boxstyle='round', facecolor='#F5F5F5', alpha=0.9, edgecolor='#999'))

    fname = "BTM_Fig12_Resumen_Ejecutivo.png"
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig(fname), dpi=160, bbox_inches='tight')
        print(f"  → Guardada: {fname}")
    plt.show(); plt.close()

# =============================================================================
# 8b. EXPORTACIÓN SISTEMÁTICA DE TABLAS DE RESULTADOS
# =============================================================================

def exportar_tabla_cei(df_cei):
    """Exporta los índices CEI nodales a CSV y al Excel maestro."""
    df = df_cei.copy()
    df.to_csv(fpath_tbl("Tabla_01_indices_CEI_nodales.csv"), index=False,
              float_format="%.4f", encoding="utf-8-sig")
    add_to_excel("T01_Indices_CEI", df)
    print(f"  → Tabla 01 (CEI nodal) exportada: {fpath_tbl('Tabla_01_indices_CEI_nodales.csv')}")

def exportar_tabla_ahorros(df_ahorro):
    """Exporta la tabla de costos evitados/diferidos a CSV y al Excel maestro."""
    df = df_ahorro.copy()
    df.to_csv(fpath_tbl("Tabla_02_costos_evitados_BTM.csv"), index=False,
              float_format="%.2f", encoding="utf-8-sig")
    add_to_excel("T02_Costos_evitados", df)
    print(f"  → Tabla 02 (costos evitados) exportada: {fpath_tbl('Tabla_02_costos_evitados_BTM.csv')}")

def exportar_lineas_construidas(resultados_uniforme, resultados_cei):
    """
    Exporta matriz (corredor × alpha) con número de líneas nuevas construidas
    para cada estrategia (Uniforme y CEI).
    """
    for nombre, res in [("Uniforme", resultados_uniforme), ("CEI", resultados_cei)]:
        filas = []
        for j in CORREDORES:
            fila = {"Corredor": j, "Descripcion": DESC[j], "N0_existentes": N0[j]}
            for alpha in ALPHAS:
                r = res.get(alpha)
                nc = r['nuevas_lineas'].get(j, 0) if r else 0
                fila[f"alpha_{int(alpha*100):02d}pct"] = nc
            filas.append(fila)
        df = pd.DataFrame(filas)
        csv_path = fpath_tbl(f"Tabla_03_lineas_nuevas_{nombre}.csv")
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        add_to_excel(f"T03_Lineas_{nombre}", df)
        print(f"  → Tabla 03 ({nombre}) exportada: {csv_path}")

def exportar_cargabilidad(resultados_uniforme, resultados_cei):
    """
    Exporta matriz (corredor × alpha) con cargabilidad de cada corredor [%]
    bajo cada escenario BTM.
    """
    for nombre, res in [("Uniforme", resultados_uniforme), ("CEI", resultados_cei)]:
        filas = []
        for j in CORREDORES:
            fila = {"Corredor": j, "Descripcion": DESC[j]}
            for alpha in ALPHAS:
                r = res.get(alpha)
                carg = r['cargabilidad'].get(j, 0) * 100 if r else None
                fila[f"alpha_{int(alpha*100):02d}pct"] = round(carg, 2) if carg is not None else None
            filas.append(fila)
        df = pd.DataFrame(filas)
        csv_path = fpath_tbl(f"Tabla_04_cargabilidad_{nombre}.csv")
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        add_to_excel(f"T04_Cargabilidad_{nombre}", df)
        print(f"  → Tabla 04 ({nombre}) exportada: {csv_path}")

def exportar_generacion_despachada(resultados_uniforme, resultados_cei):
    """
    Exporta matriz (nodo × alpha) con la generación despachada [MW] en cada
    escenario BTM, para verificar el balance generación–demanda.
    """
    for nombre, res in [("Uniforme", resultados_uniforme), ("CEI", resultados_cei)]:
        filas = []
        for i in NODOS:
            fila = {"Nodo": i, "Nombre": NOMBRE[i], "G_max_MW": G_2039[i]}
            for alpha in ALPHAS:
                r = res.get(alpha)
                g = r['generacion'].get(i, 0) if r else None
                fila[f"alpha_{int(alpha*100):02d}pct"] = round(g, 2) if g is not None else None
            filas.append(fila)
        df = pd.DataFrame(filas)
        csv_path = fpath_tbl(f"Tabla_05_generacion_{nombre}.csv")
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        add_to_excel(f"T05_Generacion_{nombre}", df)
        print(f"  → Tabla 05 ({nombre}) exportada: {csv_path}")

def exportar_demanda_efectiva(resultados_uniforme, resultados_cei):
    """
    Exporta matriz (nodo × alpha) con la demanda neta resultante [MW]
    después de aplicar la reducción BTM.
    """
    for nombre, res in [("Uniforme", resultados_uniforme), ("CEI", resultados_cei)]:
        filas = []
        for i in NODOS:
            fila = {"Nodo": i, "Nombre": NOMBRE[i], "D_2039_base_MW": D_2039[i]}
            for alpha in ALPHAS:
                r = res.get(alpha)
                d = r['demanda'].get(i, 0) if r else None
                fila[f"alpha_{int(alpha*100):02d}pct"] = round(d, 2) if d is not None else None
            filas.append(fila)
        df = pd.DataFrame(filas)
        csv_path = fpath_tbl(f"Tabla_06_demanda_efectiva_{nombre}.csv")
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        add_to_excel(f"T06_Demanda_{nombre}", df)
        print(f"  → Tabla 06 ({nombre}) exportada: {csv_path}")

def exportar_resumen_ejecutivo(resultados_uniforme, resultados_cei, df_ahorro, df_cei):
    """
    Resumen ejecutivo de una sola tabla con los KPIs principales por escenario.
    """
    filas = []
    for alpha in ALPHAS:
        ru = resultados_uniforme.get(alpha)
        rc = resultados_cei.get(alpha)
        if ru is None or rc is None:
            continue
        n_lineas_u = sum(ru['nuevas_lineas'].values())
        n_lineas_c = sum(rc['nuevas_lineas'].values())
        n_congest_u = sum(1 for j in CORREDORES if ru['cargabilidad'].get(j, 0) > 0.95)
        n_congest_c = sum(1 for j in CORREDORES if rc['cargabilidad'].get(j, 0) > 0.95)
        carg_max_u = max(ru['cargabilidad'].values()) * 100
        carg_max_c = max(rc['cargabilidad'].values()) * 100
        filas.append({
            "alpha_pct":          f"{int(alpha*100)}%",
            "Demanda_Unif_MW":    round(sum(ru['demanda'].values()), 1),
            "Demanda_CEI_MW":     round(sum(rc['demanda'].values()), 1),
            "Inv_Unif_MUSD":      round(ru['inversion'], 1),
            "Inv_CEI_MUSD":       round(rc['inversion'], 1),
            "Lineas_nuevas_Unif": n_lineas_u,
            "Lineas_nuevas_CEI":  n_lineas_c,
            "Carg_max_Unif_pct":  round(carg_max_u, 1),
            "Carg_max_CEI_pct":   round(carg_max_c, 1),
            "Corredores_congest_Unif": n_congest_u,
            "Corredores_congest_CEI":  n_congest_c,
        })
    df = pd.DataFrame(filas)
    csv_path = fpath_rep("Resumen_Ejecutivo_KPIs.csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    add_to_excel("Resumen_Ejecutivo", df)
    print(f"  → Resumen ejecutivo exportado: {csv_path}")
    return df

def exportar_multiperiodo(resultados_multi):
    """
    Exporta los resultados del modelo multiperiodo:
      - Tabla de inversión anual por escenario
      - Tabla de inversión acumulada (VPN)
      - Tabla de demanda neta por año
      - Líneas nuevas año a año por corredor
    """
    if not resultados_multi:
        return
    # Inversión anual
    filas_inv = []
    for t in PERIODOS:
        fila = {"Año": ANO_INI + t - 1, "Periodo_t": t}
        for rm in resultados_multi:
            fila[rm['titulo'][:40]] = round(rm['inversion_por_t'].get(t, 0), 2)
        filas_inv.append(fila)
    df_inv = pd.DataFrame(filas_inv)
    df_inv.to_csv(fpath_tbl("Tabla_07_inversion_anual_multiperiodo.csv"),
                  index=False, encoding="utf-8-sig")
    add_to_excel("T07_Inv_anual_multi", df_inv)

    # Demanda por año
    filas_dem = []
    for t in PERIODOS:
        fila = {"Año": ANO_INI + t - 1, "Periodo_t": t}
        for rm in resultados_multi:
            fila[rm['titulo'][:40]] = round(rm['demanda_por_t'].get(t, 0), 2)
        filas_dem.append(fila)
    df_dem = pd.DataFrame(filas_dem)
    df_dem.to_csv(fpath_tbl("Tabla_08_demanda_anual_multiperiodo.csv"),
                  index=False, encoding="utf-8-sig")
    add_to_excel("T08_Demanda_anual_multi", df_dem)

    # Inversión total VPN por escenario
    filas_vpn = []
    for rm in resultados_multi:
        filas_vpn.append({
            "Escenario": rm['titulo'],
            "Inversion_VPN_MUSD": round(rm['inversion_total'], 2),
        })
    df_vpn = pd.DataFrame(filas_vpn)
    df_vpn.to_csv(fpath_tbl("Tabla_09_VPN_multiperiodo.csv"),
                  index=False, encoding="utf-8-sig")
    add_to_excel("T09_VPN_multi", df_vpn)

    print(f"  → Tablas multiperiodo exportadas en {DIR_TABLAS}")

def exportar_excel_maestro():
    """
    Escribe el archivo Excel maestro consolidado con todas las hojas acumuladas.
    """
    if not EXCEL_SHEETS:
        print("  ⚠ No hay hojas acumuladas para el Excel maestro.")
        return
    try:
        with pd.ExcelWriter(EXCEL_MASTER, engine="openpyxl") as writer:
            for sheet, df in EXCEL_SHEETS.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
        print(f"  ✅ Excel maestro consolidado: {EXCEL_MASTER}")
        print(f"     Hojas incluidas: {len(EXCEL_SHEETS)}")
    except Exception as e:
        print(f"  ❌ Error al escribir Excel maestro: {e}")

def exportar_datos_sistema():
    """
    Exporta los datos de entrada del sistema (parámetros del backbone SIN 500 kV)
    para garantizar reproducibilidad y trazabilidad.
    """
    # Tabla de nodos
    df_nodos = pd.DataFrame([
        {"Nodo": i, "Nombre": NOMBRE[i],
         "D_2024_MW": D_2024[i], "D_2039_MW": D_2039[i],
         "G_2024_MW": G_2024[i], "G_2039_MW": G_2039[i],
         "GHI_kWh_m2_dia": GHI_SIN[i]}
        for i in NODOS
    ])
    df_nodos.to_csv(fpath_tbl("Datos_00_nodos_sistema.csv"),
                    index=False, encoding="utf-8-sig")
    add_to_excel("D00_Nodos_sistema", df_nodos)

    # Tabla de corredores
    df_corred = pd.DataFrame([
        {"Corredor": j, "Descripcion": DESC[j],
         "N0_existentes": N0[j], "Fmax_MW": FMAX[j],
         "Susceptancia_pu": B_SIN[j], "Costo_MUSD_circ": COSTO[j]}
        for j in CORREDORES
    ])
    df_corred.to_csv(fpath_tbl("Datos_00_corredores_sistema.csv"),
                     index=False, encoding="utf-8-sig")
    add_to_excel("D00_Corredores_sistema", df_corred)
    print(f"  → Datos del sistema exportados a {DIR_TABLAS}")


# =============================================================================
# 9. IMPRIMIR RESULTADOS DETALLADOS
# =============================================================================

def imprimir_resultados_detallados(resultados_uniforme, resultados_cei):
    """Imprime en consola los resultados de todos los escenarios BTM."""
    print("\n" + "="*70)
    print("  RESULTADOS DETALLADOS — ESCENARIOS BTM")
    print("  Las comunidades energéticas como recurso virtual de expansión")
    print("="*70)

    for estrategia, res_dict in [("UNIFORME", resultados_uniforme),
                                  ("CEI INTELIGENTE", resultados_cei)]:
        print(f"\n  ── Estrategia: {estrategia} ──")
        print(f"  {'α':>5} | {'Dem.Tot(MW)':>12} | {'Inv(MUSD)':>10} | "
              f"{'Líneas nuevas':>15} | {'Congest.':>10}")
        print(f"  {'─'*5}─┼─{'─'*12}─┼─{'─'*10}─┼─{'─'*15}─┼─{'─'*10}")
        for alpha in ALPHAS:
            r = res_dict[alpha]
            if r is None:
                print(f"  {int(alpha*100):4d}%  | {'Sin solución':>40}")
                continue
            lineas_str  = str(dict(sorted(r['nuevas_lineas'].items()))) if r['nuevas_lineas'] else '(ninguna)'
            n_congest   = sum(1 for j in CORREDORES
                              if r['cargabilidad'].get(j, 0) > 0.95)
            dem_total   = sum(r['demanda'].values())
            print(f"  {int(alpha*100):4d}%  | {dem_total:>12,.0f} | "
                  f"{r['inversion']:>10,.0f} | {lineas_str[:15]:>15} | "
                  f"{'🔴 '+str(n_congest) if n_congest > 0 else '🟢 OK':>10}")

# =============================================================================
# 10. MAIN
# =============================================================================

if __name__ == "__main__":

    print("\n" + "="*70)
    print("  TEP — COSTO DIFERIDO POR COMUNIDADES ENERGÉTICAS BTM")
    print("  Backbone STN Colombia 500 kV — MILP DC")
    print("  Pyomo + Gurobi | DIgSILENT PF 2025 | UPME PIEG 2025-2039")
    print("  Tesis Ingeniería Eléctrica — Uniandes 2026")
    print("="*70)

    print("\n" + "─"*70)
    print("  CONCEPTO: Las comunidades energéticas BTM (fotovoltaica + almacenamiento")
    print("  detrás del medidor) reducen la demanda NETA que el sistema de transmisión")
    print("  debe atender. Esta reducción puede diferir o evitar completamente la")
    print("  construcción de nuevas líneas 500 kV, liberando capital de inversión.")
    print()
    print("  Formulación: D_nuevo[i] = D_base[i] × (1 − α[i])")
    print("  α: penetración de comunidades energéticas (fracción de demanda cubierta)")
    print("─"*70)

    # ── Paso 0: Exportar datos de entrada del sistema (trazabilidad) ──────────
    print("\n  [0/7] Exportando datos del sistema (nodos y corredores)...")
    exportar_datos_sistema()

    # ── Paso 1: Calcular índices CEI para el escenario de alta demanda 2039 ────
    print("\n  [1/7] Calculando índices CEI para el SIN Colombia 2039...")
    df_cei = calcular_indices_CEI(D_2039, G_2039)
    exportar_tabla_cei(df_cei)
    print("\n  Índices nodales CEI — Escenario Alta Demanda 2039:")
    print(df_cei[['Node','Nombre','Demand_MW','Gen_MW','SPI','NSI','GRI','CEI']].to_string(index=False))
    print(f"\n  Nodo más prioritario para BTM: "
          f"N{df_cei.sort_values('CEI',ascending=False).iloc[0]['Node']:.0f} "
          f"({df_cei.sort_values('CEI',ascending=False).iloc[0]['Nombre']})")
    print(f"  CEI máximo: {df_cei['CEI'].max():.3f} | CEI mínimo: {df_cei['CEI'].min():.3f}")

    # ── Paso 2: Ejecutar análisis BTM (TEP estático) ──────────────────────────
    print("\n  [2/7] Ejecutando análisis TEP estático con escenarios BTM...")
    print("        Escenario base: demanda alta 2039 (Oriental 4200 MW, Caribe 3800 MW)")
    print("        Penetraciones evaluadas:", [ALPHA_LABELS[a] for a in ALPHAS])
    resultados_uniforme, resultados_cei = ejecutar_analisis_btm(df_cei, presupuesto=8000)

    # ── Paso 3: Calcular ahorros y exportar tablas detalladas ─────────────────
    print("\n  [3/7] Calculando costos evitados y diferidos...")
    df_ahorro = calcular_ahorros(resultados_uniforme, resultados_cei)
    exportar_tabla_ahorros(df_ahorro)
    exportar_lineas_construidas(resultados_uniforme, resultados_cei)
    exportar_cargabilidad(resultados_uniforme, resultados_cei)
    exportar_generacion_despachada(resultados_uniforme, resultados_cei)
    exportar_demanda_efectiva(resultados_uniforme, resultados_cei)
    df_resumen = exportar_resumen_ejecutivo(resultados_uniforme, resultados_cei,
                                            df_ahorro, df_cei)

    # ── Paso 4: Imprimir resultados detallados ────────────────────────────────
    print("\n  [4/7] Resultados detallados por escenario:")
    imprimir_resultados_detallados(resultados_uniforme, resultados_cei)

    # ── Paso 5: Generar gráficas ──────────────────────────────────────────────
    print("\n  [5/7] Generando gráficas académicas...")

    figura_01_ranking_CEI(df_cei)
    figura_02_demanda_vs_inversion(df_ahorro, resultados_uniforme, resultados_cei)
    figura_03_penetracion_vs_congestion(resultados_uniforme, resultados_cei)
    figura_04_lineas_por_escenario(resultados_uniforme, resultados_cei)
    figura_05_heatmap_cargabilidad(resultados_uniforme, resultados_cei)
    figura_06_ahorro_acumulado(df_ahorro)
    figura_07_comparacion_uniforme_vs_cei(resultados_uniforme, resultados_cei, df_ahorro)
    figura_08_red_comparativa(resultados_uniforme, alpha_ref=0.10)
    figura_09_sensibilidad_corredores(resultados_uniforme)
    figura_10_generacion_despachada(resultados_uniforme, resultados_cei)
    figura_12_resumen_ejecutivo(df_ahorro, df_cei, resultados_uniforme)

    # ── Paso 6: Modelo multiperiodo (opcional) ────────────────────────────────
    if CORRER_MULTIPERIODO:
        print("\n  [6/7] Ejecutando modelo multiperiodo BTM 2024–2039...")
        print("        Escenarios: α=0% (base), α=10% (uniforme), α=20% (uniforme)")

        resultados_multi = []
        alphas_multi = [0.00, 0.10, 0.20]

        for alpha in alphas_multi:
            alpha_dict = {i: alpha for i in NODOS}
            titulo_m   = f"Multiperiodo BTM α={int(alpha*100)}% (Uniforme)"
            rm = resolver_multiperiodo_btm(alpha_dict, titulo_m)
            if rm:
                resultados_multi.append(rm)
                print(f"    α={int(alpha*100)}%: Inversión VPN total = {rm['inversion_total']:,.0f} MUSD")

        if len(resultados_multi) >= 2:
            figura_11_multiperiodo_deferral(resultados_multi)

            # Exportar tablas del análisis multiperiodo
            exportar_multiperiodo(resultados_multi)

            print("\n  Costo diferido/evitado por BTM — Modelo Multiperiodo:")
            base_vpn = resultados_multi[0]['inversion_total']
            for rm in resultados_multi[1:]:
                ahorro_vpn = base_vpn - rm['inversion_total']
                print(f"    {rm['titulo']}: {rm['inversion_total']:,.0f} MUSD VPN  |  "
                      f"Ahorro VPN: {ahorro_vpn:+,.0f} MUSD")

    # ── Resumen final ─────────────────────────────────────────────────────────
    print("\n" + "="*70)
    print("  RESUMEN FINAL — COMUNIDADES ENERGÉTICAS BTM COMO RECURSO VIRTUAL")
    print("="*70)

    base_inv = resultados_uniforme[0.0]['inversion'] if resultados_uniforme[0.0] else 0
    for alpha in ALPHAS:
        ru = resultados_uniforme[alpha]
        rc = resultados_cei[alpha]
        if ru and rc:
            ahorro_u = base_inv - ru['inversion']
            ahorro_c = base_inv - rc['inversion']
            print(f"\n  α={int(alpha*100):02d}% → Ahorro Uniforme: {ahorro_u:+,.0f} MUSD  |  "
                  f"Ahorro CEI: {ahorro_c:+,.0f} MUSD")
            if ahorro_u > 0 or ahorro_c > 0:
                lineas_evitadas_u = set(resultados_uniforme[0.0]['nuevas_lineas'].keys()) - \
                                    set(ru['nuevas_lineas'].keys())
                lineas_evitadas_c = set(resultados_cei[0.0]['nuevas_lineas'].keys()) - \
                                    set(rc['nuevas_lineas'].keys())
                if lineas_evitadas_u:
                    print(f"         Líneas evitadas (Unif): {lineas_evitadas_u}")
                if lineas_evitadas_c:
                    print(f"         Líneas evitadas (CEI):  {lineas_evitadas_c}")

    print(f"\n  CONCLUSIÓN:")
    max_a_u = df_ahorro['Ahorro_Unif'].max() or 0
    max_a_c = df_ahorro['Ahorro_CEI'].max()  or 0
    print(f"  Con una penetración BTM del 20%, las comunidades energéticas pueden")
    print(f"  evitar hasta {max_a_u:,.0f} MUSD (uniforme) o {max_a_c:,.0f} MUSD (CEI)")
    print(f"  en inversión de transmisión en el SIN Colombia 500 kV.")
    print(f"  La estrategia CEI inteligente ahorra {max(0, max_a_c-max_a_u):,.0f} MUSD")
    print(f"  adicionales frente a la reducción uniforme, demostrando el valor de")
    print(f"  localizar las comunidades energéticas en los nodos más críticos.")
    print(f"\n  → Las comunidades BTM actúan como un RECURSO VIRTUAL DE TRANSMISIÓN")
    print(f"    capaz de diferir o evitar la construcción de líneas de 500 kV.")
    # ── Paso 7: Generar Excel maestro consolidado ─────────────────────────────
    print("\n  [7/7] Generando archivo Excel consolidado...")
    exportar_excel_maestro()

    print("\n" + "="*70)
    print("  ✅  EJECUCIÓN COMPLETADA")
    print(f"     Resultados en directorio: {os.path.abspath(OUTPUT_DIR)}")
    print(f"     ├── figuras/   ({len(os.listdir(DIR_FIGURAS))} archivos)")
    print(f"     ├── tablas/    ({len(os.listdir(DIR_TABLAS))} archivos)")
    print(f"     └── reportes/  ({len(os.listdir(DIR_REPORTES))} archivos)")
    print("="*70)