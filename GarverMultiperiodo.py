# -*- coding: utf-8 -*-
"""
=============================================================================
GARVER 6-BUS — MILP DC con análisis BTM (MULTIPERIODO con tasa de descuento)
=============================================================================

Refactorización del modelo Garver multiperiodo para producir la misma
estructura de reportes que Modelo_completo_v2.py (SIN 500 kV):

  * Datos del sistema (nodos, corredores)              [Excel + CSV]
  * Inversión anual por escenario BTM                  [Excel + CSV]
  * Demanda anual por escenario BTM                    [Excel + CSV]
  * VPN total por escenario                            [Excel + CSV]
  * Cronograma de construcción por corredor × año      [Excel + CSV]
  * Crédito diferido por escenario                     [Excel + CSV]
  * Excel maestro consolidado                          [resultados_Garver_multiperiodo.xlsx]
  * Figuras académicas                                 [PNG]

DATOS:
  Costos, reactancias y capacidades clásicos de Garver (1970),
  consistentes con el modelo estático refactorizado.

MULTIPERIODO:
  NT = 10 periodos anuales (escala temporal pedagógica del sistema Garver)
  Crecimiento de demanda: lineal de 55% a 100% del valor horizonte
  Tasa de descuento: r = 10% anual

ANÁLISIS BTM:
  α[i,t] crece linealmente desde 0 en t=1 hasta α_fin en t=NT
  D_neta[i,t] = D_base[i,t] × (1 − α[i,t])

  El crédito diferido se cuantifica como diferencia de VPN entre el plan
  base (α=0%) y los planes con BTM (α>0%).

Tesis Ingeniería Eléctrica — Uniandes 2026
=============================================================================
"""

import os
import time
import warnings
warnings.filterwarnings('ignore')

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

import pyomo.environ as pyo
from pyomo.opt import SolverFactory

# =============================================================================
# 0. CONFIGURACIÓN
# =============================================================================

GUARDAR_GRAFICAS = True
OUTPUT_DIR    = "resultados_Garver_multiperiodo"
DIR_FIGURAS   = os.path.join(OUTPUT_DIR, "figuras")
DIR_TABLAS    = os.path.join(OUTPUT_DIR, "tablas")
DIR_REPORTES  = os.path.join(OUTPUT_DIR, "reportes")
for _d in (OUTPUT_DIR, DIR_FIGURAS, DIR_TABLAS, DIR_REPORTES):
    os.makedirs(_d, exist_ok=True)

EXCEL_MASTER = os.path.join(DIR_REPORTES, "resultados_Garver_multiperiodo.xlsx")

def fpath_fig(name): return os.path.join(DIR_FIGURAS, name)
def fpath_tbl(name): return os.path.join(DIR_TABLAS,  name)
def fpath_rep(name): return os.path.join(DIR_REPORTES, name)

EXCEL_SHEETS = {}
def add_to_excel(sheet_name, df):
    EXCEL_SHEETS[sheet_name[:31]] = df

# Penetraciones BTM terminales a evaluar (α al final del horizonte)
ALPHAS_FIN = [0.00, 0.10, 0.20, 0.30, 0.40]
ALPHA_LABELS = {a: f"α_fin={int(a*100):02d}%" for a in ALPHAS_FIN}

# Horizonte temporal
NT          = 10
ANO_INI     = 2024
PERIODOS    = list(range(1, NT + 1))
TASA_DESCUENTO = 0.10

# Crecimiento de demanda (lineal de 55% a 100% — escala pedagógica)
def growth_factor(t):
    return 0.55 + (1.0 - 0.55) * (t - 1) / (NT - 1)

# =============================================================================
# 1. DATOS DEL SISTEMA GARVER 6-BUS
# =============================================================================

Sbase = 100  # MVA

NODOS  = [1, 2, 3, 4, 5, 6]
NOMBRE = {
    1: 'Nodo 1 (importador leve)',
    2: 'Nodo 2 (importador mayor)',
    3: 'Nodo 3 (exportador local)',
    4: 'Nodo 4 (importador)',
    5: 'Nodo 5 (importador mayor)',
    6: 'Nodo 6 (exportador remoto)',
}

D_HORIZ = {1: 80, 2: 240, 3: 40, 4: 160, 5: 240, 6: 0}
G_MAX   = {1: 50, 2: 0,   3: 165,4: 0,   5: 0,   6: 545}

GHI_GARVER = {i: 5.0 for i in NODOS}

CORREDORES = [
    '1,2', '1,3', '1,4', '1,5', '1,6',
    '2,3', '2,4', '2,5', '2,6',
    '3,4', '3,5', '3,6',
    '4,5', '4,6', '5,6',
]

DESC = {
    '1,2': 'Corredor 1–2', '1,3': 'Corredor 1–3', '1,4': 'Corredor 1–4',
    '1,5': 'Corredor 1–5', '1,6': 'Corredor 1–6',
    '2,3': 'Corredor 2–3', '2,4': 'Corredor 2–4', '2,5': 'Corredor 2–5',
    '2,6': 'Corredor 2–6',
    '3,4': 'Corredor 3–4', '3,5': 'Corredor 3–5', '3,6': 'Corredor 3–6',
    '4,5': 'Corredor 4–5', '4,6': 'Corredor 4–6', '5,6': 'Corredor 5–6',
}

N0 = {
    '1,2': 1, '1,3': 0, '1,4': 1, '1,5': 1, '1,6': 0,
    '2,3': 1, '2,4': 1, '2,5': 0, '2,6': 0,
    '3,4': 0, '3,5': 1, '3,6': 0,
    '4,5': 0, '4,6': 0, '5,6': 0,
}

FMAX = {
    '1,2': 100, '1,3': 100, '1,4': 80,  '1,5': 100, '1,6': 70,
    '2,3': 100, '2,4': 100, '2,5': 100, '2,6': 100,
    '3,4': 82,  '3,5': 100, '3,6': 100,
    '4,5': 75,  '4,6': 100, '5,6': 78,
}

# Susceptancia equivalente B = 1/x (p.u. en base 100 MVA)
B_GARVER = {
    '1,2': 2.50, '1,3': 0, '1,4': 1.67, '1,5': 5.00, '1,6': 0,
    '2,3': 5.00, '2,4': 2.50, '2,5': 0, '2,6': 3.33,
    '3,4': 0, '3,5': 5.00, '3,6': 0,
    '4,5': 0, '4,6': 3.33, '5,6': 1.64,
}

# Costos clásicos Garver 1970 (kUSD por circuito)
COSTO = {
    '1,2': 100000, '1,3': 100000, '1,4': 100000, '1,5': 100000, '1,6': 100000,
    '2,3': 100000, '2,4': 100000, '2,5': 100000, '2,6': 150,
    '3,4': 100000, '3,5': 100, '3,6': 100000,
    '4,5': 100000, '4,6': 150, '5,6': 350,
}

NCAND = 5
NREF  = 5
BIG_M = max(FMAX.values()) * 2.5 / Sbase

# =============================================================================
# 2. MODELO MILP DC MULTIPERIODO
# =============================================================================

def get_S(i, j_str):
    a, b = map(int, j_str.split(','))
    if i == a: return  1
    if i == b: return -1
    return 0


def demanda_t_btm(i, t, alpha_fin_dict=None):
    """
    Demanda del nodo i en periodo t, con penetración BTM creciente linealmente
    desde 0 (t=1) hasta α_fin (t=NT).
    """
    d_base_t = D_HORIZ[i] * growth_factor(t)
    if alpha_fin_dict is None:
        return d_base_t
    a_fin = alpha_fin_dict.get(i, 0.0)
    alpha_t = a_fin * (t - 1) / (NT - 1)
    return d_base_t * (1.0 - alpha_t)


def construir_modelo_multiperiodo(alpha_fin_dict=None):
    """Modelo MILP DC multiperiodo con VPN y dinámica de expansión."""
    if alpha_fin_dict is None:
        alpha_fin_dict = {i: 0.0 for i in NODOS}

    mdl = pyo.ConcreteModel()
    mdl.I = pyo.Set(initialize=NODOS)
    mdl.J = pyo.Set(initialize=CORREDORES)
    mdl.K = pyo.RangeSet(1, NCAND)
    mdl.T = pyo.RangeSet(1, NT)

    mdl.S    = pyo.Param(mdl.I, mdl.J,
                         initialize={(i,j): get_S(i,j) for i in NODOS for j in CORREDORES})
    mdl.d    = pyo.Param(mdl.I, mdl.T,
                         initialize={(i,t): demanda_t_btm(i, t, alpha_fin_dict)/Sbase
                                     for i in NODOS for t in PERIODOS})
    mdl.gmax = pyo.Param(mdl.I, initialize={i: G_MAX[i]/Sbase for i in NODOS})
    mdl.n0   = pyo.Param(mdl.J, initialize=N0)
    mdl.fmax = pyo.Param(mdl.J, initialize={j: FMAX[j]/Sbase for j in CORREDORES})
    mdl.B    = pyo.Param(mdl.J, initialize=B_GARVER)
    mdl.c    = pyo.Param(mdl.J, initialize=COSTO)

    desc = {t: 1.0/(1.0+TASA_DESCUENTO)**(t-1) for t in PERIODOS}
    mdl.desc = pyo.Param(mdl.T, initialize=desc)

    mdl.y     = pyo.Var(mdl.J, mdl.K, mdl.T, domain=pyo.Binary)
    mdl.x     = pyo.Var(mdl.J, mdl.K, mdl.T, domain=pyo.Binary)
    mdl.f     = pyo.Var(mdl.J, mdl.K, mdl.T, domain=pyo.Reals)
    mdl.g     = pyo.Var(mdl.I, mdl.T,        domain=pyo.NonNegativeReals)
    mdl.theta = pyo.Var(mdl.I, mdl.T,        domain=pyo.Reals)

    def slack_rule(m, t): return m.theta[NREF, t] == 0
    mdl.slack = pyo.Constraint(mdl.T, rule=slack_rule)

    def obj_rule(m):
        return sum(m.c[j]*m.y[j,k,t]*m.desc[t]
                   for j in m.J for k in m.K for t in m.T
                   if k > pyo.value(m.n0[j]))
    mdl.obj = pyo.Objective(rule=obj_rule, sense=pyo.minimize)

    def balance(m, i, t):
        return (sum(m.S[i,j]*m.f[j,k,t] for j in m.J for k in m.K)
                + m.g[i,t] == m.d[i,t])
    mdl.balance = pyo.Constraint(mdl.I, mdl.T, rule=balance)

    def flujo_pos(m, j, k, t):
        a, b = map(int, j.split(','))
        return m.f[j,k,t] - m.B[j]*(m.theta[a,t]-m.theta[b,t]) <=  BIG_M*(1-m.x[j,k,t])
    def flujo_neg(m, j, k, t):
        a, b = map(int, j.split(','))
        return m.f[j,k,t] - m.B[j]*(m.theta[a,t]-m.theta[b,t]) >= -BIG_M*(1-m.x[j,k,t])
    mdl.flujo_pos = pyo.Constraint(mdl.J, mdl.K, mdl.T, rule=flujo_pos)
    mdl.flujo_neg = pyo.Constraint(mdl.J, mdl.K, mdl.T, rule=flujo_neg)

    def lim_pos(m, j, k, t): return  m.f[j,k,t] <=  m.fmax[j]*m.x[j,k,t]
    def lim_neg(m, j, k, t): return -m.f[j,k,t] <=  m.fmax[j]*m.x[j,k,t]
    mdl.lim_pos = pyo.Constraint(mdl.J, mdl.K, mdl.T, rule=lim_pos)
    mdl.lim_neg = pyo.Constraint(mdl.J, mdl.K, mdl.T, rule=lim_neg)

    def gen_max(m, i, t): return m.g[i,t] <= m.gmax[i]
    mdl.gen_max = pyo.Constraint(mdl.I, mdl.T, rule=gen_max)

    def dinamica(m, j, k, t):
        base = 1 if k <= pyo.value(m.n0[j]) else 0
        if t == 1:
            return m.x[j,k,t] == base + m.y[j,k,t]
        return m.x[j,k,t] == m.x[j,k,t-1] + m.y[j,k,t]
    mdl.dinamica = pyo.Constraint(mdl.J, mdl.K, mdl.T, rule=dinamica)

    def una_vez(m, j, k):
        return sum(m.y[j,k,t] for t in m.T) <= 1
    mdl.una_vez = pyo.Constraint(mdl.J, mdl.K, rule=una_vez)

    def x_bin(m, j, k, t): return m.x[j,k,t] <= 1
    mdl.x_bin = pyo.Constraint(mdl.J, mdl.K, mdl.T, rule=x_bin)

    return mdl

# =============================================================================
# 3. RESOLUCIÓN
# =============================================================================

def resolver_multiperiodo(alpha_fin_dict, titulo):
    """Resuelve el modelo multiperiodo. Devuelve dict con resultados por periodo."""
    print(f"\n  ── Resolviendo multiperiodo: {titulo} ──")

    mdl = construir_modelo_multiperiodo(alpha_fin_dict)

    solver = SolverFactory('gurobi')
    solver.options['MIPGap']     = 0.003
    solver.options['TimeLimit']  = 200
    solver.options['OutputFlag'] = 0

    t0 = time.time()
    result = solver.solve(mdl, tee=False)
    dt = time.time() - t0

    status = result.solver.termination_condition
    cond_ok = (pyo.TerminationCondition.optimal, pyo.TerminationCondition.feasible)
    if status not in cond_ok:
        print(f"    ❌ Sin solución ({status}) — {dt:.1f}s")
        return None
    print(f"    ✅ {status} — {dt:.2f}s")

    expansion_por_t   = {}
    inversion_por_t   = {}      # nominal (no descontada)
    inversion_vpn_por_t = {}    # descontada
    demanda_por_t     = {}
    flujos_por_t      = {}
    cronograma        = {j: [] for j in CORREDORES}  # lista de (t, k) por corredor

    for t in PERIODOS:
        exp_t = {}
        inv_t = 0.0
        inv_vpn_t = 0.0
        fl_t = {}
        for j in CORREDORES:
            nc = 0
            for k in range(1, NCAND+1):
                if k > N0[j] and pyo.value(mdl.y[j,k,t]) > 0.5:
                    nc += 1
                    cronograma[j].append((t, k))
            if nc > 0:
                exp_t[j] = nc
                inv_t += COSTO[j]*nc
                inv_vpn_t += COSTO[j]*nc / ((1+TASA_DESCUENTO)**(t-1))
            fl_t[j] = sum(pyo.value(mdl.f[j,k,t])*Sbase for k in range(1, NCAND+1))
        expansion_por_t[t]     = exp_t
        inversion_por_t[t]     = inv_t
        inversion_vpn_por_t[t] = inv_vpn_t
        demanda_por_t[t]       = sum(demanda_t_btm(i, t, alpha_fin_dict) for i in NODOS)
        flujos_por_t[t]        = fl_t

    return {
        'titulo':              titulo,
        'alpha_fin':           alpha_fin_dict,
        'expansion_por_t':     expansion_por_t,
        'inversion_por_t':     inversion_por_t,
        'inversion_vpn_por_t': inversion_vpn_por_t,
        'demanda_por_t':       demanda_por_t,
        'flujos_por_t':        flujos_por_t,
        'cronograma':          cronograma,
        'inversion_total':     sum(inversion_por_t.values()),
        'inversion_vpn_total': sum(inversion_vpn_por_t.values()),
    }


# =============================================================================
# 4. EJECUTAR TODOS LOS ESCENARIOS
# =============================================================================

def ejecutar_analisis_btm_multiperiodo():
    """Resuelve el multiperiodo para cada α_fin en ALPHAS_FIN (estrategia uniforme)."""
    print("\n" + "="*65)
    print("  ANÁLISIS BTM MULTIPERIODO — GARVER 6-BUS")
    print("="*65)

    resultados_multi = []
    for alpha in ALPHAS_FIN:
        alpha_dict = {i: alpha if D_HORIZ[i] > 0 else 0.0 for i in NODOS}
        titulo = f"Multiperiodo BTM {ALPHA_LABELS[alpha]} (Uniforme)"
        rm = resolver_multiperiodo(alpha_dict, titulo)
        if rm:
            resultados_multi.append(rm)
            print(f"    {ALPHA_LABELS[alpha]}: Inv. total = {rm['inversion_total']:,.0f} kUSD "
                  f"| VPN = {rm['inversion_vpn_total']:,.1f} kUSD")
    return resultados_multi


# =============================================================================
# 5. EXPORTACIÓN DE TABLAS
# =============================================================================

def exportar_datos_sistema():
    df_nodos = pd.DataFrame([
        {"Nodo": i, "Nombre": NOMBRE[i],
         "D_horizonte_MW": D_HORIZ[i], "Gmax_MW": G_MAX[i], "GHI": GHI_GARVER[i]}
        for i in NODOS
    ])
    df_nodos.to_csv(fpath_tbl("Datos_00_nodos_sistema.csv"),
                    index=False, encoding="utf-8-sig")
    add_to_excel("D00_Nodos_sistema", df_nodos)

    df_corred = pd.DataFrame([
        {"Corredor": j, "Descripcion": DESC[j],
         "N0_existentes": N0[j], "Fmax_MW": FMAX[j],
         "Susceptancia_pu": B_GARVER[j], "Costo_kUSD_circ": COSTO[j]}
        for j in CORREDORES
    ])
    df_corred.to_csv(fpath_tbl("Datos_00_corredores_sistema.csv"),
                     index=False, encoding="utf-8-sig")
    add_to_excel("D00_Corredores_sistema", df_corred)

    # Crecimiento de demanda
    df_growth = pd.DataFrame([
        {"Periodo_t": t, "Año": ANO_INI + t - 1, "growth_factor": growth_factor(t),
         "D_total_MW": sum(D_HORIZ[i] for i in NODOS) * growth_factor(t)}
        for t in PERIODOS
    ])
    df_growth.to_csv(fpath_tbl("Datos_01_crecimiento_demanda.csv"),
                     index=False, float_format="%.4f", encoding="utf-8-sig")
    add_to_excel("D01_Crecimiento_demanda", df_growth)
    print(f"  → Datos del sistema exportados a {DIR_TABLAS}")


def exportar_inversion_anual(resultados_multi):
    """Tabla T07 análoga a la del SIN: inversión nominal anual por escenario."""
    filas = []
    for t in PERIODOS:
        fila = {"Año": ANO_INI + t - 1, "Periodo_t": t}
        for rm in resultados_multi:
            fila[rm['titulo'][:40]] = round(rm['inversion_por_t'].get(t, 0), 2)
        filas.append(fila)
    df = pd.DataFrame(filas)
    df.to_csv(fpath_tbl("Tabla_07_inversion_anual_multiperiodo.csv"),
              index=False, encoding="utf-8-sig")
    add_to_excel("T07_Inv_anual_multi", df)


def exportar_inversion_vpn_anual(resultados_multi):
    """Tabla complementaria: inversión descontada (VPN) anual por escenario."""
    filas = []
    for t in PERIODOS:
        fila = {"Año": ANO_INI + t - 1, "Periodo_t": t,
                "Factor_desc": round(1.0/(1+TASA_DESCUENTO)**(t-1), 4)}
        for rm in resultados_multi:
            fila[rm['titulo'][:40]] = round(rm['inversion_vpn_por_t'].get(t, 0), 2)
        filas.append(fila)
    df = pd.DataFrame(filas)
    df.to_csv(fpath_tbl("Tabla_07b_inversion_VPN_anual_multiperiodo.csv"),
              index=False, encoding="utf-8-sig")
    add_to_excel("T07b_Inv_VPN_anual", df)


def exportar_demanda_anual(resultados_multi):
    filas = []
    for t in PERIODOS:
        fila = {"Año": ANO_INI + t - 1, "Periodo_t": t}
        for rm in resultados_multi:
            fila[rm['titulo'][:40]] = round(rm['demanda_por_t'].get(t, 0), 2)
        filas.append(fila)
    df = pd.DataFrame(filas)
    df.to_csv(fpath_tbl("Tabla_08_demanda_anual_multiperiodo.csv"),
              index=False, encoding="utf-8-sig")
    add_to_excel("T08_Demanda_anual_multi", df)


def exportar_vpn_total(resultados_multi):
    """T09: VPN total y crédito diferido por escenario."""
    if not resultados_multi:
        return
    base_vpn = resultados_multi[0]['inversion_vpn_total']
    base_nom = resultados_multi[0]['inversion_total']
    filas = []
    for rm in resultados_multi:
        filas.append({
            "Escenario":             rm['titulo'],
            "Inversion_nominal_kUSD": round(rm['inversion_total'], 2),
            "Inversion_VPN_kUSD":     round(rm['inversion_vpn_total'], 2),
            "Credito_diferido_nominal": round(base_nom - rm['inversion_total'], 2),
            "Credito_diferido_VPN":     round(base_vpn - rm['inversion_vpn_total'], 2),
            "Reduccion_VPN_pct":        round((base_vpn - rm['inversion_vpn_total'])/base_vpn*100, 2) if base_vpn else 0,
        })
    df = pd.DataFrame(filas)
    df.to_csv(fpath_tbl("Tabla_09_VPN_multiperiodo.csv"),
              index=False, encoding="utf-8-sig")
    add_to_excel("T09_VPN_multi", df)


def exportar_cronograma(resultados_multi):
    """T10: cronograma detallado de construcción por corredor × periodo × escenario."""
    for rm in resultados_multi:
        filas = []
        for j in CORREDORES:
            fila = {"Corredor": j, "Descripcion": DESC[j], "N0_existentes": N0[j]}
            for t in PERIODOS:
                nc = rm['expansion_por_t'].get(t, {}).get(j, 0)
                fila[f"t={t}_({ANO_INI+t-1})"] = nc
            fila["Total_nuevos"] = sum(fila[f"t={t}_({ANO_INI+t-1})"] for t in PERIODOS)
            filas.append(fila)
        df = pd.DataFrame(filas)
        alpha_pct = int(rm['titulo'].split('=')[1].split('%')[0])
        csv_name = f"Tabla_10_cronograma_alpha_{alpha_pct:02d}.csv"
        df.to_csv(fpath_tbl(csv_name), index=False, encoding="utf-8-sig")
        add_to_excel(f"T10_Cronograma_a{alpha_pct:02d}", df)


def exportar_resumen_ejecutivo(resultados_multi):
    """KPIs por escenario en un solo cuadro."""
    filas = []
    if not resultados_multi:
        return None
    base_vpn = resultados_multi[0]['inversion_vpn_total']
    base_nom = resultados_multi[0]['inversion_total']
    for rm in resultados_multi:
        ap = rm['alpha_fin']
        alpha_max = max(ap.values()) if ap else 0
        n_circuitos = sum(sum(exp.values()) for exp in rm['expansion_por_t'].values())
        primer_periodo = min((t for t, exp in rm['expansion_por_t'].items() if exp),
                             default=None)
        ultimo_periodo = max((t for t, exp in rm['expansion_por_t'].items() if exp),
                             default=None)
        filas.append({
            "Escenario":                rm['titulo'],
            "alpha_fin_pct":            f"{int(alpha_max*100)}%",
            "Inversion_nominal_kUSD":    round(rm['inversion_total'], 2),
            "Inversion_VPN_kUSD":        round(rm['inversion_vpn_total'], 2),
            "Credito_diferido_nominal":  round(base_nom - rm['inversion_total'], 2),
            "Credito_diferido_VPN":      round(base_vpn - rm['inversion_vpn_total'], 2),
            "Reduccion_VPN_pct":         round((base_vpn - rm['inversion_vpn_total'])/base_vpn*100, 2) if base_vpn else 0,
            "Circuitos_construidos":     n_circuitos,
            "Primer_periodo_construc":   primer_periodo,
            "Ultimo_periodo_construc":   ultimo_periodo,
            "Demanda_inicial_MW":        round(rm['demanda_por_t'].get(1, 0), 1),
            "Demanda_final_MW":          round(rm['demanda_por_t'].get(NT, 0), 1),
        })
    df = pd.DataFrame(filas)
    df.to_csv(fpath_rep("Resumen_Ejecutivo_KPIs.csv"),
              index=False, encoding="utf-8-sig")
    add_to_excel("Resumen_Ejecutivo", df)
    return df


def exportar_excel_maestro():
    if not EXCEL_SHEETS:
        print("  ⚠ No hay hojas para el Excel maestro.")
        return
    try:
        with pd.ExcelWriter(EXCEL_MASTER, engine="openpyxl") as writer:
            for sheet, df in EXCEL_SHEETS.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
        print(f"  ✅ Excel maestro consolidado: {EXCEL_MASTER}")
        print(f"     Hojas incluidas: {len(EXCEL_SHEETS)}")
    except Exception as e:
        print(f"  ❌ Error al escribir Excel maestro: {e}")

# =============================================================================
# 6. FIGURAS
# =============================================================================

COLORES = plt.cm.viridis(np.linspace(0.05, 0.85, len(ALPHAS_FIN)))


def figura_01_inversion_anual(resultados_multi):
    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    fig.suptitle("FIG 1 — Inversión anual por escenario BTM (Garver 6-bus)",
                 fontsize=11, fontweight='bold')

    anos = [ANO_INI + t - 1 for t in PERIODOS]

    # Panel a: barras agrupadas
    ax = axes[0]
    x_base = np.arange(len(PERIODOS))
    w = 0.85 / len(resultados_multi)
    for ai, rm in enumerate(resultados_multi):
        invs = [rm['inversion_por_t'].get(t, 0) for t in PERIODOS]
        offset = (ai - len(resultados_multi)/2 + 0.5) * w
        ax.bar(x_base + offset, invs, w, color=COLORES[ai],
               alpha=0.85, label=rm['titulo'][:25])
    ax.set_xticks(x_base[::2])
    ax.set_xticklabels([anos[i] for i in range(0, len(anos), 2)], rotation=0)
    ax.set_ylabel("Inversión anual (kUSD)")
    ax.set_title("(a) Inversión nominal por año")
    ax.legend(fontsize=8); ax.grid(True, alpha=0.3, axis='y')

    # Panel b: acumulada (VPN)
    ax = axes[1]
    for ai, rm in enumerate(resultados_multi):
        inv_vpn = [rm['inversion_vpn_por_t'].get(t, 0) for t in PERIODOS]
        ac = np.cumsum(inv_vpn)
        ax.plot(anos, ac, 'o-', color=COLORES[ai], lw=2.2, ms=6,
                label=rm['titulo'][:25])
        ax.fill_between(anos, ac, alpha=0.1, color=COLORES[ai])
    ax.set_xlabel("Año"); ax.set_ylabel("Inversión acumulada VPN (kUSD)")
    ax.set_title("(b) Inversión acumulada VPN")
    ax.legend(fontsize=8); ax.grid(True, alpha=0.35)

    plt.tight_layout()
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig("Garver_MP_Fig01_Inversion_anual.png"),
                    dpi=160, bbox_inches='tight')
    plt.close()


def figura_02_demanda_evolucion(resultados_multi):
    fig, ax = plt.subplots(figsize=(9, 5))
    fig.suptitle("FIG 2 — Evolución de la demanda neta por escenario BTM (Garver)",
                 fontsize=11, fontweight='bold')

    anos = [ANO_INI + t - 1 for t in PERIODOS]
    for ai, rm in enumerate(resultados_multi):
        dems = [rm['demanda_por_t'].get(t, 0) for t in PERIODOS]
        ax.plot(anos, dems, 'o-', color=COLORES[ai], lw=2.2, ms=6,
                label=rm['titulo'][:25])
    ax.set_xlabel("Año")
    ax.set_ylabel("Demanda neta total (MW)")
    ax.legend(fontsize=8); ax.grid(True, alpha=0.35)
    plt.tight_layout()
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig("Garver_MP_Fig02_Demanda_evolucion.png"),
                    dpi=160, bbox_inches='tight')
    plt.close()


def figura_03_credito_diferido(resultados_multi):
    """Crédito diferido VPN vs α_fin."""
    if len(resultados_multi) < 2:
        return
    base_vpn = resultados_multi[0]['inversion_vpn_total']
    base_nom = resultados_multi[0]['inversion_total']

    alphas_pct = []
    creditos_vpn = []
    creditos_nom = []
    for rm in resultados_multi:
        ap = rm['alpha_fin']
        a_max = max(ap.values()) if ap else 0
        alphas_pct.append(int(a_max*100))
        creditos_vpn.append(base_vpn - rm['inversion_vpn_total'])
        creditos_nom.append(base_nom - rm['inversion_total'])

    fig, ax = plt.subplots(figsize=(9, 5))
    fig.suptitle("FIG 3 — Crédito diferido vs penetración BTM final (Garver)",
                 fontsize=11, fontweight='bold')
    x = np.arange(len(alphas_pct)); w = 0.35
    ax.bar(x - w/2, creditos_nom, w, color='#1f77b4', alpha=0.85, label='Nominal')
    ax.bar(x + w/2, creditos_vpn, w, color='#d62728', alpha=0.85, label='VPN (descontado)')
    for i, (cn, cv) in enumerate(zip(creditos_nom, creditos_vpn)):
        ax.text(i - w/2, cn+0.5, f"{cn:.0f}", ha='center', fontsize=8)
        ax.text(i + w/2, cv+0.5, f"{cv:.1f}", ha='center', fontsize=8)
    ax.set_xticks(x)
    ax.set_xticklabels([f"α_fin={a}%" for a in alphas_pct])
    ax.set_ylabel("Crédito diferido (kUSD)")
    ax.legend(fontsize=10); ax.grid(True, alpha=0.3, axis='y')
    plt.tight_layout()
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig("Garver_MP_Fig03_Credito_diferido.png"),
                    dpi=160, bbox_inches='tight')
    plt.close()


def figura_04_cronograma_corredor(resultados_multi):
    """Para cada escenario, dibuja un "Gantt" simple de cuándo se construye cada circuito."""
    n_esc = len(resultados_multi)
    fig, axes = plt.subplots(1, n_esc, figsize=(5*n_esc, 5), sharey=True)
    if n_esc == 1: axes = [axes]
    fig.suptitle("FIG 4 — Cronograma de construcción por corredor (Garver 6-bus)",
                 fontsize=11, fontweight='bold')

    corredores_con_nuevas = sorted({j for rm in resultados_multi
                                    for j, lst in rm['cronograma'].items() if lst})

    for ax, rm in zip(axes, resultados_multi):
        ax.set_title(rm['titulo'][:25], fontsize=9)
        y_pos = {j: i for i, j in enumerate(corredores_con_nuevas)}
        for j in corredores_con_nuevas:
            for (t, k) in rm['cronograma'].get(j, []):
                ax.scatter(ANO_INI + t - 1, y_pos[j],
                           s=140, c='#D32F2F', edgecolors='black', zorder=3)
                ax.text(ANO_INI + t - 1, y_pos[j], f"{k}",
                        ha='center', va='center', color='white',
                        fontsize=7, fontweight='bold', zorder=4)
        ax.set_yticks(list(y_pos.values()))
        ax.set_yticklabels([f"[{j}]" for j in corredores_con_nuevas], fontsize=8)
        ax.set_xlim(ANO_INI - 0.5, ANO_INI + NT - 0.5)
        ax.set_xlabel("Año")
        ax.grid(True, alpha=0.3, axis='x')
        ax.invert_yaxis()
    plt.tight_layout()
    if GUARDAR_GRAFICAS:
        plt.savefig(fpath_fig("Garver_MP_Fig04_Cronograma.png"),
                    dpi=160, bbox_inches='tight')
    plt.close()


# =============================================================================
# 7. IMPRIMIR RESULTADOS DETALLADOS
# =============================================================================

def imprimir_resultados_detallados(resultados_multi):
    print("\n" + "="*70)
    print("  RESULTADOS DETALLADOS — MULTIPERIODO (Garver 6-bus)")
    print("="*70)

    if not resultados_multi:
        return
    base_vpn = resultados_multi[0]['inversion_vpn_total']
    base_nom = resultados_multi[0]['inversion_total']

    for rm in resultados_multi:
        print(f"\n  ── {rm['titulo']} ──")
        print(f"     Inversión nominal total: {rm['inversion_total']:>9,.1f} kUSD")
        print(f"     Inversión VPN total:     {rm['inversion_vpn_total']:>9,.1f} kUSD")
        if rm['inversion_vpn_total'] != base_vpn:
            print(f"     Crédito diferido (nom):  {base_nom - rm['inversion_total']:>+9,.1f} kUSD")
            print(f"     Crédito diferido (VPN):  {base_vpn - rm['inversion_vpn_total']:>+9,.1f} kUSD")
        print(f"     Cronograma de construcción:")
        for t in PERIODOS:
            exp = rm['expansion_por_t'].get(t, {})
            if exp:
                exp_str = ', '.join(f"{j}:+{n}" for j, n in sorted(exp.items()))
                print(f"       Año {ANO_INI+t-1} (t={t:>2}): {exp_str}")


# =============================================================================
# 8. MAIN
# =============================================================================

if __name__ == "__main__":

    print("\n" + "="*70)
    print("  TEP — COSTO DIFERIDO POR BTM EN SISTEMA GARVER 6-BUS")
    print("  MILP DC MULTIPERIODO con tasa de descuento — Pyomo + Gurobi")
    print(f"  Horizonte: {NT} periodos ({ANO_INI}–{ANO_INI+NT-1})  |  r = {int(TASA_DESCUENTO*100)}%")
    print("  Tesis Ingeniería Eléctrica — Uniandes 2026")
    print("="*70)

    print("\n" + "─"*70)
    print("  Concepto: el modelo multiperiodo construye los circuitos del plan")
    print("  base secuencialmente con la demanda creciente. Al introducir BTM,")
    print("  algunos circuitos se POSTERGAN (diferimiento puro) y otros se")
    print("  EVITAN completamente (extensión natural del diferimiento). El")
    print("  factor de descuento traduce ambos efectos en una métrica única.")
    print("─"*70)

    # ── Paso 0: Exportar datos del sistema ────────────────────────────────────
    print("\n  [0/6] Exportando datos del sistema (nodos, corredores, crecimiento)...")
    exportar_datos_sistema()

    # ── Paso 1: Ejecutar análisis multiperiodo ────────────────────────────────
    print("\n  [1/6] Ejecutando análisis multiperiodo con escenarios BTM...")
    print("        Penetraciones α_fin evaluadas:", [ALPHA_LABELS[a] for a in ALPHAS_FIN])
    resultados_multi = ejecutar_analisis_btm_multiperiodo()

    # ── Paso 2: Exportar tablas ───────────────────────────────────────────────
    print("\n  [2/6] Exportando tablas de resultados multiperiodo...")
    exportar_inversion_anual(resultados_multi)
    exportar_inversion_vpn_anual(resultados_multi)
    exportar_demanda_anual(resultados_multi)
    exportar_vpn_total(resultados_multi)
    exportar_cronograma(resultados_multi)
    df_resumen = exportar_resumen_ejecutivo(resultados_multi)
    print(f"  → Tablas multiperiodo exportadas en {DIR_TABLAS}")

    # ── Paso 3: Imprimir resultados detallados ────────────────────────────────
    print("\n  [3/6] Resultados detallados por escenario:")
    imprimir_resultados_detallados(resultados_multi)

    # ── Paso 4: Generar figuras ───────────────────────────────────────────────
    print("\n  [4/6] Generando gráficas académicas...")
    figura_01_inversion_anual(resultados_multi)
    figura_02_demanda_evolucion(resultados_multi)
    figura_03_credito_diferido(resultados_multi)
    figura_04_cronograma_corredor(resultados_multi)

    # ── Paso 5: Excel maestro ─────────────────────────────────────────────────
    print("\n  [5/6] Generando archivo Excel consolidado...")
    exportar_excel_maestro()

    # ── Paso 6: Resumen final ─────────────────────────────────────────────────
    print("\n  [6/6] Resumen final del análisis:")
    print("\n" + "="*70)
    print("  RESUMEN — CRÉDITO DIFERIDO MULTIPERIODO (Garver 6-bus)")
    print("="*70)

    if resultados_multi:
        base_vpn = resultados_multi[0]['inversion_vpn_total']
        base_nom = resultados_multi[0]['inversion_total']
        print(f"\n  Caso base sin BTM:")
        print(f"     Inversión nominal: {base_nom:>9,.1f} kUSD")
        print(f"     Inversión VPN:     {base_vpn:>9,.1f} kUSD")
        for rm in resultados_multi[1:]:
            ap = rm['alpha_fin']
            a_max = max(ap.values()) if ap else 0
            ah_nom = base_nom - rm['inversion_total']
            ah_vpn = base_vpn - rm['inversion_vpn_total']
            pct_vpn = ah_vpn/base_vpn*100 if base_vpn else 0
            print(f"\n  α_fin = {int(a_max*100):02d}%:")
            print(f"     Inversión nominal: {rm['inversion_total']:>9,.1f} kUSD "
                  f"(crédito diferido: {ah_nom:+,.1f} kUSD)")
            print(f"     Inversión VPN:     {rm['inversion_vpn_total']:>9,.1f} kUSD "
                  f"(crédito diferido: {ah_vpn:+,.1f} kUSD, {pct_vpn:.1f}%)")

    print("\n" + "="*70)
    print("  ✅  EJECUCIÓN COMPLETADA")
    print(f"     Resultados en: {os.path.abspath(OUTPUT_DIR)}")
    print(f"     ├── figuras/   ({len(os.listdir(DIR_FIGURAS))} archivos)")
    print(f"     ├── tablas/    ({len(os.listdir(DIR_TABLAS))} archivos)")
    print(f"     └── reportes/  ({len(os.listdir(DIR_REPORTES))} archivos)")
    print("="*70)