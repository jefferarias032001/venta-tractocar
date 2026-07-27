# -*- coding: utf-8 -*-
"""
TRACTOCAR · Dashboard de Ventas Nacional
-----------------------------------------
Fuentes:
  - TRACTOCAR_UNIFICADO.xlsx (generado por procesar.py)
  - Presupuesto PLANTILLA2 (Analisis de Margen y Venta)
  - Solicitudes Nacionales_Report Data.xlsx (pendientes por planillar)

Sub-segmentos dentro de NACIONAL:
  Token CED  -> CEDI
  Token TL   -> TL
  Resto      -> NAC  (NACIONAL-NACIONAL)

USO: python procesar_ventas.py
"""

import os, json, datetime as dt, warnings, calendar, base64
import pandas as pd
import numpy as np

warnings.filterwarnings("ignore")

BASE = os.path.dirname(os.path.abspath(__file__))

# ======================== CONFIGURACION ========================

RUTA_UNIFICADO = r"C:\Users\jarias\Desktop\tractocar-ventas\TRACTOCAR_UNIFICADO.xlsx"

RUTA_PPTO = (
    r"C:\Users\jarias\OneDrive - TRACTOCAR LOGISTICS SAS"
    r"\Archivos de Data Quality Analyst Tractocar - Analisis Operacion y Venta"
    r"\Analisis de Margen y Venta - Proy Julio 2026 - jeffer.xlsx"
)
HOJA_PPTO = "PLANTILLA2"

RUTA_SOLICITUDES = (
    r"C:\Users\jarias\OneDrive - TRACTOCAR LOGISTICS SAS"
    r"\Archivos de Data Quality Analyst Tractocar - Analisis Operacion y Venta"
    r"\Automatizaciones\Despachos\Solicitudes Nacionales_Report Data.xlsx"
)

# Mapeo NIT (9 digitos) -> Codigo cliente
# Nota: JERONIMO MARTINS (900480569 / PEPS) excluido por defecto en el dashboard
NIT_A_COD = {
    "860013771": "AJOV",   # AJOVER DARNEL SAS
    "817002753": "DRYP",   # DRYPERS ANDINA SA
    "860513970": "MILP",   # C.I MILPA SA
    "860002274": "ETER",   # ETERNA SA
    "800059470": "ESEN",   # ESENTTIA BY PROPILCO
    "830006735": "ALPO",   # ALIMENTOS POLAR COLOMBIA SAS
    "860015753": "KIMB",   # COLOMBIANA KIMBERLY COLPAPEL
    "890300466": "TECN",   # TECNOQUIMICAS SA
    "860522056": "LAMI",   # LAMITECH SAS
    "890104438": "EMTR",   # EMPAQUES TRANSPARENTES SA
    "860001899": "CORP",   # CORPACERO SAS
    "800217481": "CASC",   # INVERSIONES CASCABEL SA
    "890900161": "PFAMI",  # PRODUCTOS FAMILIA SA
    "900226838": "NOCO",   # NOUVELLE COLOMBIA EU
    "860007277": "MERE",   # MEXICHEM RESINAS COLOMBIA
    # "900480569": "PEPS",  # JERONIMO MARTINS - desactivado
}

# ===============================================================


def norm_nit(x):
    if pd.isna(x):
        return ""
    s = "".join(c for c in str(x) if c.isdigit())
    return s[:9] if len(s) >= 9 else s


def get_subseg(token):
    t = str(token).strip().upper()
    if "CED" in t:
        return "CEDI"
    if t == "TL":
        return "TL"
    return "NAC"


def buscar_header_row_cod(ruta, hoja):
    raw = pd.read_excel(ruta, sheet_name=hoja, header=None, nrows=15)
    for i, row in raw.iterrows():
        vals = [str(v).strip().lower() for v in row if str(v).strip() not in ("nan", "")]
        if "cod" in vals:
            return i
    return 5


def buscar_header_row_ob(ruta, hoja):
    raw = pd.read_excel(ruta, sheet_name=hoja, header=None, nrows=10)
    for i, row in raw.iterrows():
        vals = [str(v).strip().upper() for v in row if str(v).strip() not in ("nan", "")]
        if "OB" in vals:
            return i
    return 4


def main():
    today = dt.date.today()
    mes_actual   = today.strftime("%Y-%m")
    mes_ant_date = today.replace(day=1) - dt.timedelta(days=1)
    mes_anterior = mes_ant_date.strftime("%Y-%m")
    dias_mes     = calendar.monthrange(today.year, today.month)[1]
    nombre_mes   = today.strftime("%B %Y").upper()

    print("=" * 64)
    print("TRACTOCAR - Dashboard Ventas Nacional")
    print(f"  Mes: {mes_actual}  |  Hoy: dia {today.day}/{dias_mes}")

    # ---- 1. PRESUPUESTO ----
    print(f"\n> Presupuesto: {os.path.basename(RUTA_PPTO)}")
    hrow = buscar_header_row_cod(RUTA_PPTO, HOJA_PPTO)
    df_p = pd.read_excel(RUTA_PPTO, sheet_name=HOJA_PPTO, header=hrow)
    df_p.columns = [str(c).strip() for c in df_p.columns]
    df_p = df_p[df_p["Cod"].notna()].copy()
    df_p["Cod"] = df_p["Cod"].astype(str).str.strip()

    budget = df_p.groupby("Cod").agg(
        PPTO       =("Venta total proyectada", "sum"),
        META_UTIL  =("Utilidad Total",         "sum"),
        M_VIAJES   =("Viajes totales",         "sum"),
        COMPRA_PPTO=("Compra total proyectada","sum"),
    ).reset_index()

    pct_s = pd.to_numeric(df_p.get("% INTER", 0), errors="coerce").fillna(0)
    vp_s  = pd.to_numeric(df_p.get("Venta total proyectada", 0), errors="coerce").fillna(0)
    df_p["_wp"] = pct_s * vp_s
    df_p["_vp"] = vp_s
    pct_agg = df_p.groupby("Cod").agg(_wp=("_wp","sum"), _vp=("_vp","sum")).reset_index()
    pct_agg["PCT_INTER_M"] = np.where(pct_agg["_vp"]>0, pct_agg["_wp"]/pct_agg["_vp"], 0)
    budget = budget.merge(pct_agg[["Cod","PCT_INTER_M"]], on="Cod", how="left")
    budget["PCT_INTER_M"] = budget["PCT_INTER_M"].fillna(0)
    print(f"  {len(budget)} clientes en presupuesto")

    # ---- Logo base64 ----
    logo_b64 = ""
    logo_path = r"C:\Users\jarias\Desktop\tractocar-portal\logo.png"
    try:
        with open(logo_path, "rb") as lf:
            logo_b64 = "data:image/png;base64," + base64.b64encode(lf.read()).decode()
    except Exception:
        pass

    # ---- 2. UNIFICADO (todas las fuentes) ----
    print(f"> Unificado: {os.path.basename(RUTA_UNIFICADO)}")
    u = pd.read_excel(RUTA_UNIFICADO, sheet_name="Union")
    u["_nit"] = u["ClienteNIT"].apply(norm_nit)
    u["Cod"]  = u["_nit"].map(NIT_A_COD)

    # ---- KPI por operación (mes actual, todas las fuentes) ----
    u_mes_all = u[u["Mes"] == mes_actual].copy()
    ops_kpi = {}
    for fuente, grp in u_mes_all.groupby("Fuente"):
        ops_kpi[fuente] = {
            "VENTA":    round(float(grp["AFacturar"].sum()), 0),
            "UTILIDAD": round(float(grp["Utilidad"].sum()), 0),
            "VIAJES":   int(grp["Manifiesto"].nunique()),
        }
    print(f"  Operaciones detectadas: {sorted(ops_kpi.keys())}")

    # Solo NACIONAL — usar CodCliente si existe, fallback a NIT mapping
    u_nac = u[u["Fuente"] == "NACIONAL"].copy()
    if "CodCliente" in u.columns:
        u_nac["Cod"] = u_nac["CodCliente"].astype(str).str.strip().replace({"": None, "nan": None, "<NA>": None})
    # Para filas sin CodCliente, intentar NIT mapping como respaldo
    mask_sin_cod = u_nac["Cod"].isna()
    if mask_sin_cod.any():
        u_nac.loc[mask_sin_cod, "Cod"] = u_nac.loc[mask_sin_cod, "_nit"].map(NIT_A_COD)
    u_nac = u_nac[u_nac["Cod"].notna()].copy()
    u_nac["Subseg"] = u_nac["Token"].apply(get_subseg)
    if "Dia" not in u_nac.columns:
        u_nac["Dia"] = pd.to_datetime(u_nac["Fecha"], errors="coerce").dt.day

    # Mes actual - datos diarios por (Cod, Subseg, Dia)
    m_act = u_nac[u_nac["Mes"] == mes_actual].copy()
    print(f"  {len(m_act):,} registros en {mes_actual} | dias: {sorted(m_act['Dia'].dropna().astype(int).unique())}")

    # Construir estructura diaria para JS: {Cod: {Subseg: {dia: [V, U, N]}}}
    daily_dict = {}
    for (cod, subseg, dia), grp in m_act.groupby(["Cod","Subseg","Dia"]):
        if cod not in daily_dict:
            daily_dict[cod] = {}
        if subseg not in daily_dict[cod]:
            daily_dict[cod][subseg] = {}
        daily_dict[cod][subseg][str(int(dia))] = [
            round(float(grp["AFacturar"].sum()), 0),
            round(float(grp["Utilidad"].sum()), 0),
            int(grp["Manifiesto"].nunique()),
        ]

    # Venta de ayer y hoy: usar los ultimos 2 dias con datos disponibles
    dias_disponibles = sorted(m_act["Dia"].dropna().astype(int).unique())
    dia_hoy  = dias_disponibles[-1]  if len(dias_disponibles) >= 1 else None
    dia_ayer = dias_disponibles[-2]  if len(dias_disponibles) >= 2 else None
    label_hoy  = f"dia {dia_hoy}"  if dia_hoy  else "—"
    label_ayer = f"dia {dia_ayer}" if dia_ayer else "—"
    print(f"  Ultimo dato: dia {dia_hoy} | Penultimo: dia {dia_ayer}")

    fijo_hoy  = {}
    fijo_ayer = {}
    if dia_hoy:
        hoy_df = m_act[m_act["Dia"] == dia_hoy].groupby("Cod")["AFacturar"].sum()
        fijo_hoy = hoy_df.to_dict()
    if dia_ayer:
        ayer_df = m_act[m_act["Dia"] == dia_ayer].groupby("Cod")["AFacturar"].sum()
        fijo_ayer = ayer_df.to_dict()

    # Mes anterior (total por cliente, todas las operaciones)
    m_ant = u_nac[u_nac["Mes"] == mes_anterior]
    mes_ant_agg = (m_ant.groupby("Cod")["AFacturar"]
                   .sum().rename("VENTA_MES_ANT").reset_index())

    # ---- 3. PENDIENTES POR PLANILLAR ----
    print(f"> Solicitudes: {os.path.basename(RUTA_SOLICITUDES)}")
    hrow_sol = buscar_header_row_ob(RUTA_SOLICITUDES, "Sheet1")
    sol = pd.read_excel(RUTA_SOLICITUDES, sheet_name="Sheet1", header=hrow_sol)
    sol.columns = [str(c).strip() for c in sol.columns]
    filtrado = sol[
        (sol["OB_NOTES_CANCEL_USER"] == "-") &
        (sol["SHIP_STATUS_ENROUTE"].isna())
    ].copy()
    pendiente = (filtrado
                 .groupby("OB_CUSTOMER_CODE")
                 .agg(P_PLANILLAR=("OB_RATE_RECEIVABLE","sum"),
                      N_PLANILLAR=("OB","count"))
                 .reset_index()
                 .rename(columns={"OB_CUSTOMER_CODE":"Cod"}))
    print(f"  {len(filtrado)} OB pendientes | ${filtrado['OB_RATE_RECEIVABLE'].sum():,.0f}")

    # ---- 4. TABLA BASE (para resumen Python) ----
    # Calcular ejecutado NAC completo (todos los dias) para el print
    ej_all = (m_act.groupby("Cod")
              .agg(EJECUTADO=("AFacturar","sum"), UTILIDAD=("Utilidad","sum"),
                   VIAJES=("Manifiesto","nunique"))
              .reset_index())

    df = budget.copy()
    for right in [ej_all, mes_ant_agg, pendiente]:
        df = df.merge(right, on="Cod", how="left")

    num_cols = ["PPTO","META_UTIL","M_VIAJES","COMPRA_PPTO","PCT_INTER_M",
                "EJECUTADO","UTILIDAD","VIAJES","VENTA_MES_ANT","P_PLANILLAR","N_PLANILLAR"]
    for c in num_cols:
        if c not in df.columns: df[c] = 0.0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # Agrupar clientes sin venta ni mes anterior en OTROS CLIENTES
    sin_venta = (df["EJECUTADO"] == 0) & (df["VENTA_MES_ANT"] == 0)
    otros = df[sin_venta].copy()
    df    = df[~sin_venta].copy()
    otros_count = len(otros)

    tot_pp  = float(df["PPTO"].sum()) + float(otros["PPTO"].sum())
    tot_ej  = float(df["EJECUTADO"].sum())
    tot_ut  = float(df["UTILIDAD"].sum())
    tot_pla = float(df["P_PLANILLAR"].sum()) + float(otros["P_PLANILLAR"].sum())

    print(f"\n{'='*64}")
    print(f"  PPTO total:  ${tot_pp:>22,.0f}")
    print(f"  Ejecutado:   ${tot_ej:>22,.0f}")
    print(f"  Utilidad:    ${tot_ut:>22,.0f}  ({tot_ut/tot_ej:.1%})" if tot_ej else "")
    print(f"  Planillar:   ${tot_pla:>22,.0f}")
    print(f"  Otros clit:  {otros_count} agrupados en 'OTROS CLIENTES'")
    print("=" * 64)

    # ---- 5. PREPARAR PAYLOAD PARA JS ----
    # Budget dict por Cod
    ppto_js = {}
    for _, r in budget.iterrows():
        ppto_js[r["Cod"]] = {
            "PPTO":       float(r["PPTO"]),
            "META_UTIL":  float(r["META_UTIL"]),
            "M_VIAJES":   float(r["M_VIAJES"]),
            "PCT_INTER_M":float(r["PCT_INTER_M"]),
        }

    # Mes anterior dict
    mes_ant_js = {r["Cod"]: float(r["VENTA_MES_ANT"])
                  for _, r in mes_ant_agg.iterrows()}

    # Pendiente dict
    pend_js = {}
    for _, r in pendiente.iterrows():
        pend_js[r["Cod"]] = {
            "monto": float(r["P_PLANILLAR"]),
            "n":     int(r["N_PLANILLAR"]),
        }

    # Otros clientes (agrupados)
    otros_js = {
        "PPTO":      float(otros["PPTO"].sum()),
        "META_UTIL": float(otros["META_UTIL"].sum()),
        "M_VIAJES":  float(otros["M_VIAJES"].sum()),
        "P_PLANILLAR":float(otros["P_PLANILLAR"].sum()),
        "n":         int(len(otros)),
        "nombres":   sorted(otros["Cod"].tolist()),
    } if len(otros) else None

    # Lista de clientes activos (con venta)
    clientes_activos = sorted(df["Cod"].tolist())

    payload = (
        f"window.DIARIO={json.dumps(daily_dict, ensure_ascii=False)};"
        f"window.PPTO_DATA={json.dumps(ppto_js, ensure_ascii=False)};"
        f"window.MES_ANT={json.dumps(mes_ant_js, ensure_ascii=False)};"
        f"window.PENDIENTE={json.dumps(pend_js, ensure_ascii=False)};"
        f"window.FIJO_HOY={json.dumps({k:float(v) for k,v in fijo_hoy.items()}, ensure_ascii=False)};"
        f"window.FIJO_AYER={json.dumps({k:float(v) for k,v in fijo_ayer.items()}, ensure_ascii=False)};"
        f"window.OTROS={json.dumps(otros_js, ensure_ascii=False)};"
        f"window.CLIENTES={json.dumps(clientes_activos, ensure_ascii=False)};"
        f"window.OPS_KPI={json.dumps(ops_kpi, ensure_ascii=False)};"
        f"window.LOGO='{logo_b64}';"
        f"window.META={{mes:'{mes_actual}',diaActual:{today.day},diasMes:{dias_mes},"
        f"nombreMes:'{nombre_mes}',labelHoy:'{label_hoy}',labelAyer:'{label_ayer}',"
        f"generado:'{dt.datetime.now().strftime('%Y-%m-%d %H:%M')}'}};"
    ).replace("</", "<\\/")

    html = HTML_TEMPLATE.replace("/*__DATOS__*/", payload)
    ruta_html = os.path.join(BASE, "dashboard_ventas.html")
    with open(ruta_html, "w", encoding="utf-8") as f:
        f.write(html)

    # También guardar como index.html para Netlify
    ruta_index = os.path.join(BASE, "index.html")
    with open(ruta_index, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n  HTML -> {ruta_html}")
    print(f"  HTML -> {ruta_index}  (Netlify)")
    import webbrowser
    try:
        webbrowser.open("file://" + ruta_html.replace(os.sep, "/"))
    except Exception:
        pass


# ================================================================
# HTML TEMPLATE
# ================================================================
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Ventas TRACTOCAR</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:#070d16;color:#e6edf3;min-height:100vh}

/* HEADER */
header{background:linear-gradient(135deg,#080e1a 0%,#0a1525 100%);padding:14px 24px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid rgba(249,115,22,.3);gap:12px;flex-wrap:wrap;position:relative;overflow:hidden}
header::after{content:'';position:absolute;bottom:0;left:0;right:0;height:2px;background:linear-gradient(90deg,transparent,#f97316,#fbbf24,#f97316,transparent)}
.header-left{display:flex;align-items:center;gap:14px}
.header-logo{height:36px;width:auto;filter:brightness(10);opacity:.9}
.header-logo-text{font-size:1.3rem;font-weight:800;letter-spacing:.12em;color:#fff;text-transform:uppercase}
header h1{font-size:1.1rem;font-weight:700;letter-spacing:.04em}
header .sub{font-size:.75rem;color:#64748b;margin-top:2px}
.btn{border:none;border-radius:7px;padding:7px 16px;font-size:.8rem;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:5px;white-space:nowrap}
.btn-dl{background:linear-gradient(135deg,#f97316,#ea580c);color:#fff;box-shadow:0 2px 8px rgba(249,115,22,.35)}.btn-dl:hover{background:linear-gradient(135deg,#ea580c,#c2410c)}

/* ── KPI CARDS ── */
.kpi-section{background:linear-gradient(180deg,#04090f 0%,#070d16 100%);padding:24px 24px 20px;border-bottom:1px solid #0e1e2e}
.kpi-supertitle{font-size:.65rem;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:rgba(255,255,255,.22);margin-bottom:14px;padding-left:2px}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:16px}
.kpi-card{position:relative;border-radius:18px;padding:22px 20px 18px;transition:transform .3s cubic-bezier(.22,1,.36,1),box-shadow .3s ease;overflow:hidden;transform-style:preserve-3d;cursor:default}
.kpi-card::before{content:'';position:absolute;inset:0;border-radius:18px;background:linear-gradient(135deg,rgba(255,255,255,.10) 0%,rgba(255,255,255,0) 55%);pointer-events:none;z-index:1}
.kpi-card::after{content:'';position:absolute;top:0;left:-80%;width:60%;height:100%;background:linear-gradient(90deg,transparent,rgba(255,255,255,.035),transparent);animation:shimmer 5s ease-in-out infinite;pointer-events:none}
@keyframes shimmer{0%{left:-80%}100%{left:160%}}
.kpi-card:hover{transform:perspective(700px) rotateX(-4deg) translateY(-8px) scale(1.01)}
/* Nacional - verde */
.kpi-nac{background:linear-gradient(145deg,#0e2018 0%,#071410 100%);box-shadow:0 6px 0 #041a0b,0 12px 40px rgba(0,0,0,.6),inset 0 1px 0 rgba(74,222,128,.1),0 0 0 1px rgba(34,197,94,.18)}
.kpi-nac:hover{box-shadow:0 2px 0 #041a0b,0 24px 60px rgba(0,0,0,.7),inset 0 1px 0 rgba(74,222,128,.12),0 0 0 1px rgba(34,197,94,.3)}
/* IMPO - azul */
.kpi-imp{background:linear-gradient(145deg,#0d1e33 0%,#07121f 100%);box-shadow:0 6px 0 #04122a,0 12px 40px rgba(0,0,0,.6),inset 0 1px 0 rgba(96,165,250,.1),0 0 0 1px rgba(59,130,246,.18)}
.kpi-imp:hover{box-shadow:0 2px 0 #04122a,0 24px 60px rgba(0,0,0,.7),inset 0 1px 0 rgba(96,165,250,.12),0 0 0 1px rgba(59,130,246,.3)}
/* EXPO - naranja */
.kpi-exp{background:linear-gradient(145deg,#211508 0%,#150d04 100%);box-shadow:0 6px 0 #1a0c00,0 12px 40px rgba(0,0,0,.6),inset 0 1px 0 rgba(251,146,60,.1),0 0 0 1px rgba(249,115,22,.18)}
.kpi-exp:hover{box-shadow:0 2px 0 #1a0c00,0 24px 60px rgba(0,0,0,.7),inset 0 1px 0 rgba(251,146,60,.12),0 0 0 1px rgba(249,115,22,.3)}
/* CEDIS - morado */
.kpi-ced{background:linear-gradient(145deg,#18102a 0%,#0e0918 100%);box-shadow:0 6px 0 #120840,0 12px 40px rgba(0,0,0,.6),inset 0 1px 0 rgba(192,132,252,.1),0 0 0 1px rgba(168,85,247,.18)}
.kpi-ced:hover{box-shadow:0 2px 0 #120840,0 24px 60px rgba(0,0,0,.7),inset 0 1px 0 rgba(192,132,252,.12),0 0 0 1px rgba(168,85,247,.3)}

.kpi-accent{width:36px;height:3px;border-radius:2px;margin-bottom:14px}
.kpi-nac .kpi-accent{background:linear-gradient(90deg,#22c55e,#4ade80)}
.kpi-imp .kpi-accent{background:linear-gradient(90deg,#3b82f6,#60a5fa)}
.kpi-exp .kpi-accent{background:linear-gradient(90deg,#f97316,#fb923c)}
.kpi-ced .kpi-accent{background:linear-gradient(90deg,#a855f7,#c084fc)}

.kpi-op-label{font-size:.63rem;font-weight:800;letter-spacing:.18em;text-transform:uppercase;margin-bottom:10px}
.kpi-nac .kpi-op-label{color:#4ade80}.kpi-imp .kpi-op-label{color:#60a5fa}
.kpi-exp .kpi-op-label{color:#fb923c}.kpi-ced .kpi-op-label{color:#c084fc}

.kpi-venta-val{font-size:2.1rem;font-weight:800;line-height:1;color:#fff;letter-spacing:-.03em}
.kpi-venta-unit{font-size:1rem;font-weight:600;opacity:.5;margin-left:2px}
.kpi-venta-sub{font-size:.62rem;color:rgba(255,255,255,.28);margin:4px 0 14px;text-transform:uppercase;letter-spacing:.06em}

.kpi-stats{display:flex;gap:18px}
.kpi-stat-val{font-size:.92rem;font-weight:700}
.kpi-nac .kpi-stat-val{color:#4ade80}.kpi-imp .kpi-stat-val{color:#60a5fa}
.kpi-exp .kpi-stat-val{color:#fb923c}.kpi-ced .kpi-stat-val{color:#c084fc}
.kpi-stat-lbl{font-size:.6rem;color:rgba(255,255,255,.28);text-transform:uppercase;letter-spacing:.07em;margin-top:2px}

.kpi-badge{position:absolute;top:16px;right:16px;font-size:.7rem;font-weight:700;background:rgba(0,0,0,.35);border-radius:20px;padding:3px 9px;z-index:2;letter-spacing:.03em}
.kpi-nac .kpi-badge{color:#4ade80;border:1px solid rgba(74,222,128,.2)}
.kpi-imp .kpi-badge{color:#60a5fa;border:1px solid rgba(96,165,250,.2)}
.kpi-exp .kpi-badge{color:#fb923c;border:1px solid rgba(251,146,60,.2)}
.kpi-ced .kpi-badge{color:#c084fc;border:1px solid rgba(192,132,252,.2)}

/* CHIPS RESUMEN */
.meta-bar{display:flex;gap:8px;flex-wrap:wrap;padding:9px 24px;background:#040a12;border-bottom:1px solid #0e1e2e}
.chip{background:#0e1a26;border:1px solid #1a2d3d;border-radius:18px;padding:4px 12px;font-size:.73rem;color:#64748b}
.chip b{color:#f97316}

/* BARRA DE FILTROS */
.filter-bar{display:flex;gap:10px;flex-wrap:wrap;align-items:center;padding:10px 24px;background:#0c1a27;border-bottom:1px solid #1e2d3d}
.filter-group{display:flex;align-items:center;gap:6px}
.filter-label{font-size:.73rem;color:#64748b;font-weight:600;text-transform:uppercase;letter-spacing:.04em}

/* SEGMENTO TABS */
.seg-btn{background:#1e2d3d;border:1px solid #2d3f52;color:#94a3b8;border-radius:6px;padding:5px 13px;font-size:.77rem;cursor:pointer;font-weight:600;transition:all .15s}
.seg-btn.active{background:#f97316;color:#fff;border-color:#f97316}
.seg-btn:hover:not(.active){background:#253548;color:#e2e8f0}

/* RANGO DE DIAS */
.range-wrap{display:flex;align-items:center;gap:8px;font-size:.78rem;color:#94a3b8}
.day-input{width:44px;background:#1e2d3d;border:1px solid #2d3f52;color:#e2e8f0;border-radius:5px;padding:3px 6px;text-align:center;font-size:.78rem}
.day-input:focus{outline:none;border-color:#f97316}
.range-sep{color:#64748b}

/* CLIENTE FILTER */
.client-wrap{position:relative}
.client-btn{background:#1e2d3d;border:1px solid #2d3f52;color:#94a3b8;border-radius:6px;padding:5px 12px;font-size:.77rem;cursor:pointer}
.client-btn:hover{background:#253548;color:#e2e8f0}
.client-dropdown{display:none;position:absolute;top:calc(100% + 4px);left:0;background:#0f1923;border:1px solid #2d3f52;border-radius:8px;padding:10px;z-index:100;min-width:200px;max-height:320px;overflow-y:auto;box-shadow:0 8px 24px rgba(0,0,0,.5)}
.client-dropdown.open{display:block}
.client-item{display:flex;align-items:center;gap:8px;padding:4px 2px;cursor:pointer;font-size:.78rem;color:#cbd5e1;white-space:nowrap}
.client-item:hover{color:#e2e8f0}
.client-item input[type=checkbox]{accent-color:#f97316;width:14px;height:14px;cursor:pointer}
.dd-actions{display:flex;gap:8px;margin-bottom:8px;border-bottom:1px solid #1e2d3d;padding-bottom:8px}
.dd-link{color:#f97316;font-size:.72rem;cursor:pointer;text-decoration:underline}

/* TABLA */
.container{padding:14px 10px;overflow-x:auto}
table{width:100%;border-collapse:collapse;min-width:1700px;font-size:.77rem}
thead th{background:#0c2d3f;color:#94a3b8;font-weight:600;text-align:right;padding:7px 8px;border-bottom:2px solid #1e3a4a;white-space:nowrap;position:sticky;top:0;z-index:2;cursor:pointer;user-select:none}
thead th:first-child{text-align:left}
thead th:hover{background:#0d3547;color:#e2e8f0}
thead th.sort-asc::after{content:' \25B2';color:#f97316;font-size:.68em}
thead th.sort-desc::after{content:' \25BC';color:#f97316;font-size:.68em}
tbody tr{border-bottom:1px solid #1a2533;transition:background .12s}
tbody tr:hover{background:#162130}
tbody tr.total-row{background:#0c2d3f;border-top:2px solid #f97316;font-weight:700}
tbody tr.otros-row{background:#0f1a25;color:#556778;font-style:italic}
td{padding:6px 8px;text-align:right;white-space:nowrap}
td:first-child{text-align:left;font-weight:600}
.pos{color:#22c55e}.neg{color:#ef4444}.neu{color:#3d5165}
.au::before{content:'\25B2 '}.ad::before{content:'\25BC '}
.pct-bar{display:inline-flex;align-items:center;gap:5px;min-width:105px}
.bar-bg{flex:1;background:#1e2d3d;border-radius:3px;height:6px;min-width:44px}
.bar-f{height:100%;border-radius:3px}
.b-ok{background:#22c55e}.b-wn{background:#f59e0b}.b-bd{background:#ef4444}
.pl{font-weight:600;min-width:38px;text-align:right}
.note{font-size:.68rem;color:#445566;margin-top:1px}
</style>
</head>
<body>
<header>
  <div class="header-left">
    <img class="header-logo" id="headerLogo" src="" alt="TRACTOCAR">
    <div>
      <div style="font-size:1.1rem;font-weight:700;letter-spacing:.04em">Ventas TRACTOCAR &nbsp;<span style="color:#f97316;font-size:.85rem">▸ Nacional</span></div>
      <div class="sub" id="subTitle"></div>
    </div>
  </div>
  <button class="btn btn-dl" onclick="descargarCSV()">&#8659; Descargar CSV</button>
</header>

<!-- KPI Cards por operación -->
<div class="kpi-section">
  <div class="kpi-supertitle">&#9679; Resumen por operación — mes actual</div>
  <div class="kpi-grid" id="kpiGrid"></div>
</div>

<div class="meta-bar" id="metaBar"></div>

<div class="filter-bar">
  <!-- Segmento -->
  <div class="filter-group">
    <span class="filter-label">Segmento</span>
    <button class="seg-btn active" data-seg="TODOS"    onclick="setSeg(this)">Todos</button>
    <button class="seg-btn"        data-seg="NAC"      onclick="setSeg(this)">Nacional</button>
    <button class="seg-btn"        data-seg="CEDI"     onclick="setSeg(this)">CEDI</button>
    <button class="seg-btn"        data-seg="TL"       onclick="setSeg(this)">Transporte Local</button>
  </div>

  <!-- Rango dias -->
  <div class="filter-group">
    <span class="filter-label">Rango dias</span>
    <div class="range-wrap">
      <span>Dia</span>
      <input class="day-input" id="diaDesde" type="number" min="1" value="1" onchange="recalc()">
      <span class="range-sep">al</span>
      <input class="day-input" id="diaHasta" type="number" min="1" value="1" onchange="recalc()">
      <span id="diasLabel" style="color:#64748b;font-size:.72rem"></span>
    </div>
  </div>

  <!-- Clientes -->
  <div class="filter-group">
    <span class="filter-label">Clientes</span>
    <div class="client-wrap">
      <button class="client-btn" onclick="toggleDD()" id="clientBtn">Filtrar &#9660;</button>
      <div class="client-dropdown" id="clientDD">
        <div class="dd-actions">
          <span class="dd-link" onclick="selAll(true)">Todos</span>
          <span class="dd-link" onclick="selAll(false)">Ninguno</span>
        </div>
        <div id="clientList"></div>
      </div>
    </div>
  </div>
</div>

<div class="container">
  <table id="tbl">
    <thead>
      <tr id="thr">
        <th data-k="Cod">CODIGO</th>
        <th data-k="PPTO">PPTO MES</th>
        <th data-k="PROYECCION">PROYECCION</th>
        <th data-k="DIF_PROV_PPTO">DIF PROY vs PPTO</th>
        <th data-k="PCT_CUMPL">% CUMPL</th>
        <th data-k="EJECUTADO">EJECUTADO</th>
        <th data-k="VENTA_MES_ANT">VENTA M. ANT.</th>
        <th data-k="DIF_DIAS">DIF (DIAS)</th>
        <th data-k="VIAJES">VIAJES</th>
        <th data-k="M_VIAJES">M. VIAJES</th>
        <th data-k="VENTA_AYER" id="thAyer">VENTA AYER</th>
        <th data-k="VENTA_HOY"  id="thHoy">VENTA HOY</th>
        <th data-k="META_VENTA_FINAL">META VENTA FINAL</th>
        <th data-k="META_UTIL">META UTILIDAD</th>
        <th data-k="UTILIDAD">UTILIDAD</th>
        <th data-k="PROY_UTILIDAD">PROY. UTILIDAD</th>
        <th data-k="PCT_INTER">% INTER</th>
        <th data-k="PCT_INTER_M">% INTER M</th>
        <th data-k="P_PLANILLAR">P. PLANILLAR</th>
      </tr>
    </thead>
    <tbody id="tbody"></tbody>
  </table>
</div>

<script>
/*__DATOS__*/

/* ---- formato ---- */
var COP = new Intl.NumberFormat('es-CO',{style:'currency',currency:'COP',maximumFractionDigits:0});
var NUM = new Intl.NumberFormat('es-CO',{maximumFractionDigits:0});
function mn(v){if(!v||v===0)return '<span class="neu">-</span>';return COP.format(v);}
function arr(v){
  if(!v||v===0)return '<span class="neu">-</span>';
  var c=v>0?'pos au':'neg ad';return '<span class="'+c+'">'+COP.format(Math.abs(v))+'</span>';
}
function bar(v){
  var p=Math.min((v||0)*100,100);
  var bc=v>=1?'b-ok':v>=0.8?'b-wn':'b-bd';
  var lc=v>=1?'pos':v>=0.8?'':'neg';
  return '<div class="pct-bar"><span class="pl '+lc+'">'+(v*100).toFixed(0)+'%</span><div class="bar-bg"><div class="bar-f '+bc+'" style="width:'+p+'%"></div></div></div>';
}
function pct(v){if(!v&&v!==0)return '<span class="neu">-</span>';return ((v||0)*100).toFixed(1)+'%';}
function nnum(v){if(!v||v===0)return '<span class="neu">-</span>';return NUM.format(v);}

/* ---- estado filtros ---- */
var curSeg  = 'TODOS';
var d1 = 1, d2 = 1;
var excluidos = new Set();
var sortK = 'PPTO', sortAsc = false;

/* ---- calcular fila para un cliente ---- */
function calcFila(cod){
  var dias = window.DIARIO[cod] || {};
  var pp   = window.PPTO_DATA[cod] || {};
  var ma   = window.MES_ANT[cod] || 0;
  var pend = window.PENDIENTE[cod] || {monto:0,n:0};
  var m    = window.META || {};

  // Acumular segun segmento y rango de dias
  var V=0, U=0, N=0;
  var segs = curSeg==='TODOS' ? Object.keys(dias) : (dias[curSeg] ? [curSeg] : []);
  segs.forEach(function(s){
    var sd = dias[s] || {};
    for(var d=d1; d<=d2; d++){
      var e = sd[String(d)];
      if(e){V+=e[0]; U+=e[1]; N+=e[2];}
    }
  });

  var diasRango = d2 - d1 + 1;
  var diasMes   = m.diasMes || 31;
  var PROY  = diasRango > 0 ? V / diasRango * diasMes : 0;
  var DIF_PP = PROY - (pp.PPTO||0);
  var PCT_C  = pp.PPTO > 0 ? PROY / pp.PPTO : 0;
  var DIF_D  = V - (ma / diasMes * diasRango);
  var MAR    = V > 0 ? U / V : 0;
  var PROY_U = PROY * MAR;
  var META_V = Math.max((pp.PPTO||0) - V, 0);

  return {
    Cod: cod,
    PPTO: pp.PPTO||0, META_UTIL: pp.META_UTIL||0, M_VIAJES: pp.M_VIAJES||0,
    PCT_INTER_M: pp.PCT_INTER_M||0,
    EJECUTADO: V, UTILIDAD: U, VIAJES: N,
    VENTA_MES_ANT: ma,
    VENTA_AYER: window.FIJO_AYER[cod]||0,
    VENTA_HOY:  window.FIJO_HOY[cod]||0,
    PROYECCION: PROY, DIF_PROV_PPTO: DIF_PP, PCT_CUMPL: PCT_C,
    DIF_DIAS: DIF_D, PROY_UTILIDAD: PROY_U, PCT_INTER: MAR,
    META_VENTA_FINAL: META_V,
    P_PLANILLAR: pend.monto||0, N_PLANILLAR: pend.n||0,
  };
}

function renderRow(r, cls){
  var tr = document.createElement('tr');
  if(cls) tr.className = cls;
  tr.innerHTML =
    '<td>'+r.Cod+'</td>'+
    '<td>'+mn(r.PPTO)+'</td>'+
    '<td>'+mn(r.PROYECCION)+'</td>'+
    '<td>'+arr(r.DIF_PROV_PPTO)+'</td>'+
    '<td>'+bar(r.PCT_CUMPL)+'</td>'+
    '<td>'+mn(r.EJECUTADO)+'</td>'+
    '<td>'+mn(r.VENTA_MES_ANT)+'</td>'+
    '<td>'+arr(r.DIF_DIAS)+'</td>'+
    '<td>'+nnum(r.VIAJES)+'</td>'+
    '<td>'+nnum(r.M_VIAJES)+'</td>'+
    '<td>'+mn(r.VENTA_AYER)+'</td>'+
    '<td>'+mn(r.VENTA_HOY)+'</td>'+
    '<td>'+mn(r.META_VENTA_FINAL)+'</td>'+
    '<td>'+mn(r.META_UTIL)+'</td>'+
    '<td>'+mn(r.UTILIDAD)+'</td>'+
    '<td>'+mn(r.PROY_UTILIDAD)+'</td>'+
    '<td>'+pct(r.PCT_INTER)+'</td>'+
    '<td>'+pct(r.PCT_INTER_M)+'</td>'+
    '<td>'+(r.P_PLANILLAR>0?mn(r.P_PLANILLAR):'<span class="neu">-</span>')+'</td>';
  return tr;
}

function buildTable(){
  var clientes = (window.CLIENTES||[]).filter(function(c){return !excluidos.has(c);});

  // Calcular filas
  var rows = clientes.map(calcFila);

  // Ordenar (OTROS siempre al final)
  rows.sort(function(a,b){
    var va=a[sortK], vb=b[sortK];
    if(typeof va==='string') return sortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
    return sortAsc ? va-vb : vb-va;
  });

  var tbody = document.getElementById('tbody');
  tbody.innerHTML = '';
  rows.forEach(function(r){ tbody.appendChild(renderRow(r,'')); });

  // Fila OTROS CLIENTES
  var otros = window.OTROS;
  if(otros){
    var or = {
      Cod: 'OTROS CLIENTES ('+otros.n+')',
      PPTO: otros.PPTO, META_UTIL: otros.META_UTIL, M_VIAJES: otros.M_VIAJES,
      EJECUTADO:0, UTILIDAD:0, VIAJES:0, VENTA_MES_ANT:0,
      VENTA_AYER:0, VENTA_HOY:0, PROYECCION:0,
      DIF_PROV_PPTO: -otros.PPTO, PCT_CUMPL:0, DIF_DIAS:0,
      PROY_UTILIDAD:0, PCT_INTER:0, PCT_INTER_M:0,
      META_VENTA_FINAL: otros.PPTO, P_PLANILLAR: otros.P_PLANILLAR||0,
    };
    tbody.appendChild(renderRow(or,'otros-row'));
  }

  // Totales
  var totV=0,totU=0,totPP=0,totPRY=0,totAY=0,totHY=0,totMB=0,totMA=0,totPL=0,totV0=0,totMU=0;
  rows.forEach(function(r){
    totV+=r.EJECUTADO; totU+=r.UTILIDAD; totPP+=r.PPTO; totPRY+=r.PROYECCION;
    totAY+=r.VENTA_AYER; totHY+=r.VENTA_HOY; totMB+=r.META_VENTA_FINAL;
    totMA+=r.VENTA_MES_ANT; totPL+=r.P_PLANILLAR; totV0+=r.VIAJES; totMU+=r.META_UTIL;
  });
  if(otros){totPP+=otros.PPTO; totMB+=otros.PPTO; totMU+=otros.META_UTIL;}
  var totR={
    Cod:'TOTAL', PPTO:totPP, PROYECCION:totPRY,
    DIF_PROV_PPTO:totPRY-totPP, PCT_CUMPL:totPP>0?totPRY/totPP:0,
    EJECUTADO:totV, VENTA_MES_ANT:totMA,
    DIF_DIAS:totV-(totMA/(window.META.diasMes||31)*(d2-d1+1)),
    VIAJES:totV0, M_VIAJES: rows.reduce(function(s,r){return s+r.M_VIAJES;},0)+(otros?otros.M_VIAJES:0),
    VENTA_AYER:totAY, VENTA_HOY:totHY,
    META_VENTA_FINAL:totMB, META_UTIL:totMU, UTILIDAD:totU,
    PROY_UTILIDAD:totV>0?totPRY*(totU/totV):0,
    PCT_INTER:totV>0?totU/totV:0, PCT_INTER_M:totPP>0?totMU/totPP:0,
    P_PLANILLAR:totPL,
  };
  tbody.appendChild(renderRow(totR,'total-row'));

  // Actualizar chips
  var bar2=document.getElementById('metaBar');
  bar2.innerHTML='';
  var chips=[
    ['PPTO',COP.format(totPP)],
    ['PROYECCION',COP.format(totPRY)],
    ['% CUMPL',((totPP>0?totPRY/totPP:0)*100).toFixed(1)+'%'],
    ['EJECUTADO',COP.format(totV)],
    ['UTILIDAD',COP.format(totU)+(totV>0?' ('+(totU/totV*100).toFixed(1)+'%)':'')],
    ['P. PLANILLAR',COP.format(totPL)],
  ];
  chips.forEach(function(c){
    bar2.innerHTML+='<div class="chip">'+c[0]+': <b>'+c[1]+'</b></div>';
  });
}

function setSeg(btn){
  document.querySelectorAll('.seg-btn').forEach(function(b){b.classList.remove('active');});
  btn.classList.add('active');
  curSeg = btn.getAttribute('data-seg');
  buildTable();
}

function recalc(){
  var m = window.META||{};
  d1 = Math.max(1, parseInt(document.getElementById('diaDesde').value)||1);
  d2 = Math.min(m.diasMes||31, Math.max(d1, parseInt(document.getElementById('diaHasta').value)||d1));
  document.getElementById('diaHasta').value = d2;
  document.getElementById('diasLabel').textContent = '('+( d2-d1+1)+' dias)';
  buildTable();
}

function toggleDD(){
  document.getElementById('clientDD').classList.toggle('open');
}
document.addEventListener('click',function(e){
  if(!document.getElementById('clientBtn').contains(e.target) &&
     !document.getElementById('clientDD').contains(e.target)){
    document.getElementById('clientDD').classList.remove('open');
  }
});

function buildClientList(){
  var ul = document.getElementById('clientList');
  ul.innerHTML = '';
  (window.CLIENTES||[]).forEach(function(c){
    var li = document.createElement('label');
    li.className='client-item';
    var cb = document.createElement('input');
    cb.type='checkbox'; cb.value=c; cb.checked=true;
    cb.onchange=function(){
      if(cb.checked) excluidos.delete(c); else excluidos.add(c);
      buildTable();
    };
    li.appendChild(cb);
    li.appendChild(document.createTextNode(c));
    ul.appendChild(li);
  });
}

function selAll(v){
  document.querySelectorAll('#clientList input').forEach(function(cb){
    cb.checked=v;
    if(v) excluidos.delete(cb.value); else excluidos.add(cb.value);
  });
  buildTable();
}

function descargarCSV(){
  var cols=['Cod','PPTO','PROYECCION','DIF_PROV_PPTO','PCT_CUMPL','EJECUTADO','VENTA_MES_ANT',
            'DIF_DIAS','VIAJES','M_VIAJES','VENTA_AYER','VENTA_HOY','META_VENTA_FINAL',
            'META_UTIL','UTILIDAD','PROY_UTILIDAD','PCT_INTER','PCT_INTER_M','P_PLANILLAR'];
  var rows=[cols.join(';')];
  var clientes=(window.CLIENTES||[]).filter(function(c){return !excluidos.has(c);});
  clientes.map(calcFila).forEach(function(r){
    rows.push(cols.map(function(k){
      var v=r[k];
      if(typeof v==='string') return '"'+v+'"';
      if(typeof v==='number') return v.toFixed(2).replace('.',',');
      return '';
    }).join(';'));
  });
  var blob=new Blob(['﻿'+rows.join('\n')],{type:'text/csv;charset=utf-8'});
  var a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  var m=window.META||{};
  a.download='ventas_nacional_'+m.mes+'_dia'+d1+'al'+d2+'.csv'; a.click();
}

/* ---- ordenamiento ---- */
document.querySelectorAll('#thr th').forEach(function(th){
  th.addEventListener('click',function(){
    var k=th.getAttribute('data-k');
    document.querySelectorAll('#thr th').forEach(function(t){t.classList.remove('sort-asc','sort-desc');});
    if(sortK===k){sortAsc=!sortAsc;}else{sortK=k;sortAsc=false;}
    th.classList.add(sortAsc?'sort-asc':'sort-desc');
    buildTable();
  });
});

/* ---- Formato grande ---- */
function formatBig(v){
  var abs=Math.abs(v);
  if(abs>=1e9) return {val:(v/1e9).toFixed(2),unit:'B'};
  if(abs>=1e6) return {val:(v/1e6).toFixed(1),unit:'M'};
  if(abs>=1e3) return {val:(v/1e3).toFixed(0),unit:'K'};
  return {val:Math.round(v).toString(),unit:''};
}

/* ---- KPI Cards ---- */
function buildKPICards(){
  var ops=window.OPS_KPI||{};
  var cfg=[
    {key:'NACIONAL',cls:'kpi-nac',label:'Nacional'},
    {key:'IMPO',    cls:'kpi-imp',label:'Importación'},
    {key:'EXPO',    cls:'kpi-exp',label:'Exportación'},
    {key:'CEDIS',   cls:'kpi-ced',label:'CEDIS'},
  ];
  var grid=document.getElementById('kpiGrid');
  grid.innerHTML='';
  cfg.forEach(function(op){
    var d=ops[op.key]||{VENTA:0,UTILIDAD:0,VIAJES:0};
    var margin=d.VENTA>0?((d.UTILIDAD/d.VENTA)*100).toFixed(1)+'% margen':'— margen';
    var vf=formatBig(d.VENTA);
    var uf=formatBig(d.UTILIDAD);
    var card=document.createElement('div');
    card.className='kpi-card '+op.cls;
    card.innerHTML=
      '<div class="kpi-badge">'+margin+'</div>'+
      '<div class="kpi-accent"></div>'+
      '<div class="kpi-op-label">'+op.label+'</div>'+
      '<div class="kpi-venta-val">'+vf.val+'<span class="kpi-venta-unit">'+vf.unit+'</span></div>'+
      '<div class="kpi-venta-sub">Venta ejecutada mes actual</div>'+
      '<div class="kpi-stats">'+
        '<div class="kpi-stat"><div class="kpi-stat-val">'+uf.val+' '+uf.unit+'</div><div class="kpi-stat-lbl">Utilidad</div></div>'+
        '<div class="kpi-stat"><div class="kpi-stat-val">'+NUM.format(d.VIAJES)+'</div><div class="kpi-stat-lbl">Viajes</div></div>'+
      '</div>';
    grid.appendChild(card);
  });
}

/* ---- INIT ---- */
(function(){
  var m = window.META||{};
  document.getElementById('subTitle').textContent =
    m.nombreMes+' · Actualizado: '+m.generado;

  // Logo
  var logoEl=document.getElementById('headerLogo');
  if(window.LOGO && window.LOGO.length>30){
    logoEl.src=window.LOGO;
  } else {
    logoEl.style.display='none';
  }

  // Etiquetas ayer/hoy con dia real
  document.getElementById('thAyer').textContent = 'VENTA '+m.labelAyer.toUpperCase();
  document.getElementById('thHoy').textContent  = 'VENTA '+m.labelHoy.toUpperCase();

  // Rango inicial: dia 1 al ultimo dia con datos
  d1 = 1; d2 = m.diaActual||m.diasMes;
  document.getElementById('diaDesde').value = d1;
  document.getElementById('diaHasta').value = d2;
  document.getElementById('diaHasta').max   = m.diasMes;
  document.getElementById('diaDesde').max   = m.diasMes;
  document.getElementById('diasLabel').textContent = '('+d2+' dias)';

  buildKPICards();
  buildClientList();
  buildTable();
})();
</script>
</body>
</html>"""


if __name__ == "__main__":
    main()
