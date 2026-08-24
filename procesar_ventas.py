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

import os, sys, json, datetime as dt, warnings, calendar, base64
import pandas as pd
import numpy as np

# Lector de datos local (sin filtro TL en NACIONAL)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import leer_datos as _proc

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

    # ---- 2. LEER FUENTES DIRECTAMENTE ----
    print("> Leyendo archivos fuente directamente...")
    u = _proc.obtener_union(verbose=True)
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

    # Alias: codigo fuente -> codigo presupuesto
    COD_ALIAS = {
        "COLP": "KIMB",   # KIMB cambió de nombre a COLP
    }
    # Clientes que se convierten en operación propia (se sacan de NACIONAL)
    CLIENTES_OPERACION = {"JEMA"}
    # Clientes que siempre van a OTROS CLIENTES
    SIEMPRE_OTROS = {"CPA", "GDANE", "SOCO", "YUPI",
                     "ESEN_CR_ESPE", "CRESC", "LHCO", "MOIN", "ESEN_MB", "MECO"}
    u_nac["Cod"] = u_nac["Cod"].replace(COD_ALIAS)

    if "Dia" not in u_nac.columns:
        u_nac["Dia"] = pd.to_datetime(u_nac["Fecha"], errors="coerce").dt.day

    # Mes actual - datos diarios por (Cod, Subseg, Dia)
    m_act = u_nac[u_nac["Mes"] == mes_actual].copy()
    print(f"  {len(m_act):,} registros en {mes_actual} | dias: {sorted(m_act['Dia'].dropna().astype(int).unique())}")

    # Construir estructura diaria para JS: {Cod: {Subseg: {dia: [V, U, N]}}}
    # Excluir clientes que son su propia operacion
    m_act_tabla = m_act[~m_act["Cod"].isin(CLIENTES_OPERACION)].copy()
    daily_dict = {}
    for (cod, subseg, dia), grp in m_act_tabla.groupby(["Cod","Subseg","Dia"]):
        if cod not in daily_dict:
            daily_dict[cod] = {}
        if subseg not in daily_dict[cod]:
            daily_dict[cod][subseg] = {}
        daily_dict[cod][subseg][str(int(dia))] = [
            round(float(grp["AFacturar"].sum()), 0),
            round(float(grp["Utilidad"].sum()), 0),
            int(grp["Manifiesto"].nunique()),
        ]

    # ---- OPS_DIARIO: datos diarios por operacion para proyeccion en JS ----
    ops_diario = {}
    # Nacional puro (sin JEMA, sin SIEMPRE_OTROS — AJOV_MOV SI incluido)
    nac_puro = m_act[~m_act["Cod"].isin(CLIENTES_OPERACION | SIEMPRE_OTROS)]
    for dia, grp in nac_puro.groupby("Dia"):
        ops_diario.setdefault("NACIONAL", {})[str(int(dia))] = [
            round(float(grp["AFacturar"].sum()), 0),
            round(float(grp["Utilidad"].sum()), 0),
            int(grp["Manifiesto"].nunique())]
    # AJOV_MOV separado para que JS pueda restarlo cuando esté excluido
    ajov_df = m_act[m_act["Cod"] == "AJOV_MOV"]
    for dia, grp in ajov_df.groupby("Dia"):
        ops_diario.setdefault("AJOV_MOV", {})[str(int(dia))] = [
            round(float(grp["AFacturar"].sum()), 0),
            round(float(grp["Utilidad"].sum()), 0),
            int(grp["Manifiesto"].nunique())]
    # Clientes que son su propia operacion (JEMA, AJOV_MOV)
    for cliente in CLIENTES_OPERACION:
        cl_df = m_act[m_act["Cod"] == cliente]
        for dia, grp in cl_df.groupby("Dia"):
            ops_diario.setdefault(cliente, {})[str(int(dia))] = [
                round(float(grp["AFacturar"].sum()), 0),
                round(float(grp["Utilidad"].sum()), 0),
                int(grp["Manifiesto"].nunique())]
    # IMPO, EXPO, CEDIS, NAL-TL por dia
    for fuente in ["IMPO", "EXPO", "CEDIS", "NAL-TL"]:
        f_df = u_mes_all[u_mes_all["Fuente"] == fuente]
        for dia, grp in f_df.groupby("Dia"):
            ops_diario.setdefault(fuente, {})[str(int(dia))] = [
                round(float(grp["AFacturar"].sum()), 0),
                round(float(grp["Utilidad"].sum()), 0),
                int(grp["Manifiesto"].nunique())]

    # Actualizar OPS_KPI con operaciones individuales y COMEX
    for cliente in CLIENTES_OPERACION:
        cl_df = m_act[m_act["Cod"] == cliente]
        ops_kpi[cliente] = {
            "VENTA":    round(float(cl_df["AFacturar"].sum()), 0),
            "UTILIDAD": round(float(cl_df["Utilidad"].sum()), 0),
            "VIAJES":   int(cl_df["Manifiesto"].nunique()),
        }
    nal_tl_kpi = ops_kpi.get("NAL-TL", {})
    ops_kpi["COMEX"] = {
        "VENTA":    (ops_kpi.get("IMPO",{}).get("VENTA",0) + ops_kpi.get("EXPO",{}).get("VENTA",0) + nal_tl_kpi.get("VENTA",0)),
        "UTILIDAD": (ops_kpi.get("IMPO",{}).get("UTILIDAD",0) + ops_kpi.get("EXPO",{}).get("UTILIDAD",0) + nal_tl_kpi.get("UTILIDAD",0)),
        "VIAJES":   (ops_kpi.get("IMPO",{}).get("VIAJES",0) + ops_kpi.get("EXPO",{}).get("VIAJES",0) + nal_tl_kpi.get("VIAJES",0)),
    }
    # NACIONAL puro (sin JEMA, AJOV_MOV, SIEMPRE_OTROS)
    ops_kpi["NACIONAL"] = {
        "VENTA":    round(float(nac_puro["AFacturar"].sum()), 0),
        "UTILIDAD": round(float(nac_puro["Utilidad"].sum()), 0),
        "VIAJES":   int(nac_puro["Manifiesto"].nunique()),
    }

    # Venta de ayer y hoy: usar los ultimos 2 dias con datos disponibles
    dias_disponibles = sorted(m_act_tabla["Dia"].dropna().astype(int).unique())
    dia_hoy  = dias_disponibles[-1]  if len(dias_disponibles) >= 1 else None
    dia_ayer = dias_disponibles[-2]  if len(dias_disponibles) >= 2 else None
    label_hoy  = f"dia {dia_hoy}"  if dia_hoy  else "—"
    label_ayer = f"dia {dia_ayer}" if dia_ayer else "—"
    print(f"  Ultimo dato: dia {dia_hoy} | Penultimo: dia {dia_ayer}")

    fijo_hoy  = {}
    fijo_ayer = {}
    if dia_hoy:
        hoy_df = m_act_tabla[m_act_tabla["Dia"] == dia_hoy].groupby("Cod")["AFacturar"].sum()
        fijo_hoy = hoy_df.to_dict()
    if dia_ayer:
        ayer_df = m_act_tabla[m_act_tabla["Dia"] == dia_ayer].groupby("Cod")["AFacturar"].sum()
        fijo_ayer = ayer_df.to_dict()

    # Mes anterior (total por cliente, para la tabla resumen)
    m_ant = u_nac[u_nac["Mes"] == mes_anterior]
    mes_ant_agg = (m_ant.groupby("Cod")["AFacturar"]
                   .sum().rename("VENTA_MES_ANT").reset_index())

    # Mes anterior día a día (mismo formato que DIARIO, para filtrar por rango d1-d2 en JS)
    m_ant_tabla = m_ant[~m_ant["Cod"].isin(CLIENTES_OPERACION)].copy()
    if "Dia" not in m_ant_tabla.columns:
        m_ant_tabla["Dia"] = pd.to_datetime(m_ant_tabla["Fecha"], errors="coerce").dt.day
    if "Subseg" not in m_ant_tabla.columns:
        m_ant_tabla["Subseg"] = m_ant_tabla["Token"].apply(get_subseg)
    diario_ant_dict = {}
    for (cod, subseg, dia), grp in m_ant_tabla.groupby(["Cod", "Subseg", "Dia"]):
        diario_ant_dict.setdefault(str(cod), {}).setdefault(str(subseg), {})[str(int(dia))] = [
            round(float(grp["AFacturar"].sum()), 0),
            round(float(grp["Utilidad"].sum()), 0),
            int(grp["Manifiesto"].nunique()),
        ]

    # ---- 3. PENDIENTES POR PLANILLAR ----
    print(f"> Solicitudes: {os.path.basename(RUTA_SOLICITUDES)}")
    try:
        import shutil, tempfile
        tmp_sol = os.path.join(tempfile.gettempdir(), "solicitudes_tmp.xlsx")
        shutil.copy2(RUTA_SOLICITUDES, tmp_sol)
        hrow_sol = buscar_header_row_ob(tmp_sol, "Sheet1")
        sol = pd.read_excel(tmp_sol, sheet_name="Sheet1", header=hrow_sol)
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
    except Exception as _sol_err:
        print(f"  [aviso] No se pudo leer Solicitudes ({_sol_err}). Pendientes en $0.")
        pendiente = pd.DataFrame(columns=["Cod","P_PLANILLAR","N_PLANILLAR"])

    # ---- 4. TABLA BASE ----
    ej_all = (m_act_tabla.groupby("Cod")
              .agg(EJECUTADO=("AFacturar","sum"), UTILIDAD=("Utilidad","sum"),
                   VIAJES=("Manifiesto","nunique"))
              .reset_index())

    # Outer join: incluir clientes con venta aunque no esten en PPTO
    df = budget.merge(ej_all, on="Cod", how="outer")
    for right in [mes_ant_agg, pendiente]:
        df = df.merge(right, on="Cod", how="left")

    num_cols = ["PPTO","META_UTIL","M_VIAJES","COMPRA_PPTO","PCT_INTER_M",
                "EJECUTADO","UTILIDAD","VIAJES","VENTA_MES_ANT","P_PLANILLAR","N_PLANILLAR"]
    for c in num_cols:
        if c not in df.columns: df[c] = 0.0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # Agrupar en OTROS: sin venta en ambos meses, O cliente forzado a OTROS
    sin_venta = ((df["EJECUTADO"] == 0) & (df["VENTA_MES_ANT"] == 0)) | df["Cod"].isin(SIEMPRE_OTROS)
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

    # ---- 5. MANIFIESTOS A PERDIDA (NACIONAL mes actual) ----
    m_act["COMPRA"] = m_act["AFacturar"] - m_act["Utilidad"]
    man_grp = (m_act.groupby("Manifiesto")
               .agg(Cod=("Cod","first"),
                    Fecha=("Fecha","first"),
                    Dia=("Dia","first"),
                    Subseg=("Subseg","first"),
                    VENTA=("AFacturar","sum"),
                    COMPRA=("COMPRA","sum"),
                    OBs=("OB","nunique"))
               .reset_index())
    man_grp["UTIL"] = man_grp["VENTA"] - man_grp["COMPRA"]
    man_grp["MARGEN"] = (man_grp["UTIL"] / man_grp["VENTA"].replace(0, float("nan"))).fillna(0)
    perdidas_df = man_grp[man_grp["UTIL"] < 0].sort_values("UTIL")
    perdidas_js = [
        {"man": str(r["Manifiesto"]),
         "cod": str(r["Cod"]),
         "fecha": str(r["Fecha"])[:10] if pd.notna(r["Fecha"]) else "",
         "dia": int(r["Dia"]) if pd.notna(r["Dia"]) else 0,
         "subseg": str(r["Subseg"]) if pd.notna(r["Subseg"]) else "",
         "venta": round(float(r["VENTA"]), 0),
         "compra": round(float(r["COMPRA"]), 0),
         "util": round(float(r["UTIL"]), 0),
         "margen": round(float(r["MARGEN"]), 4),
         "obs": int(r["OBs"])}
        for _, r in perdidas_df.iterrows()
    ]
    print(f"  Manifiestos a perdida: {len(perdidas_js)}")

    # ---- HISTORICO: tendencias mes a mes de clientes NACIONAL ----
    hist_dict = {}
    u_hist = u_nac[~u_nac["Cod"].isin(CLIENTES_OPERACION)].copy()
    u_hist = u_hist[u_hist["Mes"].notna() & u_hist["Dia"].notna() & u_hist["Cod"].notna()]
    for (cod, mes, dia), grp in u_hist.groupby(["Cod", "Mes", "Dia"]):
        hist_dict.setdefault(str(cod), {}).setdefault(str(mes), {})[str(int(dia))] = [
            round(float(grp["AFacturar"].sum()), 0),
            int(grp["Manifiesto"].nunique()),
        ]
    print(f"  Histórico: {len(hist_dict)} clientes, {len(set(m for c in hist_dict.values() for m in c))} meses")

    # ---- OPS_HISTORICO: tendencias mensuales por operacion {op: {mes: {dia: [V,U,N]}}} ----
    ops_hist = {}
    def _fill_ops_hist(key, df):
        df2 = df[df["Mes"].notna() & df["Dia"].notna()]
        for (mes, dia), grp in df2.groupby(["Mes", "Dia"]):
            ops_hist.setdefault(key, {}).setdefault(str(mes), {})[str(int(dia))] = [
                round(float(grp["AFacturar"].sum()), 0),
                round(float(grp["Utilidad"].sum()), 0),
                int(grp["Manifiesto"].nunique()),
            ]
    _fill_ops_hist("NACIONAL", u_nac[~u_nac["Cod"].isin(CLIENTES_OPERACION)])
    _fill_ops_hist("JEMA",     u_nac[u_nac["Cod"] == "JEMA"])
    for fuente in ["IMPO", "EXPO", "NAL-TL", "CEDIS"]:
        _fill_ops_hist(fuente, u[u["Fuente"] == fuente])
    meses_ops = set(m for k in ops_hist for m in ops_hist[k])
    print(f"  OPS Histórico: {len(ops_hist)} ops, {len(meses_ops)} meses")

    # ---- AJOVER: clasificacion por tipo de operacion (AJOV + NOCO) ----
    def _norm_tip(tip):
        t = str(tip or "").strip().upper()
        if "PATINETA"   in t: return "PT"
        if "TRACTOMULA" in t: return "TM"
        if "SENCILLO"   in t: return "SC"
        if "TURBO"      in t: return "TB"
        return t

    def clasificar_ajov_row(row):
        cod = str(row.get("Cod", "") or "").strip().upper()
        ori = str(row.get("Origen", "") or "").strip().upper()
        des = str(row.get("Destino", "") or "").strip().upper()
        tip = _norm_tip(row.get("Tipologia", ""))
        age = str(row.get("Ciudad", "") or "").strip().upper()
        ruta = f"{ori}-{des}-{tip}"

        if cod == "NOCO":
            return "TRANSFERENCIAS" if ruta in ("CARTAGENA-MADRID-PT", "MADRID-CARTAGENA-PT") else "OTROS NOCO"

        # AJOV
        if ruta == "CARTAGENA-MADRID-PT":           return "Transferencia CTG - MAD"
        if ruta == "MADRID-CARTAGENA-PT":           return "Transferencia MAD - CTG"
        if ruta == "CARTAGENA-MADRID-TM":           return "Graneles"
        if ruta == "CARTAGENA-MADRID-SC":           return "Transferencia CTG - MAD"
        if ruta in ("CARTAGENA-YUMBO-PT",
                    "CARTAGENA-YUMBO-SC"):          return "Transferencia CTG - CALI"
        if ruta == "MADRID-CARTAGENA-SC":           return "SENCILLOS MADRID"
        if ruta == "MADRID-SANTIAGO DE CALI-PT":    return "ZORROS CALI"
        if ruta == "MADRID-BOGOTA-PT":              return "TRANSFERENCIA ALAMO"
        if ruta == "MADRID-BOGOTA-TB":              return "ENTREGA A CLIENTES"
        if "CALI" in age or age in ("CLO", "CAL"):  return "FIJOS CALI"
        if ori == "LA ESTRELLA":                    return "FIJOS MEDELLIN"
        return "ENTREGA A CLIENTES"

    ajov_raw = u_nac[u_nac["Cod"].isin(["AJOV", "NOCO"])].copy()
    if not ajov_raw.empty:
        ajov_raw["TipoOp"] = ajov_raw.apply(clasificar_ajov_row, axis=1)
        ajov_hist_dict = {}
        ajov_raw2 = ajov_raw[ajov_raw["Mes"].notna() & ajov_raw["Dia"].notna()]
        for (tipo, mes, dia), grp in ajov_raw2.groupby(["TipoOp", "Mes", "Dia"]):
            ajov_hist_dict.setdefault(str(tipo), {}).setdefault(str(mes), {})[str(int(dia))] = [
                round(float(grp["AFacturar"].sum()), 0),
                round(float(grp["Utilidad"].sum()), 0),
                int(grp["Manifiesto"].nunique()),
            ]
        tipos_ajov = sorted(ajov_hist_dict.keys())
        print(f"  AJOVER: {len(ajov_raw):,} registros → {len(tipos_ajov)} tipos: {tipos_ajov}")
    else:
        ajov_hist_dict = {}
        print("  AJOVER: sin datos AJOV/NOCO")

    # ---- 6. PREPARAR PAYLOAD PARA JS ----
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
        "PPTO":       float(otros["PPTO"].sum()),
        "META_UTIL":  float(otros["META_UTIL"].sum()),
        "M_VIAJES":   float(otros["M_VIAJES"].sum()),
        "P_PLANILLAR":float(otros["P_PLANILLAR"].sum()),
        "n":          int(len(otros)),
        "nombres":    sorted(otros["Cod"].tolist()),
        "detalle":    [
            {"cod": r["Cod"],
             "ej":  round(float(r.get("EJECUTADO", 0)), 0),
             "pp":  round(float(r.get("PPTO", 0)), 0)}
            for _, r in otros.iterrows()
        ],
    } if len(otros) else None

    # Lista de clientes activos (excluir CLIENTES_OPERACION)
    clientes_activos = sorted([c for c in df["Cod"].tolist() if c not in CLIENTES_OPERACION])

    payload = (
        f"window.DIARIO={json.dumps(daily_dict, ensure_ascii=False)};"
        f"window.DIARIO_ANT={json.dumps(diario_ant_dict, ensure_ascii=False)};"
        f"window.PPTO_DATA={json.dumps(ppto_js, ensure_ascii=False)};"
        f"window.MES_ANT={json.dumps(mes_ant_js, ensure_ascii=False)};"
        f"window.PENDIENTE={json.dumps(pend_js, ensure_ascii=False)};"
        f"window.FIJO_HOY={json.dumps({k:float(v) for k,v in fijo_hoy.items()}, ensure_ascii=False)};"
        f"window.FIJO_AYER={json.dumps({k:float(v) for k,v in fijo_ayer.items()}, ensure_ascii=False)};"
        f"window.OTROS={json.dumps(otros_js, ensure_ascii=False)};"
        f"window.CLIENTES={json.dumps(clientes_activos, ensure_ascii=False)};"
        f"window.OPS_KPI={json.dumps(ops_kpi, ensure_ascii=False)};"
        f"window.OPS_DIARIO={json.dumps(ops_diario, ensure_ascii=False)};"
        f"window.PERDIDAS={json.dumps(perdidas_js, ensure_ascii=False)};"
        f"window.HISTORICO={json.dumps(hist_dict, ensure_ascii=False)};"
        f"window.OPS_HISTORICO={json.dumps(ops_hist, ensure_ascii=False)};"
        f"window.AJOV_HIST={json.dumps(ajov_hist_dict, ensure_ascii=False)};"
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
.seg-btn-sub{background:#0e1e2e;border:1px solid #1e3a4e;color:#64748b;border-radius:5px;padding:4px 10px;font-size:.72rem;cursor:pointer;font-weight:600;transition:all .15s}
.seg-btn-sub.active{background:#3b82f6;color:#fff;border-color:#3b82f6}
.seg-btn-sub:hover:not(.active){background:#1a2d3e;color:#e2e8f0}
#comexSub{display:none;flex-wrap:wrap;gap:4px;margin-top:4px;padding-left:80px;width:100%}
.obs-input{width:100%;min-width:160px;background:#0d1a26;border:1px solid #1e3a4e;color:#f0c060;border-radius:4px;padding:4px 6px;font-size:.72rem;resize:vertical;min-height:32px;font-family:inherit}
.obs-input:focus{outline:none;border-color:#f59e0b;background:#111e2e}
.hcCard{background:#060f18;border:1px solid #1e3a4e;border-radius:8px;padding:10px 12px;overflow:hidden}
.hyr-btn{background:none;border:1px solid #1e3a4e;color:#64748b;font-size:.71rem;font-weight:600;padding:3px 10px;border-radius:5px;cursor:pointer;transition:all .15s}
.hyr-btn.active{background:#1a3a5c;border-color:#2d5a8e;color:#f97316}
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

/* TABLA OPS */
.ops-th{background:#0a1520;color:#64748b;font-weight:700;padding:6px 10px;text-align:right;border-bottom:1px solid #0e1e2e;white-space:nowrap;cursor:pointer;user-select:none;font-size:.7rem;letter-spacing:.04em}
.ops-th:hover{color:#e2e8f0;background:#0d1e2e}
.ops-th.sort-asc::after{content:' \25B2';color:#f97316}
.ops-th.sort-desc::after{content:' \25BC';color:#f97316}
.ops-td{padding:6px 10px;text-align:right;border-bottom:1px solid #080f18;white-space:nowrap}
.ops-tr:hover td{background:#0d1a26}
.ops-total td{background:#0c2030;border-top:1px solid #f97316;font-weight:700}

/* OTROS expandable */
tr.otros-row td:first-child{cursor:pointer}
tr.otros-row td:first-child::before{content:'\25B6  ';font-size:.7em;color:#f97316}
tr.otros-row.open td:first-child::before{content:'\25BC  ';font-size:.7em;color:#f97316}
tr.otros-detail{background:#080f18;font-size:.72rem}
tr.otros-detail td{color:#445566;padding:4px 8px 4px 28px}

/* TABS VISTA */
.view-tabs{display:flex;gap:0;padding:0 24px;background:#040a12;border-bottom:1px solid #0e1e2e}
.view-tab{background:none;border:none;border-bottom:2px solid transparent;color:#64748b;font-size:.78rem;font-weight:600;padding:10px 18px;cursor:pointer;transition:all .15s}
.view-tab.active{color:#f97316;border-bottom-color:#f97316}
.view-tab:hover:not(.active){color:#94a3b8}
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

<!-- KPI Cards + Tabla resumen -->
<div class="kpi-section">
  <div class="kpi-supertitle">&#9679; Resumen por operación — mes actual</div>
  <div class="kpi-grid" id="kpiGrid"></div>
  <div style="margin-top:18px">
    <div class="kpi-supertitle" style="margin-bottom:8px">&#9632; Proyección por operación (según filtro de días)</div>
    <div style="overflow-x:auto">
      <table id="tblOps" style="border-collapse:collapse;font-size:.75rem;width:100%;min-width:700px">
        <thead>
          <tr id="thrOps">
            <th class="ops-th" data-ok="label" style="text-align:left">OPERACIÓN</th>
            <th class="ops-th" data-ok="ej">EJECUTADO</th>
            <th class="ops-th" data-ok="proy">PROYECCIÓN</th>
            <th class="ops-th" data-ok="util">UTILIDAD</th>
            <th class="ops-th" data-ok="margen">MARGEN</th>
            <th class="ops-th" data-ok="viajes">VIAJES</th>
          </tr>
        </thead>
        <tbody id="tbodyOps"></tbody>
      </table>
    </div>
  </div>
</div>

<div class="meta-bar" id="metaBar"></div>

<div class="filter-bar">
  <!-- Operacion -->
  <div class="filter-group" style="flex-wrap:wrap;gap:6px">
    <span class="filter-label">Operación</span>
    <button class="seg-btn active" data-op="TODOS"    onclick="setOp(this)">Todos</button>
    <button class="seg-btn"        data-op="NACIONAL" onclick="setOp(this)">Nacional</button>
    <button class="seg-btn"        data-op="JEMA"     onclick="setOp(this)">Jerónimo</button>
    <button class="seg-btn"        data-op="CEDIS"    onclick="setOp(this)">CEDIS</button>
    <button class="seg-btn"        data-op="COMEX"    onclick="setOp(this)">COMEX &#9660;</button>
    <!-- Sub-panel COMEX -->
    <div id="comexSub" style="display:none;width:100%;display:none;gap:4px;margin-top:4px;padding-left:80px">
      <button class="seg-btn-sub active" data-subop="ALL"    onclick="setSubOp(this)">Todo COMEX</button>
      <button class="seg-btn-sub"        data-subop="IMPO"   onclick="setSubOp(this)">IMPO</button>
      <button class="seg-btn-sub"        data-subop="EXPO"   onclick="setSubOp(this)">EXPO</button>
      <button class="seg-btn-sub"        data-subop="NAL-TL" onclick="setSubOp(this)">Nal-TL</button>
    </div>
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

<!-- Tabs vista -->
<div class="view-tabs">
  <button class="view-tab active" onclick="setView('tabla',this)">&#9776; Tabla Nacional</button>
  <button class="view-tab" onclick="setView('perdidas',this)">&#9888; Manifiestos a Pérdida</button>
  <button class="view-tab" onclick="setView('historico',this)">&#9196; Histórico</button>
  <button class="view-tab" onclick="setView('ajover',this)">&#9672; AJOVER</button>
</div>

<div id="viewTabla">
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
        <th data-k="VENTA_MES_ANT">M. ANT. (mismo rango)</th>
        <th data-k="DIF_DIAS">EJ. vs M. ANT.</th>
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

</div><!-- /container -->
</div><!-- /viewTabla -->

<!-- Vista Pérdidas -->
<div id="viewPerdidas" style="display:none;padding:14px 10px;overflow-x:auto">
  <div style="display:flex;align-items:center;gap:16px;margin-bottom:10px;flex-wrap:wrap">
    <span style="color:#94a3b8;font-size:.78rem">Manifiestos donde <b style="color:#ef4444">compra &gt; venta</b></span>
    <div style="display:flex;align-items:center;gap:6px;margin-left:auto">
      <span style="color:#64748b;font-size:.75rem">Tu nombre:</span>
      <input id="inputUserName" type="text" placeholder="Ej: Juan Pérez"
        style="background:#0d1a26;border:1px solid #1e3a4e;color:#f0c060;border-radius:4px;padding:4px 8px;font-size:.75rem;width:150px"
        oninput="localStorage.setItem('tc_userName',this.value)">
    </div>
  </div>
  <table style="width:100%;border-collapse:collapse;min-width:800px;font-size:.77rem">
    <thead>
      <tr id="thrPerd" style="background:#1a0d0d">
        <th class="perd-th" data-pk="man" style="text-align:left;padding:7px 8px;color:#ef4444;border-bottom:2px solid #3a1010;cursor:pointer;user-select:none">MANIFIESTO</th>
        <th class="perd-th" data-pk="cod" style="padding:7px 8px;color:#94a3b8;border-bottom:2px solid #3a1010;cursor:pointer;user-select:none">CLIENTE</th>
        <th class="perd-th" data-pk="fecha" style="padding:7px 8px;color:#94a3b8;border-bottom:2px solid #3a1010;cursor:pointer;user-select:none">FECHA</th>
        <th style="padding:7px 8px;color:#94a3b8;border-bottom:2px solid #3a1010">OBs</th>
        <th class="perd-th" data-pk="venta" style="padding:7px 8px;color:#94a3b8;border-bottom:2px solid #3a1010;cursor:pointer;user-select:none">VENTA</th>
        <th class="perd-th" data-pk="compra" style="padding:7px 8px;color:#94a3b8;border-bottom:2px solid #3a1010;cursor:pointer;user-select:none">COMPRA</th>
        <th class="perd-th" data-pk="util" style="padding:7px 8px;color:#ef4444;border-bottom:2px solid #3a1010;cursor:pointer;user-select:none">PÉRDIDA</th>
        <th class="perd-th" data-pk="margen" style="padding:7px 8px;color:#94a3b8;border-bottom:2px solid #3a1010;cursor:pointer;user-select:none">MARGEN</th>
        <th style="padding:7px 8px;color:#f59e0b;border-bottom:2px solid #3a1010;min-width:160px">OBSERVACIÓN</th>
        <th style="padding:7px 8px;color:#7dd3fc;border-bottom:2px solid #3a1010;min-width:120px">RESPONSABLE</th>
      </tr>
    </thead>
    <tbody id="tbodyPerdidas"></tbody>
  </table>
</div>

<div id="viewHistorico" style="display:none;padding:14px 10px">
  <!-- Controles superiores -->
  <div style="display:flex;align-items:center;gap:10px;margin-bottom:14px;flex-wrap:wrap">
    <span style="color:#94a3b8;font-size:.78rem;font-weight:600">Histórico</span>
    <!-- Filtro año -->
    <div id="histYearBtns" style="display:flex;gap:4px"></div>
    <!-- Modo venta/viajes -->
    <div style="display:flex;gap:0;margin-left:auto;border:1px solid #1e3a4e;border-radius:6px;overflow:hidden">
      <button class="hist-mode-btn" id="histBtnVenta" onclick="setHistMode('venta',this)" style="background:#1a3a5c;border:none;color:#f97316;font-size:.73rem;font-weight:700;padding:5px 14px;cursor:pointer">$ Venta</button>
      <button class="hist-mode-btn" id="histBtnViajes" onclick="setHistMode('viajes',this)" style="background:none;border:none;color:#64748b;font-size:.73rem;font-weight:600;padding:5px 14px;cursor:pointer">&#9201; Viajes</button>
    </div>
  </div>
  <!-- Gráficos de operaciones -->
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:18px">
    <div class="hcCard" id="hcNac"></div>
    <div class="hcCard" id="hcJema"></div>
    <div class="hcCard" id="hcComex"></div>
    <div class="hcCard" id="hcCedis"></div>
  </div>
  <!-- Análisis Inteligente -->
  <div id="histInsights"></div>
  <!-- Tabla clientes -->
  <div style="overflow-x:auto">
    <table id="tblHistorico" style="border-collapse:collapse;font-size:.76rem;min-width:600px;width:100%">
      <thead id="theadHistorico"></thead>
      <tbody id="tbodyHistorico"></tbody>
    </table>
  </div>
</div>

<!-- ===== AJOVER ===== -->
<div id="viewAjover" style="display:none;padding:14px 10px">
  <div style="display:flex;align-items:center;gap:14px;margin-bottom:14px;flex-wrap:wrap">
    <span style="color:#60a5fa;font-size:.85rem;font-weight:700;letter-spacing:.05em">AJOVER · Detalle por Tipo de Operación</span>
    <span style="color:#334155;font-size:.72rem">(mismo rango días filtrado en la tabla principal)</span>
    <div id="ajovKpi" style="display:flex;gap:20px;margin-left:auto;flex-wrap:wrap"></div>
  </div>
  <div style="overflow-x:auto">
    <table style="border-collapse:collapse;font-size:.76rem;min-width:700px;width:100%">
      <thead id="theadAjover"></thead>
      <tbody id="tbodyAjover"></tbody>
    </table>
  </div>
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

  // Acumular segun segmento y rango de dias — mes actual
  var V=0, U=0, N=0;
  var segs = curSeg==='TODOS' ? Object.keys(dias) : (dias[curSeg] ? [curSeg] : []);
  segs.forEach(function(s){
    var sd = dias[s] || {};
    for(var d=d1; d<=d2; d++){
      var e = sd[String(d)];
      if(e){V+=e[0]; U+=e[1]; N+=e[2];}
    }
  });

  // Mes anterior — mismo rango de dias d1-d2
  var maRng=0;
  var diasAnt=window.DIARIO_ANT[cod]||{};
  var segsAnt=curSeg==='TODOS'?Object.keys(diasAnt):(diasAnt[curSeg]?[curSeg]:[]);
  segsAnt.forEach(function(s){
    var sd=diasAnt[s]||{};
    for(var d=d1;d<=d2;d++){var e=sd[String(d)];if(e)maRng+=e[0];}
  });

  var diasRango = d2 - d1 + 1;
  var diasMes   = m.diasMes || 31;
  var PROY  = diasRango > 0 ? V / diasRango * diasMes : 0;
  var DIF_PP = PROY - (pp.PPTO||0);
  var PCT_C  = pp.PPTO > 0 ? PROY / pp.PPTO : 0;
  var DIF_D  = V - maRng;
  var MAR    = V > 0 ? U / V : 0;
  var PROY_U = PROY * MAR;
  var META_V = Math.max((pp.PPTO||0) - V, 0);

  return {
    Cod: cod,
    PPTO: pp.PPTO||0, META_UTIL: pp.META_UTIL||0, M_VIAJES: pp.M_VIAJES||0,
    PCT_INTER_M: pp.PCT_INTER_M||0,
    EJECUTADO: V, UTILIDAD: U, VIAJES: N,
    VENTA_MES_ANT: maRng,
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
  // Operaciones sin tabla de clientes → mostrar mensaje
  var OPS_SIN_TABLA = {'CEDIS':1,'COMEX':1,'JEMA':1};
  var tbody = document.getElementById('tbody');
  if(OPS_SIN_TABLA[curOp]){
    tbody.innerHTML='<tr><td colspan="19" style="text-align:center;padding:30px;color:#445566;font-size:.8rem">'+
      '&#8593; Ver tarjeta de operación arriba para los datos de <b style="color:#94a3b8">'+curOp+'</b></td></tr>';
    document.getElementById('metaBar').innerHTML='';
    return;
  }
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

  // Fila OTROS CLIENTES (expandable)
  var otros = window.OTROS;
  if(otros){
    var or = {
      Cod: '<span onclick="toggleOtros()" style="cursor:pointer;user-select:none" title="Ver clientes"><span id="otrosArrow" style="margin-right:4px">&#9654;</span>OTROS CLIENTES ('+otros.n+')</span>',
      PPTO: otros.PPTO, META_UTIL: otros.META_UTIL, M_VIAJES: otros.M_VIAJES,
      EJECUTADO:0, UTILIDAD:0, VIAJES:0, VENTA_MES_ANT:0,
      VENTA_AYER:0, VENTA_HOY:0, PROYECCION:0,
      DIF_PROV_PPTO: -otros.PPTO, PCT_CUMPL:0, DIF_DIAS:0,
      PROY_UTILIDAD:0, PCT_INTER:0, PCT_INTER_M:0,
      META_VENTA_FINAL: otros.PPTO, P_PLANILLAR: otros.P_PLANILLAR||0,
    };
    var orTr=renderRow(or,'otros-row');
    tbody.appendChild(orTr);
    // Filas detalle (ocultas inicialmente)
    var det=otros.detalle||[];
    det.sort(function(a,b){return b.ej-a.ej;});
    det.forEach(function(c){
      var dtr=document.createElement('tr');
      dtr.className='otros-detail'; dtr.style.display='none';
      var ef=formatBig(c.ej), pf=formatBig(c.pp||0);
      dtr.innerHTML=
        '<td style="padding:4px 8px 4px 32px;text-align:left;color:#7aa8cc">&#9492; '+c.cod+'</td>'+
        '<td style="padding:4px 8px;color:#445566">'+pf.val+pf.unit+'</td>'+
        '<td colspan="17" style="padding:4px 8px;color:#56789a">'+ef.val+ef.unit+'</td>';
      tbody.appendChild(dtr);
    });
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
    DIF_DIAS:totV-totMA,
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

function isPerdidasActive(){
  return document.getElementById('viewPerdidas').style.display !== 'none';
}
function isHistoricoActive(){
  return document.getElementById('viewHistorico').style.display !== 'none';
}

/* Mapa operacion → curSeg para la tabla Nacional */
var OP_SEG_MAP = {
  'TODOS':'TODOS','NACIONAL':'NAC','JEMA':'JEMA',
  'CEDIS':'__OP__','COMEX':'__OP__'
};
var curOp='TODOS', curSubOp='ALL';

function isAjoverActive(){
  return document.getElementById('viewAjover').style.display !== 'none';
}
function rebuildAll(){
  buildKPICards();
  buildOpsTable();
  buildTable();
  if(isPerdidasActive()) buildPerdidas();
  if(isHistoricoActive()) buildHistorico();
  if(isAjoverActive()) buildAjover();
}

function setOp(btn){
  document.querySelectorAll('.seg-btn').forEach(function(b){b.classList.remove('active');});
  btn.classList.add('active');
  curOp = btn.getAttribute('data-op');
  curSeg = OP_SEG_MAP[curOp] || 'TODOS';
  // Mostrar sub-panel COMEX
  var sub = document.getElementById('comexSub');
  if(curOp==='COMEX'){
    sub.style.display='flex';
  } else {
    sub.style.display='none';
    curSubOp='ALL';
  }
  rebuildAll();
}

function setSubOp(btn){
  document.querySelectorAll('.seg-btn-sub').forEach(function(b){b.classList.remove('active');});
  btn.classList.add('active');
  curSubOp = btn.getAttribute('data-subop');
  rebuildAll();
}

function recalc(){
  var m = window.META||{};
  d1 = Math.max(1, parseInt(document.getElementById('diaDesde').value)||1);
  d2 = Math.min(m.diasMes||31, Math.max(d1, parseInt(document.getElementById('diaHasta').value)||d1));
  document.getElementById('diaHasta').value = d2;
  document.getElementById('diasLabel').textContent = '('+( d2-d1+1)+' dias)';
  rebuildAll();
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
      rebuildAll();
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
  rebuildAll();
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

/* ---- Config operaciones ---- */
var OPS_CFG = [
  {key:'NACIONAL', cls:'kpi-nac', label:'Nacional',    sub:null,      clientKey:null},
  {key:'JEMA',     cls:'kpi-ced', label:'Jeronimo M.', sub:null,      clientKey:'JEMA'},
  {key:'COMEX',    cls:'kpi-exp', label:'COMEX',       clientKey:null, sub:[
    {key:'IMPO',label:'IMPO'},{key:'EXPO',label:'EXPO'},{key:'NAL-TL',label:'Nal-TL'}
  ]},
  {key:'CEDIS',    cls:'kpi-ced', label:'CEDIS',       sub:null,      clientKey:null},
];

/* ---- KPI Cards ---- */
function calcNacKPIDynamic(){
  /* Suma el mes completo de todos los clientes NACIONAL no excluidos (sin SIEMPRE_OTROS) */
  var V=0,U=0,N=0;
  var clts=window.CLIENTES||[];
  clts.forEach(function(c){
    if(excluidos.has(c)) return;
    var cdata=window.DIARIO[c]||{};
    var subs=Object.keys(cdata);
    for(var si=0;si<subs.length;si++){
      var dias=cdata[subs[si]];
      var dks=Object.keys(dias);
      for(var di=0;di<dks.length;di++){
        var e=dias[dks[di]];
        if(e){V+=e[0];U+=e[1];N+=e[2];}
      }
    }
  });
  return {VENTA:V,UTILIDAD:U,VIAJES:N};
}

function buildKPICards(){
  var ops=window.OPS_KPI||{};
  var grid=document.getElementById('kpiGrid');
  grid.innerHTML='';
  OPS_CFG.forEach(function(op){
    if(op.clientKey && excluidos.has(op.clientKey)) return;
    var d=(op.key==='NACIONAL')?calcNacKPIDynamic():(ops[op.key]||{VENTA:0,UTILIDAD:0,VIAJES:0});
    var margin=d.VENTA>0?((d.UTILIDAD/d.VENTA)*100).toFixed(1)+'% margen':'— margen';
    var vf=formatBig(d.VENTA);
    var uf=formatBig(d.UTILIDAD);
    var subHtml='';
    if(op.sub){
      subHtml='<div style="display:flex;gap:10px;margin-top:6px;flex-wrap:wrap">';
      op.sub.forEach(function(s){
        var sd=ops[s.key]||{VENTA:0};
        var sf=formatBig(sd.VENTA);
        subHtml+='<span style="font-size:.65rem;opacity:.6">'+s.label+': '+sf.val+sf.unit+'</span>';
      });
      subHtml+='</div>';
    }
    var ps=calcPerdidasStats(op.key);
    var perdHtml='';
    if(ps&&ps.total>0){
      perdHtml='<div style="margin-top:8px;padding:5px 8px;background:rgba(239,68,68,.1);border-radius:5px;font-size:.68rem">'+
        '<span style="color:#ef4444;font-weight:700">&#9888; '+ps.total+' pérd.</span>'+
        ' <span style="color:#4ade80">&#10003; '+ps.conC+' comentados</span>'+
        ' <span style="color:#ef4444;opacity:.7">&#9888; '+ps.sinC+' sin comentario</span>'+
      '</div>';
    }
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
      '</div>'+subHtml+perdHtml;
    grid.appendChild(card);
  });
}

/* ---- Calculo proyeccion por operacion ---- */
function calcOpData(opKey){
  var keys=opKey==='COMEX'?['IMPO','EXPO','NAL-TL']:[opKey];
  var V=0,U=0,N=0;
  keys.forEach(function(k){
    var daily=window.OPS_DIARIO[k]||{};
    for(var d=d1;d<=d2;d++){var e=daily[String(d)];if(e){V+=e[0];U+=e[1];N+=e[2];}}
  });
  // Si AJOV_MOV excluido, restar su contribucion de NACIONAL
  if(opKey==='NACIONAL' && excluidos.has('AJOV_MOV')){
    var ajovD=window.OPS_DIARIO['AJOV_MOV']||{};
    for(var d=d1;d<=d2;d++){var e=ajovD[String(d)];if(e){V-=e[0];U-=e[1];N-=e[2];}}
  }
  var dias=d2-d1+1;
  var proy=dias>0?V/dias*(window.META.diasMes||31):0;
  return {V:V,U:U,N:N,proy:proy};
}

/* ---- Tabla resumen operaciones ---- */
var opsSortK='ej', opsSortAsc=false;
function buildOpsTable(){
  var m=window.META||{};
  // Si hay sub-op de COMEX activa, calcOpData especial
  function calcOpFiltered(op){
    if(op.key==='COMEX' && curOp==='COMEX' && curSubOp!=='ALL'){
      var keys=[curSubOp];
      var V=0,U=0,N=0;
      keys.forEach(function(k){
        var daily=window.OPS_DIARIO[k]||{};
        for(var d=d1;d<=d2;d++){var e=daily[String(d)];if(e){V+=e[0];U+=e[1];N+=e[2];}}
      });
      var dias=d2-d1+1;
      var proy=dias>0?V/dias*(window.META.diasMes||31):0;
      return {V:V,U:U,N:N,proy:proy};
    }
    return calcOpData(op.key);
  }
  var rows=OPS_CFG.filter(function(op){
    if(op.clientKey && excluidos.has(op.clientKey)) return false;
    if(curOp==='TODOS') return true;
    if(curOp==='COMEX') return op.key==='COMEX';
    return op.key===curOp;
  }).map(function(op){
    var r=calcOpFiltered(op);
    return {label:op.label,key:op.key,ej:r.V,util:r.U,viajes:r.N,proy:r.proy,
            margen:r.V>0?r.U/r.V:0};
  });
  rows.sort(function(a,b){
    var va=a[opsSortK],vb=b[opsSortK];
    if(typeof va==='string') return opsSortAsc?va.localeCompare(vb):vb.localeCompare(va);
    return opsSortAsc?va-vb:vb-va;
  });
  var tbody=document.getElementById('tbodyOps');
  tbody.innerHTML='';
  var totEj=0,totProy=0,totUtil=0,totV=0;
  rows.forEach(function(r){
    totEj+=r.ej;totProy+=r.proy;totUtil+=r.util;totV+=r.ej;
    var tr=document.createElement('tr');
    tr.className='ops-tr';
    tr.innerHTML=
      '<td class="ops-td" style="text-align:left;font-weight:600;color:#e2e8f0">'+r.label+'</td>'+
      '<td class="ops-td">'+COP.format(r.ej)+'</td>'+
      '<td class="ops-td" style="color:#f97316">'+COP.format(r.proy)+'</td>'+
      '<td class="ops-td" style="color:#4ade80">'+COP.format(r.util)+'</td>'+
      '<td class="ops-td">'+(r.margen*100).toFixed(1)+'%</td>'+
      '<td class="ops-td">'+NUM.format(r.viajes)+'</td>';
    tbody.appendChild(tr);
  });
  var tf=document.createElement('tr');
  tf.className='ops-total';
  var totM=totEj>0?totUtil/totEj:0;
  tf.innerHTML=
    '<td class="ops-td" style="text-align:left;color:#f97316">TOTAL</td>'+
    '<td class="ops-td" style="color:#f97316">'+COP.format(totEj)+'</td>'+
    '<td class="ops-td" style="color:#f97316">'+COP.format(totProy)+'</td>'+
    '<td class="ops-td" style="color:#4ade80">'+COP.format(totUtil)+'</td>'+
    '<td class="ops-td">'+(totM*100).toFixed(1)+'%</td>'+
    '<td class="ops-td"></td>';
  tbody.appendChild(tf);
}

// Ordenar tabla ops
document.addEventListener('DOMContentLoaded',function(){
  document.querySelectorAll('.ops-th').forEach(function(th){
    th.addEventListener('click',function(){
      var k=th.getAttribute('data-ok');
      document.querySelectorAll('.ops-th').forEach(function(t){t.classList.remove('sort-asc','sort-desc');});
      if(opsSortK===k){opsSortAsc=!opsSortAsc;}else{opsSortK=k;opsSortAsc=false;}
      th.classList.add(opsSortAsc?'sort-asc':'sort-desc');
      buildOpsTable();
    });
  });
  // Ordenar tabla perdidas
  document.querySelectorAll('.perd-th').forEach(function(th){
    th.addEventListener('click',function(){
      var k=th.getAttribute('data-pk');
      document.querySelectorAll('.perd-th').forEach(function(t){t.style.color='';});
      if(perdSortK===k){perdSortAsc=!perdSortAsc;}else{perdSortK=k;perdSortAsc=false;}
      th.style.color='#f97316';
      buildPerdidas();
    });
  });
});

/* ---- Cambio de vista ---- */
function setView(v, btn){
  document.querySelectorAll('.view-tab').forEach(function(b){b.classList.remove('active');});
  btn.classList.add('active');
  document.getElementById('viewTabla').style.display     = v==='tabla'     ? '' : 'none';
  document.getElementById('viewPerdidas').style.display  = v==='perdidas'  ? '' : 'none';
  document.getElementById('viewHistorico').style.display = v==='historico' ? '' : 'none';
  document.getElementById('viewAjover').style.display    = v==='ajover'    ? '' : 'none';
  if(v==='perdidas'){
    var inp=document.getElementById('inputUserName');
    if(inp) inp.value=localStorage.getItem('tc_userName')||'';
    buildPerdidas();
  }
  if(v==='historico') buildHistorico();
  if(v==='ajover') buildAjover();
}

/* ---- Toggle OTROS detalle ---- */
var otrosOpen=false;
function toggleOtros(){
  otrosOpen=!otrosOpen;
  var arr=document.getElementById('otrosArrow');
  if(arr) arr.innerHTML=otrosOpen?'&#9660;':'&#9654;';
  document.querySelectorAll('tr.otros-detail').forEach(function(tr){
    tr.style.display=otrosOpen?'':'none';
  });
}

/* ---- Perdidas sort + observaciones ---- */
var perdSortK='util', perdSortAsc=true;
var perdidasObs = JSON.parse(localStorage.getItem('tc_perdidasObs')||'{}');
// Migrar formato antiguo (string → objeto)
Object.keys(perdidasObs).forEach(function(k){
  if(typeof perdidasObs[k]==='string') perdidasObs[k]={text:perdidasObs[k],user:'',ts:''};
});
function saveObs(man,field,val){
  if(!perdidasObs[man]) perdidasObs[man]={text:'',user:'',ts:''};
  perdidasObs[man][field]=val;
  perdidasObs[man].ts=new Date().toLocaleString('es-CO');
  localStorage.setItem('tc_perdidasObs',JSON.stringify(perdidasObs));
}
function perdHasComment(man){return !!(perdidasObs[man]&&perdidasObs[man].text&&perdidasObs[man].text.trim());}
function calcPerdidasStats(opKey){
  var perd=window.PERDIDAS||[];
  var CLIENTES_OP_S={'JEMA':1,'AJOV_MOV':1};
  var filtered=perd.filter(function(r){
    if(excluidos.has(r.cod)) return false;
    if(opKey==='JEMA') return r.cod==='JEMA';
    if(opKey==='NACIONAL'){
      if(CLIENTES_OP_S[r.cod]) return false;
      return !r.subseg||r.subseg.indexOf('CED')<0&&r.subseg.indexOf('TL')<0;
    }
    return false;
  });
  var conC=filtered.filter(function(r){return perdHasComment(r.man);}).length;
  return {total:filtered.length,conC:conC,sinC:filtered.length-conC};
}

function buildPerdidas(){
  var raw=(window.PERDIDAS||[]);
  var CLIENTES_OP={'JEMA':1,'AJOV_MOV':1};
  var data=raw.filter(function(r){
    if(r.dia && (r.dia<d1||r.dia>d2)) return false;
    if(excluidos.has(r.cod)) return false;
    // Filtro por operacion
    if(curOp==='JEMA'||curOp==='AJOV_MOV') return r.cod===curOp;
    if(curOp==='NACIONAL'){
      // Excluir clientes de operacion y cuentas CED/TL
      if(CLIENTES_OP[r.cod]) return false;
      return !r.subseg||r.subseg.indexOf('CED')<0&&r.subseg.indexOf('TL')<0;
    }
    if(curOp==='CEDIS'||curOp==='COMEX') return false;
    // TODOS: excluir solo los clientes que no están en ninguna op visible
    return true;
  }).slice();

  var tbody=document.getElementById('tbodyPerdidas');
  tbody.innerHTML='';
  if(!data.length){
    tbody.innerHTML='<tr><td colspan="9" style="text-align:center;padding:20px;color:#445566">Sin manifiestos a pérdida con los filtros actuales</td></tr>';
    return;
  }
  data.sort(function(a,b){
    var va=a[perdSortK],vb=b[perdSortK];
    if(typeof va==='string') return perdSortAsc?va.localeCompare(vb):vb.localeCompare(va);
    return perdSortAsc?va-vb:vb-va;
  });
  var totalPerd=0, totalVenta=0, totalCompra=0;
  data.forEach(function(r){
    totalPerd+=r.util; totalVenta+=r.venta; totalCompra+=r.compra;
    var tr=document.createElement('tr');
    tr.style.borderBottom='1px solid #1a1010';
    var obsVal=perdidasObs[r.man]||'';
    var saved=perdidasObs[r.man]||{text:'',user:'',ts:''};
    var tdObs=document.createElement('td');
    tdObs.style.cssText='padding:4px 6px;vertical-align:top';
    var ta=document.createElement('textarea');
    ta.className='obs-input'; ta.rows=1; ta.value=saved.text||'';
    ta.placeholder='Observación...';
    ta.addEventListener('input',function(){saveObs(r.man,'text',ta.value);});
    tdObs.appendChild(ta);
    if(saved.ts) tdObs.appendChild(Object.assign(document.createElement('div'),
      {style:'font-size:.62rem;color:#334155;margin-top:2px',textContent:saved.ts}));
    var tdUser=document.createElement('td');
    tdUser.style.cssText='padding:4px 6px;vertical-align:top';
    var inp=document.createElement('input');
    inp.type='text'; inp.className='obs-input'; inp.value=saved.user||'';
    inp.placeholder='Nombre...';
    inp.addEventListener('input',function(){saveObs(r.man,'user',inp.value);});
    tdUser.appendChild(inp);
    tr.innerHTML=
      '<td style="text-align:left;padding:6px 8px;font-weight:600;color:#fca5a5">'+r.man+'</td>'+
      '<td style="padding:6px 8px;color:#94a3b8">'+r.cod+'</td>'+
      '<td style="padding:6px 8px;color:#64748b">'+r.fecha+'</td>'+
      '<td style="padding:6px 8px;color:#64748b">'+r.obs+'</td>'+
      '<td style="padding:6px 8px;color:#e2e8f0">'+COP.format(r.venta)+'</td>'+
      '<td style="padding:6px 8px;color:#fbbf24">'+COP.format(r.compra)+'</td>'+
      '<td style="padding:6px 8px;color:#ef4444;font-weight:700">'+COP.format(r.util)+'</td>'+
      '<td style="padding:6px 8px;color:#ef4444">'+(r.margen*100).toFixed(1)+'%</td>';
    tr.appendChild(tdObs);
    tr.appendChild(tdUser);
    tbody.appendChild(tr);
  });
  var tf=document.createElement('tr');
  tf.style.cssText='background:#1a0505;border-top:2px solid #ef4444;font-weight:700';
  tf.innerHTML=
    '<td colspan="4" style="text-align:left;padding:6px 8px;color:#ef4444">TOTAL ('+data.length+' manifiestos)</td>'+
    '<td style="padding:6px 8px;color:#e2e8f0">'+COP.format(totalVenta)+'</td>'+
    '<td style="padding:6px 8px;color:#fbbf24">'+COP.format(totalCompra)+'</td>'+
    '<td style="padding:6px 8px;color:#ef4444">'+COP.format(totalPerd)+'</td>'+
    '<td style="padding:6px 8px;color:#ef4444">'+(totalVenta>0?(totalPerd/totalVenta*100).toFixed(1):'0')+'%</td>'+
    '<td colspan="2"></td>';
  tbody.appendChild(tf);
}

/* ---- HISTORICO ---- */
var histMode='venta';
var histYear='2026';
var histSortIdx=-1, histSortAsc=false;

function setHistMode(mode, btn){
  histMode=mode;
  document.querySelectorAll('.hist-mode-btn').forEach(function(b){
    b.style.background='none'; b.style.color='#64748b';
  });
  btn.style.background='#1a3a5c'; btn.style.color='#f97316';
  buildHistorico();
}

function setHistYear(yr, btn){
  histYear=yr;
  document.querySelectorAll('.hyr-btn').forEach(function(b){b.classList.remove('active');});
  btn.classList.add('active');
  histSortIdx=-1;
  buildHistorico();
}

function sparkline(vals){
  if(!vals||vals.length<2) return '';
  var max=Math.max.apply(null,vals), min=Math.min.apply(null,vals);
  var w=60,h=20,pad=2;
  if(max===min) return '<svg width="'+w+'" height="'+h+'"><line x1="2" y1="10" x2="58" y2="10" stroke="#4ade80" stroke-width="1.5"/></svg>';
  var pts=vals.map(function(v,i){
    var x=pad+(i/(vals.length-1))*(w-2*pad);
    var y=(h-pad)-((v-min)/(max-min))*(h-2*pad);
    return x.toFixed(1)+','+y.toFixed(1);
  });
  var last=vals[vals.length-1], prev=vals[vals.length-2];
  var col=last>=prev?'#4ade80':'#ef4444';
  var lp=pts[pts.length-1].split(',');
  return '<svg width="'+w+'" height="'+h+'" style="display:block;overflow:visible">'+
    '<polyline points="'+pts.join(' ')+'" fill="none" stroke="'+col+'" stroke-width="1.5" stroke-linejoin="round" stroke-linecap="round"/>'+
    '<circle cx="'+lp[0]+'" cy="'+lp[1]+'" r="2.5" fill="'+col+'"/>'+
    '</svg>';
}

function mesLabel(m){
  var mNames=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  var y,mo;
  if(m.indexOf('-')>0){var p=m.split('-');y=p[0];mo=parseInt(p[1])-1;}
  else{y=m.slice(0,4);mo=parseInt(m.slice(4))-1;}
  return mNames[mo]+' '+y.slice(2);
}

function mesYear(m){
  return m.indexOf('-')>0?m.split('-')[0]:m.slice(0,4);
}

function buildHistorico(){
  var raw=window.HISTORICO||{};
  // Obtener todos los meses disponibles
  var mesesSet={};
  Object.keys(raw).forEach(function(cod){
    Object.keys(raw[cod]).forEach(function(m){ mesesSet[m]=1; });
  });
  var allMeses=Object.keys(mesesSet).sort();

  // Botones de año (dinámicos)
  var yearContainer=document.getElementById('histYearBtns');
  if(yearContainer && !yearContainer.dataset.built){
    var yearsSet={};
    allMeses.forEach(function(m){ yearsSet[mesYear(m)]=1; });
    var years=Object.keys(yearsSet).sort();
    var yHtml='<button class="hyr-btn'+(histYear===''?' active':'')+'" onclick="setHistYear(\'\',this)">Todos</button>';
    years.forEach(function(y){
      yHtml+='<button class="hyr-btn'+(histYear===y?' active':'')+'" onclick="setHistYear(\''+y+'\',this)">'+y+'</button>';
    });
    yearContainer.innerHTML=yHtml;
    yearContainer.dataset.built='1';
  }

  // Filtrar por año seleccionado
  var meses=histYear?allMeses.filter(function(m){return mesYear(m)===histYear;}):allMeses;
  if(meses.length>13) meses=meses.slice(meses.length-13);

  // Clientes filtrados (excluidos, misma logica que tabla)
  var clientes=(window.CLIENTES||[]).filter(function(c){ return !excluidos.has(c); });

  // Sumar valor por cliente por mes (respetando rango d1-d2)
  function getVal(cod,mes){
    var mdata=(raw[cod]||{})[mes]||{};
    var v=0;
    for(var d=d1;d<=d2;d++){
      var e=mdata[String(d)];
      if(e) v+=histMode==='venta'?e[0]:e[1];
    }
    return v;
  }

  // Construir filas
  var rows=clientes.map(function(cod){
    var vals=meses.map(function(m){ return getVal(cod,m); });
    return {cod:cod, vals:vals};
  }).filter(function(r){ return r.vals.some(function(v){return v>0;}); });

  // Ordenar
  if(histSortIdx===-1){
    rows.sort(function(a,b){
      var va=a.vals[a.vals.length-1]||0, vb=b.vals[b.vals.length-1]||0;
      return vb-va;
    });
  } else if(histSortIdx===0){
    rows.sort(function(a,b){ return histSortAsc?a.cod.localeCompare(b.cod):b.cod.localeCompare(a.cod); });
  } else {
    var si=histSortIdx-1;
    rows.sort(function(a,b){ var d=(a.vals[si]||0)-(b.vals[si]||0); return histSortAsc?d:-d; });
  }

  // Header
  var thead=document.getElementById('theadHistorico');
  var thHtml='<tr style="background:#0d1a26">';
  thHtml+='<th style="text-align:left;padding:6px 10px;color:#64748b;font-size:.72rem;border-bottom:2px solid #1e3a4e;cursor:pointer;white-space:nowrap" onclick="histColSort(0)">CLIENTE</th>';
  meses.forEach(function(m,i){
    thHtml+='<th style="text-align:right;padding:6px 8px;color:#94a3b8;font-size:.72rem;border-bottom:2px solid #1e3a4e;cursor:pointer;white-space:nowrap" onclick="histColSort('+(i+1)+')">'+mesLabel(m)+'</th>';
  });
  thHtml+='<th style="text-align:right;padding:6px 8px;color:#f97316;font-size:.72rem;border-bottom:2px solid #1e3a4e;white-space:nowrap">vs Ant.</th>';
  thHtml+='<th style="text-align:right;padding:6px 8px;color:#64748b;font-size:.72rem;border-bottom:2px solid #1e3a4e;white-space:nowrap">Dif.</th>';
  thHtml+='<th style="text-align:center;padding:6px 10px;color:#64748b;font-size:.72rem;border-bottom:2px solid #1e3a4e">Tendencia</th>';
  thHtml+='</tr>';
  thead.innerHTML=thHtml;

  // Body
  var tbody=document.getElementById('tbodyHistorico');
  tbody.innerHTML='';
  if(!rows.length){
    tbody.innerHTML='<tr><td colspan="'+(meses.length+4)+'" style="text-align:center;padding:20px;color:#445566">Sin datos históricos disponibles</td></tr>';
    return;
  }

  // Totales por mes
  var totals=meses.map(function(_,i){ return rows.reduce(function(s,r){return s+(r.vals[i]||0);},0); });

  rows.forEach(function(r,ri){
    var isMes=ri%2===0;
    var tr=document.createElement('tr');
    tr.style.cssText='border-bottom:1px solid #0e1e2e;'+(isMes?'background:#060f18':'background:#040a12');
    var last=r.vals[r.vals.length-1]||0;
    var prev=r.vals.length>1?(r.vals[r.vals.length-2]||0):0;
    var dif=last-prev;
    var pct=prev>0?(dif/prev*100):0;
    var difCol=dif>=0?'#4ade80':'#ef4444';
    var pctStr=(dif>=0?'+':'')+pct.toFixed(1)+'%';
    var html='<td style="text-align:left;padding:5px 10px;font-weight:600;color:#e2e8f0;white-space:nowrap">'+r.cod+'</td>';
    r.vals.forEach(function(v){
      html+='<td style="text-align:right;padding:5px 8px;color:#cbd5e1;white-space:nowrap">';
      if(v>0) html+=histMode==='venta'?formatCopCompact(v):NUM.format(v);
      else html+='<span style="color:#334155">-</span>';
      html+='</td>';
    });
    html+='<td style="text-align:right;padding:5px 8px;color:'+difCol+';font-weight:700;white-space:nowrap">'+pctStr+'</td>';
    html+='<td style="text-align:right;padding:5px 8px;color:'+difCol+';white-space:nowrap">';
    if(dif!==0) html+=(histMode==='venta'?formatCopCompact(Math.abs(dif)):NUM.format(Math.abs(dif)));
    html+='</td>';
    html+='<td style="text-align:center;padding:5px 8px">'+sparkline(r.vals.filter(function(v){return v>0;}))+' </td>';
    tr.innerHTML=html;
    tbody.appendChild(tr);
  });

  // Fila totales
  var tf=document.createElement('tr');
  tf.style.cssText='background:#1a2d40;border-top:2px solid #2d4a6a;font-weight:700';
  var tlast=totals[totals.length-1]||0;
  var tprev=totals.length>1?(totals[totals.length-2]||0):0;
  var tdif=tlast-tprev;
  var tcol=tdif>=0?'#4ade80':'#ef4444';
  var tpct=tprev>0?(tdif/tprev*100):0;
  var thtml='<td style="text-align:left;padding:6px 10px;color:#f97316">TOTAL</td>';
  totals.forEach(function(v){
    thtml+='<td style="text-align:right;padding:6px 8px;color:#f0c060">'+(histMode==='venta'?formatCopCompact(v):NUM.format(v))+'</td>';
  });
  thtml+='<td style="text-align:right;padding:6px 8px;color:'+tcol+'">'+(tdif>=0?'+':'')+tpct.toFixed(1)+'%</td>';
  thtml+='<td style="text-align:right;padding:6px 8px;color:'+tcol+'">'+(histMode==='venta'?formatCopCompact(Math.abs(tdif)):NUM.format(Math.abs(tdif)))+'</td>';
  thtml+='<td></td>';
  tf.innerHTML=thtml;
  tbody.appendChild(tf);

  buildHistCharts();
  buildInsights();
}

function histColSort(idx){
  if(histSortIdx===idx) histSortAsc=!histSortAsc;
  else{histSortIdx=idx; histSortAsc=idx===0;}
  buildHistorico();
}

function formatCopCompact(v){
  if(!v&&v!==0) return '-';
  v=Math.round(v);
  if(Math.abs(v)>=1000000000) return (v/1000000000).toFixed(1)+'B';
  if(Math.abs(v)>=1000000) return (v/1000000).toFixed(1)+'M';
  if(Math.abs(v)>=1000) return (v/1000).toFixed(0)+'K';
  return NUM.format(v);
}

/* ---- Gráficos de operaciones ---- */
function opsHistMeses(keys){
  var ms={}, oph=window.OPS_HISTORICO||{};
  keys.forEach(function(k){Object.keys(oph[k]||{}).forEach(function(m){ms[m]=1;});});
  var arr=Object.keys(ms).sort();
  if(histYear) arr=arr.filter(function(m){return mesYear(m)===histYear;});
  return arr;
}

function opsHistVal(key, mes){
  var mdata=((window.OPS_HISTORICO||{})[key]||{})[mes]||{};
  var v=0;
  for(var d=d1;d<=d2;d++){var e=mdata[String(d)];if(e)v+=histMode==='venta'?e[0]:e[2];}
  return v;
}

function getClientMonthVal(cod, mes){
  var cdata=((window.HISTORICO||{})[cod]||{})[mes]||{};
  var v=0;
  for(var d=d1;d<=d2;d++){var e=cdata[String(d)];if(e)v+=histMode==='venta'?e[0]:e[1];}
  return v;
}

function drawBarSvg(meses, series, gid){
  var VW=280,VH=100,pL=6,pB=19,pT=16,pR=10;
  var cW=VW-pL-pR, cH=VH-pT-pB;
  var n=meses.length;
  var noData='<svg viewBox="0 0 '+VW+' '+VH+'" width="100%"><text x="'+(VW/2)+'" y="'+(VH/2)+'" text-anchor="middle" fill="#334155" font-size="10">Sin datos</text></svg>';
  if(!n) return noData;
  var totals=meses.map(function(_,i){var s=0;series.forEach(function(sr){s+=sr.vals[i]||0;});return s;});
  var maxV=Math.max.apply(null,totals);
  if(!maxV) return noData;
  var avg=totals.reduce(function(a,b){return a+b;},0)/n;
  var avgY=(VH-pB)-((avg/maxV)*cH);
  var gW=cW/n, bW=Math.max(3,gW*0.58);
  var uid=gid||'g0';

  var svg='<svg viewBox="0 0 '+VW+' '+VH+'" width="100%" style="display:block;overflow:visible">';
  /* degradados */
  svg+='<defs>';
  series.forEach(function(sr,si){
    svg+='<linearGradient id="'+uid+'_'+si+'" x1="0" y1="0" x2="0" y2="1">'+
      '<stop offset="0%" stop-color="'+sr.color+'" stop-opacity="0.9"/>'+
      '<stop offset="100%" stop-color="'+sr.color+'" stop-opacity="0.2"/>'+
      '</linearGradient>';
    svg+='<filter id="'+uid+'_glow'+si+'"><feGaussianBlur stdDeviation="2" result="b"/><feMerge><feMergeNode in="b"/><feMergeNode in="SourceGraphic"/></feMerge></filter>';
  });
  svg+='</defs>';
  /* grid sutil */
  for(var g=1;g<=3;g++){var gy=VH-pB-(cH/3*g);svg+='<line x1="'+pL+'" y1="'+gy.toFixed(1)+'" x2="'+(VW-pR+8)+'" y2="'+gy.toFixed(1)+'" stroke="#0d1a26" stroke-width="0.5"/>';}
  /* linea promedio */
  svg+='<line x1="'+pL+'" y1="'+avgY.toFixed(1)+'" x2="'+(VW-pR+6)+'" y2="'+avgY.toFixed(1)+'" stroke="#475569" stroke-width="0.8" stroke-dasharray="2.5,2.5"/>';
  svg+='<text x="'+(VW-pR+8)+'" y="'+(avgY+3).toFixed(1)+'" fill="#475569" font-size="7">Ø</text>';

  for(var i=0;i<n;i++){
    var cx=pL+i*gW+gW/2, bx=cx-bW/2, bot=VH-pB, ys=bot, tH=0;
    var isLast=(i===n-1);
    series.forEach(function(sr,si){
      var v=sr.vals[i]||0; if(v<=0)return;
      var bh=(v/maxV)*cH; ys-=bh; tH+=bh;
      var fill=series.length===1?'url(#'+uid+'_'+si+')':sr.color+(isLast?'dd':'88');
      var extra=isLast?' stroke="'+sr.color+'" stroke-width="1"':'';
      var filt=isLast&&series.length===1?' filter="url(#'+uid+'_glow'+si+')"':'';
      svg+='<rect x="'+bx.toFixed(1)+'" y="'+ys.toFixed(1)+'" width="'+bW.toFixed(1)+'" height="'+bh.toFixed(1)+'" fill="'+fill+'" rx="2"'+extra+filt+'/>';
    });
    /* etiqueta valor encima de TODAS las barras */
    if(tH>0){
      var sum=0; series.forEach(function(sr){sum+=sr.vals[i]||0;});
      var lCol=isLast?'#f0c060':'#475569';
      var lW=isLast?' font-weight="bold"':'';
      var lSz=isLast?'8.5':'7.5';
      svg+='<text x="'+cx.toFixed(1)+'" y="'+(bot-tH-3).toFixed(1)+'" text-anchor="middle" fill="'+lCol+'" font-size="'+lSz+'"'+lW+'>'+formatCopCompact(sum)+'</text>';
    }
    /* etiqueta mes */
    var mCol=isLast?'#64748b':'#2d3d4d';
    svg+='<text x="'+cx.toFixed(1)+'" y="'+(VH-5)+'" text-anchor="middle" fill="'+mCol+'" font-size="7.5">'+mesLabel(meses[i])+'</text>';
  }
  svg+='</svg>';
  return svg;
}

function trendBadge(cur, prev){
  if(!prev||!cur) return '';
  var pct=(cur-prev)/prev*100;
  var up=pct>=0, col=up?'#4ade80':'#ef4444', arr=up?'▲':'▼';
  return '<span style="font-size:.65rem;font-weight:700;color:'+col+';margin-left:6px">'+arr+' '+Math.abs(pct).toFixed(1)+'%</span>';
}

function buildHistCharts(){
  function lastTwo(vals){return vals.length>=2?[vals[vals.length-1],vals[vals.length-2]]:[vals[0]||0,0];}

  /* NACIONAL — suma solo clientes no excluidos (igual que la tabla) */
  var nacM=opsHistMeses(['NACIONAL']);
  var nacV=nacM.map(function(m){
    var clts=(window.CLIENTES||[]).filter(function(c){return !excluidos.has(c);});
    return clts.reduce(function(s,cod){return s+getClientMonthVal(cod,m);},0);
  });
  var nlt=lastTwo(nacV);
  document.getElementById('hcNac').innerHTML=
    '<div style="display:flex;align-items:baseline;justify-content:space-between;margin-bottom:4px">'+
    '<span style="font-size:.7rem;font-weight:700;color:#f97316;letter-spacing:.03em">NACIONAL</span>'+
    '<span style="font-size:.85rem;font-weight:800;color:#f0c060">'+formatCopCompact(nlt[0])+trendBadge(nlt[0],nlt[1])+'</span>'+
    '</div>'+drawBarSvg(nacM,[{vals:nacV,color:'#f97316'}],'gnac')+
    '<div style="font-size:.65rem;color:#334155;margin-top:2px">Ø prom: '+formatCopCompact(nacV.length?nacV.reduce(function(a,b){return a+b;},0)/nacV.length:0)+'</div>';

  /* JEMA */
  var jemM=opsHistMeses(['JEMA']), jemV=jemM.map(function(m){return opsHistVal('JEMA',m);});
  var jlt=lastTwo(jemV);
  document.getElementById('hcJema').innerHTML=
    '<div style="display:flex;align-items:baseline;justify-content:space-between;margin-bottom:4px">'+
    '<span style="font-size:.7rem;font-weight:700;color:#22c55e;letter-spacing:.03em">JERÓNIMO</span>'+
    '<span style="font-size:.85rem;font-weight:800;color:#f0c060">'+formatCopCompact(jlt[0])+trendBadge(jlt[0],jlt[1])+'</span>'+
    '</div>'+drawBarSvg(jemM,[{vals:jemV,color:'#22c55e'}],'gjem')+
    '<div style="font-size:.65rem;color:#334155;margin-top:2px">Ø prom: '+formatCopCompact(jemV.length?jemV.reduce(function(a,b){return a+b;},0)/jemV.length:0)+'</div>';

  /* COMEX */
  var cxM=opsHistMeses(['IMPO','EXPO','NAL-TL']);
  var imV=cxM.map(function(m){return opsHistVal('IMPO',m);}),
      exV=cxM.map(function(m){return opsHistVal('EXPO',m);}),
      tlV=cxM.map(function(m){return opsHistVal('NAL-TL',m);});
  var cxTot=cxM.map(function(_,i){return (imV[i]||0)+(exV[i]||0)+(tlV[i]||0);});
  var clt=lastTwo(cxTot);
  var legH='<span style="font-size:.62rem;color:#475569">'+
    '<span style="background:#3b82f6;display:inline-block;width:7px;height:7px;border-radius:1px;margin-right:2px;vertical-align:middle"></span>IMPO '+
    '<span style="background:#8b5cf6;display:inline-block;width:7px;height:7px;border-radius:1px;margin:0 2px 0 6px;vertical-align:middle"></span>EXPO '+
    '<span style="background:#06b6d4;display:inline-block;width:7px;height:7px;border-radius:1px;margin:0 2px 0 6px;vertical-align:middle"></span>Nal-TL</span>';
  document.getElementById('hcComex').innerHTML=
    '<div style="display:flex;align-items:baseline;justify-content:space-between;margin-bottom:2px">'+
    '<span style="font-size:.7rem;font-weight:700;color:#8b5cf6;letter-spacing:.03em">COMEX</span>'+
    '<span style="font-size:.85rem;font-weight:800;color:#f0c060">'+formatCopCompact(clt[0])+trendBadge(clt[0],clt[1])+'</span>'+
    '</div><div style="margin-bottom:3px">'+legH+'</div>'+
    drawBarSvg(cxM,[{vals:imV,color:'#3b82f6'},{vals:exV,color:'#8b5cf6'},{vals:tlV,color:'#06b6d4'}],'gcx')+
    '<div style="font-size:.65rem;color:#334155;margin-top:2px">Ø prom: '+formatCopCompact(cxTot.length?cxTot.reduce(function(a,b){return a+b;},0)/cxTot.length:0)+'</div>';

  /* CEDIS */
  var cdM=opsHistMeses(['CEDIS']), cdV=cdM.map(function(m){return opsHistVal('CEDIS',m);});
  var cdlt=lastTwo(cdV);
  document.getElementById('hcCedis').innerHTML=
    '<div style="display:flex;align-items:baseline;justify-content:space-between;margin-bottom:4px">'+
    '<span style="font-size:.7rem;font-weight:700;color:#fbbf24;letter-spacing:.03em">CEDIS</span>'+
    '<span style="font-size:.85rem;font-weight:800;color:#f0c060">'+formatCopCompact(cdlt[0])+trendBadge(cdlt[0],cdlt[1])+'</span>'+
    '</div>'+drawBarSvg(cdM,[{vals:cdV,color:'#fbbf24'}],'gcd')+
    '<div style="font-size:.65rem;color:#334155;margin-top:2px">Ø prom: '+formatCopCompact(cdV.length?cdV.reduce(function(a,b){return a+b;},0)/cdV.length:0)+'</div>';
}

/* ---- Análisis Inteligente ---- */
function buildInsights(){
  var el=document.getElementById('histInsights');
  if(!el) return;
  var nacM=opsHistMeses(['NACIONAL']);
  if(nacM.length<2){el.innerHTML='';return;}
  var lastM=nacM[nacM.length-1], prevM=nacM[nacM.length-2];
  var lLbl=mesLabel(lastM), pLbl=mesLabel(prevM);

  /* Totales por operacion */
  function opTotal(keys, mes){return keys.reduce(function(s,k){return s+opsHistVal(k,mes);},0);}
  var ops=[
    {label:'Nacional', keys:['NACIONAL'], col:'#f97316'},
    {label:'Jerónimo', keys:['JEMA'], col:'#22c55e'},
    {label:'COMEX',    keys:['IMPO','EXPO','NAL-TL'], col:'#8b5cf6'},
    {label:'CEDIS',    keys:['CEDIS'], col:'#fbbf24'}
  ];

  /* Análisis por cliente de NACIONAL */
  var clientes=(window.CLIENTES||[]).filter(function(c){return !excluidos.has(c);});
  var cdata=clientes.map(function(cod){
    var last=getClientMonthVal(cod,lastM), prev=getClientMonthVal(cod,prevM);
    var dif=last-prev, pct=prev>0?(dif/prev*100):(last>0?100:0);
    return {cod:cod,last:last,prev:prev,dif:dif,pct:pct};
  }).filter(function(c){return c.last>0||c.prev>0;});

  /* Totales NACIONAL filtrados por excluidos */
  function nacFilteredTotal(mes){
    var clts=(window.CLIENTES||[]).filter(function(c){return !excluidos.has(c);});
    return clts.reduce(function(s,cod){return s+getClientMonthVal(cod,mes);},0);
  }
  var nacLast=nacFilteredTotal(lastM), nacPrev=nacFilteredTotal(prevM);
  var nacDif=nacLast-nacPrev, nacPct=nacPrev>0?(nacDif/nacPrev*100):0;
  var nacUp=nacDif>=0;

  /* Top declines & growths */
  var declining=cdata.filter(function(c){return c.dif<-500000;})
    .sort(function(a,b){return a.dif-b.dif;}).slice(0,5);
  var growing=cdata.filter(function(c){return c.dif>500000&&c.prev>0;})
    .sort(function(a,b){return b.dif-a.dif;}).slice(0,3);
  var newC=cdata.filter(function(c){return c.prev<500000&&c.last>2000000;}).slice(0,3);
  var lostC=cdata.filter(function(c){return c.prev>2000000&&c.last<500000;}).slice(0,3);

  /* Mes con mejor venta en el año */
  var allV=nacM.map(function(m){return opTotal(['NACIONAL'],m);});
  var bestIdx=allV.indexOf(Math.max.apply(null,allV));
  var bestM=nacM[bestIdx];

  /* Calcular cuántos meses consecutivos baja */
  var streak=0;
  for(var i=allV.length-1;i>=1;i--){if(allV[i]<allV[i-1])streak++;else break;}

  /* Construir HTML */
  var nacSign=nacUp?'+':'';
  var nacCol=nacUp?'#4ade80':'#ef4444';
  var nacArrow=nacUp?'▲':'▼';

  var html='<div style="background:linear-gradient(135deg,#060f18 0%,#0b1928 100%);border:1px solid '+(nacUp?'#1e4a2e':'#4a1e1e')+';border-left:3px solid '+nacCol+';border-radius:8px;padding:14px 16px;margin-bottom:16px">';
  html+='<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">';
  html+='<span style="font-size:.85rem;font-weight:800;color:'+nacCol+'">'+nacArrow+' Nacional '+lLbl+' vs '+pLbl+'</span>';
  html+='<span style="font-size:1rem;font-weight:900;color:'+nacCol+'">'+nacSign+formatCopCompact(nacDif)+'</span>';
  html+='<span style="font-size:.75rem;color:'+nacCol+';opacity:.8">'+nacSign+nacPct.toFixed(1)+'%</span>';
  html+='</div>';

  /* Resumen general de operaciones */
  html+='<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px">';
  ops.forEach(function(op){
    var last=opTotal(op.keys,lastM), prev=opTotal(op.keys,prevM);
    var d=last-prev, p=prev>0?(d/prev*100):0;
    var c=d>=0?'#4ade80':'#ef4444', a=d>=0?'▲':'▼';
    html+='<div style="background:#0d1a26;border:1px solid #1e3a4e;border-radius:5px;padding:4px 10px;font-size:.68rem">';
    html+='<span style="color:'+op.col+';font-weight:700">'+op.label+'</span>';
    html+=' <span style="color:'+c+'">'+a+' '+Math.abs(p).toFixed(1)+'%</span>';
    html+='</div>';
  });
  html+='</div>';

  /* Clientes que más bajaron */
  if(declining.length){
    var totalDecl=declining.reduce(function(s,c){return s+c.dif;},0);
    html+='<div style="margin-bottom:10px">';
    html+='<div style="font-size:.7rem;font-weight:700;color:#ef4444;margin-bottom:6px">📉 Principales caídas en Nacional ('+lLbl+')</div>';
    declining.forEach(function(c){
      var pct=Math.abs(c.pct), barW=Math.min(100,Math.abs(c.dif/declining[0].dif)*100);
      var share=nacDif!==0?(c.dif/nacDif*100):0;
      html+='<div style="margin-bottom:5px">';
      html+='<div style="display:flex;justify-content:space-between;font-size:.7rem;margin-bottom:2px">';
      html+='<span style="color:#e2e8f0;font-weight:600">'+c.cod+'</span>';
      html+='<span style="color:#ef4444">'+formatCopCompact(c.dif)+' &nbsp;<span style="color:#64748b">-'+pct.toFixed(0)+'% · '+Math.abs(share).toFixed(0)+'% del total</span></span>';
      html+='</div>';
      html+='<div style="background:#1a0d0d;border-radius:2px;height:4px"><div style="background:#ef4444;border-radius:2px;height:4px;width:'+barW.toFixed(0)+'%;transition:width .3s"></div></div>';
      html+='</div>';
    });
    html+='<div style="font-size:.65rem;color:#475569;margin-top:4px">Estos '+declining.length+' clientes explican '+formatCopCompact(totalDecl)+' de la variación ('+nacDif!==0?(totalDecl/nacDif*100).toFixed(0)+'%':'—'+')</div>';
    html+='</div>';
  }

  /* Clientes que subieron */
  if(growing.length){
    html+='<div style="margin-bottom:10px">';
    html+='<div style="font-size:.7rem;font-weight:700;color:#4ade80;margin-bottom:6px">📈 Clientes que subieron</div>';
    html+='<div style="display:flex;flex-wrap:wrap;gap:6px">';
    growing.forEach(function(c){
      html+='<div style="background:#0d2a1a;border:1px solid #1e4a2e;border-radius:5px;padding:3px 10px;font-size:.68rem">';
      html+='<span style="color:#e2e8f0;font-weight:600">'+c.cod+'</span>';
      html+=' <span style="color:#4ade80">+'+formatCopCompact(c.dif)+'</span>';
      html+='</div>';
    });
    html+='</div></div>';
  }

  /* Alertas */
  var alerts=[];
  if(lostC.length) alerts.push('⚠️ Sin actividad este mes (tenían el anterior): <b style="color:#fbbf24">'+lostC.map(function(c){return c.cod;}).join(', ')+'</b>');
  if(newC.length)  alerts.push('✅ Actividad nueva o recuperada: <b style="color:#4ade80">'+newC.map(function(c){return c.cod;}).join(', ')+'</b>');
  if(streak>=2)    alerts.push('🔴 Nacional lleva <b style="color:#ef4444">'+streak+' meses consecutivos</b> a la baja');
  if(bestIdx===nacM.length-1) alerts.push('🏆 <b style="color:#f0c060">'+lLbl+' es el mejor mes del año</b> en Nacional');
  else if(bestM)  alerts.push('📌 El mejor mes del año fue <b style="color:#f0c060">'+mesLabel(bestM)+'</b> con '+formatCopCompact(Math.max.apply(null,allV)));
  if(alerts.length){
    html+='<div style="border-top:1px solid #1e3a4e;padding-top:8px;display:flex;flex-direction:column;gap:4px">';
    alerts.forEach(function(a){html+='<div style="font-size:.68rem;color:#94a3b8">'+a+'</div>';});
    html+='</div>';
  }

  html+='</div>';
  el.innerHTML=html;
}

/* ======================================================
   AJOVER — Detalle por Tipo de Operación
   ====================================================== */
function buildAjover(){
  var raw = window.AJOV_HIST || {};
  var m   = window.META || {};

  // Obtener todos los meses disponibles en los datos AJOV
  var mesesSet = {};
  Object.keys(raw).forEach(function(t){
    Object.keys(raw[t]).forEach(function(mes){ mesesSet[mes] = 1; });
  });
  var allMeses = Object.keys(mesesSet).sort();
  if(!allMeses.length){ document.getElementById('tbodyAjover').innerHTML='<tr><td colspan="20" style="text-align:center;padding:24px;color:#334155">Sin datos AJOV disponibles.</td></tr>'; return; }
  // Limitar a los 8 meses más recientes
  if(allMeses.length > 8) allMeses = allMeses.slice(allMeses.length - 8);

  var mesActIdx = allMeses.length - 1;

  // Suma de V, U, N para un tipo en un mes restringida a d1-d2
  function getVUN(tipo, mes){
    var mdata = (raw[tipo] || {})[mes] || {};
    var V=0, U=0, N=0;
    for(var day=d1; day<=d2; day++){
      var e = mdata[String(day)];
      if(e){ V+=e[0]; U+=e[1]; N+=e[2]; }
    }
    return {V:V, U:U, N:N};
  }

  // Tipos de operación con algún dato
  var tipos = Object.keys(raw).filter(function(t){
    return allMeses.some(function(mes){ var r=getVUN(t,mes); return r.V>0||r.N>0; });
  });
  // Orden fijo deseado; el resto al final
  var ORDER = ['Transferencia CTG - MAD','Transferencia MAD - CTG','Graneles',
               'Transferencia CTG - CALI','SENCILLOS MADRID','ZORROS CALI',
               'TRANSFERENCIA ALAMO','FIJOS CALI','FIJOS MEDELLIN',
               'ENTREGA A CLIENTES','TRANSFERENCIAS','OTROS NOCO'];
  tipos.sort(function(a,b){
    var ia=ORDER.indexOf(a), ib=ORDER.indexOf(b);
    if(ia<0) ia=99; if(ib<0) ib=99;
    if(ia!==ib) return ia-ib;
    return a.localeCompare(b);
  });

  // KPI strip: mes actual totals
  var kpiV=0, kpiU=0, kpiN=0, kpiVant=0;
  tipos.forEach(function(t){
    var cur=getVUN(t,allMeses[mesActIdx]);
    kpiV+=cur.V; kpiU+=cur.U; kpiN+=cur.N;
    if(mesActIdx>0){ var ant=getVUN(t,allMeses[mesActIdx-1]); kpiVant+=ant.V; }
  });
  var kpiDif=kpiV-kpiVant, kpiPct=kpiVant>0?(kpiDif/kpiVant*100):0;
  var kpiCol=kpiDif>=0?'#4ade80':'#ef4444';
  var kpiEl=document.getElementById('ajovKpi');
  if(kpiEl){
    kpiEl.innerHTML=
      '<div style="background:#0d1a26;border:1px solid #1e3a4e;border-radius:8px;padding:8px 14px">'+
        '<div style="color:#60a5fa;font-size:.62rem;font-weight:700;letter-spacing:.08em">VENTA '+mesLabel(allMeses[mesActIdx])+'</div>'+
        '<div style="color:#f0c060;font-size:1rem;font-weight:800">'+formatCopCompact(kpiV)+'</div>'+
      '</div>'+
      '<div style="background:#0d1a26;border:1px solid #1e3a4e;border-radius:8px;padding:8px 14px">'+
        '<div style="color:#60a5fa;font-size:.62rem;font-weight:700;letter-spacing:.08em">vs MES ANT.</div>'+
        '<div style="color:'+kpiCol+';font-size:1rem;font-weight:800">'+(kpiDif>=0?'+':'')+formatCopCompact(kpiDif)+'</div>'+
        '<div style="color:'+kpiCol+';font-size:.65rem">'+(kpiPct>=0?'+':'')+kpiPct.toFixed(1)+'%</div>'+
      '</div>'+
      '<div style="background:#0d1a26;border:1px solid #1e3a4e;border-radius:8px;padding:8px 14px">'+
        '<div style="color:#60a5fa;font-size:.62rem;font-weight:700;letter-spacing:.08em">VIAJES</div>'+
        '<div style="color:#e2e8f0;font-size:1rem;font-weight:800">'+kpiN+'</div>'+
      '</div>'+
      '<div style="background:#0d1a26;border:1px solid #1e3a4e;border-radius:8px;padding:8px 14px">'+
        '<div style="color:#60a5fa;font-size:.62rem;font-weight:700;letter-spacing:.08em">MARGEN</div>'+
        '<div style="color:#4ade80;font-size:1rem;font-weight:800">'+(kpiV>0?(kpiU/kpiV*100).toFixed(1)+'%':'—')+'</div>'+
      '</div>';
  }

  // ---- HEADER ----
  var thBg  = 'background:#0a1520;color:#64748b;font-size:.7rem;font-weight:700;padding:7px 10px;border-bottom:2px solid #1e3a4e;white-space:nowrap;';
  var thAct = 'background:#0d1e2e;color:#f0c060;font-size:.7rem;font-weight:700;padding:7px 10px;border-bottom:2px solid #f97316;white-space:nowrap;';
  var th2   = 'background:#060f18;color:#334155;font-size:.65rem;padding:3px 8px;border-bottom:1px solid #0e1e2e;white-space:nowrap;';
  var h1='<tr>', h2='<tr>';
  h1 += '<th style="'+thBg+'text-align:left;min-width:190px">TIPO OPERACIÓN</th>';
  h2 += '<th style="'+th2+'"></th>';
  allMeses.forEach(function(mes, mi){
    var isAct = (mi===mesActIdx);
    var sty   = isAct ? thAct : thBg;
    h1 += '<th colspan="2" style="'+sty+'text-align:center">'+mesLabel(mes)+'</th>';
    h2 += '<th style="'+th2+'text-align:right">VENTA</th><th style="'+th2+'text-align:right">VJE</th>';
  });
  h1 += '<th style="'+thBg+'text-align:right;color:#f97316">vs Ant.</th>';
  h1 += '<th style="'+thBg+'text-align:right;color:#f97316">%</th>';
  h2 += '<th style="'+th2+'"></th><th style="'+th2+'"></th>';
  h1 += '</tr>'; h2 += '</tr>';
  document.getElementById('theadAjover').innerHTML = h1 + h2;

  // ---- ROWS ----
  var tbody = document.getElementById('tbodyAjover');
  tbody.innerHTML = '';

  // Totals per month
  var totals = allMeses.map(function(mes){
    var V=0,U=0,N=0;
    tipos.forEach(function(t){ var r=getVUN(t,mes); V+=r.V; U+=r.U; N+=r.N; });
    return {V:V,U:U,N:N};
  });

  tipos.forEach(function(tipo, ri){
    var tr = document.createElement('tr');
    tr.style.cssText='border-bottom:1px solid #0e1e2e;'+(ri%2===0?'background:#060f18':'background:#040a12');

    var actR = getVUN(tipo, allMeses[mesActIdx]);
    var antR = mesActIdx > 0 ? getVUN(tipo, allMeses[mesActIdx-1]) : {V:0,U:0,N:0};
    var dif  = actR.V - antR.V;
    var pct  = antR.V > 0 ? (dif/antR.V*100) : (actR.V>0?100:0);
    var difCol = dif >= 0 ? '#4ade80' : '#ef4444';

    var html = '<td style="text-align:left;padding:6px 10px;font-weight:600;color:#e2e8f0;white-space:nowrap">'+tipo+'</td>';
    allMeses.forEach(function(mes, mi){
      var r = getVUN(tipo, mes);
      var isAct = (mi === mesActIdx);
      var vCol = isAct ? '#f0c060' : '#7a8fa8';
      var nCol = '#334155';
      html += '<td style="text-align:right;padding:5px 10px;color:'+vCol+';white-space:nowrap;font-weight:'+(isAct?'700':'400')+'">';
      html += r.V > 0 ? formatCopCompact(r.V) : '<span style="color:#1a2a3a">—</span>';
      html += '</td>';
      html += '<td style="text-align:right;padding:5px 6px;color:'+nCol+';font-size:.7rem;white-space:nowrap">';
      html += r.N > 0 ? r.N : '<span style="color:#1a2a3a">—</span>';
      html += '</td>';
    });
    html += '<td style="text-align:right;padding:6px 10px;color:'+difCol+';font-weight:700;white-space:nowrap">'+(dif>=0?'+':'')+formatCopCompact(dif)+'</td>';
    html += '<td style="text-align:right;padding:6px 8px;color:'+difCol+';font-weight:700;white-space:nowrap">'+(pct>=0?'+':'')+pct.toFixed(1)+'%</td>';
    tr.innerHTML = html;
    tbody.appendChild(tr);
  });

  // Total row
  var tf = document.createElement('tr');
  tf.style.cssText = 'background:#1a2d40;border-top:2px solid #2d4a6a;font-weight:700';
  var tAct = totals[mesActIdx], tAnt = mesActIdx>0?totals[mesActIdx-1]:{V:0,N:0};
  var tDif = tAct.V-tAnt.V, tPct=tAnt.V>0?(tDif/tAnt.V*100):0, tCol=tDif>=0?'#4ade80':'#ef4444';
  var tHtml = '<td style="text-align:left;padding:7px 10px;color:#f97316">TOTAL</td>';
  totals.forEach(function(t, ti){
    var isAct = (ti===mesActIdx);
    tHtml += '<td style="text-align:right;padding:7px 10px;color:'+(isAct?'#f0c060':'#94a3b8')+';font-weight:700">'+formatCopCompact(t.V)+'</td>';
    tHtml += '<td style="text-align:right;padding:7px 6px;color:#64748b;font-size:.7rem">'+t.N+'</td>';
  });
  tHtml += '<td style="text-align:right;padding:7px 10px;color:'+tCol+'">'+(tDif>=0?'+':'')+formatCopCompact(tDif)+'</td>';
  tHtml += '<td style="text-align:right;padding:7px 8px;color:'+tCol+'">'+(tPct>=0?'+':'')+tPct.toFixed(1)+'%</td>';
  tf.innerHTML = tHtml;
  tbody.appendChild(tf);
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
  buildOpsTable();
  buildClientList();
  buildTable();
})();
</script>
</body>
</html>"""


if __name__ == "__main__":
    main()
