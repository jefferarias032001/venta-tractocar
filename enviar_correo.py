# -*- coding: utf-8 -*-
"""
TRACTOCAR · Envio automatico de tabla VENTA NACIONAL por correo.

Uso:
    python enviar_correo.py

El email se envia con asunto "VENTA NACIONAL" y contiene la tabla de clientes
con datos del mes actual, cortados al DIA ANTERIOR (d2 = hoy - 1).
Se envia via Outlook instalado localmente (win32com).

Configura DESTINATARIOS con las direcciones que deben recibir el correo.
"""

import os, sys, datetime as dt, calendar, warnings
import pandas as pd
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import leer_datos as _proc

warnings.filterwarnings("ignore")

# ======================== CONFIGURACION ========================

DESTINATARIOS = [
    "jarias@tractocar.com",
    # agrega mas correos aqui separados por coma, por ejemplo:
    # "gerencia@tractocar.com",
]

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

NIT_A_COD = {
    "860013771": "AJOV",
    "817002753": "DRYP",
    "860513970": "MILP",
    "860002274": "ETER",
    "800059470": "ESEN",
    "830006735": "ALPO",
    "860015753": "KIMB",
    "890300466": "TECN",
    "860522056": "LAMI",
    "890104438": "EMTR",
    "860001899": "CORP",
    "800217481": "CASC",
    "890900161": "PFAMI",
    "900226838": "NOCO",
    "860007277": "MERE",
}

CLIENTES_OPERACION = {"JEMA"}
SIEMPRE_OTROS = {"CPA", "GDANE", "SOCO", "YUPI",
                 "ESEN_CR_ESPE", "CRESC", "LHCO", "MOIN", "ESEN_MB", "MECO"}
EXCLUIR_CORREO = {"AJOV_MOV"}  # clientes excluidos del correo (pero si en el dashboard)
COD_ALIAS = {"COLP": "KIMB"}

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


def fmt_cop(v):
    """Formato $ 1.234.567"""
    if pd.isna(v) or v == 0:
        return "$ -"
    return f"$ {int(v):,}".replace(",", ".")


def fmt_pct(v):
    if pd.isna(v):
        return "-"
    return f"{v * 100:.1f}%"


def build_tabla(d2):
    today = dt.date.today()
    mes_actual   = today.strftime("%Y-%m")
    mes_ant_date = today.replace(day=1) - dt.timedelta(days=1)
    mes_anterior = mes_ant_date.strftime("%Y-%m")
    dias_mes     = calendar.monthrange(today.year, today.month)[1]
    d1 = 1

    print(f"  Leyendo datos — mes {mes_actual}, rango dia {d1} al {d2}...")
    u = _proc.obtener_union(verbose=False)
    u["_nit"] = u["ClienteNIT"].apply(norm_nit)
    u["Cod"]  = u["_nit"].map(NIT_A_COD)

    u_nac = u[u["Fuente"] == "NACIONAL"].copy()
    if "CodCliente" in u.columns:
        u_nac["Cod"] = u_nac["CodCliente"].astype(str).str.strip().replace({"": None, "nan": None, "<NA>": None})
    mask_sin = u_nac["Cod"].isna()
    if mask_sin.any():
        u_nac.loc[mask_sin, "Cod"] = u_nac.loc[mask_sin, "_nit"].map(NIT_A_COD)
    u_nac = u_nac[u_nac["Cod"].notna()].copy()
    u_nac["Cod"] = u_nac["Cod"].replace(COD_ALIAS)
    if "Dia" not in u_nac.columns:
        u_nac["Dia"] = pd.to_datetime(u_nac["Fecha"], errors="coerce").dt.day
    if "Subseg" not in u_nac.columns:
        u_nac["Subseg"] = u_nac["Token"].apply(get_subseg)

    # Presupuesto
    hrow = buscar_header_row_cod(RUTA_PPTO, HOJA_PPTO)
    df_p = pd.read_excel(RUTA_PPTO, sheet_name=HOJA_PPTO, header=hrow)
    df_p.columns = [str(c).strip() for c in df_p.columns]
    df_p = df_p[df_p["Cod"].notna()].copy()
    df_p["Cod"] = df_p["Cod"].astype(str).str.strip()
    budget = df_p.groupby("Cod").agg(
        PPTO        =("Venta total proyectada", "sum"),
        META_UTIL   =("Utilidad Total",         "sum"),
        M_VIAJES    =("Viajes totales",          "sum"),
        COMPRA_PPTO =("Compra total proyectada", "sum"),
    ).reset_index()

    pct_s = pd.to_numeric(df_p.get("% INTER", 0), errors="coerce").fillna(0)
    vp_s  = pd.to_numeric(df_p.get("Venta total proyectada", 0), errors="coerce").fillna(0)
    df_p["_wp"] = pct_s * vp_s
    df_p["_vp"] = vp_s
    pct_agg = df_p.groupby("Cod").agg(_wp=("_wp","sum"), _vp=("_vp","sum")).reset_index()
    pct_agg["PCT_INTER_M"] = np.where(pct_agg["_vp"] > 0, pct_agg["_wp"] / pct_agg["_vp"], 0)
    budget = budget.merge(pct_agg[["Cod","PCT_INTER_M"]], on="Cod", how="left")
    budget["PCT_INTER_M"] = budget["PCT_INTER_M"].fillna(0)

    # Mes actual — rango d1 a d2
    m_act = u_nac[(u_nac["Mes"] == mes_actual) &
                  (u_nac["Dia"] >= d1) & (u_nac["Dia"] <= d2)].copy()
    m_act_tabla = m_act[~m_act["Cod"].isin(CLIENTES_OPERACION)].copy()

    ej_all = (m_act_tabla.groupby("Cod")
              .agg(EJECUTADO=("AFacturar", "sum"),
                   UTILIDAD=("Utilidad", "sum"),
                   VIAJES=("Manifiesto", "nunique"))
              .reset_index())

    # Mes anterior — mismo rango
    m_ant = u_nac[(u_nac["Mes"] == mes_anterior) &
                  (u_nac["Dia"] >= d1) & (u_nac["Dia"] <= d2)]
    m_ant = m_ant[~m_ant["Cod"].isin(CLIENTES_OPERACION)]
    mes_ant_agg = (m_ant.groupby("Cod")["AFacturar"]
                   .sum().rename("VENTA_MES_ANT").reset_index())

    # Pendientes por planillar
    try:
        import shutil, tempfile
        tmp = os.path.join(tempfile.gettempdir(), "sol_correo_tmp.xlsx")
        shutil.copy2(RUTA_SOLICITUDES, tmp)
        hr = buscar_header_row_ob(tmp, "Sheet1")
        sol = pd.read_excel(tmp, sheet_name="Sheet1", header=hr)
        sol.columns = [str(c).strip() for c in sol.columns]
        filtrado = sol[(sol["OB_NOTES_CANCEL_USER"] == "-") & (sol["SHIP_STATUS_ENROUTE"].isna())]
        pendiente = (filtrado.groupby("OB_CUSTOMER_CODE")
                     .agg(P_PLANILLAR=("OB_RATE_RECEIVABLE", "sum"),
                          N_PLANILLAR=("OB", "count"))
                     .reset_index().rename(columns={"OB_CUSTOMER_CODE": "Cod"}))
    except Exception:
        pendiente = pd.DataFrame(columns=["Cod", "P_PLANILLAR", "N_PLANILLAR"])

    # Tabla base
    df = budget.merge(ej_all, on="Cod", how="outer")
    for right in [mes_ant_agg, pendiente]:
        df = df.merge(right, on="Cod", how="left")

    num_cols = ["PPTO", "META_UTIL", "M_VIAJES", "COMPRA_PPTO", "PCT_INTER_M",
                "EJECUTADO", "UTILIDAD", "VIAJES", "VENTA_MES_ANT",
                "P_PLANILLAR", "N_PLANILLAR"]
    for c in num_cols:
        if c not in df.columns:
            df[c] = 0.0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # Excluir clientes que no van en el correo (ej: AJOV_MOV)
    df = df[~df["Cod"].isin(EXCLUIR_CORREO)].copy()

    # Excluir clientes sin venta en ambos meses o forzados a OTROS
    sin_venta = ((df["EJECUTADO"] == 0) & (df["VENTA_MES_ANT"] == 0)) | df["Cod"].isin(SIEMPRE_OTROS)
    df_otros = df[sin_venta].copy()
    df = df[~sin_venta].copy()

    # Calcular columnas derivadas (mismo orden que el dashboard)
    rango_dias = d2 - d1 + 1
    df["PROYECCION"]    = np.where(rango_dias > 0, df["EJECUTADO"] / rango_dias * dias_mes, 0)
    df["DIF_PROY_PPTO"] = df["PROYECCION"] - df["PPTO"]
    df["PCT_CUMPL"]     = np.where(df["PPTO"] > 0, df["EJECUTADO"] / df["PPTO"], 0)
    df["DIF_DIAS"]      = df["EJECUTADO"] - df["VENTA_MES_ANT"]
    df["PCT_MARGEN"]    = np.where(df["EJECUTADO"] > 0, df["UTILIDAD"] / df["EJECUTADO"], 0)
    df["PROY_UTILIDAD"] = np.where(rango_dias > 0, df["UTILIDAD"] / rango_dias * dias_mes, 0)
    # META_VENTA_FINAL: presupuesto ajustado por % cumplimiento proyectado (simplificado = PPTO)
    df["META_VENTA_FINAL"] = df["PPTO"]

    # Orden igual al dashboard: PPTO descendente (sortK='PPTO', sortAsc=false)
    df = df.sort_values("PPTO", ascending=False)

    # Totales
    t_ppto   = df["PPTO"].sum()     + df_otros["PPTO"].sum()
    t_ej     = df["EJECUTADO"].sum()
    t_proy   = df["PROYECCION"].sum()
    t_util   = df["UTILIDAD"].sum()
    t_viajes = df["VIAJES"].sum()
    t_mant   = df["VENTA_MES_ANT"].sum() + df_otros["VENTA_MES_ANT"].sum()
    t_meta_util = df["META_UTIL"].sum()  + df_otros["META_UTIL"].sum()
    t_proy_util = df["PROY_UTILIDAD"].sum()
    t_plan   = df["P_PLANILLAR"].sum()  + df_otros["P_PLANILLAR"].sum()
    t_m_viajes = df["M_VIAJES"].sum()   + df_otros["M_VIAJES"].sum()
    totales = {
        "PPTO":          t_ppto,
        "PROYECCION":    t_proy,
        "DIF_PROY_PPTO": t_proy - t_ppto,
        "PCT_CUMPL":     t_ej / t_ppto if t_ppto > 0 else 0,
        "EJECUTADO":     t_ej,
        "VENTA_MES_ANT": t_mant,
        "DIF_DIAS":      t_ej - t_mant,
        "VIAJES":        t_viajes,
        "M_VIAJES":      t_m_viajes,
        "META_UTIL":     t_meta_util,
        "UTILIDAD":      t_util,
        "PROY_UTILIDAD": t_proy_util,
        "PCT_MARGEN":    t_util / t_ej if t_ej > 0 else 0,
        "PCT_INTER_M":   0,
        "P_PLANILLAR":   t_plan,
    }

    return df, df_otros, totales, dias_mes, mes_actual, mes_anterior


def build_html(df, df_otros, totales, dias_mes, mes_actual, mes_anterior, d2, today):
    nombre_mes = dt.date(today.year, today.month, 1).strftime("%B %Y").upper()
    mes_ant_label = dt.datetime.strptime(mes_anterior, "%Y-%m").strftime("%B %Y").upper()

    # Colores base (email-safe, fondo blanco)
    COL_HEADER  = "#1e3a5f"
    COL_TOTAL   = "#0f2a45"
    COL_ODD     = "#f8fafc"
    COL_EVEN    = "#ffffff"
    COL_POS     = "#166534"   # verde
    COL_NEG     = "#991b1b"   # rojo
    COL_NARANJA = "#c2410c"

    def row_color(cumpl):
        if cumpl >= 1.0:
            return "#f0fdf4"
        if cumpl >= 0.8:
            return "#fffbeb"
        return "#fff1f2"

    def val_color(v):
        return COL_POS if v >= 0 else COL_NEG

    td = "padding:7px 10px;border-bottom:1px solid #e2e8f0;white-space:nowrap;font-size:12px;"
    th = (f"background:{COL_HEADER};color:#ffffff;padding:8px 10px;white-space:nowrap;"
          "font-size:11px;font-weight:700;letter-spacing:.03em;border-bottom:2px solid #0f2a45;")

    # Columnas en el mismo orden que el dashboard
    col_names = [
        "CÓDIGO",
        "PPTO MES",
        "PROYECCIÓN",
        "DIF PROY vs PPTO",
        "% CUMPL",
        f"EJECUTADO<br>día 1–{d2}",
        f"M. ANT.<br>(mismo rango)",
        "EJ. vs M. ANT.",
        "VIAJES",
        "M. VIAJES",
        "META UTILIDAD",
        "UTILIDAD",
        "PROY. UTILIDAD",
        "% INTER M",
        "P. PLANILLAR",
    ]

    def make_header():
        cells = "".join(f'<th style="{th}text-align:{"left" if i==0 else "right"}">{h}</th>'
                        for i, h in enumerate(col_names))
        return f"<tr>{cells}</tr>"

    def make_row(r, bg):
        cumpl   = float(r.get("PCT_CUMPL", 0))
        row_bg  = row_color(cumpl)
        bg_use  = row_bg if bg == "alt" else bg
        dif_pp  = float(r.get("DIF_PROY_PPTO", 0))
        dif_ant = float(r.get("DIF_DIAS", 0))
        pct_m   = float(r.get("PCT_MARGEN", 0))
        pct_im  = float(r.get("PCT_INTER_M", 0))
        c_cumpl = "#166534" if cumpl >= 1 else "#92400e" if cumpl >= 0.8 else "#991b1b"
        cells = [
            # 1. CODIGO
            f'<td style="{td}background:{bg_use};font-weight:700;text-align:left;color:#1e293b">{r["Cod"]}</td>',
            # 2. PPTO MES
            f'<td style="{td}background:{bg_use};text-align:right;color:#334155">{fmt_cop(r["PPTO"])}</td>',
            # 3. PROYECCION
            f'<td style="{td}background:{bg_use};text-align:right;color:#1e40af">{fmt_cop(r["PROYECCION"])}</td>',
            # 4. DIF PROY vs PPTO
            f'<td style="{td}background:{bg_use};text-align:right;color:{val_color(dif_pp)}">{fmt_cop(dif_pp)}</td>',
            # 5. % CUMPL
            f'<td style="{td}background:{bg_use};text-align:right;color:{c_cumpl};font-weight:700">{fmt_pct(cumpl)}</td>',
            # 6. EJECUTADO
            f'<td style="{td}background:{bg_use};text-align:right;font-weight:700;color:#1e293b">{fmt_cop(r["EJECUTADO"])}</td>',
            # 7. M. ANT.
            f'<td style="{td}background:{bg_use};text-align:right;color:#64748b">{fmt_cop(r["VENTA_MES_ANT"])}</td>',
            # 8. EJ. vs M. ANT.
            f'<td style="{td}background:{bg_use};text-align:right;color:{val_color(dif_ant)}">{fmt_cop(dif_ant)}</td>',
            # 9. VIAJES
            f'<td style="{td}background:{bg_use};text-align:right;color:#374151">{int(r["VIAJES"])}</td>',
            # 10. M. VIAJES
            f'<td style="{td}background:{bg_use};text-align:right;color:#64748b">{int(r.get("M_VIAJES", 0))}</td>',
            # 11. META UTILIDAD
            f'<td style="{td}background:{bg_use};text-align:right;color:#64748b">{fmt_cop(r.get("META_UTIL", 0))}</td>',
            # 12. UTILIDAD
            f'<td style="{td}background:{bg_use};text-align:right;color:{"#166534" if pct_m>=0.12 else "#92400e"}">{fmt_cop(r["UTILIDAD"])}</td>',
            # 13. PROY. UTILIDAD
            f'<td style="{td}background:{bg_use};text-align:right;color:#1e40af">{fmt_cop(r.get("PROY_UTILIDAD", 0))}</td>',
            # 14. % INTER M
            f'<td style="{td}background:{bg_use};text-align:right;color:#64748b">{fmt_pct(pct_im)}</td>',
            # 15. P. PLANILLAR
            f'<td style="{td}background:{bg_use};text-align:right;color:{COL_NARANJA if r["P_PLANILLAR"]>0 else "#64748b"}">{fmt_cop(r["P_PLANILLAR"])}</td>',
        ]
        return f'<tr>{"".join(cells)}</tr>'

    def make_total_row(t):
        td_tot = td + f"background:{COL_TOTAL};color:#ffffff;font-weight:700;"
        cells = [
            f'<td style="{td_tot}text-align:left">TOTAL</td>',
            f'<td style="{td_tot}text-align:right">{fmt_cop(t["PPTO"])}</td>',
            f'<td style="{td_tot}text-align:right">{fmt_cop(t["PROYECCION"])}</td>',
            f'<td style="{td_tot}text-align:right">{fmt_cop(t["DIF_PROY_PPTO"])}</td>',
            f'<td style="{td_tot}text-align:right">{fmt_pct(t["PCT_CUMPL"])}</td>',
            f'<td style="{td_tot}text-align:right">{fmt_cop(t["EJECUTADO"])}</td>',
            f'<td style="{td_tot}text-align:right">{fmt_cop(t["VENTA_MES_ANT"])}</td>',
            f'<td style="{td_tot}text-align:right">{fmt_cop(t["DIF_DIAS"])}</td>',
            f'<td style="{td_tot}text-align:right">{int(t["VIAJES"])}</td>',
            f'<td style="{td_tot}text-align:right">{int(t["M_VIAJES"])}</td>',
            f'<td style="{td_tot}text-align:right">{fmt_cop(t["META_UTIL"])}</td>',
            f'<td style="{td_tot}text-align:right">{fmt_cop(t["UTILIDAD"])}</td>',
            f'<td style="{td_tot}text-align:right">{fmt_cop(t["PROY_UTILIDAD"])}</td>',
            f'<td style="{td_tot}text-align:right">-</td>',
            f'<td style="{td_tot}text-align:right">{fmt_cop(t["P_PLANILLAR"])}</td>',
        ]
        return f'<tr>{"".join(cells)}</tr>'

    rows_html = ""
    for _, r in df.iterrows():
        rows_html += make_row(r, "alt")

    # Fila OTROS
    if len(df_otros) > 0:
        otros_bg = "#f1f5f9"
        td_ot = td + f"background:{otros_bg};color:#64748b;font-style:italic;"
        n_otros = len(df_otros)
        guion = f'<td style="{td_ot}text-align:right">-</td>'
        rows_html += (
            f'<tr>'
            f'<td style="{td_ot}text-align:left">OTROS CLIENTES ({n_otros})</td>'
            f'<td style="{td_ot}text-align:right">{fmt_cop(df_otros["PPTO"].sum())}</td>'
            + guion  # PROYECCION
            + guion  # DIF PROY vs PPTO
            + guion  # % CUMPL
            + f'<td style="{td_ot}text-align:right">{fmt_cop(df_otros["EJECUTADO"].sum())}</td>'
            + f'<td style="{td_ot}text-align:right">{fmt_cop(df_otros["VENTA_MES_ANT"].sum())}</td>'
            + guion  # EJ. vs M. ANT.
            + guion  # VIAJES
            + f'<td style="{td_ot}text-align:right">{int(df_otros["M_VIAJES"].sum())}</td>'
            + f'<td style="{td_ot}text-align:right">{fmt_cop(df_otros["META_UTIL"].sum())}</td>'
            + f'<td style="{td_ot}text-align:right">{fmt_cop(df_otros["UTILIDAD"].sum())}</td>'
            + guion  # PROY. UTILIDAD
            + guion  # % INTER M
            + f'<td style="{td_ot}text-align:right">{fmt_cop(df_otros["P_PLANILLAR"].sum())}</td>'
            f'</tr>'
        )

    rows_html += make_total_row(totales)

    tabla_html = f"""
<table cellpadding="0" cellspacing="0" border="0"
       style="border-collapse:collapse;width:100%;min-width:900px;font-family:'Segoe UI',Arial,sans-serif;">
  <thead>{make_header()}</thead>
  <tbody>{rows_html}</tbody>
</table>
"""

    fecha_str = today.strftime("%A %d de %B de %Y").upper()
    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#f8fafc;font-family:'Segoe UI',Arial,sans-serif;">
<div style="max-width:1100px;margin:0 auto;background:#ffffff;">

  <!-- Header -->
  <div style="background:#1e3a5f;padding:20px 24px;">
    <div style="color:#ffffff;font-size:22px;font-weight:800;letter-spacing:.04em">TRACTOCAR</div>
    <div style="color:#93c5fd;font-size:13px;margin-top:4px">Venta Nacional · Corte día {d2} de {nombre_mes}</div>
    <div style="color:#64748b;font-size:11px;margin-top:2px">Generado: {fecha_str}</div>
  </div>

  <!-- KPI bar -->
  <div style="background:#0f2a45;padding:12px 24px;display:flex;gap:40px;flex-wrap:wrap;">
    <div>
      <div style="color:#93c5fd;font-size:9px;font-weight:700;letter-spacing:.1em">EJECUTADO (1–{d2})</div>
      <div style="color:#ffffff;font-size:20px;font-weight:800">{fmt_cop(totales["EJECUTADO"])}</div>
    </div>
    <div>
      <div style="color:#93c5fd;font-size:9px;font-weight:700;letter-spacing:.1em">PROYECCIÓN MES</div>
      <div style="color:#60a5fa;font-size:20px;font-weight:800">{fmt_cop(totales["PROYECCION"])}</div>
    </div>
    <div>
      <div style="color:#93c5fd;font-size:9px;font-weight:700;letter-spacing:.1em">PPTO MES</div>
      <div style="color:#94a3b8;font-size:20px;font-weight:800">{fmt_cop(totales["PPTO"])}</div>
    </div>
    <div>
      <div style="color:#93c5fd;font-size:9px;font-weight:700;letter-spacing:.1em">UTILIDAD</div>
      <div style="color:#4ade80;font-size:20px;font-weight:800">{fmt_cop(totales["UTILIDAD"])}</div>
    </div>
    <div>
      <div style="color:#93c5fd;font-size:9px;font-weight:700;letter-spacing:.1em">% MARGEN</div>
      <div style="color:#4ade80;font-size:20px;font-weight:800">{fmt_pct(totales["PCT_MARGEN"])}</div>
    </div>
    <div>
      <div style="color:#93c5fd;font-size:9px;font-weight:700;letter-spacing:.1em">P. PLANILLAR</div>
      <div style="color:#fb923c;font-size:20px;font-weight:800">{fmt_cop(totales["P_PLANILLAR"])}</div>
    </div>
  </div>

  <!-- Tabla -->
  <div style="padding:16px 16px 24px;overflow-x:auto;">
    {tabla_html}
  </div>

  <!-- Footer -->
  <div style="background:#f1f5f9;padding:12px 24px;border-top:1px solid #e2e8f0;">
    <div style="color:#64748b;font-size:10px;">
      Este reporte se genera automáticamente cada día. Datos acumulados del 1 al {d2} de {nombre_mes}.
      La proyección se calcula como: ejecutado ÷ {d2 - 1 + 1} días × {today.day and calendar.monthrange(today.year, today.month)[1]} días del mes.
    </div>
  </div>

</div>
</body>
</html>"""
    return html


def enviar_outlook(html_body, asunto, destinatarios):
    try:
        import win32com.client
    except ImportError:
        print("[ERROR] win32com no instalado. Ejecuta: pip install pywin32")
        return False

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = asunto
        mail.HTMLBody = html_body
        mail.To = "; ".join(destinatarios)
        mail.Send()
        print(f"  Correo enviado a: {', '.join(destinatarios)}")
        return True
    except Exception as e:
        print(f"[ERROR] No se pudo enviar el correo: {e}")
        return False


ARCHIVO_LOCK = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".correo_enviado")


def ya_enviado_hoy():
    if not os.path.exists(ARCHIVO_LOCK):
        return False
    try:
        with open(ARCHIVO_LOCK) as f:
            fecha = f.read().strip()
        return fecha == dt.date.today().isoformat()
    except Exception:
        return False


def marcar_enviado():
    with open(ARCHIVO_LOCK, "w") as f:
        f.write(dt.date.today().isoformat())


def main():
    today = dt.date.today()
    d2 = today.day - 1

    if d2 < 1:
        print(f"  Hoy es dia 1 del mes — no hay datos del dia anterior. No se envia correo.")
        return

    if ya_enviado_hoy():
        print(f"  Correo ya enviado hoy ({today}). Saltando.")
        return

    print("=" * 60)
    print(f"TRACTOCAR · Correo VENTA NACIONAL")
    print(f"  Fecha: {today}  |  Corte: dia 1 al {d2}")
    print("=" * 60)

    df, df_otros, totales, dias_mes, mes_actual, mes_anterior = build_tabla(d2)
    print(f"  Clientes en tabla: {len(df)} + {len(df_otros)} en OTROS")
    print(f"  Ejecutado: {fmt_cop(totales['EJECUTADO'])}  |  Margen: {fmt_pct(totales['PCT_MARGEN'])}")

    html = build_html(df, df_otros, totales, dias_mes, mes_actual, mes_anterior, d2, today)

    nombre_mes = dt.date(today.year, today.month, 1).strftime("%B %Y").upper()
    asunto = f"VENTA NACIONAL — Corte día {d2} {nombre_mes}"

    enviado = enviar_outlook(html, asunto, DESTINATARIOS)
    if enviado:
        marcar_enviado()
    else:
        # Guardar HTML como fallback para revision manual
        out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "correo_venta_nacional.html")
        with open(out, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  HTML guardado en: {out}")

    print("Listo.")


if __name__ == "__main__":
    main()
