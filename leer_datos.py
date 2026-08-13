# -*- coding: utf-8 -*-
"""
Lector de datos de ventas TRACTOCAR — version local para venta-tractocar.

Diferencia clave vs procesar.py de tractocar-ventas:
  limpiar_nacional() NO filtra a solo filas TL; retiene TODO el domestico (TN, AC, etc.)
  para que Fuente="NACIONAL" llegue correctamente al dashboard.
"""

import os, sys, glob, unicodedata, warnings
import pandas as pd
import numpy as np

warnings.filterwarnings("ignore")

# ====================== CARPETAS DE DATOS ======================
CARPETA  = r"C:\Users\jarias\OneDrive - TRACTOCAR LOGISTICS SAS\POWER BI JEFFER\ARCHIVOS"
CARPETA2 = r"C:\Users\jarias\OneDrive - TRACTOCAR LOGISTICS SAS\POWER BI JEFFER\ACHIVOS 2"
# ===============================================================

NIT_PROPIO = "9005033252"


# ---------------------------------------------------------------- utilidades

def _norm(s):
    s = str(s).replace("\xa0", " ").strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return " ".join(s.split())


def norm_nit(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.replace("-", "").replace(" ", "").replace(".", "")


def split_cuenta(cc):
    """De 'BTA-TN-0022' saca ciudad='BTA', token='TN'. De 'CEDI-BMED-0043' -> token='CED'."""
    if pd.isna(cc) or cc == "":
        return ("", "")
    p = str(cc).strip().split("-")
    if len(p) >= 2:
        if p[0].upper().startswith("CED"):
            return (p[0], "CED")
        return (p[0], p[1].upper())
    return (str(cc), "")


def seg_from(fuente, token):
    if fuente == "IMPO":
        return "Comex Impo"
    if fuente == "EXPO":
        return "Comex Expo"
    if token == "AC":
        return "Alto Cubicaje"
    if token == "CED":
        return "Cedis"
    if token == "TN":
        return "Nacional"
    return f"Nacional ({token})" if token else "Nacional"


class Cols:
    """Resuelve columnas por nombre flexible (ignora tildes/mayusculas/espacios y acepta alias)."""
    def __init__(self, df, tipo, archivo, avisos):
        self.df, self.tipo, self.archivo, self.avisos = df, tipo, archivo, avisos
        self.map = {}
        for c in df.columns:
            self.map.setdefault(_norm(c), c)

    def serie(self, *alias, requerido=False, defecto=pd.NA, aviso=None):
        for a in alias:
            n = _norm(a)
            if n in self.map:
                return self.df[self.map[n]]
        if requerido:
            cols = "\n  ".join(repr(str(c)) for c in self.df.columns)
            raise SystemExit(
                f"\n[ERROR] En '{self.archivo}' ({self.tipo}) no encontre la columna [{alias[0]}].\n"
                f"Columnas disponibles:\n  {cols}")
        if aviso:
            self.avisos.append(f"{self.tipo} ({self.archivo}): {aviso}")
        return pd.Series([defecto] * len(self.df), index=self.df.index, dtype="object")


# ---------------------------------------------------------------- lectura
SIG = {"orden base": "COMEX", "contable (man)": "CEDIS", "cuenta contable": "NACIONAL"}


def leer_archivo(path):
    """Detecta tipo de archivo (COMEX/NACIONAL/CEDIS) buscando columna signature."""
    try:
        xls = pd.ExcelFile(path)
    except Exception:
        return None, None
    for sheet in xls.sheet_names:
        try:
            raw = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=12)
        except Exception:
            continue
        hrow = None
        for i in range(len(raw)):
            vals = {_norm(x) for x in raw.iloc[i].tolist() if pd.notna(x)}
            if vals & set(SIG):
                hrow = i
                break
        if hrow is None:
            continue
        df = pd.read_excel(xls, sheet_name=sheet, header=hrow)
        df.columns = [str(c).strip() for c in df.columns]
        nc = {_norm(c) for c in df.columns}
        for sig, tipo in SIG.items():
            if sig in nc:
                return tipo, df
    return None, None


# ---------------------------------------------------------------- limpieza

def limpiar_comex(df, stats, archivo):
    nmap = {_norm(c): c for c in df.columns}
    n0 = len(df)
    for c in ["Orden Borrada?", "Venta Borrada?", "Compra Borrada?", "Orden Borrada", "Venta Borrada"]:
        if _norm(c) in nmap:
            df = df[pd.to_numeric(df[nmap[_norm(c)]], errors="coerce").fillna(0) != 1]
    stats["borradas"] += n0 - len(df)

    C = Cols(df, "COMEX", archivo, stats["avisos"])
    fecha = pd.to_datetime(C.serie("Fecha Creacion", "Fecha Creacion", "Fecha de Creacion", "Fecha",
                                   aviso="no encontre columna de fecha"), errors="coerce")
    af = pd.to_numeric(C.serie("A Facturar($)", "A Facturar ($)", "A Facturar", "Valor a Facturar",
                               requerido=True).astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
    ap = pd.to_numeric(C.serie("A Pagar($)", "A Pagar ($)", "A Pagar", "Valor a Pagar",
                               requerido=True).astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
    oper = C.serie("Operacion", "Operacion").astype("string")
    fuente = np.where(oper.astype(str).str.upper().str.contains("EXPO"), "EXPO", "IMPO")
    return pd.DataFrame({
        "Fuente": fuente, "EsComex": True, "Operacion": oper,
        "Segmento": np.where(fuente == "EXPO", "Comex Expo", "Comex Impo"),
        "OB": C.serie("Orden Base", requerido=True).astype("string"),
        "Manifiesto": C.serie("Envio(compra)", "Envio (compra)", "Envio(compra)", "Envio Compra").astype("string"),
        "Fecha": fecha.values,
        "ClienteNIT": C.serie("Cliente").astype("string"),
        "ClienteNombre": C.serie("Cliente Nombre").astype("string"),
        "ProvNIT": C.serie("Proveedor").astype("string"),
        "ProvNombre": C.serie("Proveedor Nombre").astype("string"),
        "Propiedad": np.where(C.serie("Proveedor").map(norm_nit) == NIT_PROPIO, "PROPIO", "TERCERO"),
        "Tipologia": C.serie("Tipologia", "Tipologia", defecto="(Sin tipologia)").fillna("(Sin tipologia)").astype("string"),
        "CuentaContable": "", "Token": "TL", "Ciudad": C.serie("Agencia").astype("string"),
        "AltoCubicaje": False,
        "CodCliente": pd.Series([""] * len(df), dtype="string"),
        "Placa": C.serie("Placa").astype("string"),
        "Origen": C.serie("Ciudad Origen", "Origen").astype("string"),
        "Destino": C.serie("Ciudad Destino", "Destino").astype("string"),
        "AFacturar": af.values, "APagar": ap.values,
    })


def limpiar_nacional(df, stats, archivo):
    # NOTA: NO se filtra por token TL — se retienen TODOS los registros domesticos (TN, AC, etc.)
    # para que Fuente="NACIONAL" llegue al dashboard de ventas nacional.
    C = Cols(df, "NACIONAL", archivo, stats["avisos"])
    fecha = pd.to_datetime(C.serie("Fecha Creacion", "Fecha Creacion", "Fecha de Creacion", "Fecha",
                                   aviso="no encontre columna de fecha"), errors="coerce")
    af = pd.to_numeric(C.serie("Envio venta (Total)", "Envio venta (Total)", "Envio Venta (Total)",
                               "Venta (Total)", "Total Venta", "Valor Venta", "A Facturar",
                               requerido=True), errors="coerce").fillna(0)
    ap = pd.to_numeric(C.serie("Envio compra (Total)", "Envio compra (Total)", "Compra (Total)",
                               "Total Compra", "A Pagar", requerido=True), errors="coerce").fillna(0)
    cuenta = C.serie("Cuenta Contable", "Cuenta contable", "Cuenta", requerido=True).astype("string")
    sc = cuenta.map(split_cuenta)
    ciudad = sc.map(lambda t: t[0])
    token  = sc.map(lambda t: t[1])
    # CED -> CEDIS, TL -> NAL-TL, resto (TN, AC, etc.) -> NACIONAL
    fuente_col = np.where(token == "CED", "CEDIS",
                 np.where(token == "TL",  "NAL-TL", "NACIONAL"))
    oper_col = fuente_col
    es_comex = (token == "TL")
    return pd.DataFrame({
        "Fuente": fuente_col, "EsComex": es_comex, "Operacion": oper_col,
        "Segmento": [seg_from("NACIONAL", t) for t in token],
        "OB": pd.NA,
        "Manifiesto": C.serie("Envio", "Envio", "Manifiesto").astype("string"),
        "Fecha": fecha.values,
        "ClienteNIT": C.serie("Doc Cliente", "Documento Cliente", "NIT Cliente").astype("string"),
        "ClienteNombre": C.serie("Cliente", "Nombre Cliente").astype("string"),
        "ProvNIT": C.serie("Doc Afiliado", "Documento Afiliado", "NIT Afiliado").astype("string"),
        "ProvNombre": C.serie("Afiliado", "Nombre Afiliado").astype("string"),
        "Propiedad": np.where(C.serie("Doc Afiliado", "Documento Afiliado", "NIT Afiliado").map(norm_nit) == NIT_PROPIO, "PROPIO", "TERCERO"),
        "Tipologia": C.serie("Tipologia", "Tipologia", defecto="(Sin tipologia)").fillna("(Sin tipologia)").astype("string"),
        "CuentaContable": cuenta, "Token": token.astype("string"), "Ciudad": ciudad.astype("string"),
        "AltoCubicaje": (token == "AC").values,
        "CodCliente": C.serie("Cod Cliente", "Codigo Cliente", "Codigo Cliente", defecto="").fillna("").astype("string"),
        "Placa": C.serie("Placa").astype("string"),
        "Origen": C.serie("Ciudad Origen", "Origen").astype("string"),
        "Destino": C.serie("Ciudad Destino", "Destino").astype("string"),
        "AFacturar": af.values, "APagar": ap.values,
    })


def limpiar_cedis(df, stats, archivo):
    C = Cols(df, "CEDIS", archivo, stats["avisos"])
    fecha = pd.to_datetime(C.serie("Creacion (Man)", "Creacion (Man)", "Fecha Creacion", "Fecha (Man)", "Fecha",
                                   aviso="no encontre columna de fecha"), errors="coerce")
    af = pd.to_numeric(C.serie("Costo Total (Venta)", "Costo total (Venta)", "Total (Venta)", "Valor Venta",
                               "A Facturar", "Envio venta (Total)", requerido=True), errors="coerce").fillna(0)
    ap = pd.to_numeric(C.serie("Costo Total (Man)", "Costo total (Man)", "Total (Man)", "A Pagar",
                               "Envio compra (Total)", requerido=True), errors="coerce").fillna(0)
    cont = C.serie("Contable (Man)", "Contable", "Cuenta Contable", requerido=True).astype("string")
    sc = cont.map(split_cuenta)
    ciudad = sc.map(lambda t: t[0])
    token  = sc.map(lambda t: t[1])
    return pd.DataFrame({
        "Fuente": "CEDIS", "EsComex": False, "Operacion": "CEDIS", "Segmento": "Cedis",
        "OB": pd.NA,
        "Manifiesto": C.serie("Manifiesto", "Envio").astype("string"),
        "Fecha": fecha.values,
        "ClienteNIT": C.serie("NIT (Cliente)", "Nit (Cliente)").astype("string"),
        "ClienteNombre": C.serie("Nombre (Cliente)").astype("string"),
        "ProvNIT": C.serie("Doc Afiliado").astype("string"),
        "ProvNombre": C.serie("Nom Afiliado", "Nombre Afiliado", "Afiliado").astype("string"),
        "Propiedad": np.where(C.serie("Doc Afiliado").map(norm_nit) == NIT_PROPIO, "PROPIO", "TERCERO"),
        "Tipologia": "(Sin tipologia)",
        "CuentaContable": cont, "Token": token.astype("string"), "Ciudad": ciudad.astype("string"),
        "AltoCubicaje": False,
        "CodCliente": pd.Series([""] * len(df), dtype="string"),
        "Placa": C.serie("Placa (Veh)", "Placa").astype("string"),
        "Origen": C.serie("Origen (Man)", "Origen", "Ciudad Origen").astype("string"),
        "Destino": C.serie("Destino (Man)", "Destino", "Ciudad Destino").astype("string"),
        "AFacturar": af.values, "APagar": ap.values,
    })


LIMPIADORES = {"COMEX": limpiar_comex, "NACIONAL": limpiar_nacional, "CEDIS": limpiar_cedis}


# ---------------------------------------------------------------- union principal

def obtener_union(verbose=True):
    """Lee todos los archivos de CARPETA y CARPETA2 y retorna el DataFrame unificado."""
    carps = [c for c in [CARPETA, CARPETA2] if c and os.path.isdir(c)]
    if not carps:
        raise RuntimeError(f"No encuentro las carpetas de datos. Verifica CARPETA en leer_datos.py")

    archivos = []
    for carp in carps:
        archivos += [f for f in glob.glob(os.path.join(carp, "**", "*.xls*"), recursive=True)
                     if not os.path.basename(f).startswith("~$")]
    archivos = sorted(set(archivos))
    if not archivos:
        raise RuntimeError("No encontre archivos .xlsx en ninguna carpeta.")

    stats = {"borradas": 0, "avisos": []}
    frames = []
    carp0 = carps[0]

    for path in sorted(archivos):
        try:
            tipo, df = leer_archivo(path)
        except Exception as e:
            if verbose:
                print(f"  · {os.path.relpath(path, carp0)}: omitido ({e})")
            continue
        if tipo not in LIMPIADORES:
            continue
        # Si viene de carpeta cedis pero se detecta como NACIONAL, forzar CEDIS
        if tipo == "NACIONAL" and "cedis" in path.lower():
            tipo = "CEDIS"
        rel = os.path.relpath(path, carp0)
        limpio = LIMPIADORES[tipo](df, stats, rel)
        frames.append(limpio)
        if verbose:
            print(f"  · {rel:44s} -> {tipo:8s} {len(limpio):>6,} filas")

    if not frames:
        raise RuntimeError("Ningun archivo reconocido (COMEX/NACIONAL/CEDIS).")

    U = pd.concat(frames, ignore_index=True)
    U["Utilidad"] = U["AFacturar"] - U["APagar"]
    U["Margen"]   = np.where(U["AFacturar"] != 0, U["Utilidad"] / U["AFacturar"], 0.0)
    U["Mes"]      = U["Fecha"].dt.strftime("%Y-%m").fillna("(sin fecha)")
    U["Anio"]     = U["Fecha"].dt.year
    U["Dia"]      = U["Fecha"].dt.day
    U["FechaISO"] = U["Fecha"].dt.strftime("%Y-%m-%d")
    return U
