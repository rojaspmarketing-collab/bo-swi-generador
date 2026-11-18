# -*- coding: utf-8 -*-
"""Lee el Excel/CSV del usuario, normaliza, corrige lo corregible y valida contra la guia SWI 1.4."""
import csv
import io
import os
import re
from collections import defaultdict, OrderedDict
from decimal import Decimal, InvalidOperation

from .tablas import (FIELDS, SIZE, PCT, ADM_FIELDS, SOC_FIELDS, PUB_UNK, WRI_UNK, ROLES,
                     SOCIEDADES, PAISES, TERRITORIOS, SOC_POR_NOMBRE, TERR_POR_NOMBRE)

EJEMPLO_CODE = "EJ0001"               # filas de ejemplo que trae la plantilla
EJEMPLO_TITULO = "CANCION DE EJEMPLO"
CIEN = Decimal("100.00")
CERO = Decimal("0.00")


class Resultado:
    def __init__(self):
        self.registros = []
        self.errores = []       # dict(fila, campo, msg)
        self.avisos = []
        self.correcciones = []

    def err(self, fila, campo, msg):
        self.errores.append({"fila": fila, "campo": campo, "msg": msg})

    def aviso(self, fila, campo, msg):
        self.avisos.append({"fila": fila, "campo": campo, "msg": msg})

    def corr(self, fila, campo, msg):
        self.correcciones.append({"fila": fila, "campo": campo, "msg": msg})


# ------------------------------------------------------------------ lectura
def _celda_excel(v, formato=""):
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        if "%" in (formato or ""):
            v = v * 100
        if isinstance(v, float):
            v = round(v, 6)
            if v.is_integer():
                v = int(v)
        return str(v)
    return str(v)


def leer_archivo(nombre, contenido):
    """Devuelve lista de filas (listas de str). Acepta .xlsx/.xlsm/.xls/.csv/.txt(csv)."""
    ext = os.path.splitext(nombre)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(contenido), data_only=True)
        ws = wb["SWI"] if "SWI" in wb.sheetnames else wb.worksheets[0]
        filas = []
        for row in ws.iter_rows():
            filas.append([_celda_excel(c.value, c.number_format) for c in row])
        return filas
    if ext == ".xls":
        import xlrd
        wb = xlrd.open_workbook(file_contents=contenido, formatting_info=False)
        sh = wb.sheet_by_name("SWI") if "SWI" in wb.sheet_names() else wb.sheet_by_index(0)
        return [[_celda_excel(sh.cell_value(r, c)) for c in range(sh.ncols)] for r in range(sh.nrows)]
    # CSV / texto delimitado
    texto = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            texto = contenido.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    lineas = texto.splitlines()
    muestra = "\n".join(lineas[:5])
    sep = max([";", ",", "\t"], key=muestra.count)
    return list(csv.reader(io.StringIO(texto), delimiter=sep))


# ------------------------------------------------------------------ normalizacion
def _pct(valor):
    v = valor.strip().replace("%", "").replace(" ", "")
    if v == "":
        return None
    if "," in v and "." not in v:
        v = v.replace(",", ".")
    elif "," in v and "." in v:          # 1.234,56 -> 1234.56
        v = v.replace(".", "").replace(",", ".")
    try:
        return Decimal(v).quantize(Decimal("0.01"))
    except InvalidOperation:
        raise ValueError(f"'{valor}' no es un número")


def _sociedad(valor):
    v = valor.strip().upper()
    if not v:
        return ""
    m = re.fullmatch(r"(\d+)(\.0+)?", v)
    if m:
        return str(int(m.group(1)))
    m = re.match(r"^(\d+)\s*[-–(]", v) or re.search(r"\((\d+)\)\s*$", v)
    if m:
        return str(int(m.group(1)))
    return SOC_POR_NOMBRE.get(v, v)


def _territorio(valor):
    v = valor.strip().upper()
    if not v:
        return ""
    if v in TERRITORIOS:
        return v
    primero = re.split(r"\s*[-–]\s*|\s+", v)[0]
    if primero in TERRITORIOS:
        return primero
    return TERR_POR_NOMBRE.get(v, v)


def _rol(valor):
    v = valor.strip().upper()
    if not v:
        return ""
    primero = re.split(r"\s*[-–]\s*|\s+", v)[0]
    return primero


def _encontrar_encabezado(filas):
    for i, fila in enumerate(filas[:10]):
        cab = [re.sub(r"[\s\-]+", "_", c.strip().upper().lstrip("﻿")) for c in fila]
        if "SONG_CODE" in cab and "SONG_TITLE" in cab:
            return i, cab
    return None, None


def construir(filas, res):
    i_cab, cab = _encontrar_encabezado(filas)
    if cab is None:
        res.err(None, None, "No encontré el encabezado. La primera fila debe tener los nombres de "
                            "columna del layout (SONG_CODE, SONG_TITLE, ...). Descarga la plantilla.")
        return
    faltan = [f for f in FIELDS if f not in cab]
    if faltan:
        res.err(None, None, "Faltan columnas en el encabezado: " + ", ".join(faltan))
        return
    idx = {f: cab.index(f) for f in FIELDS}
    for n, fila in enumerate(filas[i_cab + 1:], start=i_cab + 2):
        if not any((c or "").strip() for c in fila):
            continue
        r = OrderedDict()
        for f in FIELDS:
            i = idx[f]
            v = (fila[i] if i < len(fila) else "") or ""
            v = re.sub(r"\s+", " ", v).strip()
            try:
                if f in PCT:
                    r[f] = _pct(v)
                elif f in SOC_FIELDS:
                    r[f] = _sociedad(v)
                elif f.endswith("TERRITORY_CODE"):
                    r[f] = _territorio(v)
                elif f == "ROLE":
                    r[f] = _rol(v)
                elif f.endswith("_CAE"):
                    r[f] = re.sub(r"\.0+$", "", v)
                else:
                    r[f] = v
            except ValueError as e:
                res.err(n, f, str(e))
                r[f] = None if f in PCT else v
        if r["SONG_CODE"].upper() == EJEMPLO_CODE and r["SONG_TITLE"].upper() == EJEMPLO_TITULO:
            res.corr(n, None, "Fila de ejemplo de la plantilla: se ignoró")
            continue
        if r["WRITER_LAST_NAME"].upper() == WRI_UNK:
            r["WRITER_LAST_NAME"] = WRI_UNK
        if r["ORI_PUBLISHER_NAME"].upper() == PUB_UNK:
            r["ORI_PUBLISHER_NAME"] = PUB_UNK
        r["_fila"] = n
        res.registros.append(r)


# ------------------------------------------------------------------ correcciones
def _es_unk_pub(r):
    return r["ORI_PUBLISHER_NAME"] == PUB_UNK


def aplicar_mi_editorial(res, cfg):
    """Completa CAE / sociedades / territorio del editor propio cuando estan vacios."""
    nombre = (cfg.get("nombre") or "").strip().upper()
    if not nombre:
        return
    datos = {k: (cfg.get(k) or "").strip() for k in ("cae", "pr_soc", "mr_soc", "territorio")}
    for r in res.registros:
        if r["ORI_PUBLISHER_NAME"].upper() == nombre:
            for campo, clave in (("ORI_PUBLISHER_CAE", "cae"), ("ORI_PR_SOC_CODE", "pr_soc"),
                                 ("ORI_MR_SOC_CODE", "mr_soc")):
                if not r[campo] and datos[clave]:
                    r[campo] = datos[clave]
                    res.corr(r["_fila"], campo, f"Completado con el dato de tu editorial ({datos[clave]})")
        if r["ADM_PUBLISHER_NAME"].upper() == nombre:
            for campo, clave in (("ADM_PUBLISHER_CAE", "cae"), ("ADM_PR_SOC_CODE", "pr_soc"),
                                 ("ADM_MR_SOC_CODE", "mr_soc"), ("ADM_TERRITORY_CODE", "territorio")):
                if not r[campo] and datos[clave]:
                    r[campo] = datos[clave]
                    res.corr(r["_fila"], campo, f"Completado con el dato de tu editorial ({datos[clave]})")


def corregir(res, completar_100=True):
    por_obra = OrderedDict()
    for r in res.registros:
        por_obra.setdefault(r["SONG_CODE"], []).append(r)

    nuevos = []
    for code, rs in por_obra.items():
        terr_ok = [r["ORI_TERRITORY_CODE"] for r in rs
                   if not _es_unk_pub(r) and r["ORI_TERRITORY_CODE"] in TERRITORIOS]
        terr = terr_ok[0] if terr_ok else ""
        for r in rs:
            if _es_unk_pub(r):
                if terr and r["ORI_TERRITORY_CODE"] != terr:
                    ant = r["ORI_TERRITORY_CODE"] or "(vacío)"
                    r["ORI_TERRITORY_CODE"] = terr
                    res.corr(r["_fila"], "ORI_TERRITORY_CODE",
                             f"PUBLISHER UNKNOWN debe llevar el mismo territorio que los demás editores: {ant} → {terr}")
                for f in ADM_FIELDS + ["ORI_PUBLISHER_CAE", "ORI_PR_SOC_CODE", "ORI_MR_SOC_CODE"]:
                    if r[f] not in ("", None):
                        r[f] = None if f in PCT else ""
                        res.corr(r["_fila"], f, "Con PUBLISHER UNKNOWN este campo debe ir vacío: se borró")
            if r["WRITER_LAST_NAME"] == WRI_UNK:
                for f in ("WRITER_PR_SOC_CODE", "WRITER_MR_SOC_CODE"):
                    if r[f]:
                        r[f] = ""
                        res.corr(r["_fila"], f, "Con WRITER UNKNOWN este campo debe ir vacío: se borró")

        if completar_100:
            total = suma_share(rs)
            if total is not None and CERO < total < CIEN:
                falta = CIEN - total
                base = rs[0]
                nuevo = OrderedDict((f, None if f in PCT else "") for f in FIELDS)
                for f in ("SONG_CODE", "ISWC", "SONG_TITLE", "SONG_SUBTITLE", "PERFORMER_NAME"):
                    nuevo[f] = base[f]
                nuevo["WRITER_LAST_NAME"] = WRI_UNK
                nuevo["COPYRIGHT_SHARE"] = falta
                nuevo["ORI_PUBLISHER_NAME"] = PUB_UNK
                nuevo["ORI_TERRITORY_CODE"] = terr
                nuevo["_fila"] = rs[-1]["_fila"]
                nuevo["_nuevo"] = True
                nuevos.append((rs[-1], nuevo))
                res.corr(rs[-1]["_fila"], "COPYRIGHT_SHARE",
                         f"Obra {code}: los autores sumaban {total}%. Se agregó WRITER UNKNOWN / "
                         f"PUBLISHER UNKNOWN con {falta}% para llegar a 100%")
    # insertar las filas nuevas justo despues de la ultima fila de su obra
    for despues_de, nuevo in nuevos:
        i = res.registros.index(despues_de)
        res.registros.insert(i + 1, nuevo)


def suma_share(rs):
    """Suma de COPYRIGHT_SHARE contando una sola vez a cada autor (en coediciones se repite)."""
    vistos, total = set(), CERO
    for r in rs:
        if r["COPYRIGHT_SHARE"] is None:
            continue
        clave = (r["WRITER_CODE"] or r["WRITER_LAST_NAME"].upper(), r["WRITER_FIRST_NAME"].upper(),
                 r["ROLE"], r["COPYRIGHT_SHARE"])
        if r["WRITER_LAST_NAME"] == WRI_UNK:
            clave += (id(r),)
        if clave in vistos:
            continue
        vistos.add(clave)
        total += r["COPYRIGHT_SHARE"]
    return total


# ------------------------------------------------------------------ validacion
def validar(res):
    por_obra = OrderedDict()
    for r in res.registros:
        por_obra.setdefault(r["SONG_CODE"], []).append(r)

    for r in res.registros:
        n = r["_fila"]
        for f in ("SONG_CODE", "SONG_TITLE", "WRITER_LAST_NAME", "ORI_PUBLISHER_NAME", "ORI_TERRITORY_CODE"):
            if not r[f]:
                res.err(n, f, "Campo obligatorio vacío")
        sh = r["COPYRIGHT_SHARE"]
        if sh is None:
            res.err(n, "COPYRIGHT_SHARE", "Campo obligatorio vacío")
        elif not (CERO < sh <= CIEN):
            res.err(n, "COPYRIGHT_SHARE", f"{sh} está fuera de rango (debe ser mayor que 0 y hasta 100)")

        for f in FIELDS:
            if f not in PCT and r[f] and len(r[f]) > SIZE[f]:
                res.err(n, f, f"Tiene {len(r[f])} caracteres; el máximo es {SIZE[f]}")
            if f not in PCT and r[f]:
                try:
                    r[f].encode("cp1252")
                except UnicodeEncodeError:
                    res.err(n, f, f"Tiene caracteres que no existen en ANSI: '{r[f]}'")

        if r["ROLE"] and r["ROLE"] not in ROLES:
            res.err(n, "ROLE", f"'{r['ROLE']}' no es un rol válido ({', '.join(ROLES)})")
        for f in SOC_FIELDS:
            if r[f] and r[f] not in SOCIEDADES:
                res.err(n, f, f"'{r[f]}' no está en la tabla de sociedades (usa el código, ej. 84 = SAYCO)")
        t = r["ORI_TERRITORY_CODE"]
        if t and t not in TERRITORIOS:
            res.err(n, "ORI_TERRITORY_CODE", f"'{t}' no es un código de territorio válido (ej. CO, 2WL)")

        tiene_adm = any(r[f] not in ("", None) for f in ADM_FIELDS)
        if tiene_adm:
            if not r["ADM_PUBLISHER_NAME"]:
                res.err(n, "ADM_PUBLISHER_NAME", "Hay datos de editor local pero falta su nombre")
            for f in ("ADM_PR_COLLECT", "ADM_MR_COLLECT"):
                if r[f] is None:
                    res.err(n, f, "Obligatorio cuando hay editor local (pon 0 si no cobra)")
                elif not (CERO <= r[f] <= CIEN):
                    res.err(n, f, f"{r[f]} fuera de rango (0 a 100)")
            ta = r["ADM_TERRITORY_CODE"]
            if not ta:
                res.err(n, "ADM_TERRITORY_CODE", "Obligatorio cuando hay editor local")
            elif ta not in PAISES:
                res.err(n, "ADM_TERRITORY_CODE", f"'{ta}' debe ser UN país (ej. CO), no un grupo de países")
            if r["COPYRIGHT_SHARE"] is not None:
                for f in ("ADM_PR_COLLECT", "ADM_MR_COLLECT"):
                    if r[f] is not None and r[f] > r["COPYRIGHT_SHARE"]:
                        res.aviso(n, f, f"El collect ({r[f]}) es mayor que el % del autor ({r['COPYRIGHT_SHARE']})")
        elif not _es_unk_pub(r):
            res.aviso(n, "ADM_PUBLISHER_NAME",
                      f"'{r['ORI_PUBLISHER_NAME']}' no tiene editor local. Si tú lo administras, llena las columnas ADM_")

        if r["WRITER_LAST_NAME"] != WRI_UNK and not r["WRITER_PR_SOC_CODE"]:
            res.aviso(n, "WRITER_PR_SOC_CODE", "Autor sin sociedad PR. Si está afiliado (ej. SAYCO = 84), ponla")

    for code, rs in por_obra.items():
        if not code:
            continue
        titulos = {r["SONG_TITLE"] for r in rs}
        if len(titulos) > 1:
            res.err(rs[0]["_fila"], "SONG_TITLE", f"La obra {code} tiene títulos distintos: " + " / ".join(titulos))
        total = suma_share(rs)
        if total != CIEN:
            res.err(rs[0]["_fila"], "COPYRIGHT_SHARE",
                    f"Obra {code} ({rs[0]['SONG_TITLE']}): los autores suman {total}%, debe ser 100%")
        for campo in ("ADM_PR_COLLECT", "ADM_MR_COLLECT"):
            suma = defaultdict(Decimal)
            for r in rs:
                if r[campo] is not None and r["ADM_TERRITORY_CODE"]:
                    suma[r["ADM_TERRITORY_CODE"]] += r[campo]
            for pais, s in suma.items():
                if s > CIEN:
                    res.err(rs[0]["_fila"], campo, f"Obra {code}: {campo} suma {s} en {pais} (máximo 100)")
        vistos = set()
        for r in rs:
            clave = tuple(r[f] for f in FIELDS)
            if clave in vistos:
                res.err(r["_fila"], None, "Fila duplicada (idéntica a otra de la misma obra)")
            vistos.add(clave)


# ------------------------------------------------------------------ API
def procesar(nombre, contenido, cfg=None, completar_100=True):
    res = Resultado()
    try:
        filas = leer_archivo(nombre, contenido)
    except Exception as e:  # archivo ilegible
        res.err(None, None, f"No pude leer el archivo: {e}")
        return res
    construir(filas, res)
    if res.errores:
        return res
    if not res.registros:
        res.err(None, None, "El archivo no tiene obras (solo encabezado, filas vacías o las filas de ejemplo)")
        return res
    aplicar_mi_editorial(res, cfg or {})
    corregir(res, completar_100=completar_100)
    validar(res)
    return res


def fmt_pct(d):
    if d is None:
        return ""
    s = f"{d:.2f}"
    return s[:-3] if s.endswith(".00") else s.rstrip("0")
