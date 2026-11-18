# -*- coding: utf-8 -*-
"""Plantilla Excel amigable para llenar el SWI (con listas desplegables y ayudas)."""
import io

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .tablas import LAYOUT, ROLES, SOCIEDADES, PAISES, GRUPOS

AYUDA = {
    "SONG_CODE": "Código único de la obra (ej. DF019851). Todas las filas de la misma obra llevan el mismo código.",
    "ISWC": "Opcional. Ej: T0345246801",
    "SONG_TITLE": "Título original de la obra.",
    "PERFORMER_NAME": "Intérprete original.",
    "WRITER_LAST_NAME": "Apellido(s) del autor. Si no se conoce: WRITER UNKNOWN",
    "WRITER_CAE": "Código CAE/IPI del autor.",
    "WRITER_PR_SOC_CODE": "Sociedad del autor (derechos de ejecución). Elige de la lista, ej. 84 - SAYCO",
    "WRITER_MR_SOC_CODE": "Sociedad del autor (derechos mecánicos). Elige de la lista.",
    "ROLE": "Rol del autor. CA = compositor y autor.",
    "COPYRIGHT_SHARE": "% del autor sobre la obra. Solo el número (70, no 70%). La obra debe sumar 100.",
    "ORI_PUBLISHER_NAME": "Editor original. Si no se conoce: PUBLISHER UNKNOWN",
    "ORI_TERRITORY_CODE": "Territorio del editor original. Ej: 2WL - WORLD o CO - COLOMBIA",
    "ADM_PUBLISHER_NAME": "Editor local que administra (tú). Vacío si el editor original es PUBLISHER UNKNOWN.",
    "ADM_PR_COLLECT": "% de ejecución que cobra el editor local (0 a 100).",
    "ADM_MR_COLLECT": "% mecánico que cobra el editor local (0 a 100).",
    "ADM_TERRITORY_CODE": "UN solo país. Ej: CO - COLOMBIA",
}

COLOR_GRUPO = {
    "Obra": "1F4E79", "Intérprete": "2E75B6", "Autor": "548235",
    "Editor original": "BF8F00", "Editor local": "C55A11",
}


def generar_plantilla(cfg=None):
    cfg = cfg or {}
    wb = Workbook()

    # ---------------- hojas de referencia
    ws_soc = wb.active
    ws_soc.title = "Sociedades"
    ws_soc.append(["LISTA", "CODIGO", "SOCIEDAD", "PAIS"])
    for code, (nombre, pais) in sorted(SOCIEDADES.items(), key=lambda x: x[1][0]):
        ws_soc.append([f"{code} - {nombre}", int(code), nombre, pais])
    ws_ter = wb.create_sheet("Territorios")
    ws_ter.append(["LISTA", "CODIGO", "NOMBRE", "TIPO"])
    for code, nombre in sorted(PAISES.items(), key=lambda x: x[1]):
        ws_ter.append([f"{code} - {nombre}", code, nombre, "PAIS"])
    n_paises = len(PAISES) + 1
    for code, nombre in sorted(GRUPOS.items(), key=lambda x: x[1]):
        ws_ter.append([f"{code} - {nombre}", code, nombre, "GRUPO"])
    ws_rol = wb.create_sheet("Roles")
    ws_rol.append(["LISTA", "CODIGO", "ROL"])
    for code, nombre in ROLES.items():
        ws_rol.append([f"{code} - {nombre}", code, nombre])
    for ws in (ws_soc, ws_ter, ws_rol):
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="404040")
        ws.column_dimensions["A"].width = 34
        ws.column_dimensions["C"].width = 28
        ws.column_dimensions["D"].width = 18
        ws.freeze_panes = "A2"

    # ---------------- hoja principal
    ws = wb.create_sheet("SWI", 0)
    borde = Border(bottom=Side(style="thin", color="999999"))
    for col, (campo, _i, largo, tipo, req, grupo) in enumerate(LAYOUT, start=1):
        c = ws.cell(row=1, column=col, value=campo)
        c.font = Font(bold=True, color="FFFFFF", underline="single" if req == "Y" else None)
        c.fill = PatternFill("solid", fgColor=COLOR_GRUPO[grupo])
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = borde
        texto = f"{grupo} · " + ("OBLIGATORIO" if req == "Y" else
                                 "Obligatorio si administras al editor" if req == "Y*" else "Opcional")
        texto += f" · máx {largo}" if tipo == "T" else " · número 0-100"
        if campo in AYUDA:
            texto += "\n\n" + AYUDA[campo]
        com = Comment(texto, "BO SWI")
        com.width, com.height = 260, 120
        c.comment = com
        ws.column_dimensions[get_column_letter(col)].width = max(12, min(30, len(campo) + 2))
        if tipo == "P":
            for fila in range(2, 3001):
                ws.cell(row=fila, column=col).number_format = "0.00"
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(LAYOUT))}1"

    col = {campo: get_column_letter(i) for i, (campo, *_x) in enumerate(LAYOUT, start=1)}
    def lista(formula, campos, msg):
        dv = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=False)
        dv.prompt = msg
        dv.showInputMessage = True
        ws.add_data_validation(dv)
        for campo in campos:
            dv.add(f"{col[campo]}2:{col[campo]}3000")
    n_soc = len(SOCIEDADES) + 1
    n_ter = len(PAISES) + len(GRUPOS) + 1
    lista(f"=Sociedades!$A$2:$A${n_soc}",
          ["WRITER_PR_SOC_CODE", "WRITER_MR_SOC_CODE", "ORI_PR_SOC_CODE", "ORI_MR_SOC_CODE",
           "ADM_PR_SOC_CODE", "ADM_MR_SOC_CODE"], "Elige la sociedad o escribe su código")
    lista(f"=Territorios!$A$2:$A${n_ter}", ["ORI_TERRITORY_CODE"], "País o grupo (ej. 2WL - WORLD)")
    lista(f"=Territorios!$A$2:$A${n_paises}", ["ADM_TERRITORY_CODE"], "Un solo país (ej. CO - COLOMBIA)")
    lista(f"=Roles!$A$2:$A${len(ROLES) + 1}", ["ROLE"], "Rol del autor")
    pct = DataValidation(type="decimal", operator="between", formula1="0", formula2="100",
                         allow_blank=True, showErrorMessage=True,
                         error="Escribe solo el número entre 0 y 100 (sin %)", errorTitle="Porcentaje")
    ws.add_data_validation(pct)
    for campo in ("COPYRIGHT_SHARE", "ADM_PR_COLLECT", "ADM_MR_COLLECT"):
        pct.add(f"{col[campo]}2:{col[campo]}3000")

    # ---------------- ejemplo (ficticio) para guiar
    ej = ws.cell(row=2, column=1)
    ej.comment = Comment("Las filas 2-3 son un EJEMPLO. Bórralas y escribe tus obras.", "BO SWI")
    mi = (cfg.get("nombre") or "MI EDITORIAL").upper()
    terr = cfg.get("territorio") or "CO"
    filas = [
        {"SONG_CODE": "EJ0001", "SONG_TITLE": "CANCION DE EJEMPLO", "PERFORMER_NAME": "ARTISTA",
         "WRITER_LAST_NAME": "PEREZ", "WRITER_FIRST_NAME": "JUAN", "ROLE": "CA - Composer/Author",
         "COPYRIGHT_SHARE": 60, "ORI_PUBLISHER_NAME": mi, "ORI_TERRITORY_CODE": "2WL - WORLD",
         "ADM_PUBLISHER_NAME": mi, "ADM_PR_COLLECT": 30, "ADM_MR_COLLECT": 60,
         "ADM_TERRITORY_CODE": f"{terr} - {PAISES.get(terr, '')}"},
        {"SONG_CODE": "EJ0001", "SONG_TITLE": "CANCION DE EJEMPLO", "PERFORMER_NAME": "ARTISTA",
         "WRITER_LAST_NAME": "WRITER UNKNOWN", "COPYRIGHT_SHARE": 40,
         "ORI_PUBLISHER_NAME": "PUBLISHER UNKNOWN", "ORI_TERRITORY_CODE": "2WL - WORLD"},
    ]
    gris = Font(italic=True, color="808080")
    for i, datos in enumerate(filas, start=2):
        for campo, v in datos.items():
            c = ws[f"{col[campo]}{i}"]
            c.value = v
            c.font = gris

    # ---------------- instrucciones
    wi = wb.create_sheet("Instrucciones", 1)
    pasos = [
        "CÓMO LLENAR ESTA PLANTILLA",
        "",
        "1. Una fila por cada autor de cada obra. Si una obra tiene 2 autores, son 2 filas con el mismo SONG_CODE.",
        "2. Los títulos de columna subrayados son obligatorios. Pasa el mouse por el título para ver la ayuda.",
        "3. Porcentajes: solo el número (70 o 33.33), nunca con el símbolo %.",
        "4. Los autores de cada obra deben sumar 100. Si falta un autor, el programa agrega WRITER UNKNOWN solo.",
        "5. Sociedades, territorios y roles: usa la lista desplegable (o escribe el código, ej. 84).",
        "6. Si el editor original es PUBLISHER UNKNOWN, deja vacías las columnas ADM_ (editor local).",
        "7. Si tú administras la obra, llena ADM_PUBLISHER_NAME, ADM_PR_COLLECT, ADM_MR_COLLECT y ADM_TERRITORY_CODE.",
        "8. Coediciones: repite la fila del autor (mismo %) una vez por cada editor original.",
        "",
        "Cuando termines, guarda el Excel normal (.xlsx) y arrástralo al programa BO SWI Generador.",
        "El programa revisa todo y te entrega el archivo en Access (.mdb), TXT y CSV listos para BO.",
    ]
    for p in pasos:
        wi.append([p])
    wi["A1"].font = Font(bold=True, size=14)
    wi.column_dimensions["A"].width = 120

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
