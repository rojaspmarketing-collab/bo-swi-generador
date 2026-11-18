# -*- coding: utf-8 -*-
import os
import sys
import unittest
from decimal import Decimal

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from swi.procesador import procesar
from swi.escritores import generar_csv, generar_txt, generar_mdb
from swi.plantilla import generar_plantilla
from swi.tablas import FIELDS, TXT_LINE_LEN, WRI_UNK, PUB_UNK

CAB = ";".join(FIELDS)


def fila(**kw):
    return ";".join(str(kw.get(f, "")) for f in FIELDS)


def csv_bytes(*filas, sep=";"):
    txt = "\r\n".join([CAB] + list(filas))
    return txt.replace(";", sep).encode("utf-8-sig")


BASE = dict(SONG_TITLE="Canción Ñandú", PERFORMER_NAME="Artista", WRITER_LAST_NAME="Perez",
            WRITER_FIRST_NAME="Juan", ROLE="CA", ORI_PUBLISHER_NAME="MI EDITORIAL", ORI_TERRITORY_CODE="2WL")


class TestProcesador(unittest.TestCase):
    def test_archivo_como_el_original_con_problemas(self):
        """Punto y coma, % y PUBLISHER UNKNOWN en el territorio: se corrige solo."""
        datos = csv_bytes(
            fila(SONG_CODE="X1", **{**BASE, "COPYRIGHT_SHARE": "70%", "ORI_TERRITORY_CODE": "CO"}),
            fila(SONG_CODE="X1", SONG_TITLE=BASE["SONG_TITLE"], WRITER_LAST_NAME="WRITER UNKNOWN",
                 COPYRIGHT_SHARE="30%", ORI_PUBLISHER_NAME="PUBLISHER UNKNOWN",
                 ORI_TERRITORY_CODE="PUBLISHER UNKNOWN"),
        )
        res = procesar("x.csv", datos)
        self.assertEqual(res.errores, [])
        self.assertEqual(res.registros[0]["COPYRIGHT_SHARE"], Decimal("70.00"))
        self.assertEqual(res.registros[1]["ORI_TERRITORY_CODE"], "CO")
        self.assertTrue(res.correcciones)

    def test_completa_a_100(self):
        res = procesar("x.csv", csv_bytes(fila(SONG_CODE="X2", COPYRIGHT_SHARE="30", **BASE)))
        self.assertEqual(res.errores, [])
        self.assertEqual(len(res.registros), 2)
        nuevo = res.registros[1]
        self.assertEqual(nuevo["WRITER_LAST_NAME"], WRI_UNK)
        self.assertEqual(nuevo["ORI_PUBLISHER_NAME"], PUB_UNK)
        self.assertEqual(nuevo["COPYRIGHT_SHARE"], Decimal("70.00"))
        self.assertEqual(nuevo["ORI_TERRITORY_CODE"], "2WL")

    def test_sin_completar_da_error(self):
        res = procesar("x.csv", csv_bytes(fila(SONG_CODE="X3", COPYRIGHT_SHARE="50", **BASE)),
                       completar_100=False)
        self.assertTrue(any("suman 50" in e["msg"] for e in res.errores))

    def test_coedicion_no_suma_doble(self):
        a = fila(SONG_CODE="X4", COPYRIGHT_SHARE="100", **BASE)
        b = fila(SONG_CODE="X4", COPYRIGHT_SHARE="100", **{**BASE, "ORI_PUBLISHER_NAME": "OTRO EDITOR"})
        res = procesar("x.csv", csv_bytes(a, b))
        self.assertEqual(res.errores, [])

    def test_nombres_de_sociedad_y_pais(self):
        res = procesar("x.csv", csv_bytes(fila(
            SONG_CODE="X5", COPYRIGHT_SHARE="100", WRITER_PR_SOC_CODE="SAYCO",
            ADM_PUBLISHER_NAME="MI EDITORIAL", ADM_PR_COLLECT="50", ADM_MR_COLLECT="100",
            ADM_TERRITORY_CODE="COLOMBIA", ADM_PR_SOC_CODE="84 - SAYCO", **BASE)))
        self.assertEqual(res.errores, [])
        r = res.registros[0]
        self.assertEqual((r["WRITER_PR_SOC_CODE"], r["ADM_PR_SOC_CODE"], r["ADM_TERRITORY_CODE"]), ("84", "84", "CO"))

    def test_errores_bloqueantes(self):
        res = procesar("x.csv", csv_bytes(fila(
            SONG_CODE="X6", COPYRIGHT_SHARE="100", WRITER_PR_SOC_CODE="9999",
            ADM_PUBLISHER_NAME="MI EDITORIAL", ADM_PR_COLLECT="50", ADM_TERRITORY_CODE="2WL", **BASE)))
        campos = {e["campo"] for e in res.errores}
        self.assertIn("WRITER_PR_SOC_CODE", campos)     # sociedad inexistente
        self.assertIn("ADM_MR_COLLECT", campos)         # falta collect
        self.assertIn("ADM_TERRITORY_CODE", campos)     # grupo en vez de pais

    def test_mi_editorial(self):
        cfg = {"nombre": "MI EDITORIAL", "cae": "00123456789", "pr_soc": "84", "territorio": "CO"}
        res = procesar("x.csv", csv_bytes(fila(
            SONG_CODE="X7", COPYRIGHT_SHARE="100", ADM_PUBLISHER_NAME="Mi Editorial",
            ADM_PR_COLLECT="50", ADM_MR_COLLECT="100", **BASE)), cfg=cfg)
        self.assertEqual(res.errores, [])
        r = res.registros[0]
        self.assertEqual((r["ORI_PUBLISHER_CAE"], r["ORI_PR_SOC_CODE"]), ("00123456789", "84"))
        self.assertEqual((r["ADM_TERRITORY_CODE"], r["ADM_PR_SOC_CODE"]), ("CO", "84"))

    def test_coma_decimal(self):
        a = fila(SONG_CODE="X8", COPYRIGHT_SHARE="33,33", **BASE)
        b = fila(SONG_CODE="X8", COPYRIGHT_SHARE="66,67%", **{**BASE, "WRITER_LAST_NAME": "Gomez"})
        res = procesar("x.csv", csv_bytes(a, b))
        self.assertEqual(res.errores, [])

    def test_csv_separado_por_comas(self):
        res = procesar("x.csv", csv_bytes(fila(SONG_CODE="X9", COPYRIGHT_SHARE="100", **BASE), sep=","))
        self.assertEqual(res.errores, [])
        self.assertEqual(res.registros[0]["SONG_TITLE"], "Canción Ñandú")


class TestFilasSobrantes(unittest.TestCase):
    def test_filas_vacias_al_final_y_en_medio(self):
        vacia = ";" * (len(FIELDS) - 1)
        datos = csv_bytes(fila(SONG_CODE="V1", COPYRIGHT_SHARE="100", **BASE), vacia, "  ;  ; ",
                          fila(SONG_CODE="V2", COPYRIGHT_SHARE="100", **BASE), vacia, vacia, "", "")
        res = procesar("x.csv", datos)
        self.assertEqual(res.errores, [])
        self.assertEqual([r["SONG_CODE"] for r in res.registros], ["V1", "V2"])
        txt = generar_csv(res.registros).decode("utf-8-sig")
        self.assertEqual(txt.count("\r\n"), 3)              # encabezado + 2 filas, nada más
        self.assertFalse(txt.endswith("\r\n\r\n"))

    def test_titulo_arriba_y_columnas_extra(self):
        cab = CAB + ";NOTAS INTERNAS"
        datos = ("REPORTE DE OBRAS\r\n\r\n" + cab + "\r\n" +
                 fila(SONG_CODE="V3", COPYRIGHT_SHARE="100", **BASE) + ";nota\r\n").encode("utf-8")
        res = procesar("x.csv", datos)
        self.assertEqual(res.errores, [])
        self.assertEqual(len(res.registros), 1)
        linea = generar_csv(res.registros).decode("utf-8-sig").split("\r\n")[1]
        self.assertEqual(len(linea.split(",")), len(FIELDS))
        self.assertNotIn("nota", linea)

    def test_fila_que_no_es_obra_se_bloquea(self):
        datos = csv_bytes(fila(SONG_CODE="V4", COPYRIGHT_SHARE="100", **BASE), "TOTAL;;;;;;;;;;;;;100")
        res = procesar("x.csv", datos)
        self.assertTrue(any(e["fila"] == 3 for e in res.errores))

    def test_filas_de_ejemplo_de_la_plantilla_se_ignoran(self):
        import io, openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(generar_plantilla({"nombre": "MI EDITORIAL"})))
        ws = wb["SWI"]
        for c, v in enumerate([fila(SONG_CODE="V5", COPYRIGHT_SHARE="100", **BASE).split(";")][0], start=1):
            ws.cell(row=4, column=c, value=v or None)
        buf = io.BytesIO(); wb.save(buf)
        res = procesar("mio.xlsx", buf.getvalue())
        self.assertEqual(res.errores, [])
        self.assertEqual([r["SONG_CODE"] for r in res.registros], ["V5"])


class TestFormatos(unittest.TestCase):
    def setUp(self):
        self.res = procesar("x.csv", csv_bytes(
            fila(SONG_CODE="F1", COPYRIGHT_SHARE="33.33", ADM_PUBLISHER_NAME="MI EDITORIAL",
                 ADM_PR_COLLECT="16.67", ADM_MR_COLLECT="33.33", ADM_TERRITORY_CODE="CO", **BASE)))
        self.assertEqual(self.res.errores, [])

    def test_txt(self):
        lineas = generar_txt(self.res.registros).decode("cp1252").split("\r\n")[:-1]
        self.assertEqual(len(lineas), 2)
        for l in lineas:
            self.assertEqual(len(l), TXT_LINE_LEN)
        l = lineas[0]
        self.assertEqual(l[712:717], "03333")          # COPYRIGHT_SHARE pos 713
        self.assertEqual(l[997:1002], "01667")         # ADM_PR_COLLECT pos 998
        self.assertEqual(l[1010:1013].strip(), "CO")   # ADM_TERRITORY_CODE pos 1011
        self.assertIn("Canción Ñandú", l)

    def test_csv(self):
        txt = generar_csv(self.res.registros).decode("utf-8-sig")
        self.assertTrue(txt.startswith("SONG_CODE,ISWC,"))
        self.assertIn(",33.33,", txt)
        self.assertIn(",66.67,", txt)          # WRITER UNKNOWN agregado
        self.assertIn(",16.67,", txt)
        self.assertNotIn("%", txt)

    def test_csv_siempre_dos_decimales(self):
        res = procesar("x.csv", csv_bytes(fila(SONG_CODE="D1", COPYRIGHT_SHARE="50", ADM_PUBLISHER_NAME="MI EDITORIAL",
                                               ADM_PR_COLLECT="25", ADM_MR_COLLECT="0", ADM_TERRITORY_CODE="CO", **BASE)))
        linea = generar_csv(res.registros).decode("utf-8-sig").split("\r\n")[1].split(",")
        from swi.tablas import FIELDS as F
        self.assertEqual(linea[F.index("COPYRIGHT_SHARE")], "50.00")
        self.assertEqual(linea[F.index("ADM_PR_COLLECT")], "25.00")
        self.assertEqual(linea[F.index("ADM_MR_COLLECT")], "0.00")

    @unittest.skipUnless(os.name == "nt", "MDB solo en Windows")
    def test_mdb(self):
        datos = generar_mdb(self.res.registros)
        self.assertTrue(datos[4:19] == b"Standard Jet DB")


class TestPlantilla(unittest.TestCase):
    def test_plantilla_sin_llenar_no_genera_nada(self):
        res = procesar("Plantilla SWI.xlsx", generar_plantilla({"nombre": "MI EDITORIAL"}))
        self.assertEqual(res.registros, [])
        self.assertTrue(any("no tiene obras" in e["msg"] for e in res.errores))

    def test_el_ejemplo_de_la_plantilla_es_valido(self):
        import io, openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(generar_plantilla({"nombre": "MI EDITORIAL", "territorio": "CO"})))
        for fila_ in (2, 3):
            wb["SWI"].cell(row=fila_, column=1, value="EJ0002")   # que no se ignore como ejemplo
        buf = io.BytesIO(); wb.save(buf)
        res = procesar("Plantilla SWI.xlsx", buf.getvalue())
        self.assertEqual(res.errores, [], res.errores)
        self.assertEqual(len(res.registros), 2)
        self.assertEqual(res.registros[0]["ROLE"], "CA")
        self.assertEqual(res.registros[0]["ORI_TERRITORY_CODE"], "2WL")


if __name__ == "__main__":
    unittest.main()
