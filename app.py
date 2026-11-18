# -*- coding: utf-8 -*-
"""
BO SWI Generador - aplicacion local.
Abre una pagina en el navegador donde arrastras tu Excel y obtienes el SWI en Access, TXT y CSV.
Todo corre en tu computador (127.0.0.1); ningun archivo sale a internet.
"""
import io
import json
import os
import re
import socket
import sys
import threading
import traceback
import uuid
import webbrowser
import zipfile
from collections import OrderedDict
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import quote, unquote

from swi import __version__
from swi.procesador import procesar, fmt_pct, suma_share
from swi.escritores import generar_csv, generar_txt, generar_mdb
from swi.plantilla import generar_plantilla
from swi.tablas import SOCIEDADES, PAISES, GRUPOS, ROLES, WRI_UNK

BASE = os.path.dirname(os.path.abspath(__file__))
WEB = os.path.join(BASE, "web")
SALIDA = os.path.join(BASE, "salida")
CONFIG = os.path.join(BASE, "config.json")
PLANTILLA_OFICIAL = os.path.join(BASE, "plantillas", "SWI_oficial.mdb")
PUERTO = 8765
MAX_SUBIDA = 30 * 1024 * 1024

TRABAJOS = {}


def leer_config():
    try:
        with open(CONFIG, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {"nombre": "", "cae": "", "pr_soc": "", "mr_soc": "", "territorio": "CO", "completar_100": True}


def guardar_config(cfg):
    limpio = {k: str(cfg.get(k, "") or "").strip() for k in ("nombre", "cae", "pr_soc", "mr_soc", "territorio")}
    limpio["completar_100"] = bool(cfg.get("completar_100", True))
    with open(CONFIG, "w", encoding="utf-8") as fh:
        json.dump(limpio, fh, ensure_ascii=False, indent=2)
    return limpio


def nombre_base(nombre):
    base = os.path.splitext(os.path.basename(nombre))[0]
    base = re.sub(r"\s*\((comas|punto y coma)\)\s*$", "", base, flags=re.I)
    base = re.sub(r'[<>:"/\\|?*]+', " ", base).strip() or "SWI"
    return base[:80]


def resumen_obras(registros, errores):
    filas_con_error = {e["fila"] for e in errores if e["fila"]}
    obras = OrderedDict()
    for r in registros:
        o = obras.setdefault(r["SONG_CODE"], {"codigo": r["SONG_CODE"], "titulo": r["SONG_TITLE"],
                                              "interprete": r["PERFORMER_NAME"], "autores": [],
                                              "total": "0", "error": False, "agregado": False})
        nombre = (r["WRITER_LAST_NAME"] + (", " + r["WRITER_FIRST_NAME"] if r["WRITER_FIRST_NAME"] else ""))
        o["autores"].append({"nombre": nombre, "pct": fmt_pct(r["COPYRIGHT_SHARE"]),
                             "editor": r["ORI_PUBLISHER_NAME"], "nuevo": bool(r.get("_nuevo"))})
        o.setdefault("_rs", []).append(r)
        if r.get("_nuevo"):
            o["agregado"] = True
        if r["_fila"] in filas_con_error and not r.get("_nuevo"):
            o["error"] = True
    for o in obras.values():
        o["total"] = fmt_pct(suma_share(o.pop("_rs")))
    return list(obras.values())


def procesar_subida(nombre, contenido, opciones):
    cfg = leer_config()
    completar = opciones.get("completar_100", cfg.get("completar_100", True))
    res = procesar(nombre, contenido, cfg=cfg, completar_100=completar)
    obras = resumen_obras(res.registros, res.errores)
    salida = {
        "ok": not res.errores,
        "archivo": nombre,
        "filas": len(res.registros),
        "obras": obras,
        "errores": res.errores,
        "avisos": res.avisos,
        "correcciones": res.correcciones,
    }
    if res.errores:
        return salida

    base = nombre_base(nombre)
    archivos = OrderedDict()
    archivos[base + ".csv"] = generar_csv(res.registros)
    archivos[base + ".txt"] = generar_txt(res.registros)
    mdb_error = None
    try:
        plantilla = PLANTILLA_OFICIAL if os.path.exists(PLANTILLA_OFICIAL) else None
        archivos[base + ".mdb"] = generar_mdb(res.registros, plantilla)
    except Exception as e:
        mdb_error = str(e)
    reporte = ["BO SWI Generador " + __version__, f"Archivo: {nombre}", f"Fecha: {datetime.now():%Y-%m-%d %H:%M}",
               f"Obras: {len(obras)}  Filas: {len(res.registros)}", "", "CORRECCIONES AUTOMÁTICAS:"]
    reporte += [f"  Fila {c['fila']} {c['campo'] or ''}: {c['msg']}" for c in res.correcciones] or ["  (ninguna)"]
    reporte += ["", "AVISOS:"]
    reporte += [f"  Fila {a['fila']} {a['campo'] or ''}: {a['msg']}" for a in res.avisos] or ["  (ninguno)"]
    archivos["reporte.txt"] = "\r\n".join(reporte).encode("utf-8-sig")

    carpeta = os.path.join(SALIDA, f"{datetime.now():%Y-%m-%d_%H%M%S} {base}")
    os.makedirs(carpeta, exist_ok=True)
    for n, datos in archivos.items():
        with open(os.path.join(carpeta, n), "wb") as fh:
            fh.write(datos)
    job = uuid.uuid4().hex[:12]
    TRABAJOS[job] = {"carpeta": carpeta, "archivos": archivos, "base": base}
    salida.update({
        "job": job,
        "carpeta": carpeta,
        "descargas": [{"nombre": n, "tam": len(d)} for n, d in archivos.items()],
        "mdb_error": mdb_error,
        "plantilla_oficial": os.path.exists(PLANTILLA_OFICIAL),
    })
    return salida


class Handler(BaseHTTPRequestHandler):
    server_version = "BOSWI/" + __version__

    def log_message(self, fmt, *args):
        pass

    def _json(self, datos, codigo=200):
        cuerpo = json.dumps(datos, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(codigo)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(cuerpo)))
        self.end_headers()
        self.wfile.write(cuerpo)

    def _archivo(self, datos, nombre, tipo="application/octet-stream"):
        self.send_response(200)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(datos)))
        self.send_header("Content-Disposition",
                         "attachment; filename*=UTF-8''" + quote(nombre))
        self.end_headers()
        self.wfile.write(datos)

    def _leer_cuerpo(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n > MAX_SUBIDA:
            raise ValueError("El archivo es demasiado grande")
        return self.rfile.read(n)

    def _mismo_origen(self):
        origen = self.headers.get("Origin")
        return origen is None or origen.startswith(("http://127.0.0.1:", "http://localhost:"))

    def do_GET(self):
        try:
            ruta = unquote(self.path.split("?")[0])
            if ruta in ("/", "/index.html"):
                with open(os.path.join(WEB, "index.html"), "rb") as fh:
                    datos = fh.read()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(datos)))
                self.end_headers()
                self.wfile.write(datos)
            elif ruta == "/api/config":
                cfg = leer_config()
                cfg["version"] = __version__
                cfg["plantilla_oficial"] = os.path.exists(PLANTILLA_OFICIAL)
                self._json(cfg)
            elif ruta == "/api/tablas":
                self._json({
                    "sociedades": [{"codigo": c, "nombre": n, "pais": p} for c, (n, p) in SOCIEDADES.items()],
                    "territorios": [{"codigo": c, "nombre": n, "tipo": "País"} for c, n in PAISES.items()]
                                   + [{"codigo": c, "nombre": n, "tipo": "Grupo"} for c, n in GRUPOS.items()],
                    "roles": [{"codigo": c, "nombre": n} for c, n in ROLES.items()],
                })
            elif ruta == "/api/plantilla":
                self._archivo(generar_plantilla(leer_config()), "Plantilla SWI.xlsx",
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            elif ruta.startswith("/api/descargar/"):
                _, _, _, job, nombre = ruta.split("/", 4)
                t = TRABAJOS.get(job)
                if not t:
                    return self._json({"error": "Ese resultado ya no existe; vuelve a procesar el archivo"}, 404)
                if nombre == "todo.zip":
                    buf = io.BytesIO()
                    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                        for n, d in t["archivos"].items():
                            z.writestr(n, d)
                    return self._archivo(buf.getvalue(), t["base"] + ".zip", "application/zip")
                if nombre not in t["archivos"]:
                    return self._json({"error": "Archivo no encontrado"}, 404)
                self._archivo(t["archivos"][nombre], nombre)
            else:
                self._json({"error": "No encontrado"}, 404)
        except Exception as e:
            traceback.print_exc()
            self._json({"error": str(e)}, 500)

    def do_POST(self):
        if not self._mismo_origen():
            return self._json({"error": "Origen no permitido"}, 403)
        try:
            ruta = self.path.split("?")[0]
            if ruta == "/api/procesar":
                nombre = unquote(self.headers.get("X-Filename") or "archivo.xlsx")
                opciones = json.loads(unquote(self.headers.get("X-Opciones") or "{}"))
                self._json(procesar_subida(nombre, self._leer_cuerpo(), opciones))
            elif ruta == "/api/config":
                self._json(guardar_config(json.loads(self._leer_cuerpo() or b"{}")))
            elif ruta.startswith("/api/abrir/"):
                t = TRABAJOS.get(ruta.rsplit("/", 1)[-1])
                if t:
                    os.startfile(t["carpeta"])
                self._json({"ok": bool(t)})
            else:
                self._json({"error": "No encontrado"}, 404)
        except Exception as e:
            traceback.print_exc()
            self._json({"error": str(e)}, 500)


def puerto_libre(preferido):
    for p in [preferido] + list(range(preferido + 1, preferido + 20)):
        with socket.socket() as s:
            try:
                s.bind(("127.0.0.1", p))
                return p
            except OSError:
                continue
    return 0


def main():
    os.makedirs(SALIDA, exist_ok=True)
    puerto = puerto_libre(PUERTO)
    srv = ThreadingHTTPServer(("127.0.0.1", puerto), Handler)
    url = f"http://127.0.0.1:{srv.server_port}/"
    print("=" * 56)
    print(f"  BO SWI Generador {__version__}")
    print(f"  Abierto en: {url}")
    print("  No cierres esta ventana mientras uses el programa.")
    print("=" * 56)
    if "--no-browser" not in sys.argv:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
