# -*- coding: utf-8 -*-
"""Genera los 3 formatos aceptados por BO: CSV UTF-8 (comas), TXT ANSI ancho fijo y Access MDB."""
import csv
import io
import os
import shutil
import subprocess
import tempfile

from .tablas import LAYOUT, FIELDS, PCT, TXT_LINE_LEN
from .procesador import fmt_pct


def _pct_csv(d):
    """Siempre 2 decimales con punto, como la guia ('100.00'). Nunca coma: chocaria con el separador."""
    return "" if d is None else f"{d:.2f}"


def _fila_texto(r):
    return [_pct_csv(r[f]) if f in PCT else (r[f] or "") for f in FIELDS]


def generar_csv(registros):
    """
    CSV UTF-8 (delimitado por comas) con BOM y CRLF, como 'CSV UTF-8' de Excel.
    El archivo termina justo en la ultima obra, SIN salto de linea final: asi ningun lector
    puede interpretar un renglon vacio al final (que BO reportaria como error de copyright).
    """
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=",", lineterminator="\r\n", quoting=csv.QUOTE_MINIMAL)
    w.writerow(FIELDS)
    for r in registros:
        w.writerow(_fila_texto(r))
    return buf.getvalue().rstrip("\r\n").encode("utf-8-sig")


def generar_txt(registros):
    """TXT de ancho fijo (1013 caracteres por linea), codificacion ANSI (Windows-1252), sin encabezado."""
    lineas = []
    for r in registros:
        partes = []
        for f, _ini, largo, tipo, _req, _g in LAYOUT:
            if tipo == "P":
                d = r[f]
                partes.append("00000" if d is None else f"{int((d * 100).to_integral_value()):05d}")
            else:
                partes.append((r[f] or "").ljust(largo)[:largo])
        linea = "".join(partes)
        assert len(linea) == TXT_LINE_LEN, len(linea)
        lineas.append(linea)
    return "\r\n".join(lineas).encode("cp1252")   # sin salto final: ningun renglon vacio


_PS_MDB = r'''
param([string]$Mdb, [string]$Csv, [string]$Crear, [string]$Cols, [string]$Pct)
$ErrorActionPreference = "Stop"
$cs = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$Mdb"
if ($Crear -eq "1") {
    $cat = New-Object -ComObject ADOX.Catalog
    $cat.Create($cs) | Out-Null
    $cat.ActiveConnection.Execute("CREATE TABLE SWI ($Cols)") | Out-Null
    $cat.ActiveConnection.Close()
}
$pctCols = $Pct.Split(",")
$cn = New-Object -ComObject ADODB.Connection
$cn.Open($cs)
$filas = Import-Csv -Path $Csv -Encoding UTF8
foreach ($f in $filas) {
    $nombres = @(); $valores = @()
    foreach ($p in $f.PSObject.Properties) {
        if ([string]::IsNullOrEmpty($p.Value)) { continue }
        $nombres += "[" + $p.Name + "]"
        if ($pctCols -contains $p.Name) {
            if ($p.Value -notmatch '^\d+(\.\d+)?$') { throw "Numero invalido en $($p.Name): $($p.Value)" }
            $valores += $p.Value            # literal SQL: siempre con punto, sin importar la region
        } else {
            $valores += "'" + $p.Value.Replace("'", "''") + "'"
        }
    }
    $cn.Execute("INSERT INTO SWI (" + ($nombres -join ",") + ") VALUES (" + ($valores -join ",") + ")") | Out-Null
}
$cn.Close()
Write-Output $filas.Count
'''


def _powershell_32():
    raiz = os.environ.get("SystemRoot", r"C:\Windows")
    p = os.path.join(raiz, r"SysWOW64\WindowsPowerShell\v1.0\powershell.exe")
    return p if os.path.exists(p) else os.path.join(raiz, r"System32\WindowsPowerShell\v1.0\powershell.exe")


def generar_mdb(registros, plantilla=None):
    """
    Access 2000 (.mdb) con la tabla SWI. Usa el motor Jet 4.0 que trae Windows (32 bits),
    asi que no hace falta tener Access instalado. Si se pasa la plantilla oficial de BO,
    se copia y se llenan sus filas en vez de crear la tabla.
    """
    cols = ", ".join(f"[{f}] DOUBLE" if tipo == "P" else f"[{f}] TEXT({largo})"
                     for f, _i, largo, tipo, _r, _g in LAYOUT)
    tmp = tempfile.mkdtemp(prefix="swi_")
    try:
        ps_path = os.path.join(tmp, "mdb.ps1")
        csv_path = os.path.join(tmp, "datos.csv")
        mdb_path = os.path.join(tmp, "SWI.mdb")
        with open(ps_path, "w", encoding="utf-8-sig") as fh:
            fh.write(_PS_MDB)
        with open(csv_path, "wb") as fh:
            fh.write(generar_csv(registros))
        crear = "1"
        if plantilla and os.path.exists(plantilla):
            shutil.copyfile(plantilla, mdb_path)
            crear = "0"
        out = subprocess.run(
            [_powershell_32(), "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps_path,
             "-Mdb", mdb_path, "-Csv", csv_path, "-Crear", crear, "-Cols", cols, "-Pct", ",".join(sorted(PCT))],
            capture_output=True, text=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        if out.returncode != 0:
            raise RuntimeError((out.stderr or out.stdout).strip())
        with open(mdb_path, "rb") as fh:
            return fh.read()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
