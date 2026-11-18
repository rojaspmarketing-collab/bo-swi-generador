# BO SWI Generador

Convierte un Excel de obras al archivo **SWI** de BackOffice Music Services (BO LATAM) en los
tres formatos que acepta BO, después de revisarlo contra todas las reglas del
*Layout SWI File v1.4*:

| Formato | Detalle |
|---|---|
| **Access (.mdb)** | Tabla `SWI`. No necesita Access instalado (usa el motor Jet que trae Windows). |
| **TXT** | Ancho fijo, 1013 caracteres por línea, codificación ANSI. |
| **CSV** | UTF-8 delimitado por comas. |

## Cómo usarlo

1. Doble clic en **`Iniciar.bat`**. Se abre el programa en el navegador.
2. La primera vez, entra a **Mi editorial** (arriba a la derecha) y pon el nombre, el CAE/IPI,
   las sociedades y el país de tu editorial.
3. Descarga la **plantilla Excel**, llénala y arrástrala al programa.
4. Si todo está bien, descarga el `.mdb`, el `.txt` o el `.csv` y súbelo a BO.
   Si hay errores, el programa te dice la fila y el campo exactos.

Todo corre en tu computador (`127.0.0.1`). Ningún archivo sale a internet.
Los resultados también quedan guardados en la carpeta `salida/`.

## Qué revisa y qué corrige

**Corrige solo** (y te lo muestra en la pestaña *Correcciones*):

- Separador `;` o `,`, porcentajes con `%` o con coma decimal, y porcentajes en formato % de Excel.
- Sociedades escritas por nombre (`SAYCO`) o como `84 - SAYCO`, que pasan a `84`.
- Territorios escritos por nombre (`COLOMBIA`), que pasan a `CO`.
- Obras cuyos autores no suman 100 %: agrega `WRITER UNKNOWN` / `PUBLISHER UNKNOWN` con el resto
  (se puede apagar).
- Al `PUBLISHER UNKNOWN` le pone el mismo territorio que a los demás editores y le borra los datos de editor local.
- Completa CAE, sociedades y país de **tu** editorial cuando están vacíos.

**Bloquea** (errores que BO rechazaría):

- Campos obligatorios vacíos, textos más largos que el layout, caracteres no ANSI.
- Códigos de sociedad, territorio o rol que no están en las tablas de BO.
- `ADM_TERRITORY_CODE` que no es un solo país, y collect fuera de 0 a 100 o que suma más de 100 por país.
- Datos de editor local incompletos, filas duplicadas y títulos distintos para el mismo código.

## Plantilla MDB oficial

Si tienes la plantilla `.mdb` de la Intranet de BO, guárdala como `plantillas/SWI_oficial.mdb` y
el programa la usará en lugar de crear la tabla.

## Desarrollo

```
python -m unittest discover -s tests -v
```

Estructura:

```
app.py              servidor local + endpoints
web/index.html      interfaz
swi/tablas.py       layout, sociedades, territorios y roles de la guía v1.4
swi/procesador.py   lectura, normalización, correcciones y validación
swi/escritores.py   CSV, TXT y MDB
swi/plantilla.py    plantilla Excel con listas desplegables
```
