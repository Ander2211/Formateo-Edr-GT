# Formateo de EDRs de ICT: contexto y funcionamiento del script

## 1. Contexto

Cada mes, ICT genera los EDRs (CDRs de tráfico de interconexión) de Guatemala en archivos
CSV con la estructura `CDR_EXT`. Esos archivos se transforman a tablas DBF que se envían a
los operadores (ATEL, CLARO, COMCEL, TELGUA, CABNET, ATT, etc.).

El proceso que genera el DBF necesita CSV limpios, en el formato de la tabla de referencia, y
con los valores de cada llamada **en quetzales (GTQ), con IVA incluido**, que es lo acordado
para el intercambio de información con los operadores.

**Problema detectado en la revisión de septiembre 2026:** los CSV traen la mayoría de las
llamadas en dólares (USD, con IVA incluido), pero algunas transacciones de COMCEL ya vienen
en quetzales. El valor no es homogéneo, y todos deben ir en quetzales.

**Qué confirmó el área de interconexión (Ronaldo):**

- ICT entrega casi todo en USD con IVA incluido; basta aplicar el tipo de cambio del mes.
- ATEL: todas las transacciones vienen en USD, se convierten todas.
- COMCEL: la mayoría viene en USD, pero **ya vienen en GTQ y no deben alterarse**:
  - TELGUA: transacción **83** solo tráfico entrante; transacción **94** entrante y saliente.
  - CLARO: transacción **94** entrante y saliente.
- Tipo de cambio de septiembre 2026: **7.63856** (usar los 5 decimales).
- Los decimales se **truncan** (no se redondean). El DBF guarda 4 decimales, pero si el CSV
  trae 5, mejor.

## 2. Qué es el script

Una herramienta en **Python 2.7.5** (solo librería estándar, sin internet) que:

1. Lee los CSV `CDR_EXT` originales de la carpeta de entrada.
2. Les quita las líneas técnicas y los deja como una tabla que se abre en Excel.
3. Convierte a quetzales lo que viene en dólares.
4. Escribe el resultado en la carpeta de salida, con el mismo nombre de archivo.
5. Genera un log con todo lo procesado y los errores.

Se compone de tres archivos:

| Archivo | Para qué sirve |
|---|---|
| `formatear_cdr.py` | El script principal |
| `config.txt` | Carpetas, exclusiones, tipo de cambio y reglas de excepción |
| `ejecutar_formatear.sh` | Lanzador en bash: busca Python 2, ejecuta el script, guarda el log y programa la ejecución diaria (variable `HORAS`) |

**El script no modifica los originales.** Los archivos de `ENTRADA` solo se leen. Todo se
escribe en `SALIDA` (primero como `.tmp` y luego se renombra) y en la carpeta `logs`.

## 3. Qué hace con cada archivo

### 3.1 Estructura

Entrada (`CDR_EXT`):

```
AARHDR,1.0,CDR_EXTRACT,CSV,20261005171829
AARTAB,CEXT,20261005171829,CEXT
AARCOL,PARAMETER_MATRIX,RATING_COMPONENT,ANUM,BNUM,EVENT_START_DATE,EVENT_START_TIME,EVENT_DURATION,AMOUNT,...
AARSIZ,...  AARTYP,...  AARDEC,...
,*,0092,22232028,24581309,20260901,10415900,625.00000000000,0.17500000000,,AT1GINB,USD,24581309
AARCNT,245
AARTRA,20261005175030
```

Salida (formato de la tabla de referencia):

```
Cod_tra,Tel_ori,Tel_des,Fec_lla,Hor_lla,Seg_rea,Val_lla,Rut_ent,Rut_sal
92,22232028,24581309,20260901,104159,625,0.13474,,AT1GINB
```

### 3.2 Cambios de formato

| Campo de salida | Origen | Cambio |
|---|---|---|
| `Cod_tra` | `RATING_COMPONENT` | Sin ceros a la izquierda (`0092` → `92`) |
| `Tel_ori` | `ANUM` | Igual |
| `Tel_des` | `BNUM` | Igual |
| `Fec_lla` | `EVENT_START_DATE` | Igual (`AAAAMMDD`) |
| `Hor_lla` | `EVENT_START_TIME` | `HHMMSS`, se quitan los 2 últimos dígitos (`10415900` → `104159`) |
| `Seg_rea` | `EVENT_DURATION` | Sin decimales (`625.00000000000` → `625`) |
| `Val_lla` | `AMOUNT` | En quetzales, ver sección 4 |
| `Rut_ent` | `INCOMING_PATH` | Igual |
| `Rut_sal` | `OUTGOING_PATH` | Igual |

Se omiten la columna `*` (`PARAMETER_MATRIX`), la moneda y `USER_SUMMARISATION`. Las dos
últimas se pueden agregar con `incluir_moneda = si` e `incluir_resumen = si` en el config.

### 3.3 Carpetas y nombres

- `entrada` del config acepta una o varias carpetas, cada una con sus subcarpetas.
- En `SALIDA` se crea una carpeta por cada entrada (`CABGUA`, `CLARO/...`, `TELGUA`) con la
  misma estructura.
- El archivo de salida conserva exactamente el nombre del de entrada.
- `excluir_carpetas` (por ejemplo `DESCARGA`) no se procesan. Además se omite cualquier
  archivo que no empiece con `AARHDR` (reportes que no son CDR).

## 4. Conversión a quetzales

Se activa con `convertir_a_gtq = si`. Para cada registro, el script decide así:

| Situación | Qué hace |
|---|---|
| La moneda no es USD ni GTQ | Error: ese archivo no se genera y queda en la lista de errores |
| Está en `excepciones_gtq` | **No se convierte ni se altera** (ya viene en GTQ). Si su moneda decía USD, queda como aviso |
| Moneda USD | `monto × tipo_cambio`, truncado a 5 decimales |
| Moneda GTQ sin regla | No se convierte (ya está en GTQ) y queda como **aviso** para revisar |

- **No se suma IVA:** los montos de ICT ya lo incluyen, en dólares y en quetzales.
- **Truncado, no redondeo.** Ejemplo: `0.01764 USD × 7.63856 = 0.1347441984` → `0.13474`.
  En el DBF de 4 decimales quedaría `0.1347`.
- **Cómo identifica la transacción, la franquicia, el operador y el sentido:**
  - Transacción: `Cod_tra` del registro (se compara sin ceros: `0094` = `94`).
  - Franquicia y operador: del nombre del archivo (`FRAN-TELGUA_NOP-COMCEL`).
  - Sentido: de la letra `SDIR-` del nombre. Por la estructura de carpetas se asumió
    **`SDIR-R` = entrante** y **`SDIR-S` = saliente** (configurable).

### Excepciones configuradas para septiembre 2026

```
excepciones_gtq = TELGUA:COMCEL:83:entrante,
    TELGUA:COMCEL:94:*,
    CLARO:COMCEL:94:*
```

Formato: `FRANQUICIA:OPERADOR:TRANSACCIÓN:SENTIDO` (`entrante`, `saliente` o `*` para ambos).

## 5. El log

En cada ejecución se crea `logs/formatear_cdr_AAAAMMDD_HHMMSS.log` con:

- Fecha y hora de inicio y fin, configuración usada (tipo de cambio, excepciones, carpetas).
- Carpetas y archivos procesados, con el número de registros y cuántos se convirtieron.
- **Detalle por franquicia / operador / sentido / transacción / moneda:** cuántos registros
  se convirtieron, cuántos son excepción y cuántos vienen en GTQ sin regla.
- **Avisos de moneda:** casos que conviene revisar antes de entregar.
- Resumen final y **lista de errores**.

Con el detalle se puede verificar que lo que dice Ronaldo coincide con los datos reales (por
ejemplo, que COMCEL en las transacciones 83 y 94 realmente venga en GTQ).

## 6. Uso

```bash
chmod +x ejecutar_formatear.sh
./ejecutar_formatear.sh                    # ejecuta una vez
./ejecutar_formatear.sh --instalar-cron    # programa la ejecución diaria (variable HORAS)
./ejecutar_formatear.sh --estado           # muestra la programación
./ejecutar_formatear.sh --quitar-cron      # quita la programación
```

El bash busca un Python 2 en el servidor (`python2.7`, `python2`, `python`); si está en otra
ruta: `PYTHON_BIN=/ruta/python2.7 ./ejecutar_formatear.sh`.

**Cada mes:** cambiar `tipo_cambio` y las carpetas `EDRS_MM_AAAA` en `config.txt`.

## 7. Puntos pendientes de confirmar

1. **Sentido:** que `SDIR-R` sea entrante y `SDIR-S` saliente. Solo afecta a la excepción de
   la transacción 83 de TELGUA.
2. **CABGUA → CLARO:** viene en GTQ en el archivo de muestra y no está en la lista de
   excepciones. No se convierte y sale como aviso.
3. **Doble conversión:** confirmar que el extract de ICT no aplique ya el tipo de cambio.
   Los CSV de muestra traen USD.
4. **Validación:** comparar el resultado contra el DBF de muestra de TELGUA → ATEL
   (`2026090200_1D.xlsx`), que aún no se ha recibido.
5. **Moneda en el archivo:** el formato de referencia no lleva columna de moneda.
6. **`Cod_tra` en FoxPro:** valores de 4 dígitos (`1801`) aparecían como 0 en la tabla,
   probablemente por el ancho del campo en el DBF (no es del script).

## 8. Pruebas realizadas

Probado con datos de muestra (CSV propios y el archivo real de CABGUA → TELGUA con 245
registros): montos truncados verificados contra el cálculo a mano, excepciones de COMCEL sin
alterar, error con moneda desconocida sin dejar archivos a medias, y archivos originales
idénticos antes y después (suma md5). **No se probó con Python 2.7.5 ni con los datos reales
del servidor:** se probó con Python 3 mediante una copia de prueba. Conviene hacer una
corrida de prueba con una central en el servidor antes de procesar todo.
