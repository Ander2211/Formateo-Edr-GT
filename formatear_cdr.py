# -*- coding: utf-8 -*-
"""
Convierte archivos CDR_EXT a un CSV legible en Excel.

Modo "tabla" (por defecto), formato de la tabla de referencia:
    Cod_tra, Tel_ori, Tel_des, Fec_lla, Hor_lla, Seg_rea, Val_lla, Rut_ent, Rut_sal
  - Se omite la primera columna (PARAMETER_MATRIX, el "*").
  - Cod_tra: se quitan los ceros a la izquierda (0092 -> 92).
  - Hor_lla: HHMMSS (se quitan los 2 ultimos digitos de 10415900 -> 104159).
  - Seg_rea: sin decimales (625.00000000000 -> 625).
  - Val_lla: sin ceros finales (0.17500000000 -> 0.175).

Modo "original": solo quita las lineas tecnicas, los datos no se modifican.

Conversion a quetzales (convertir_a_gtq = si, solo modo tabla):
  - Los registros en USD se multiplican por tipo_cambio y el resultado se TRUNCA
    (no se redondea) a decimales_monto decimales.
  - Los montos ya incluyen IVA: NO se suma IVA.
  - Los registros que ya vienen en quetzales (excepciones COMCEL, lista
    excepciones_gtq del config: franquicia:operador:transaccion:sentido) NO se
    convierten ni se alteran.
  - Los registros en GTQ que no estan en la lista de excepciones tampoco se
    convierten (ya estan en quetzales) pero se reportan como AVISO en el log.
  - La franquicia, el operador y el sentido (SDIR-R / SDIR-S) se toman del
    nombre del archivo; la transaccion es Cod_tra (RATING_COMPONENT).

Los archivos originales en ENTRADA solo se leen; nunca se modifican.

Compatible con Python 2.7.5 (solo libreria estandar, sin internet).

Version optimizada + multiprocessing (misma logica, mismas salidas):
  - Cada archivo se procesa en un proceso aparte (multiprocessing.Pool).
  - Solo el proceso principal escribe el log y acumula las estadisticas, en el
    mismo orden de siempre (carpetas y archivos en orden alfabetico).
  - Opcional en [archivos] del config:  procesos = auto | 1 | 2 | 4 ...
      auto = cantidad de CPU de la maquina (por defecto)
      1    = sin multiprocessing (modo secuencial)
  - Mejoras por archivo: plan de columnas preparado una vez, cache de codigo/
    excepcion/moneda, quantum de Decimal calculado una vez y escritura por
    bloques con buffer grande.

Uso:
    python formatear_cdr_mp.py [ruta_config.txt]

'entrada' en el config puede ser una carpeta o varias (separadas por coma).
Cada carpeta se recorre con sus subcarpetas y se escribe en SALIDA conservando el
nombre de cada archivo. Con varias carpetas de entrada (por ejemplo una por
central: CABGUA, CLARO, TELGUA...), en SALIDA se crea una carpeta por cada una
con su mismo nombre.

En cada ejecucion se genera un .log (carpeta 'logs' junto al script, o la que
indique 'carpeta_log' en [carpetas] del config, opcional) con: fecha y hora,
carpetas y archivos procesados, detalle por franquicia/operador/transaccion/
moneda, avisos, resumen y lista de errores.
"""
import os
import re
import sys
import csv
import time
import ConfigParser
import multiprocessing
from decimal import Decimal, ROUND_DOWN


class NoEsCDR(Exception):
    pass


SEPARADORES = {
    'coma': ',',
    'punto_y_coma': ';',
    'tab': '\t',
    'barra': '|',
}

# FRAN-TELGUA_NOP-COMCEL_SDIR-R  ->  franquicia, operador, letra de sentido
PATRON_NOMBRE = re.compile(r'FRAN-([A-Za-z0-9]+)_NOP-([A-Za-z0-9]+)_SDIR-([RSrs])')

# Tamanos de buffer / bloque
BUFFER_ARCHIVO = 1024 * 1024
BLOQUE_FILAS = 5000

# Opciones disponibles dentro de cada proceso trabajador
_OPC = None


# ----------------------------------------------------------------------------
# Log
# ----------------------------------------------------------------------------
def ahora():
    return time.strftime('%Y-%m-%d %H:%M:%S')


def hora():
    return time.strftime('%H:%M:%S')


class Log(object):
    """Escribe cada linea en pantalla y en el archivo .log."""

    def __init__(self, ruta):
        self.ruta = ruta
        self.f = open(ruta, 'w')

    def escribir(self, texto=''):
        print(texto)
        self.f.write(texto + '\n')
        self.f.flush()

    def cerrar(self):
        self.f.close()


# ----------------------------------------------------------------------------
# Configuracion
# ----------------------------------------------------------------------------
def a_bool(valor):
    return valor.strip().lower() in ('si', 's', 'yes', 'y', 'true', '1')


def a_lista(valor):
    return [x.strip() for x in valor.split(',') if x.strip()]


def obtener(cfg, seccion, opcion, defecto):
    if cfg.has_option(seccion, opcion):
        return cfg.get(seccion, opcion).strip()
    return defecto


def leer_excepciones(texto):
    """
    Lee 'excepciones_gtq': FRANQUICIA:OPERADOR:TRANSACCION:SENTIDO separados por
    coma o salto de linea. SENTIDO = entrante, saliente o * (ambos).
    Devuelve lista de (FRANQUICIA, OPERADOR, int(transaccion), sentido).
    """
    reglas = []
    for token in [x.strip() for x in texto.replace('\n', ',').split(',') if x.strip()]:
        partes = [p.strip() for p in token.split(':')]
        if len(partes) != 4:
            raise ValueError("excepciones_gtq: '%s' no tiene el formato "
                             "FRANQUICIA:OPERADOR:TRANSACCION:SENTIDO" % token)
        fran, oper, trans, sentido = partes
        if not trans.isdigit():
            raise ValueError("excepciones_gtq: la transaccion '%s' debe ser un numero" % trans)
        sentido = sentido.lower()
        if sentido not in ('entrante', 'saliente', '*'):
            raise ValueError("excepciones_gtq: sentido '%s' invalido en '%s' "
                             "(use entrante, saliente o *)" % (sentido, token))
        reglas.append((fran.upper(), oper.upper(), int(trans), sentido))
    return reglas


def leer_sentido(cfg, opcion, defecto):
    valor = obtener(cfg, 'formato', opcion, defecto).lower()
    if valor not in ('entrante', 'saliente'):
        raise ValueError("%s invalido '%s'. Use: entrante o saliente" % (opcion, valor))
    return valor


def leer_procesos(texto):
    """'auto' = cantidad de CPU; o un entero >= 1."""
    t = texto.strip().lower()
    if t in ('', 'auto'):
        try:
            return max(1, multiprocessing.cpu_count())
        except NotImplementedError:
            return 1
    try:
        n = int(t)
        if n < 1:
            raise ValueError()
    except ValueError:
        raise ValueError("procesos invalido '%s'. Use: auto o un entero >= 1" % texto)
    return n


def leer_config(ruta_config, base_dir):
    cfg = ConfigParser.RawConfigParser()
    if not cfg.read(ruta_config):
        raise IOError("No se pudo leer el archivo de configuracion: %s" % ruta_config)

    entrada_txt = obtener(cfg, 'carpetas', 'entrada', '')
    salida = obtener(cfg, 'carpetas', 'salida', '')
    # 'entrada' admite una o varias carpetas, separadas por coma o por salto de linea
    pedidas = [x.strip() for x in entrada_txt.replace('\n', ',').split(',') if x.strip()]
    if not pedidas or not salida:
        raise ValueError("Faltan 'entrada' o 'salida' en la seccion [carpetas]")

    entradas = []
    for e in pedidas:
        if not os.path.isabs(e):
            e = os.path.join(base_dir, e)
        e = os.path.normpath(e)
        if e not in entradas:
            entradas.append(e)

    if not os.path.isabs(salida):
        salida = os.path.join(base_dir, salida)

    # carpeta_por_entrada (opcional): auto = una carpeta por entrada solo si hay varias
    modo_carpeta = obtener(cfg, 'carpetas', 'carpeta_por_entrada', 'auto').lower()
    if modo_carpeta not in ('auto', 'si', 'no'):
        raise ValueError("carpeta_por_entrada invalido '%s'. Use: auto, si o no" % modo_carpeta)
    envolver = (modo_carpeta == 'si') or (modo_carpeta == 'auto' and len(entradas) > 1)
    if envolver:
        nombres = [os.path.basename(e).upper() for e in entradas]
        if len(set(nombres)) != len(nombres):
            raise ValueError("hay carpetas de entrada con el mismo nombre; "
                             "no se pueden juntar dentro de SALIDA")
    elif len(entradas) > 1:
        raise ValueError("con varias carpetas de entrada, carpeta_por_entrada no puede ser 'no'")

    # Opcional: carpeta de logs (por defecto 'logs' junto al script)
    dir_log = obtener(cfg, 'carpetas', 'carpeta_log', 'logs')
    if not os.path.isabs(dir_log):
        dir_log = os.path.join(base_dir, dir_log)

    nombre_sep = obtener(cfg, 'formato', 'separador', 'coma').lower()
    if nombre_sep not in SEPARADORES:
        raise ValueError("separador invalido '%s'. Use: %s" %
                         (nombre_sep, ', '.join(sorted(SEPARADORES.keys()))))

    modo = obtener(cfg, 'formato', 'modo', 'tabla').lower()
    if modo not in ('tabla', 'original'):
        raise ValueError("modo invalido '%s'. Use: tabla u original" % modo)

    dec_txt = obtener(cfg, 'formato', 'decimales_monto', '')
    dec_monto = None
    if dec_txt:
        try:
            dec_monto = int(dec_txt)
            if dec_monto < 0:
                raise ValueError()
        except ValueError:
            raise ValueError("decimales_monto debe ser un entero >= 0 (o quedar vacio)")

    # Conversion USD -> GTQ (opcional)
    convertir = a_bool(obtener(cfg, 'formato', 'convertir_a_gtq', 'no'))
    tipo_cambio = None
    excepciones = []
    if convertir:
        if modo != 'tabla':
            raise ValueError("convertir_a_gtq solo funciona con modo = tabla")
        try:
            tipo_cambio = Decimal(obtener(cfg, 'formato', 'tipo_cambio', ''))
            if tipo_cambio <= 0:
                raise ValueError()
        except Exception:
            raise ValueError("con convertir_a_gtq = si, 'tipo_cambio' debe ser un numero "
                             "mayor a 0 (ejemplo: 7.63856)")
        if dec_monto is None:
            dec_monto = 5
        excepciones = leer_excepciones(obtener(cfg, 'formato', 'excepciones_gtq', ''))

    # Quantum de Decimal calculado una sola vez (antes se recalculaba por registro)
    quantum = Decimal(1).scaleb(-dec_monto) if dec_monto is not None else None

    procesos = leer_procesos(obtener(cfg, 'archivos', 'procesos', 'auto'))

    return {
        'entradas': entradas,
        'envolver': envolver,
        'salida': os.path.normpath(salida),
        'dir_log': os.path.normpath(dir_log),
        'excluir': [x.upper() for x in a_lista(obtener(cfg, 'carpetas', 'excluir_carpetas', ''))],
        'extensiones': [x.lower() for x in a_lista(obtener(cfg, 'archivos', 'extensiones', '.csv'))],
        'sobrescribir': a_bool(obtener(cfg, 'archivos', 'sobrescribir', 'si')),
        'procesos': procesos,
        'separador': SEPARADORES[nombre_sep],
        'modo': modo,
        'dec_monto': dec_monto,
        'quantum': quantum,
        'convertir': convertir,
        'tipo_cambio': tipo_cambio,
        'excepciones': excepciones,
        'sentidos': {'R': leer_sentido(cfg, 'sentido_sdir_r', 'entrante'),
                     'S': leer_sentido(cfg, 'sentido_sdir_s', 'saliente')},
        'moneda': a_bool(obtener(cfg, 'formato', 'incluir_moneda', 'no')),
        'resumen': a_bool(obtener(cfg, 'formato', 'incluir_resumen', 'no')),
    }


# ----------------------------------------------------------------------------
# Formateo de campos (modo tabla)
# ----------------------------------------------------------------------------
def _es_numero(v):
    try:
        float(v)
        return True
    except ValueError:
        return False


def quitar_ceros_finales(v):
    """625.00000000000 -> 625 ; 0.17500000000 -> 0.175 (no redondea)."""
    if '.' not in v:
        return v
    return v.rstrip('0').rstrip('.')


def truncar(valor, decimales):
    """Trunca (no redondea) un Decimal a N decimales y quita ceros finales."""
    q = Decimal(1).scaleb(-decimales)
    return quitar_ceros_finales(format(valor.quantize(q, rounding=ROUND_DOWN), 'f'))


def truncar_q(valor, q):
    """Igual que truncar() pero con el quantum ya calculado."""
    return quitar_ceros_finales(format(valor.quantize(q, rounding=ROUND_DOWN), 'f'))


def sin_ceros_izq(valor):
    v = valor.strip()
    if not v:
        return v
    sin_ceros = v.lstrip('0')
    return sin_ceros if sin_ceros else '0'


def fmt_texto(valor, opc):
    return valor


def fmt_cod(valor, opc):
    # 0092 -> 92 (solo ceros a la izquierda)
    return sin_ceros_izq(valor)


def fmt_hora(valor, opc):
    # HHMMSScc -> HHMMSS
    v = valor.strip()
    if v.isdigit():
        return v.zfill(8)[0:6]
    return v


def fmt_duracion(valor, opc):
    # 625.00000000000 -> 625
    v = valor.strip()
    if not _es_numero(v):
        return v
    return quitar_ceros_finales(v)


def fmt_monto(valor, opc):
    v = valor.strip()
    if not _es_numero(v):
        return v
    q = opc['quantum']
    if q is not None:
        return truncar_q(Decimal(v), q)
    # Sin decimales definidos: solo se quitan los ceros finales
    return quitar_ceros_finales(v)


# (titulo de salida, columna de origen, funcion de formato)
PLAN_TABLA = [
    ('Cod_tra', 'RATING_COMPONENT', fmt_cod),
    ('Tel_ori', 'ANUM', fmt_texto),
    ('Tel_des', 'BNUM', fmt_texto),
    ('Fec_lla', 'EVENT_START_DATE', fmt_texto),
    ('Hor_lla', 'EVENT_START_TIME', fmt_hora),
    ('Seg_rea', 'EVENT_DURATION', fmt_duracion),
    ('Val_lla', 'AMOUNT', fmt_monto),
    ('Rut_ent', 'INCOMING_PATH', fmt_texto),
    ('Rut_sal', 'OUTGOING_PATH', fmt_texto),
]


def construir_plan(columnas, opc):
    """Devuelve lista de (titulo, indice_en_origen, funcion_o_None)."""
    if opc['modo'] == 'original':
        return [(c, i, None) for i, c in enumerate(columnas)]

    pos = {}
    for i, c in enumerate(columnas):
        pos[c.strip().upper()] = i

    definicion = list(PLAN_TABLA)
    if opc['moneda']:
        definicion.append(('Moneda', 'CURRENCY', fmt_texto))
    if opc['resumen']:
        definicion.append(('Resumen', 'USER_SUMMARISATION', fmt_texto))

    plan = []
    for titulo, origen, fn in definicion:
        if origen not in pos:
            raise ValueError("falta la columna %s en la linea AARCOL" % origen)
        # fmt_texto no hace nada: se marca como None para copiar directo
        plan.append((titulo, pos[origen], None if fn is fmt_texto else fn))
    return plan


# ----------------------------------------------------------------------------
# Conversion a quetzales
# ----------------------------------------------------------------------------
def datos_del_nombre(nombre, opc):
    """
    Devuelve (franquicia, operador, sentido) tomados del nombre del archivo.
    sentido = 'entrante' / 'saliente' segun sentido_sdir_r / sentido_sdir_s.
    Si el nombre no sigue el patron devuelve ('?', '?', '?').
    """
    m = PATRON_NOMBRE.search(nombre)
    if not m:
        return '?', '?', '?'
    letra = m.group(3).upper()
    return m.group(1).upper(), m.group(2).upper(), opc['sentidos'][letra]


def es_excepcion(meta, cod_num, opc):
    """True si (franquicia, operador, transaccion, sentido) esta en excepciones_gtq."""
    fran, oper, sentido = meta
    for r_fran, r_oper, r_trans, r_sent in opc['excepciones']:
        if r_fran == fran and r_oper == oper and r_trans == cod_num \
                and (r_sent == '*' or r_sent == sentido):
            return True
    return False


# ----------------------------------------------------------------------------
# Procesamiento de un archivo
# ----------------------------------------------------------------------------
def procesar_archivo(ruta_in, ruta_out, opc, det):
    """
    Devuelve (registros_escritos, registros_declarados, registros_convertidos).
    'det' acumula el detalle: (franquicia, operador, sentido, transaccion, moneda)
    -> [registros, convertidos, excepcion_sin_convertir, gtq_sin_regla].
    """
    ruta_tmp = ruta_out + '.tmp'
    meta = datos_del_nombre(os.path.basename(ruta_in), opc)
    convertir = opc['convertir']
    tipo_cambio = opc['tipo_cambio']
    quantum = opc['quantum']

    f_in = open(ruta_in, 'rb')
    f_out = None
    writer = None
    bloque = []
    plan = []
    plan_ops = []           # [(indice_origen, funcion_o_None)]
    pos_monto_salida = -1   # posicion de AMOUNT dentro de la fila de salida
    num_cols = 0
    escritos = 0
    convertidos = 0
    idx_monto = idx_mon = idx_cod = None
    declarado = None
    primera = True
    cache_cod = {}          # codigo crudo -> (cod, es_excepcion)
    cache_det = {}          # (cod, moneda) -> acumulador del detalle

    try:
        for fila in csv.reader(f_in):
            if not fila:
                continue
            etiqueta = fila[0].strip()

            if primera:
                if etiqueta != 'AARHDR':
                    raise NoEsCDR("no empieza con AARHDR (no es un CDR_EXT)")
                primera = False

            if etiqueta == 'AARCOL':
                columnas = fila[1:]
                num_cols = len(columnas)
                plan = construir_plan(columnas, opc)
                plan_ops = [(p[1], p[2]) for p in plan]
                if convertir:
                    pos = dict((c.strip().upper(), i) for i, c in enumerate(columnas))
                    for req in ('AMOUNT', 'CURRENCY', 'RATING_COMPONENT'):
                        if req not in pos:
                            raise ValueError("para convertir a GTQ falta la columna %s" % req)
                    idx_monto = pos['AMOUNT']
                    idx_mon = pos['CURRENCY']
                    idx_cod = pos['RATING_COMPONENT']
                    pos_monto_salida = -1
                    for j, (i, fn) in enumerate(plan_ops):
                        if i == idx_monto:
                            pos_monto_salida = j
                f_out = open(ruta_tmp, 'wb', BUFFER_ARCHIVO)
                writer = csv.writer(f_out, delimiter=opc['separador'], lineterminator='\n')
                writer.writerow([p[0] for p in plan])

            elif etiqueta == 'AARCNT':
                try:
                    declarado = int(fila[1])
                except (IndexError, ValueError):
                    pass

            elif etiqueta == '':
                if writer is None:
                    raise ValueError("hay registros antes de la linea AARCOL")
                datos = fila[1:]
                if len(datos) < num_cols:
                    datos = datos + [''] * (num_cols - len(datos))
                monto_txt = None
                if convertir:
                    moneda = datos[idx_mon].strip().upper()
                    if moneda != 'USD' and moneda != 'GTQ':
                        raise ValueError("moneda no reconocida '%s' en el registro %d "
                                         "(solo USD o GTQ)" % (moneda, escritos + 1))
                    cod_crudo = datos[idx_cod]
                    info = cache_cod.get(cod_crudo)
                    if info is None:
                        cod = sin_ceros_izq(cod_crudo)
                        cod_num = int(cod) if cod.isdigit() else -1
                        info = (cod, es_excepcion(meta, cod_num, opc))
                        cache_cod[cod_crudo] = info
                    cod = info[0]
                    v = datos[idx_monto].strip()
                    clave_c = (cod, moneda)
                    d = cache_det.get(clave_c)
                    if d is None:
                        d = det.setdefault((meta[0], meta[1], meta[2], cod, moneda),
                                           [0, 0, 0, 0])
                        cache_det[clave_c] = d
                    d[0] += 1
                    if info[1]:
                        # Ya viene en quetzales desde ICT: no se altera
                        monto_txt = quitar_ceros_finales(v) if _es_numero(v) else v
                        d[2] += 1
                    elif moneda == 'USD':
                        try:
                            monto_txt = truncar_q(Decimal(v) * tipo_cambio, quantum)
                        except Exception:
                            raise ValueError("monto no numerico '%s' en el registro %d"
                                             % (v, escritos + 1))
                        d[1] += 1
                        convertidos += 1
                    else:
                        # GTQ fuera de la lista de excepciones: ya esta en quetzales
                        monto_txt = quitar_ceros_finales(v) if _es_numero(v) else v
                        d[3] += 1
                    datos[idx_mon] = 'GTQ'

                salida = [datos[i] if fn is None else fn(datos[i], opc)
                          for i, fn in plan_ops]
                if monto_txt is not None and pos_monto_salida >= 0:
                    salida[pos_monto_salida] = monto_txt
                bloque.append(salida)
                escritos += 1
                if len(bloque) >= BLOQUE_FILAS:
                    writer.writerows(bloque)
                    bloque = []
            # AARHDR, AARTAB, AARSIZ, AARTYP, AARDEC, AARTRA: se ignoran

        if writer is None:
            raise ValueError("no se encontro la linea AARCOL con los nombres de columna")
        if bloque:
            writer.writerows(bloque)
            bloque = []
    except Exception:
        f_in.close()
        if f_out is not None:
            f_out.close()
        if os.path.exists(ruta_tmp):
            os.remove(ruta_tmp)
        raise

    f_in.close()
    f_out.close()
    if os.path.exists(ruta_out):
        # En Windows os.rename falla si el destino existe (sobrescribir = si)
        os.remove(ruta_out)
    os.rename(ruta_tmp, ruta_out)
    return escritos, declarado, convertidos


# ----------------------------------------------------------------------------
# Trabajador (corre en cada proceso del Pool)
# ----------------------------------------------------------------------------
def _iniciar_trabajador(opc):
    global _OPC
    _OPC = opc


def trabajar(tarea):
    """
    Procesa un archivo. Nunca lanza excepciones (se devuelven como resultado).
    Devuelve:
      ('ok', escritos, declarado, convertidos, det_local)
      ('nocdr', mensaje)
      ('err', mensaje)
    """
    ruta_in, ruta_out, destino_dir = tarea
    det = {}
    try:
        if not os.path.isdir(destino_dir):
            try:
                os.makedirs(destino_dir)
            except OSError:
                # otro proceso pudo crearla al mismo tiempo
                if not os.path.isdir(destino_dir):
                    raise
        escritos, declarado, convertidos = procesar_archivo(ruta_in, ruta_out, _OPC, det)
        return ('ok', escritos, declarado, convertidos, det)
    except NoEsCDR as e:
        return ('nocdr', str(e))
    except Exception as e:
        return ('err', str(e))


# ----------------------------------------------------------------------------
# Planificacion: recorre carpetas y arma la lista ordenada de pasos
# ----------------------------------------------------------------------------
def planificar(opc):
    """
    Devuelve la lista ordenada de pasos a registrar/procesar:
      ('error_ruta', entrada, mensaje)
      ('excluida', ruta_mostrada)
      ('carpeta', nombre_carpeta, n_archivos)
      ('archivo', nombre, etiqueta, ruta_in, ruta_out, destino_dir, omitir_existe)
    """
    items = []
    salida = opc['salida']
    extensiones = opc['extensiones']
    excluir = opc['excluir']

    for entrada in opc['entradas']:
        # Con varias entradas, cada una va en su propia carpeta dentro de SALIDA
        prefijo = os.path.basename(entrada) if opc['envolver'] else ''
        destino_base = os.path.join(salida, prefijo) if prefijo else salida

        if not os.path.isdir(entrada):
            items.append(('error_ruta', entrada,
                          "la carpeta de entrada no existe: %s" % entrada))
            continue
        if entrada == salida:
            items.append(('error_ruta', entrada,
                          "entrada y salida no pueden ser la misma carpeta: %s" % entrada))
            continue

        for raiz, dirs, archivos in os.walk(entrada):
            rel = os.path.relpath(raiz, entrada)
            rel = '' if rel == '.' else rel
            if prefijo and rel:
                rel_total = os.path.join(prefijo, rel)
            else:
                rel_total = prefijo or rel
            nombre_carpeta = rel_total if rel_total else '(raiz de entrada)'

            # Carpetas excluidas por config (se informan y no se recorren)
            for d in sorted(dirs):
                if d.upper() in excluir:
                    items.append(('excluida',
                                  os.path.join(rel_total, d) if rel_total else d))
            dirs[:] = sorted([d for d in dirs
                              if d.upper() not in excluir
                              and os.path.normpath(os.path.join(raiz, d)) != salida])

            destino_dir = os.path.join(destino_base, rel) if rel else destino_base

            candidatos = [n for n in sorted(archivos)
                          if os.path.splitext(n)[1].lower() in extensiones]
            if not candidatos:
                continue

            items.append(('carpeta', nombre_carpeta, len(candidatos)))
            for nombre in candidatos:
                ruta_in = os.path.join(raiz, nombre)
                etiqueta = os.path.join(rel_total, nombre) if rel_total else nombre
                # El archivo de salida conserva el nombre del archivo de entrada
                ruta_out = os.path.join(destino_dir, nombre)
                omitir = (not opc['sobrescribir']) and os.path.exists(ruta_out)
                items.append(('archivo', nombre, etiqueta, ruta_in, ruta_out,
                              destino_dir, omitir))
    return items


# ----------------------------------------------------------------------------
# Consumo de resultados (solo el proceso principal escribe el log)
# ----------------------------------------------------------------------------
def sumar_detalle(total, parcial):
    for clave, valores in parcial.items():
        d = total.get(clave)
        if d is None:
            total[clave] = list(valores)
        else:
            for k in range(4):
                d[k] += valores[k]


def consumir(items, resultados, opc, log, est):
    """Registra los pasos en orden, tomando del iterador el resultado de cada archivo."""
    actual = [None, 0, 0]   # nombre_carpeta, ok_carpeta, reg_carpeta

    def cerrar_carpeta():
        if actual[0] is not None and actual[1]:
            est['por_carpeta'].append([actual[0], actual[1], actual[2]])
        actual[0], actual[1], actual[2] = None, 0, 0

    for item in items:
        tipo = item[0]

        if tipo == 'error_ruta':
            log.escribir("[%s] ERROR %s" % (hora(), item[2]))
            est['errores'].append((item[1], item[2]))

        elif tipo == 'excluida':
            log.escribir("[%s] EXCLUIDA por config: %s" % (hora(), item[1]))

        elif tipo == 'carpeta':
            cerrar_carpeta()
            actual[0] = item[1]
            log.escribir("-" * 70)
            log.escribir("Carpeta: %s  (%d archivos)" % (item[1], item[2]))
            log.escribir("-" * 70)

        elif tipo == 'archivo':
            _, nombre, etiqueta, ruta_in, ruta_out, destino_dir, omitir = item

            if omitir:
                log.escribir("[%s] OMIT  %s (ya existe en salida)" % (hora(), nombre))
                est['omit'] += 1
                continue

            res = resultados.next(9999999)   # con timeout para poder interrumpir con Ctrl+C
            estado = res[0]

            if estado == 'ok':
                escritos, declarado, convertidos, det_local = res[1:]
                sumar_detalle(est['det'], det_local)
                aviso = ''
                if opc['convertir']:
                    aviso += '  [%d de %d registros convertidos USD->GTQ]' % (convertidos, escritos)
                    est['conv'] += convertidos
                    if datos_del_nombre(nombre, opc)[0] == '?':
                        aviso += '  [AVISO: el nombre no tiene FRAN-/NOP-/SDIR-; no se aplican excepciones]'
                        est['avisos_nombre'].append(etiqueta)
                if declarado is not None and declarado != escritos:
                    aviso += '  [AVISO: el archivo declara %d registros]' % declarado
                log.escribir("[%s] OK    %s (%d registros)%s" %
                             (hora(), nombre, escritos, aviso))
                est['ok'] += 1
                est['reg'] += escritos
                actual[1] += 1
                actual[2] += escritos
            elif estado == 'nocdr':
                log.escribir("[%s] OMIT  %s (%s)" % (hora(), nombre, res[1]))
                est['omit'] += 1
            else:
                log.escribir("[%s] ERROR %s: %s" % (hora(), nombre, res[1]))
                est['errores'].append((etiqueta, res[1]))
                est['err'] += 1

    cerrar_carpeta()


# ----------------------------------------------------------------------------
# Reportes
# ----------------------------------------------------------------------------
def escribir_detalle_moneda(log, est):
    """Detalle por franquicia/operador/sentido/transaccion/moneda y avisos."""
    det = est['det']
    log.escribir("")
    log.escribir("DETALLE POR FRANQUICIA / OPERADOR / SENTIDO / TRANSACCION / MONEDA ORIGINAL")
    log.escribir("(convertidos = USD -> GTQ; excepcion = ya en GTQ por regla, no se altera;")
    log.escribir(" GTQ sin regla = vienen en GTQ pero no estan en excepciones_gtq)")
    if not det:
        log.escribir("  (sin registros)")
        return
    avisos = []
    for clave in sorted(det.keys()):
        fran, oper, sentido, cod, moneda = clave
        reg, conv, exc, gtq = det[clave]
        log.escribir("  %-7s %-8s %-8s trans %-5s %s | registros %d | convertidos %d | "
                     "excepcion %d | GTQ sin regla %d"
                     % (fran, oper, sentido, cod, moneda, reg, conv, exc, gtq))
        if exc and moneda == 'USD':
            avisos.append("%s/%s/%s transaccion %s: %d registros estan en la lista de "
                          "excepciones pero su moneda dice USD (no se convirtieron; revisar)"
                          % (fran, oper, sentido, cod, exc))
        if gtq:
            avisos.append("%s/%s/%s transaccion %s: %d registros vienen en GTQ y no estan en "
                          "excepciones_gtq (no se convirtieron; revisar si es correcto)"
                          % (fran, oper, sentido, cod, gtq))
    log.escribir("")
    log.escribir("AVISOS DE MONEDA (%d)" % len(avisos))
    if avisos:
        for a in avisos:
            log.escribir("  - %s" % a)
    else:
        log.escribir("  Ninguno")


def ejecutar(opc, ruta_config, log):
    """Hace todo el recorrido y devuelve el codigo de salida (0 ok, 1 con errores)."""
    inicio = time.time()
    fecha_ini = ahora()

    entradas = opc['entradas']
    salida = opc['salida']

    log.escribir("=" * 70)
    log.escribir("PROCESO DE FORMATEO DE CDR")
    log.escribir("=" * 70)
    log.escribir("Inicio  : %s" % fecha_ini)
    log.escribir("Config  : %s" % ruta_config)
    if len(entradas) == 1:
        log.escribir("Entrada : %s" % entradas[0])
    else:
        log.escribir("Entradas (%d):" % len(entradas))
        for e in entradas:
            log.escribir("  - %s" % e)
    log.escribir("Salida  : %s" % salida)
    log.escribir("Modo    : %s" % opc['modo'])
    if opc['convertir']:
        log.escribir("Moneda  : USD -> GTQ a tipo de cambio %s (montos con IVA incluido, sin sumar IVA)"
                     % opc['tipo_cambio'])
        log.escribir("Decimales de Val_lla: %d (se TRUNCA, no se redondea)" % opc['dec_monto'])
        log.escribir("Sentido : SDIR-R = %s | SDIR-S = %s"
                     % (opc['sentidos']['R'], opc['sentidos']['S']))
        if opc['excepciones']:
            log.escribir("Excepciones (ya en GTQ, no se convierten):")
            for fran, oper, trans, sentido in opc['excepciones']:
                log.escribir("  - %s / %s / transaccion %d / %s" % (fran, oper, trans, sentido))
        else:
            log.escribir("Excepciones: (ninguna)")
    else:
        log.escribir("Moneda  : sin conversion")
    log.escribir("Excluye : %s" % (', '.join(opc['excluir']) if opc['excluir'] else '(ninguna)'))
    log.escribir("Procesos: %d" % opc['procesos'])
    log.escribir("Log     : %s" % log.ruta)
    log.escribir("")

    est = {'ok': 0, 'err': 0, 'omit': 0, 'reg': 0, 'conv': 0,
           'errores': [], 'por_carpeta': [], 'det': {}, 'avisos_nombre': []}

    try:
        if not os.path.isdir(salida):
            os.makedirs(salida)
        listo = True
    except Exception as e:
        msg = "no se pudo crear la carpeta de salida %s: %s" % (salida, e)
        log.escribir("[%s] ERROR %s" % (hora(), msg))
        est['errores'].append(('(configuracion)', msg))
        listo = False

    if listo:
        items = planificar(opc)
        tareas = [(it[3], it[4], it[5]) for it in items
                  if it[0] == 'archivo' and not it[6]]

        n_proc = min(opc['procesos'], len(tareas)) if tareas else 1
        pool = None
        try:
            if n_proc > 1:
                pool = multiprocessing.Pool(processes=n_proc,
                                            initializer=_iniciar_trabajador,
                                            initargs=(opc,))
                resultados = pool.imap(trabajar, tareas, 1)   # ordenado, de 1 en 1
            else:
                # Modo secuencial (procesos = 1): misma ruta de codigo, sin Pool
                _iniciar_trabajador(opc)
                resultados = _IteradorSecuencial(tareas)
            consumir(items, resultados, opc, log, est)
            if pool is not None:
                pool.close()
                pool.join()
        except BaseException:
            if pool is not None:
                pool.terminate()
                pool.join()
            raise

    fin = time.time()

    if opc['convertir']:
        escribir_detalle_moneda(log, est)
        if est['avisos_nombre']:
            log.escribir("")
            log.escribir("ARCHIVOS CUYO NOMBRE NO TIENE FRAN-/NOP-/SDIR- (sin excepciones): %d"
                         % len(est['avisos_nombre']))
            for n in est['avisos_nombre']:
                log.escribir("  - %s" % n)

    log.escribir("")
    log.escribir("=" * 70)
    log.escribir("RESUMEN")
    log.escribir("=" * 70)
    log.escribir("Inicio  : %s" % fecha_ini)
    log.escribir("Fin     : %s" % ahora())
    log.escribir("Duracion: %d segundos" % int(fin - inicio))
    log.escribir("")
    log.escribir("Carpetas procesadas: %d" % len(est['por_carpeta']))
    for nombre_carpeta, n_ok, n_reg in est['por_carpeta']:
        log.escribir("  %s: %d archivos, %d registros" % (nombre_carpeta, n_ok, n_reg))
    log.escribir("")
    log.escribir("Archivos OK: %d  |  Omitidos: %d  |  Con error: %d  |  Registros: %d"
                 % (est['ok'], est['omit'], est['err'], est['reg']))
    if opc['convertir']:
        log.escribir("Registros convertidos de USD a GTQ: %d de %d"
                     % (est['conv'], est['reg']))
    log.escribir("")
    log.escribir("ERRORES (%d)" % len(est['errores']))
    if est['errores']:
        for archivo, mensaje in est['errores']:
            log.escribir("  - %s: %s" % (archivo, mensaje))
    else:
        log.escribir("  Ninguno")
    log.escribir("=" * 70)

    return 1 if est['errores'] else 0


class _IteradorSecuencial(object):
    """Imita pool.imap(...).next(timeout) pero procesando en el mismo proceso."""

    def __init__(self, tareas):
        self.it = iter(tareas)

    def next(self, timeout=None):
        return trabajar(next(self.it))


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    ruta_config = sys.argv[1] if len(sys.argv) > 1 else os.path.join(base_dir, 'config.txt')

    try:
        opc = leer_config(ruta_config, base_dir)
    except Exception as e:
        print("ERROR de configuracion: %s" % e)
        return 1

    try:
        if not os.path.isdir(opc['dir_log']):
            os.makedirs(opc['dir_log'])
        ruta_log = os.path.join(opc['dir_log'],
                                'formatear_cdr_%s.log' % time.strftime('%Y%m%d_%H%M%S'))
        log = Log(ruta_log)
    except Exception as e:
        print("ERROR: no se pudo crear el archivo de log en %s: %s" % (opc['dir_log'], e))
        return 1

    try:
        return ejecutar(opc, ruta_config, log)
    finally:
        log.cerrar()


if __name__ == '__main__':
    multiprocessing.freeze_support()   # necesario en Windows (spawn / .exe)
    sys.exit(main())
