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

Uso:
    python formatear_cdr.py [ruta_config.txt]

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

    return {
        'entradas': entradas,
        'envolver': envolver,
        'salida': os.path.normpath(salida),
        'dir_log': os.path.normpath(dir_log),
        'excluir': [x.upper() for x in a_lista(obtener(cfg, 'carpetas', 'excluir_carpetas', ''))],
        'extensiones': [x.lower() for x in a_lista(obtener(cfg, 'archivos', 'extensiones', '.csv'))],
        'sobrescribir': a_bool(obtener(cfg, 'archivos', 'sobrescribir', 'si')),
        'separador': SEPARADORES[nombre_sep],
        'modo': modo,
        'dec_monto': dec_monto,
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
    if '.' in v:
        ent, dec = v.split('.', 1)
    else:
        ent, dec = v, ''
    dec = dec.rstrip('0')
    return ent if not dec else ent + '.' + dec


def truncar(valor, decimales):
    """Trunca (no redondea) un Decimal a N decimales y quita ceros finales."""
    q = Decimal(1).scaleb(-decimales)
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
    n = opc['dec_monto']
    if n is not None:
        return truncar(Decimal(v), n)
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
        plan.append((titulo, pos[origen], fn))
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
    f_in = open(ruta_in, 'rb')
    f_out = None
    writer = None
    plan = []
    num_cols = 0
    escritos = 0
    convertidos = 0
    idx_monto = idx_mon = idx_cod = None
    declarado = None
    primera = True

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
                if opc['convertir']:
                    pos = dict((c.strip().upper(), i) for i, c in enumerate(columnas))
                    for req in ('AMOUNT', 'CURRENCY', 'RATING_COMPONENT'):
                        if req not in pos:
                            raise ValueError("para convertir a GTQ falta la columna %s" % req)
                    idx_monto = pos['AMOUNT']
                    idx_mon = pos['CURRENCY']
                    idx_cod = pos['RATING_COMPONENT']
                f_out = open(ruta_tmp, 'wb')
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
                if opc['convertir']:
                    n_reg = escritos + 1
                    moneda = datos[idx_mon].strip().upper()
                    if moneda not in ('USD', 'GTQ'):
                        raise ValueError("moneda no reconocida '%s' en el registro %d "
                                         "(solo USD o GTQ)" % (moneda, n_reg))
                    cod = sin_ceros_izq(datos[idx_cod])
                    cod_num = int(cod) if cod.isdigit() else -1
                    v = datos[idx_monto].strip()
                    clave = (meta[0], meta[1], meta[2], cod, moneda)
                    d = det.setdefault(clave, [0, 0, 0, 0])
                    d[0] += 1
                    if es_excepcion(meta, cod_num, opc):
                        # Ya viene en quetzales desde ICT: no se altera
                        monto_txt = quitar_ceros_finales(v) if _es_numero(v) else v
                        d[2] += 1
                    elif moneda == 'USD':
                        try:
                            monto_txt = truncar(Decimal(v) * opc['tipo_cambio'],
                                                opc['dec_monto'])
                        except Exception:
                            raise ValueError("monto no numerico '%s' en el registro %d"
                                             % (v, n_reg))
                        d[1] += 1
                        convertidos += 1
                    else:
                        # GTQ fuera de la lista de excepciones: ya esta en quetzales
                        monto_txt = quitar_ceros_finales(v) if _es_numero(v) else v
                        d[3] += 1
                    datos[idx_mon] = 'GTQ'
                salida = []
                for titulo, i, fn in plan:
                    if monto_txt is not None and i == idx_monto:
                        salida.append(monto_txt)
                    else:
                        salida.append(fn(datos[i], opc) if fn else datos[i])
                writer.writerow(salida)
                escritos += 1
            # AARHDR, AARTAB, AARSIZ, AARTYP, AARDEC, AARTRA: se ignoran

        if writer is None:
            raise ValueError("no se encontro la linea AARCOL con los nombres de columna")
    except Exception:
        f_in.close()
        if f_out is not None:
            f_out.close()
        if os.path.exists(ruta_tmp):
            os.remove(ruta_tmp)
        raise

    f_in.close()
    f_out.close()
    os.rename(ruta_tmp, ruta_out)
    return escritos, declarado, convertidos


# ----------------------------------------------------------------------------
# Recorrido de carpetas
# ----------------------------------------------------------------------------
def procesar_raiz(entrada, opc, log, est):
    """Recorre una carpeta de entrada (con subcarpetas) y escribe en SALIDA."""
    salida = opc['salida']
    # Con varias entradas, cada una va en su propia carpeta dentro de SALIDA
    prefijo = os.path.basename(entrada) if opc['envolver'] else ''
    destino_base = os.path.join(salida, prefijo) if prefijo else salida

    if not os.path.isdir(entrada):
        msg = "la carpeta de entrada no existe: %s" % entrada
        log.escribir("[%s] ERROR %s" % (hora(), msg))
        est['errores'].append((entrada, msg))
        return
    if entrada == salida:
        msg = "entrada y salida no pueden ser la misma carpeta: %s" % entrada
        log.escribir("[%s] ERROR %s" % (hora(), msg))
        est['errores'].append((entrada, msg))
        return

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
            if d.upper() in opc['excluir']:
                log.escribir("[%s] EXCLUIDA por config: %s" %
                             (hora(), os.path.join(rel_total, d) if rel_total else d))
        dirs[:] = sorted([d for d in dirs
                          if d.upper() not in opc['excluir']
                          and os.path.normpath(os.path.join(raiz, d)) != salida])

        destino_dir = os.path.join(destino_base, rel) if rel else destino_base

        candidatos = [n for n in sorted(archivos)
                      if os.path.splitext(n)[1].lower() in opc['extensiones']]
        if not candidatos:
            continue

        log.escribir("-" * 70)
        log.escribir("Carpeta: %s  (%d archivos)" % (nombre_carpeta, len(candidatos)))
        log.escribir("-" * 70)

        ok_carpeta = 0
        reg_carpeta = 0

        for nombre in candidatos:
            ruta_in = os.path.join(raiz, nombre)
            etiqueta = os.path.join(rel_total, nombre) if rel_total else nombre

            # El archivo de salida conserva el nombre del archivo de entrada
            aviso = ''
            ruta_out = os.path.join(destino_dir, nombre)

            if os.path.exists(ruta_out) and not opc['sobrescribir']:
                log.escribir("[%s] OMIT  %s (ya existe en salida)" % (hora(), nombre))
                est['omit'] += 1
                continue

            try:
                if not os.path.isdir(destino_dir):
                    os.makedirs(destino_dir)
                escritos, declarado, convertidos = procesar_archivo(
                    ruta_in, ruta_out, opc, est['det'])
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
                ok_carpeta += 1
                reg_carpeta += escritos
            except NoEsCDR as e:
                log.escribir("[%s] OMIT  %s (%s)" % (hora(), nombre, e))
                est['omit'] += 1
            except Exception as e:
                log.escribir("[%s] ERROR %s: %s" % (hora(), nombre, e))
                est['errores'].append((etiqueta, str(e)))
                est['err'] += 1

        if ok_carpeta:
            est['por_carpeta'].append([nombre_carpeta, ok_carpeta, reg_carpeta])


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
        for entrada in entradas:
            procesar_raiz(entrada, opc, log, est)

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
    sys.exit(main())
