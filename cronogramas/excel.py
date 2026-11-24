# -*- coding: utf-8 -*-
"""El Excel de seguimiento: se arma completo en cada ejecución desde la base local.

Las columnas con encabezado VERDE las llena el equipo (nombre corto, responsable, estado de gestión,
observaciones, hito gestionado). Antes de rehacer el archivo, el programa las lee y las guarda en la
base (con su historia), y las vuelve a poner en el Excel nuevo.
"""
import datetime as dt

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from . import __version__, reglas
from .catalogo import CODIGOS, CUMPLIDO, EJECUCION, ETAPAS, HITOS, SIN_FECHA_ANTERIOR

AZUL = PatternFill("solid", fgColor="1F4E78")
VERDE_CAB = PatternFill("solid", fgColor="375623")
VERDE_CELDA = PatternFill("solid", fgColor="F3F9EE")
GRIS = PatternFill("solid", fgColor="F2F2F2")
BLANCO = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
NEGRITA = Font(name="Calibri", size=10, bold=True)
NORMAL = Font(name="Calibri", size=10)
ROJA = Font(name="Calibri", size=10, color="C00000", bold=True)
ENLACE = Font(name="Calibri", size=10, color="0563C1", underline="single")
TITULO = Font(name="Calibri", size=14, bold=True, color="1F4E78")
BORDE = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
PESOS = '"$" #,##0'
FECHA = "dd/mm/yyyy"
DIAS = ("lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo")
DIAS_CORTOS = ("lun", "mar", "mié", "jue", "vie", "sáb", "dom")
COLOR_ETAPA = {"Pre-pliego": "DDEBF7", "Pliego definitivo": "9BC2E6", "Ofertas y evaluación": "FFE699",
               "Adjudicación": "F8CBAD", "Contrato": "C6E0B4", CUMPLIDO: "D9D9D9"}
ALERTAS = [(0, "Hoy", "F8CBAD"), (3, "En 1 a 3 días", "FCE4D6"), (7, "En 4 a 7 días", "FFF2CC"),
           (15, "En 8 a 15 días", "E2EFDA"), (10 ** 6, "Más adelante", None)]
ESTADOS_GESTION = ("Pendiente", "En revisión", "Gestionado", "No aplica")
PROC_GESTION = [("Nombre corto", "nombre_corto"), ("Responsable", "responsable"),
                ("Estado de gestión", "estado"), ("Observaciones de gestión", "observacion")]
HITO_GESTION = [("Gestionado", "estado"), ("Observación del hito", "observacion")]
CLAVE_PROC, CLAVE_CAMPO, CLAVE_FECHA = "ID del expediente", "Código del hito", "Fecha y hora (clave)"


def relleno(color):
    return PatternFill("solid", fgColor=color) if color else None


def alerta(dias):
    for tope, txt, color in ALERTAS:
        if dias <= tope:
            return txt, relleno(color)


def fecha(fh):
    return dt.datetime.strptime(fh[:10], "%Y-%m-%d") if fh else None


def hora(fh):
    return fh[11:16] if fh and len(fh) >= 16 else None


def fmt(fh):
    return f"{fh[8:10]}/{fh[5:7]}/{fh[:4]} {fh[11:16]}" if fh else None


def recortar(t, n=30):
    t = " ".join(str(t or "").split())
    if len(t) <= n:
        return t
    return t[:n].rsplit(" ", 1)[0] + "…"


def enlace(u, texto="Ver en SECOP II"):
    return {"url": u, "texto": texto} if u else None


# ───────────────────────── datos de cada proceso ─────────────────────────
def armar(almacen, ahora, gestion):
    """Todo lo que necesitan las hojas, calculado una vez."""
    hoy = ahora.date().isoformat()
    m = ahora.strftime("%Y-%m-%d %H:%M")
    out = {}
    for k, p in almacen.procesos(solo_seguimiento=True).items():
        hs = []
        for h in almacen.cronograma(k):
            nombre, etapa, actor, orden, cod = HITOS.get(h["campo"], (h["campo"], "Otro", "—", 99, ""))
            ant = [] if h["campo"] in SIN_FECHA_ANTERIOR else h["anteriores"]
            hs.append(dict(h, nombre=nombre, etapa=etapa, actor=actor, orden=orden, cod=cod or "•", anteriores=ant))
        hs.sort(key=lambda h: (h["fecha_hora"], h["orden"]))
        propios = [h for h in hs if h["etapa"] != EJECUCION]
        pend = [h for h in propios if h["fecha_hora"][:10] >= hoy]

        def primero(*cods):
            for c in cods:
                for h in propios:
                    if h["cod"] == c:
                        return h["fecha_hora"]
            return None

        g = gestion.get((k, "", ""), {})
        es_estudio = (p["modalidad"] or "").startswith("Solicitud de información")
        out[k] = dict(
            p=p, hitos=hs, propios=propios, pendientes=pend, es_estudio=es_estudio,
            etapa=reglas.etapa_en([(h["fecha_hora"], h["campo"]) for h in hs], ahora),
            proximo=next((h for h in pend if h["fecha_hora"] >= m), pend[0] if pend else None),
            ultimo=propios[-1] if propios else None,
            apertura=primero("AP"), cierre=primero("C"), informe=primero("IE"), adjudicacion=primero("A", "AA"),
            firma=primero("F"), cambios=sum(1 for h in hs if h["anteriores"]),
            adendas=almacen.adendas(k), gestion=g,
            etiqueta=(p["codigo"] or "") + (" · " + g["nombre_corto"] if g.get("nombre_corto") else ""))
    return out


# ───────────────────────── hojas ─────────────────────────
def hoja(wb, nombre, titulo, notas, cols, filas, colores=None, fuentes=None, validaciones=None, ancho_alto=None):
    """cols: [(encabezado, ancho, formato, tipo)] con tipo None · 'editable' · 'oculta'.
    Un valor {'url','texto'} se escribe como enlace."""
    ws = wb.create_sheet(nombre)
    n = len(cols)
    c = ws.cell(row=1, column=1, value=titulo)
    c.font, c.fill = BLANCO, AZUL
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(n, 2))
    f = 2
    for nota in notas:
        c = ws.cell(row=f, column=1, value=nota)
        c.font, c.fill = NORMAL, GRIS
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=f, start_column=1, end_row=f, end_column=max(n, 2))
        ws.row_dimensions[f].height = 15 * (len(nota) // 170 + 1)
        f += 1
    for j, (h, ancho, _, tipo) in enumerate(cols, 1):
        c = ws.cell(row=f, column=j, value=h)
        c.font, c.border = BLANCO, BORDE
        c.fill = VERDE_CAB if tipo == "editable" else AZUL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = ancho
        if tipo == "oculta":
            ws.column_dimensions[get_column_letter(j)].hidden = True
    ws.row_dimensions[f].height = 36
    ws.freeze_panes = ws.cell(row=f + 1, column=1)
    ws.auto_filter.ref = f"A{f}:{get_column_letter(n)}{f + max(len(filas), 1)}"
    for i, fila in enumerate(filas):
        r = f + 1 + i
        for j, v in enumerate(fila, 1):
            if isinstance(v, dict):
                c = ws.cell(row=r, column=j, value=v["texto"])
                c.hyperlink, c.font = v["url"], ENLACE
            else:
                c = ws.cell(row=r, column=j, value=v)
                c.font = NORMAL
            c.border = BORDE
            c.alignment = Alignment(vertical="top", wrap_text=isinstance(v, str) and len(v) > 40)
            if cols[j - 1][2]:
                c.number_format = cols[j - 1][2]
            if cols[j - 1][3] == "editable":
                c.fill = VERDE_CELDA
            elif colores and colores[i]:
                c.fill = colores[i]
            if fuentes and fuentes[i] and fuentes[i].get(j):
                c.font = fuentes[i][j]
    if not filas:
        ws.cell(row=f + 1, column=1, value="(ninguno en esta actualización)").font = NORMAL
    for j, opciones in (validaciones or {}).items():
        dv = DataValidation(type="list", formula1='"' + ",".join(opciones) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{get_column_letter(j)}{f + 1}:{get_column_letter(j)}{f + max(len(filas), 1) + 200}")
    return ws


def gest_cols():
    return [(h, 22 if c != "observacion" else 40, None, "editable") for h, c in PROC_GESTION]


def gest_vals(g):
    return [g.get(c) for _, c in PROC_GESTION]


def construir(ruta, almacen, eid, ahora, cfg, avisos):
    gestion = almacen.gestion_actual()
    datos = armar(almacen, ahora, gestion)
    hoy = ahora.date()
    hoy_s = hoy.isoformat()
    desde = (hoy - dt.timedelta(days=cfg.dias_abiertos_sin_fechas)).isoformat()
    en_curso = sorted((k for k, d in datos.items() if d["pendientes"] and not d["es_estudio"]),
                      key=lambda k: datos[k]["pendientes"][0]["fecha_hora"])
    estudios = sorted((k for k, d in datos.items() if d["pendientes"] and d["es_estudio"]),
                      key=lambda k: datos[k]["pendientes"][0]["fecha_hora"])
    sin_fechas = sorted((k for k, d in datos.items() if not d["pendientes"] and d["ultimo"]
                         and d["ultimo"]["fecha_hora"][:10] >= desde),
                        key=lambda k: datos[k]["ultimo"]["fecha_hora"], reverse=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    cuentas = {}

    # PRÓXIMOS HITOS
    filas, colores, fuentes = [], [], []
    for k in en_curso:
        d = datos[k]
        for h in d["pendientes"]:
            dias = (fecha(h["fecha_hora"]).date() - hoy).days
            txt, color = alerta(dias)
            g = gestion.get((k, h["campo"], h["fecha_hora"]), {})
            filas.append([txt, fecha(h["fecha_hora"]), hora(h["fecha_hora"]), DIAS[fecha(h["fecha_hora"]).weekday()],
                          dias, h["nombre"], h["etapa"], h["actor"], d["etiqueta"], d["p"]["modalidad"],
                          d["p"]["objeto"], d["p"]["estado"], " · ".join(fmt(x) for x in h["anteriores"]) or None,
                          g.get("estado"), g.get("observacion"), enlace(d["p"]["url"]), k, h["campo"], h["fecha_hora"]])
            colores.append(color)
            fuentes.append({8: ROJA} if h["actor"] == "Entidad" else None)
    orden = sorted(range(len(filas)), key=lambda i: (filas[i][18], filas[i][8]))
    filas, colores, fuentes = [filas[i] for i in orden], [colores[i] for i in orden], [fuentes[i] for i in orden]
    cuentas["hitos"] = len(filas)
    cuentas["hoy"] = sum(1 for f in filas if f[4] == 0)
    cuentas["hoy_entidad"] = sum(1 for f in filas if f[4] == 0 and f[7] == "Entidad")
    cuentas["semana"] = sum(1 for f in filas if f[4] <= 7)
    hoja(wb, "PROXIMOS_HITOS", f"Hitos por venir · actualizado {ahora:%d/%m/%Y %H:%M}",
         ["Una fila por hito pendiente, en orden de fecha y hora (hora de Bogotá). «Días que faltan» son días calendario. "
          "«Quién actúa» en rojo = le toca a la Entidad. Si la fecha cambió (adenda o ajuste del cronograma), la anterior "
          "está en «Fecha antes del cambio».",
          "Columnas verdes: las llena el equipo y se conservan en la siguiente actualización."],
         [("Alerta", 13, None, None), ("Fecha", 11, FECHA, None), ("Hora", 7, None, None), ("Día", 10, None, None),
          ("Días que faltan", 8, "0", None), ("Hito", 42, None, None), ("Etapa", 14, None, None),
          ("Quién actúa", 14, None, None), ("Proceso", 26, None, None), ("Modalidad", 20, None, None),
          ("Objeto", 50, None, None), ("Estado en SECOP II", 12, None, None), ("Fecha antes del cambio", 20, None, None),
          ("Gestionado", 11, None, "editable"), ("Observación del hito", 34, None, "editable"),
          ("Enlace", 14, None, None), (CLAVE_PROC, 18, None, "oculta"), (CLAVE_CAMPO, 18, None, "oculta"),
          (CLAVE_FECHA, 16, None, "oculta")],
         filas, colores, fuentes, validaciones={14: ("Sí", "No")})

    # TABLERO DIARIO
    fin = max((fecha(datos[k]["pendientes"][-1]["fecha_hora"]).date() for k in en_curso), default=hoy)
    dias, d0 = [], hoy
    con_hito = {fecha(h["fecha_hora"]).date() for k in en_curso for h in datos[k]["pendientes"]}
    while d0 <= fin:
        if d0.weekday() < 5 or d0 in con_hito:
            dias.append(d0)
        d0 += dt.timedelta(days=1)
    ws = wb.create_sheet("TABLERO_DIARIO")
    ws.cell(row=1, column=1, value=f"Tablero diario · {len(en_curso)} procesos · del {hoy:%d/%m/%Y} al {fin:%d/%m/%Y}"
            ).font = TITULO
    ws.cell(row=2, column=1, value="Color = etapa del proceso al final de ese día. Códigos: " + " · ".join(
        f"{c} {t}" for c, t in CODIGOS.items()) + ". Los sábados y domingos solo aparecen si hay un hito.").font = NORMAL
    for j, (et, col) in enumerate(COLOR_ETAPA.items(), 2):
        c = ws.cell(row=3, column=j * 2 - 2, value=et)
        c.fill, c.font = relleno(col), NORMAL
    ws.cell(row=5, column=1, value="Proceso").font = BLANCO
    ws.cell(row=5, column=1).fill = AZUL
    ws.column_dimensions["A"].width = 34
    for j, d in enumerate(dias, 2):
        c = ws.cell(row=5, column=j, value=f"{DIAS_CORTOS[d.weekday()]} {d:%d/%m}")
        c.font, c.fill, c.alignment = BLANCO, AZUL, Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = 7
    ws.row_dimensions[5].height = 30
    for i, k in enumerate(en_curso, 6):
        dd = datos[k]
        c = ws.cell(row=i, column=1, value=dd["etiqueta"] + ("" if dd["gestion"].get("nombre_corto") else
                                                          " · " + recortar(dd["p"]["nombre"]).capitalize()))
        c.font, c.border = NORMAL, BORDE
        if dd["p"]["url"]:
            c.hyperlink = dd["p"]["url"]
        for j, d in enumerate(dias, 2):
            fin_dia = dt.datetime.combine(d, dt.time(23, 59))
            et = reglas.etapa_en([(h["fecha_hora"], h["campo"]) for h in dd["hitos"]], fin_dia)
            cods = []
            for h in dd["propios"]:
                if fecha(h["fecha_hora"]).date() == d and h["cod"] not in cods:
                    cods.append(h["cod"])
            if "•" in cods and len(cods) > 1:
                cods.remove("•")
            c = ws.cell(row=i, column=j, value=" ".join(cods) or None)
            c.fill, c.border, c.font = relleno(COLOR_ETAPA.get(et)), BORDE, NEGRITA
            c.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "B6"

    # PROCESOS EN CURSO
    filas, colores = [], []
    for k in en_curso:
        d = datos[k]
        p, px = d["p"], d["proximo"]
        filas.append([p["codigo"], p["referencia"], p["modalidad"], p["objeto"], p["precio_base"], d["etapa"],
                      p["estado"], px["nombre"] if px else None, fecha(px["fecha_hora"]) if px else None,
                      hora(px["fecha_hora"]) if px else None, (fecha(px["fecha_hora"]).date() - hoy).days if px else None,
                      px["actor"] if px else None, fecha(d["apertura"]), fecha(d["cierre"]), hora(d["cierre"]),
                      fecha(d["informe"]), fecha(d["adjudicacion"]), fecha(d["firma"]),
                      p["proveedores_manifestaron"], p["respuestas"], len(d["adendas"]), d["cambios"],
                      *gest_vals(d["gestion"]), enlace(p["url"]), k])
        colores.append(None)
    cuentas["en_curso"] = len(filas)
    cuentas["presupuesto"] = sum(datos[k]["p"]["precio_base"] or 0 for k in en_curso)
    ws = hoja(wb, "PROCESOS_EN_CURSO", f"Procesos abiertos sin adjudicar, con hitos pendientes · {ahora:%d/%m/%Y %H:%M}",
              ["Un proceso por fila. «Etapa hoy» = la del último hito ya alcanzado. Fechas clave del cronograma vigente "
               "(vacías = esa modalidad no trae ese hito). «Adjudicación» = publicación del acto o, si no hay, la audiencia.",
               "Columnas verdes: las llena el equipo (el «Nombre corto» se usa en el tablero) y se conservan."],
              [("Código", 10, None, None), ("Proceso", 26, None, None), ("Modalidad", 20, None, None),
               ("Objeto", 50, None, None), ("Presupuesto oficial", 16, PESOS, None), ("Etapa hoy", 15, None, None),
               ("Estado en SECOP II", 12, None, None), ("Próximo hito", 36, None, None), ("Fecha", 11, FECHA, None),
               ("Hora", 7, None, None), ("Días que faltan", 8, "0", None), ("Quién actúa", 14, None, None),
               ("Acto de apertura", 11, FECHA, None), ("Cierre de ofertas", 11, FECHA, None),
               ("Hora de cierre", 8, None, None), ("Informe de evaluación", 11, FECHA, None),
               ("Adjudicación", 11, FECHA, None), ("Firma", 11, FECHA, None),
               ("Interesados que manifestaron", 11, "0", None), ("Respuestas en el proceso", 10, "0", None),
               ("Adendas", 8, "0", None), ("Fechas cambiadas", 8, "0", None), *gest_cols(),
               ("Enlace", 14, None, None), (CLAVE_PROC, 18, None, "oculta")],
              filas, validaciones={25: ESTADOS_GESTION})
    for i, k in enumerate(en_curso):
        c = ws.cell(row=ws.freeze_panes and int(ws.freeze_panes[1:]) + i, column=6)
        c.fill = relleno(COLOR_ETAPA.get(datos[k]["etapa"]))

    # EVOLUCIÓN
    ws = wb.create_sheet("EVOLUCION")
    etapas = list(ETAPAS) + [CUMPLIDO]
    ws.cell(row=1, column=1, value=f"Evolución diaria: procesos en cada etapa (los {len(en_curso)} en curso)").font = TITULO
    ws.cell(row=2, column=1, value="Según el cronograma vigente; si SECOP II cambia fechas, la curva cambia.").font = NORMAL
    for j, h in enumerate(["Día"] + etapas, 1):
        c = ws.cell(row=4, column=j, value=h)
        c.font, c.fill, c.alignment = BLANCO, AZUL, Alignment(wrap_text=True, horizontal="center")
        ws.column_dimensions[get_column_letter(j)].width = 13
    for i, d in enumerate(dias, 5):
        ws.cell(row=i, column=1, value=f"{d:%d/%m}").font = NORMAL
        fin_dia = dt.datetime.combine(d, dt.time(23, 59))
        cuenta = {e: 0 for e in etapas}
        for k in en_curso:
            cuenta[reglas.etapa_en([(h["fecha_hora"], h["campo"]) for h in datos[k]["hitos"]], fin_dia)] += 1
        for j, e in enumerate(etapas, 2):
            ws.cell(row=i, column=j, value=cuenta[e]).font = NORMAL
    if dias:
        g = BarChart()
        g.type, g.grouping, g.overlap = "col", "stacked", 100
        g.title, g.y_axis.title = "Procesos por etapa, día a día", "Procesos"
        g.add_data(Reference(ws, min_col=2, max_col=len(etapas) + 1, min_row=4, max_row=4 + len(dias)), titles_from_data=True)
        g.set_categories(Reference(ws, min_col=1, min_row=5, max_row=4 + len(dias)))
        for s, e in zip(g.series, etapas):
            s.graphicalProperties.solidFill = COLOR_ETAPA[e]
        g.height, g.width = 10, 26
        ws.add_chart(g, "I4")

    # CRONOGRAMA COMPLETO
    filas, colores = [], []
    for k in en_curso + estudios:
        d = datos[k]
        for h in d["hitos"]:
            pend = h["fecha_hora"][:10] >= hoy_s
            filas.append([d["etiqueta"], d["p"]["modalidad"], h["orden"], h["nombre"], h["etapa"], h["actor"],
                          h["cod"] if h["cod"] != "•" else None, fecha(h["fecha_hora"]), hora(h["fecha_hora"]),
                          "Pendiente" if pend else "Fecha ya pasó", " · ".join(fmt(x) for x in h["anteriores"]) or None,
                          enlace(d["p"]["url"]), k])
            colores.append(None if pend else GRIS)
    hoja(wb, "CRONOGRAMA_COMPLETO", "Cronograma completo de los procesos en curso y estudios de mercado",
         ["Todos los hitos, en orden. «Fecha ya pasó» solo dice que la fecha quedó atrás: SECOP II no publica aquí "
          "si la actuación se cumplió."],
         [("Proceso", 26, None, None), ("Modalidad", 20, None, None), ("Orden", 6, "0", None), ("Hito", 44, None, None),
          ("Etapa", 15, None, None), ("Quién actúa", 14, None, None), ("Código", 7, None, None),
          ("Fecha", 11, FECHA, None), ("Hora", 7, None, None), ("Situación", 13, None, None),
          ("Fecha antes del cambio", 20, None, None), ("Enlace", 14, None, None), (CLAVE_PROC, 18, None, "oculta")],
         filas, colores)

    # CAMBIOS DE FECHA
    filas = []
    for k in en_curso + estudios + sin_fechas:
        d = datos[k]
        for h in d["hitos"]:
            for a in h["anteriores"]:
                filas.append([d["etiqueta"], h["nombre"], fmt(a), fmt(h["fecha_hora"]),
                              "Se aplazó" if a < h["fecha_hora"] else "Se adelantó", len(d["adendas"]),
                              enlace(d["p"]["url"]), k])
    cuentas["cambios"] = len(filas)
    hoja(wb, "CAMBIOS_DE_FECHA", "Hitos cuya fecha cambió",
         ["Fecha anterior → fecha vigente. Sale de las versiones del cronograma que SECOP II guarda por fase y de las "
          "actualizaciones anteriores de este archivo. Puede venir de una adenda o de un ajuste del cronograma en la "
          "plataforma: no siempre hay una adenda publicada con ese nombre."],
         [("Proceso", 26, None, None), ("Hito", 44, None, None), ("Fecha anterior", 16, None, None),
          ("Fecha vigente", 16, None, None), ("Sentido", 12, None, None), ("Adendas del proceso", 9, "0", None),
          ("Enlace", 14, None, None), (CLAVE_PROC, 18, None, "oculta")], filas)

    # PUNTOS PARA REVISAR
    filas = []
    for k in en_curso:
        d = datos[k]
        by = {h["campo"]: h["fecha_hora"] for h in d["hitos"]}
        ap, ie = by.get("fecha_de_apertura_de_ofertas"), by.get("publicaci_n_del_informe_de")
        if ap and ie and ap > ie and ap[:10] >= hoy_s:
            filas.append([d["etiqueta"], "Orden de hitos", f"La «apertura de ofertas» ({fmt(ap)}) va después del informe "
                          f"de evaluación ({fmt(ie)}): suele ser la apertura del sobre económico o la subasta.",
                          "Confirmar en SECOP II", enlace(d["p"]["url"]), k])
        if d["firma"] and not d["adjudicacion"]:
            filas.append([d["etiqueta"], "Sin fecha de adjudicación", "El cronograma pasa a la firma del contrato "
                          f"({fmt(d['firma'])}) sin fecha de adjudicación.", "Confirmar en SECOP II",
                          enlace(d["p"]["url"]), k])
        finde = [h for h in d["pendientes"] if fecha(h["fecha_hora"]).weekday() >= 5]
        if finde:
            filas.append([d["etiqueta"], "Hito en fin de semana", "; ".join(f"{h['nombre']} ({fmt(h['fecha_hora'])})"
                          for h in finde[:3]), "Revisar si es día hábil", enlace(d["p"]["url"]), k])
        fin_ano = [h for h in d["pendientes"] if h["fecha_hora"][5:10] >= "12-23" or h["fecha_hora"][5:10] <= "01-06"]
        if fin_ano:
            filas.append([d["etiqueta"], "Fin de año", "; ".join(f"{h['nombre']} ({fmt(h['fecha_hora'])})"
                          for h in fin_ano[:4]), "Prever vacancias y cierre de vigencia", enlace(d["p"]["url"]), k])
        if d["cambios"] >= 3 or d["adendas"]:
            filas.append([d["etiqueta"], "Cronograma modificado", f"{d['cambios']} fechas cambiadas y "
                          f"{len(d['adendas'])} adendas; {d['p']['respuestas'] or 0} respuestas publicadas. Próximo: "
                          + (f"{d['proximo']['nombre']} ({fmt(d['proximo']['fecha_hora'])})" if d["proximo"] else "—"),
                          "Leer la última adenda", enlace(d["p"]["url"]), k])
    for k in en_curso:
        d = datos[k]
        c = next((h["fecha_hora"] for h in d["hitos"] if h["campo"] == "presentaci_n_de_ofertas"), None)
        rec = d["p"]["fecha_recepcion_ofertas"]
        if c and rec and c != rec:
            filas.append([d["etiqueta"], "Dos fechas de cierre", f"El cronograma dice {fmt(c)}; la ficha del proceso "
                          f"dice {fmt(rec)}.", "Vale lo que muestre la página del proceso", enlace(d["p"]["url"]), k])
    cuentas["revisar"] = len(filas)
    hoja(wb, "PUNTOS_PARA_REVISAR", "Puntos para revisar antes de actuar",
         ["Avisos automáticos sobre el cronograma publicado. No son errores confirmados: son cosas que conviene mirar "
          "en SECOP II."],
         [("Proceso", 26, None, None), ("Punto", 22, None, None), ("Detalle", 80, None, None),
          ("Qué hacer", 30, None, None), ("Enlace", 14, None, None), (CLAVE_PROC, 18, None, "oculta")], filas)

    # NOVEDADES
    nov = almacen.novedades(eid)
    previa = almacen.ultima_ejecucion_ok(eid)
    procs = almacen.procesos(solo_seguimiento=False)
    def f2(v):
        return fmt(v) if v and len(v) == 16 and v[4] == "-" and v[13] == ":" else v
    filas = [[n["tipo"], procs.get(n["id_portafolio"], {}).get("codigo"), n["detalle"], f2(n["antes"]), f2(n["despues"]),
              enlace(procs.get(n["id_portafolio"], {}).get("url")), n["id_portafolio"]] for n in nov]
    cuentas["novedades"] = len(filas)
    hoja(wb, "NOVEDADES", "Qué cambió desde la actualización anterior",
         [("Primera actualización: esta es la línea base; desde la próxima aparecerán aquí los procesos nuevos, las "
           "fechas cambiadas, las adendas y los procesos que salen (adjudicados, cancelados, desiertos).") if not previa
          else f"Comparado con la actualización del {previa['inicio'][:16]}."],
         [("Novedad", 26, None, None), ("Proceso", 12, None, None), ("Detalle", 60, None, None),
          ("Antes", 18, None, None), ("Ahora", 18, None, None), ("Enlace", 14, None, None),
          (CLAVE_PROC, 18, None, "oculta")], filas)

    # ABIERTOS SIN FECHAS
    filas = []
    for k in sin_fechas:
        d = datos[k]
        u = d["ultimo"]
        filas.append([d["p"]["codigo"], d["p"]["referencia"], d["p"]["modalidad"], d["p"]["objeto"], d["p"]["estado"],
                      d["p"]["fase"], u["nombre"], fecha(u["fecha_hora"]), (hoy - fecha(u["fecha_hora"]).date()).days,
                      *gest_vals(d["gestion"]), enlace(d["p"]["url"]), k])
    cuentas["sin_fechas"] = len(filas)
    hoja(wb, "ABIERTOS_SIN_FECHAS", f"Abiertos sin adjudicar y sin fechas pendientes (último hito en {cfg.dias_abiertos_sin_fechas} días)",
         ["SECOP II los muestra abiertos, pero su cronograma ya terminó: puede ser evaluación en curso, un paso que no "
          "está en el cronograma o un trámite pendiente en la plataforma."],
         [("Código", 10, None, None), ("Proceso", 26, None, None), ("Modalidad", 20, None, None), ("Objeto", 50, None, None),
          ("Estado en SECOP II", 12, None, None), ("Fase en SECOP II", 20, None, None),
          ("Último hito del cronograma", 36, None, None), ("Fecha", 11, FECHA, None), ("Días desde ese hito", 8, "0", None),
          *gest_cols(), ("Enlace", 14, None, None), (CLAVE_PROC, 18, None, "oculta")],
         filas, validaciones={12: ESTADOS_GESTION})

    # ESTUDIOS DE MERCADO
    filas = []
    for k in estudios:
        d = datos[k]
        px = d["proximo"]
        filas.append([d["p"]["codigo"], d["p"]["referencia"], d["p"]["objeto"], d["p"]["estado"],
                      px["nombre"] if px else None, fecha(px["fecha_hora"]) if px else None,
                      hora(px["fecha_hora"]) if px else None, enlace(d["p"]["url"]), k])
    cuentas["estudios"] = len(filas)
    hoja(wb, "ESTUDIOS_DE_MERCADO", "Solicitudes de información a proveedores con fechas pendientes",
         ["Son estudios de mercado, no procesos de selección: no van en el tablero."],
         [("Código", 10, None, None), ("Solicitud", 30, None, None), ("Objeto", 60, None, None),
          ("Estado", 12, None, None), ("Próximo hito", 30, None, None), ("Fecha", 11, FECHA, None),
          ("Hora", 7, None, None), ("Enlace", 14, None, None), (CLAVE_PROC, 18, None, "oculta")], filas)

    # ADENDAS
    filas = []
    for k in en_curso + sin_fechas:
        for a in datos[k]["adendas"]:
            filas.append([datos[k]["etiqueta"], a["nombre"], fecha(a["fecha_carga"]), enlace(a["url"], "Descargar"),
                          enlace(datos[k]["p"]["url"]), k])
    hoja(wb, "ADENDAS", "Adendas publicadas",
         ["Documentos del expediente con «adenda» en el nombre (incluye anexos de la adenda). Una modificación "
          "publicada con otro nombre no aparece aquí."],
         [("Proceso", 26, None, None), ("Documento", 70, None, None), ("Fecha de carga", 12, FECHA, None),
          ("Documento", 12, None, None), ("Proceso en SECOP II", 14, None, None), (CLAVE_PROC, 18, None, "oculta")], filas)

    # REGISTRO DE PROCESOS
    filas = []
    for k, p in sorted(procs.items(), key=lambda x: (-x[1]["en_seguimiento"], x[1]["primera_vez"] or "")):
        g = gestion.get((k, "", ""), {})
        filas.append(["Sí" if p["en_seguimiento"] else "No", p["codigo"], p["referencia"], p["modalidad"],
                      p["precio_base"], p["estado"], p["motivo_fuera"], (p["primera_vez"] or "")[:16],
                      (p["ultima_vez"] or "")[:16], (p["salio"] or "")[:16] or None,
                      p["proveedor_adjudicado"] if p["proveedor_adjudicado"] != "No Adjudicado" else None,
                      p["valor_adjudicado"] or None, g.get("responsable"), g.get("estado"), enlace(p["url"]), k])
    hoja(wb, "REGISTRO_PROCESOS", "Todos los procesos que han estado en seguimiento",
         ["Lo que guarda la base local: cuándo entró cada proceso, cuándo salió y por qué (adjudicado, cancelado, "
          "desierto), con el adjudicatario y el último estado de gestión del equipo."],
         [("En seguimiento", 10, None, None), ("Código", 10, None, None), ("Proceso", 26, None, None),
          ("Modalidad", 20, None, None), ("Presupuesto oficial", 16, PESOS, None), ("Estado en SECOP II", 12, None, None),
          ("Motivo de salida", 16, None, None), ("Primera vez", 16, None, None), ("Última consulta", 16, None, None),
          ("Salió", 16, None, None), ("Adjudicatario", 30, None, None), ("Valor adjudicado", 16, PESOS, None),
          ("Responsable", 18, None, None), ("Estado de gestión", 14, None, None), ("Enlace", 14, None, None),
          (CLAVE_PROC, 18, None, "oculta")], filas)

    # HISTORIAL DE GESTIÓN
    filas = [[g["momento"][:16], g["codigo"], HITOS.get(g["campo"], (g["campo"],))[0] if g["campo"] else "(proceso)",
              fmt(g["fecha_hora"]) if g["fecha_hora"] else None, g["nombre_corto"], g["responsable"], g["estado"],
              g["observacion"]] for g in almacen.historial_gestion()]
    hoja(wb, "HISTORIAL_GESTION", "Lo que ha anotado el equipo (más reciente primero)",
         ["Cada cambio en las columnas verdes queda aquí con la fecha de la actualización en que se guardó."],
         [("Guardado", 16, None, None), ("Proceso", 12, None, None), ("Hito", 40, None, None),
          ("Fecha del hito", 16, None, None), ("Nombre corto", 20, None, None), ("Responsable", 18, None, None),
          ("Estado", 14, None, None), ("Observación", 60, None, None)], filas)

    # REGISTRO DE EJECUCIONES
    filas = [[e["inicio"][:16], (e["fin"] or "")[:16] or None, e["estado"], e["version"], e["en_seguimiento"],
              e["filas_procesos"], e["filas_cronograma"], e["adendas"], e["novedades"], e["consultas"],
              e["avisos"]] for e in almacen.ejecuciones()]
    hoja(wb, "REGISTRO_EJECUCIONES", "Cada actualización del archivo",
         ["Qué trajo cada consulta a SECOP II. «Avisos» incluye hitos que SECOP agregó y el programa no conoce."],
         [("Inicio", 16, None, None), ("Fin", 16, None, None), ("Resultado", 11, None, None), ("Versión", 8, None, None),
          ("En seguimiento", 10, "0", None), ("Filas de procesos", 10, "0", None), ("Cronogramas", 10, "0", None),
          ("Adendas", 8, "0", None), ("Novedades", 9, "0", None), ("Consultas a la API", 9, "0", None),
          ("Avisos", 60, None, None)], filas)

    resumen(wb, datos, en_curso, cuentas, ahora, cfg, len(estudios))
    orden = ["RESUMEN", "NOVEDADES", "PROXIMOS_HITOS", "TABLERO_DIARIO", "PROCESOS_EN_CURSO", "EVOLUCION",
             "PUNTOS_PARA_REVISAR", "CAMBIOS_DE_FECHA", "CRONOGRAMA_COMPLETO", "ABIERTOS_SIN_FECHAS",
             "ESTUDIOS_DE_MERCADO", "ADENDAS", "REGISTRO_PROCESOS", "HISTORIAL_GESTION", "REGISTRO_EJECUCIONES"]
    wb._sheets = [wb[n] for n in orden]
    wb.active = 0
    ruta.parent.mkdir(parents=True, exist_ok=True)
    wb.save(ruta)
    return cuentas


def resumen(wb, datos, en_curso, cuentas, ahora, cfg, n_estudios):
    ws = wb.create_sheet("RESUMEN")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 100
    ws.cell(row=1, column=1, value=f"Procesos de selección · {cfg.nombre_mostrar}").font = TITULO
    ws.cell(row=2, column=1, value=f"Actualizado el {ahora:%d/%m/%Y a las %H:%M} con los datos abiertos de SECOP II "
            "(datos.gov.co)").font = NORMAL
    r = 4

    def titulo(t):
        nonlocal r
        c = ws.cell(row=r, column=1, value=t)
        c.font, c.fill = BLANCO, AZUL
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        r += 1

    def linea(a, b=None, cc=None, color=None, fuente=NORMAL):
        nonlocal r
        x = ws.cell(row=r, column=1, value=a)
        x.font = NEGRITA
        if color:
            x.fill = relleno(color)
        ws.cell(row=r, column=2, value=b).font = fuente
        y = ws.cell(row=r, column=3, value=cc)
        y.font, y.alignment = NORMAL, Alignment(wrap_text=True, vertical="top")
        r += 1

    titulo("Hoy: dónde está cada proceso (etapa del último hito alcanzado)")
    for e in list(ETAPAS) + [CUMPLIDO]:
        ks = [k for k in en_curso if datos[k]["etapa"] == e]
        linea(e, len(ks), " · ".join(datos[k]["etiqueta"] for k in ks) or "Ninguno", COLOR_ETAPA[e])
    r += 1
    titulo("En cifras")
    linea("Procesos en curso", cuentas["en_curso"], "Abiertos y sin adjudicar, con hitos pendientes.")
    linea("Presupuesto oficial en juego", cuentas["presupuesto"])
    ws.cell(row=r - 1, column=2).number_format = PESOS
    grandes = sorted(en_curso, key=lambda k: -(datos[k]["p"]["precio_base"] or 0))[:3]
    linea("Los tres más grandes", None, " · ".join(
        f"{datos[k]['etiqueta']}: ${datos[k]['p']['precio_base'] or 0:,.0f}".replace(",", ".") for k in grandes))
    linea("Hitos hoy", cuentas["hoy"], f"{cuentas['hoy_entidad']} a cargo de la Entidad.")
    linea("Hitos en los próximos 7 días", cuentas["semana"])
    linea("Novedades desde la última actualización", cuentas["novedades"], "Ver hoja NOVEDADES.")
    linea("Puntos para revisar", cuentas["revisar"], "Ver hoja PUNTOS_PARA_REVISAR.")
    linea("Fuera del tablero", None, f"{n_estudios} estudios de mercado con fechas pendientes · "
          f"{cuentas['sin_fechas']} abiertos sin fechas pendientes.")
    r += 1
    titulo("Cómo se usa")
    for a, b in [
        ("Actualizar", "Doble clic en «Actualizar cronogramas.bat». Cerrar este archivo antes: el programa lo rehace."),
        ("Columnas verdes", "Nombre corto, responsable, estado de gestión, observaciones y «gestionado» de cada hito: "
                            "se escriben aquí y se conservan en la siguiente actualización (historia en HISTORIAL_GESTION)."),
        ("Qué entra", "Procesos publicados (desde el proyecto de pliego) hasta antes de la adjudicación: estado "
                      "Publicado, Abierto o Evaluación en SECOP II. No entran adjudicados, cancelados, desiertos ni borradores."),
        ("Hora", "Hora de Bogotá. SECOP II publica el cronograma en hora universal; aquí ya está convertida."),
        ("Fecha antes del cambio", "SECOP II guarda una versión del cronograma por cada fase del proceso; rige la más "
                                   "reciente. Las anteriores muestran la fecha que tenía el hito antes de una adenda o ajuste."),
        ("Lo que no trae", "El cronograma publicado no incluye todos los hitos del pliego (p. ej. el informe final de "
                           "evaluación). Antes de actuar, confirmar en la página del proceso en SECOP II."),
        ("Copias", "Cada actualización deja una copia fechada en la carpeta «historial»."),
    ]:
        linea(a, None, b)
    r += 1
    titulo("Alertas y códigos")
    for _, t, col in ALERTAS[:-1]:
        linea(t, None, "Días calendario desde hoy.", col)
    linea("Códigos del tablero", None, " · ".join(f"{c} = {t}" for c, t in CODIGOS.items()))
    ws.cell(row=r + 1, column=1, value=f"Versión {__version__}").font = Font(name="Calibri", size=8, color="808080")


# ───────────────────────── leer lo que anotó el equipo ─────────────────────────
def leer_gestion(ruta):
    """[(expediente, campo, fecha_hora, {valores})] de las columnas verdes del Excel anterior."""
    if not ruta.exists():
        return []
    wb = openpyxl.load_workbook(ruta, data_only=True)
    out = []
    for nombre, cols, nivel in (("PROCESOS_EN_CURSO", PROC_GESTION, "proceso"),
                                ("ABIERTOS_SIN_FECHAS", PROC_GESTION, "proceso"),
                                ("PROXIMOS_HITOS", HITO_GESTION, "hito")):
        if nombre not in wb.sheetnames:
            continue
        ws = wb[nombre]
        cab, fila_cab = None, None
        for i, fila in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), 1):
            if CLAVE_PROC in fila:
                cab, fila_cab = [str(x or "") for x in fila], i
                break
        if not cab:
            continue
        idx = {h: cab.index(h) for h, _ in cols if h in cab}
        for fila in ws.iter_rows(min_row=fila_cab + 1, values_only=True):
            k = fila[cab.index(CLAVE_PROC)] if len(fila) > cab.index(CLAVE_PROC) else None
            if not k:
                continue
            vals = {c: fila[idx[h]] for h, c in cols if h in idx}
            campo = fila[cab.index(CLAVE_CAMPO)] if nivel == "hito" and CLAVE_CAMPO in cab else ""
            fh = fila[cab.index(CLAVE_FECHA)] if nivel == "hito" and CLAVE_FECHA in cab else ""
            out.append((k, campo, fh, vals))
    wb.close()
    return out
