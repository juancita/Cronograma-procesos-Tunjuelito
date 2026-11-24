# -*- coding: utf-8 -*-
"""Base local (SQLite, datos/cronogramas.db): guarda lo consultado y lo que hace el equipo.

  ejecucion     cada vez que se corre el programa (cuándo, qué trajo, qué cambió, avisos)
  proceso       cada expediente que alguna vez estuvo en seguimiento, con su último estado
  fase          las filas CO1.REQ de cada expediente
  hito          cada valor que tuvo cada hito en cada fila (primera y última vez que se vio)
  hito_vigente  el cronograma que regía en cada momento (desde / hasta): de aquí salen los cambios de fecha
  adenda        documentos del expediente con «adenda» en el nombre
  novedad       lo que cambió de una ejecución a la siguiente
  gestion       lo que el equipo anota en el Excel (responsable, estado, observaciones), con su historia
Nada se borra: lo viejo queda con su fecha.
"""
import json
import sqlite3

from . import reglas
from .catalogo import HITOS, NO_HITO

ESQUEMA = """
CREATE TABLE IF NOT EXISTS ejecucion (
    id INTEGER PRIMARY KEY, inicio TEXT NOT NULL, fin TEXT, version TEXT, estado TEXT,
    filas_procesos INTEGER, filas_cronograma INTEGER, adendas INTEGER, en_seguimiento INTEGER,
    novedades INTEGER, consultas INTEGER, avisos TEXT);
CREATE TABLE IF NOT EXISTS proceso (
    id_portafolio TEXT PRIMARY KEY, referencia TEXT, codigo TEXT, nombre TEXT, objeto TEXT,
    modalidad TEXT, tipo_contrato TEXT, precio_base REAL, id_req_vigente TEXT, estado TEXT, fase TEXT,
    adjudicado TEXT, fecha_adjudicacion TEXT, valor_adjudicado REAL, proveedor_adjudicado TEXT,
    nit_proveedor TEXT, fecha_recepcion_ofertas TEXT, unidad_contratacion TEXT,
    proveedores_manifestaron INTEGER, respuestas INTEGER, url TEXT,
    en_seguimiento INTEGER NOT NULL DEFAULT 0, motivo_fuera TEXT,
    primera_vez TEXT, ultima_vez TEXT, salio TEXT);
CREATE TABLE IF NOT EXISTS fase (
    id_req TEXT PRIMARY KEY, id_portafolio TEXT NOT NULL, referencia TEXT, fase TEXT, estado TEXT,
    fecha_ultima_publicacion TEXT, url TEXT, ultima_vez TEXT);
CREATE TABLE IF NOT EXISTS hito (
    id INTEGER PRIMARY KEY, id_req TEXT NOT NULL, campo TEXT NOT NULL, fecha_hora TEXT NOT NULL,
    valor_fuente TEXT, primera_vez TEXT, ultima_vez TEXT, UNIQUE (id_req, campo, fecha_hora));
CREATE TABLE IF NOT EXISTS hito_vigente (
    id INTEGER PRIMARY KEY, id_portafolio TEXT NOT NULL, campo TEXT NOT NULL, fecha_hora TEXT NOT NULL,
    desde TEXT NOT NULL, hasta TEXT);
CREATE INDEX IF NOT EXISTS ix_hv ON hito_vigente (id_portafolio, campo, hasta);
CREATE TABLE IF NOT EXISTS adenda (
    id_documento TEXT PRIMARY KEY, id_portafolio TEXT NOT NULL, nombre TEXT, fecha_carga TEXT,
    url TEXT, primera_vez TEXT);
CREATE TABLE IF NOT EXISTS novedad (
    id INTEGER PRIMARY KEY, ejecucion_id INTEGER, momento TEXT, id_portafolio TEXT, tipo TEXT,
    detalle TEXT, antes TEXT, despues TEXT);
CREATE TABLE IF NOT EXISTS gestion (
    id INTEGER PRIMARY KEY, momento TEXT, ejecucion_id INTEGER, id_portafolio TEXT NOT NULL,
    campo TEXT, fecha_hora TEXT, nombre_corto TEXT, responsable TEXT, estado TEXT, observacion TEXT);
CREATE INDEX IF NOT EXISTS ix_gestion ON gestion (id_portafolio, campo, fecha_hora, id);
"""


def abrir(ruta):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(ruta)
    db.row_factory = sqlite3.Row
    db.executescript(ESQUEMA)
    return db


def txt(v):
    return v.strftime("%Y-%m-%d %H:%M:%S") if hasattr(v, "strftime") else v


def num(v):
    try:
        return float(v) if v not in (None, "") else None
    except (TypeError, ValueError):
        return None


def entero(v):
    n = num(v)
    return int(n) if n is not None else None


def url(v):
    return v.get("url") if isinstance(v, dict) else v


class Almacen:
    def __init__(self, db):
        self.db = db

    # ── ejecuciones ──
    def iniciar(self, momento, version):
        cur = self.db.execute("INSERT INTO ejecucion (inicio, version, estado) VALUES (?,?,?)",
                              (txt(momento), version, "en curso"))
        self.db.commit()
        return cur.lastrowid

    def terminar(self, eid, momento, estado, **cuentas):
        cuentas = {k: (json.dumps(v, ensure_ascii=False) if isinstance(v, list) else v) for k, v in cuentas.items()}
        sets = ", ".join(f"{k}=?" for k in cuentas)
        self.db.execute(f"UPDATE ejecucion SET fin=?, estado=?{', ' + sets if sets else ''} WHERE id=?",
                        (txt(momento), estado, *cuentas.values(), eid))
        self.db.commit()

    def ultima_ejecucion_ok(self, antes_de):
        r = self.db.execute("SELECT * FROM ejecucion WHERE estado='completa' AND id<? ORDER BY id DESC LIMIT 1",
                            (antes_de,)).fetchone()
        return dict(r) if r else None

    def expedientes_en_seguimiento(self):
        return [r[0] for r in self.db.execute("SELECT id_portafolio FROM proceso WHERE en_seguimiento=1")]

    def _novedad(self, eid, momento, k, tipo, detalle, antes=None, despues=None):
        self.db.execute("INSERT INTO novedad (ejecucion_id, momento, id_portafolio, tipo, detalle, antes, despues) "
                        "VALUES (?,?,?,?,?,?,?)", (eid, txt(momento), k, tipo, detalle, antes, despues))

    # ── guardar una consulta ──
    def guardar(self, eid, momento, procesos, cronogramas, adendas, avisos):
        """Devuelve el número de novedades. En la primera ejecución no hay con qué comparar."""
        m = txt(momento)
        primera = self.ultima_ejecucion_ok(eid) is None
        cr = {r["id_del_proceso"]: r for r in cronogramas}
        grupos = {}
        for f in procesos:
            grupos.setdefault(reglas.expediente(f), []).append(f)
        n0 = self.db.execute("SELECT COUNT(*) FROM novedad").fetchone()[0]
        for k, filas in grupos.items():
            v = reglas.fila_vigente(filas, cr)
            seguir, motivo = reglas.en_seguimiento(filas)
            previo = self.db.execute("SELECT * FROM proceso WHERE id_portafolio=?", (k,)).fetchone()
            if not previo and not seguir:
                continue                         # nunca estuvo en seguimiento: no interesa guardarlo
            ref = min((f.get("referencia_del_proceso") or "" for f in filas), key=len) or None
            fila = dict(
                id_portafolio=k, referencia=ref, codigo=reglas.codigo_corto(ref),
                nombre=v.get("nombre_del_procedimiento"), objeto=v.get("descripci_n_del_procedimiento"),
                modalidad=v.get("modalidad_de_contratacion"), tipo_contrato=v.get("tipo_de_contrato"),
                precio_base=num(v.get("precio_base")), id_req_vigente=v["id_del_proceso"],
                estado=v.get("estado_del_procedimiento"), fase=v.get("fase"),
                adjudicado=v.get("adjudicado"), fecha_adjudicacion=(v.get("fecha_adjudicacion") or "")[:10] or None,
                valor_adjudicado=num(v.get("valor_total_adjudicacion")),
                proveedor_adjudicado=v.get("nombre_del_proveedor"), nit_proveedor=v.get("nit_del_proveedor_adjudicado"),
                fecha_recepcion_ofertas=reglas.local(v.get("fecha_de_recepcion_de")),
                unidad_contratacion=v.get("nombre_de_la_unidad_de"),
                proveedores_manifestaron=entero(v.get("proveedores_que_manifestaron")),
                respuestas=entero(v.get("respuestas_al_procedimiento")), url=url(v.get("urlproceso")),
                en_seguimiento=int(seguir), motivo_fuera=None if seguir else motivo, ultima_vez=m)
            if previo:
                if previo["en_seguimiento"] and not seguir:
                    fila["salio"] = m
                    det = motivo
                    if fila["proveedor_adjudicado"] and fila["proveedor_adjudicado"] != "No Adjudicado":
                        det += f" · {fila['proveedor_adjudicado']}"
                    self._novedad(eid, m, k, "Salió del seguimiento", det, previo["estado"], fila["estado"])
                elif not previo["en_seguimiento"] and seguir:
                    fila["salio"] = None
                    self._novedad(eid, m, k, "Volvió al seguimiento", fila["estado"])
                elif previo["estado"] != fila["estado"]:
                    self._novedad(eid, m, k, "Cambio de estado en SECOP II", "", previo["estado"], fila["estado"])
                cols = ", ".join(f"{c}=?" for c in fila if c != "id_portafolio")
                self.db.execute(f"UPDATE proceso SET {cols} WHERE id_portafolio=?",
                                (*[fila[c] for c in fila if c != "id_portafolio"], k))
            else:
                fila["primera_vez"] = m
                self.db.execute(f"INSERT INTO proceso ({', '.join(fila)}) VALUES ({', '.join('?' * len(fila))})",
                                tuple(fila.values()))
                if not primera:
                    self._novedad(eid, m, k, "Proceso nuevo en seguimiento", f"{fila['modalidad']} · {fila['estado']}")
            for f in filas:
                self.db.execute(
                    "INSERT INTO fase VALUES (?,?,?,?,?,?,?,?) ON CONFLICT(id_req) DO UPDATE SET "
                    "id_portafolio=excluded.id_portafolio, referencia=excluded.referencia, fase=excluded.fase, "
                    "estado=excluded.estado, fecha_ultima_publicacion=excluded.fecha_ultima_publicacion, "
                    "url=excluded.url, ultima_vez=excluded.ultima_vez",
                    (f["id_del_proceso"], k, f.get("referencia_del_proceso"), f.get("fase"),
                     f.get("estado_del_procedimiento"), (f.get("fecha_de_ultima_publicaci") or "")[:10] or None,
                     url(f.get("urlproceso")), m))
                for campo, valor in (cr.get(f["id_del_proceso"]) or {}).items():
                    if campo in NO_HITO:
                        continue
                    if campo not in HITOS and f"hito desconocido: {campo}" not in avisos:
                        avisos.append(f"hito desconocido: {campo}")
                    fh = reglas.bogota(valor)
                    if not fh:
                        continue
                    self.db.execute(
                        "INSERT INTO hito (id_req, campo, fecha_hora, valor_fuente, primera_vez, ultima_vez) "
                        "VALUES (?,?,?,?,?,?) ON CONFLICT(id_req, campo, fecha_hora) DO UPDATE SET "
                        "ultima_vez=excluded.ultima_vez", (f["id_del_proceso"], campo, fh, valor, m, m))
            # el cronograma que rige hoy, contra el que regía en la ejecución anterior
            hoy = {c: reglas.bogota(x) for c, x in (cr.get(v["id_del_proceso"]) or {}).items()
                   if c not in NO_HITO and reglas.bogota(x)}
            antes = {r["campo"]: r for r in self.db.execute(
                "SELECT * FROM hito_vigente WHERE id_portafolio=? AND hasta IS NULL", (k,))}
            for campo, fh in hoy.items():
                a = antes.get(campo)
                if a and a["fecha_hora"] == fh:
                    continue
                avisar = seguir and previo and previo["en_seguimiento"] and not primera
                if a:
                    self.db.execute("UPDATE hito_vigente SET hasta=? WHERE id=?", (m, a["id"]))
                    if avisar:
                        self._novedad(eid, m, k, "Cambio de fecha", HITOS.get(campo, (campo,))[0], a["fecha_hora"], fh)
                elif avisar and antes:
                    self._novedad(eid, m, k, "Hito nuevo en el cronograma", HITOS.get(campo, (campo,))[0], None, fh)
                self.db.execute("INSERT INTO hito_vigente (id_portafolio, campo, fecha_hora, desde) VALUES (?,?,?,?)",
                                (k, campo, fh, m))
            for campo, a in antes.items():
                if campo not in hoy and hoy:
                    self.db.execute("UPDATE hito_vigente SET hasta=? WHERE id=?", (m, a["id"]))
                    if seguir:
                        self._novedad(eid, m, k, "Hito retirado del cronograma", HITOS.get(campo, (campo,))[0],
                                  a["fecha_hora"], None)
        for d in adendas:
            k = d.get("proceso")
            if not self.db.execute("SELECT 1 FROM proceso WHERE id_portafolio=?", (k,)).fetchone():
                continue
            nueva = self.db.execute(
                "INSERT OR IGNORE INTO adenda VALUES (?,?,?,?,?,?)",
                (d["id_documento"], k, d.get("nombre_archivo"), (d.get("fecha_carga") or "")[:16].replace("T", " "),
                 url(d.get("url_descarga_documento")), m)).rowcount
            if nueva and not primera:
                self._novedad(eid, m, k, "Adenda publicada", d.get("nombre_archivo"))
        self.db.commit()
        return self.db.execute("SELECT COUNT(*) FROM novedad").fetchone()[0] - n0

    # ── lecturas para el Excel ──
    def procesos(self, solo_seguimiento=True):
        q = "SELECT * FROM proceso" + (" WHERE en_seguimiento=1" if solo_seguimiento else "")
        return {r["id_portafolio"]: dict(r) for r in self.db.execute(q)}

    def cronograma(self, k):
        """[{campo, fecha_hora, anteriores}] del cronograma vigente de un expediente."""
        p = self.db.execute("SELECT id_req_vigente FROM proceso WHERE id_portafolio=?", (k,)).fetchone()
        if not p:
            return []
        ult = self.db.execute("SELECT MAX(ultima_vez) FROM hito WHERE id_req=?", (p[0],)).fetchone()[0]
        out = []
        for r in self.db.execute("SELECT campo, fecha_hora FROM hito WHERE id_req=? AND ultima_vez=?", (p[0], ult)):
            otras = {x[0] for x in self.db.execute(
                "SELECT h.fecha_hora FROM hito h JOIN fase f ON f.id_req=h.id_req "
                "WHERE f.id_portafolio=? AND h.campo=? AND h.fecha_hora<>?", (k, r["campo"], r["fecha_hora"]))}
            otras |= {x[0] for x in self.db.execute(
                "SELECT fecha_hora FROM hito_vigente WHERE id_portafolio=? AND campo=? AND fecha_hora<>?",
                (k, r["campo"], r["fecha_hora"]))}
            out.append({"campo": r["campo"], "fecha_hora": r["fecha_hora"], "anteriores": sorted(otras)})
        return out

    def adendas(self, k):
        return [dict(r) for r in self.db.execute(
            "SELECT * FROM adenda WHERE id_portafolio=? ORDER BY fecha_carga DESC", (k,))]

    def novedades(self, eid):
        return [dict(r) for r in self.db.execute("SELECT * FROM novedad WHERE ejecucion_id=? ORDER BY id", (eid,))]

    def ejecuciones(self):
        return [dict(r) for r in self.db.execute("SELECT * FROM ejecucion ORDER BY id DESC")]

    # ── gestión del equipo ──
    def gestion_actual(self):
        """{(expediente, campo, fecha_hora): última anotación}; campo/fecha vacíos = del proceso."""
        out = {}
        for r in self.db.execute("SELECT * FROM gestion ORDER BY id"):
            out[(r["id_portafolio"], r["campo"] or "", r["fecha_hora"] or "")] = dict(r)
        return out

    def registrar_gestion(self, eid, momento, k, campo, fecha_hora, **valores):
        actual = self.gestion_actual().get((k, campo or "", fecha_hora or ""), {})
        limpios = {c: (str(v).strip() if v not in (None, "") else None) for c, v in valores.items()}
        if all(limpios[c] == (actual.get(c) or None) for c in limpios):
            return False
        self.db.execute("INSERT INTO gestion (momento, ejecucion_id, id_portafolio, campo, fecha_hora, nombre_corto, "
                        "responsable, estado, observacion) VALUES (?,?,?,?,?,?,?,?,?)",
                        (txt(momento), eid, k, campo or None, fecha_hora or None, limpios.get("nombre_corto"),
                         limpios.get("responsable"), limpios.get("estado"), limpios.get("observacion")))
        return True

    def historial_gestion(self):
        return [dict(r) for r in self.db.execute(
            "SELECT g.*, p.codigo, p.referencia FROM gestion g LEFT JOIN proceso p USING (id_portafolio) "
            "ORDER BY g.id DESC")]
