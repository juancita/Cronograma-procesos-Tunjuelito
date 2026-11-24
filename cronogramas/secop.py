# -*- coding: utf-8 -*-
"""Consultas a los datos abiertos de SECOP II (datos.gov.co). Solo LEE información pública.

Conjuntos:
  p6dx-8zbt   SECOP II - Procesos de Contratación   (una fila por proceso y fase; hora local)
  f3ev-294c   SECOPII - Fechas Cronograma            (una fila por proceso; hora UTC)
  Archivos    SECOP II - Archivos Descarga Historico <año>  (índice de documentos; para las adendas)
"""
import json
import re
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request

from . import __version__
from .config import ESTADOS_ABIERTOS

DOMINIO = "https://www.datos.gov.co"
PROCESOS = "p6dx-8zbt"
CRONOGRAMA = "f3ev-294c"
LOTE = 80
AGENTE = f"cronogramas-secop/{__version__}"


class ErrorConsulta(Exception):
    pass


def contexto_ssl():
    """Primero los certificados de Windows (funciona en redes con inspección HTTPS),
    luego los de certifi, luego los del sistema."""
    try:
        import truststore
        return truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    except ImportError:
        pass
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        return ssl.create_default_context()


class Cliente:
    def __init__(self, cfg, avisar=print):
        self.cfg = cfg
        self.avisar = avisar
        self.ctx = contexto_ssl()
        self.consultas = 0

    def _pedir(self, url):
        cab = {"Accept": "application/json", "User-Agent": AGENTE}
        if self.cfg.token_app:
            cab["X-App-Token"] = self.cfg.token_app
        ultimo = None
        for intento in range(5):
            try:
                with urllib.request.urlopen(urllib.request.Request(url, headers=cab), timeout=120,
                                            context=self.ctx) as r:
                    self.consultas += 1
                    return json.load(r)
            except urllib.error.HTTPError as e:
                ultimo = f"el servidor respondió {e.code}"
                if e.code in (400, 403, 404):
                    break
            except urllib.error.URLError as e:
                ultimo = f"sin conexión ({e.reason})"
                if isinstance(e.reason, ssl.SSLError):
                    ultimo = ("la conexión segura falló (certificado). Si la red de la oficina inspecciona "
                              "HTTPS, avisar a sistemas o probar desde otra red")
            except (TimeoutError, OSError, ValueError) as e:
                ultimo = str(e)
            espera = 5 * (intento + 1)
            self.avisar(f"   … reintento {intento + 1} en {espera} s: {ultimo}")
            time.sleep(espera)
        raise ErrorConsulta(f"No se pudo consultar datos.gov.co: {ultimo}")

    def consulta(self, recurso, where, select=None):
        filas, desde = [], 0
        while True:
            p = {"$where": where, "$limit": 20000, "$offset": desde, "$order": ":id"}
            if select:
                p["$select"] = select
            d = self._pedir(f"{DOMINIO}/resource/{recurso}.json?" + urllib.parse.urlencode(p))
            filas += d
            if len(d) < 20000:
                return filas
            desde += 20000

    def por_lotes(self, recurso, campo, valores, extra=""):
        """Consulta por lotes de ID, hasta 4 lotes a la vez (el servidor es lento por consulta)."""
        from concurrent.futures import ThreadPoolExecutor
        valores = sorted({v for v in valores if v})
        wheres = [f"{campo} in (" + ",".join("'" + v.replace("'", "''") + "'" for v in valores[i:i + LOTE]) + f"){extra}"
                  for i in range(0, len(valores), LOTE)]
        if not wheres:
            return []
        with ThreadPoolExecutor(max_workers=min(4, len(wheres))) as ex:
            partes = list(ex.map(lambda w: self.consulta(recurso, w), wheres))
        return [f for parte in partes for f in parte]

    # ── lo que necesita el programa ──
    def _entidad(self):
        nombres = ",".join("'" + n.replace("'", "''") + "'" for n in self.cfg.nombres_entidad)
        return f"(codigo_entidad='{self.cfg.codigo_entidad}' OR entidad in ({nombres}))"

    def procesos_abiertos(self):
        estados = ",".join(f"'{e}'" for e in ESTADOS_ABIERTOS)
        return self.consulta(PROCESOS, f"{self._entidad()} AND estado_del_procedimiento in ({estados})")

    def procesos_de(self, expedientes):
        """Todas las filas (fases) de esos expedientes, de la entidad."""
        return self.por_lotes(PROCESOS, "id_del_portafolio", expedientes, extra=f" AND {self._entidad()}")

    def cronogramas(self, ids_proceso):
        return self.por_lotes(CRONOGRAMA, "id_del_proceso", ids_proceso)

    def datasets_archivos(self):
        """Busca en el catálogo los índices «Archivos Descarga Historico <año>» del año actual y el
        anterior. Si el catálogo no responde, usa los de config.ini."""
        try:
            d = self._pedir(f"{DOMINIO}/api/catalog/v1?" + urllib.parse.urlencode(
                {"q": "SECOP II Archivos Descarga Historico", "domains": "www.datos.gov.co", "limit": 40}))
            hallados = {}
            for r in d.get("results", []):
                m = re.fullmatch(r"SECOP II - Archivos Descarga Historico (\d{4})", r["resource"]["name"].strip())
                if m:
                    hallados[int(m.group(1))] = r["resource"]["id"]
            if hallados:
                ultimos = sorted(hallados)[-2:]
                return [hallados[a] for a in reversed(ultimos)]
        except ErrorConsulta:
            pass
        return list(self.cfg.datasets_archivos)

    def adendas(self, expedientes):
        filas = []
        for rec in self.datasets_archivos():
            try:
                filas += self.por_lotes(rec, "proceso", expedientes, extra=(
                    " AND (upper(nombre_archivo) like '%ADENDA%' OR upper(descripci_n) like '%ADENDA%')"))
            except ErrorConsulta as e:
                self.avisar(f"   ! no se pudieron leer las adendas de {rec}: {e}")
        return filas
