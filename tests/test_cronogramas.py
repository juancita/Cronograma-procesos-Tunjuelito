# -*- coding: utf-8 -*-
"""Pruebas sin internet: reglas y ciclo completo con datos de ejemplo.   python -m unittest -v"""
import datetime as dt
import pathlib
import sqlite3
import tempfile
import unittest

import openpyxl

from cronogramas import almacen, config, excel, reglas


def fila(req, bdos, estado="Publicado", ultpub="2026-11-01", adjudicado="No", ref="FDLT-LP-379-2026 (163980)"):
    return {"id_del_proceso": req, "id_del_portafolio": bdos, "estado_del_procedimiento": estado,
            "fecha_de_ultima_publicaci": ultpub, "adjudicado": adjudicado, "referencia_del_proceso": ref,
            "modalidad_de_contratacion": "Licitación pública", "nombre_del_procedimiento": "MALLA VIAL",
            "descripci_n_del_procedimiento": "Obra de malla vial", "precio_base": "1000000",
            "fecha_de_recepcion_de": "2026-12-01T10:00:00.000",
            "urlproceso": {"url": "https://community.secop.gov.co/x"}}


class Reglas(unittest.TestCase):
    def test_hora_utc_a_bogota(self):
        self.assertEqual(reglas.bogota("2026-11-11 15:00:00.0000000"), "2026-11-11 10:00")
        self.assertEqual(reglas.bogota("2026-11-25 04:59:00.0000000"), "2026-11-24 23:59")
        self.assertIsNone(reglas.bogota(None))

    def test_fila_vigente_es_la_mas_reciente_con_cronograma(self):
        a, b = fila("CO1.REQ.1", "B", ultpub="2026-09-21"), fila("CO1.REQ.2", "B", ultpub="2026-10-27")
        self.assertEqual(reglas.fila_vigente([a, b], {"CO1.REQ.1", "CO1.REQ.2"})["id_del_proceso"], "CO1.REQ.2")
        self.assertEqual(reglas.fila_vigente([a, b], {"CO1.REQ.1"})["id_del_proceso"], "CO1.REQ.1")

    def test_seguimiento_solo_abiertos_sin_adjudicar(self):
        self.assertTrue(reglas.en_seguimiento([fila("1", "B", "Publicado")])[0])
        self.assertTrue(reglas.en_seguimiento([fila("1", "B", "Evaluación")])[0])
        self.assertFalse(reglas.en_seguimiento([fila("1", "B", "Evaluación"), fila("2", "B", "Seleccionado")])[0])
        self.assertFalse(reglas.en_seguimiento([fila("1", "B", "Abierto", adjudicado="Si")])[0])
        self.assertFalse(reglas.en_seguimiento([fila("1", "B", "Borrador")])[0])
        self.assertFalse(reglas.en_seguimiento([fila("1", "B", "Cancelado")])[0])

    def test_codigo_corto(self):
        self.assertEqual(reglas.codigo_corto("FDLT-SAMC-426-2026 (166972)"), "SAMC-426")
        self.assertEqual(reglas.codigo_corto("FDLT-LP 429-2026 (167708)"), "LP-429")
        self.assertEqual(reglas.codigo_corto("FDLT-CMA-434-2026(163984)"), "CMA-434")
        self.assertEqual(reglas.codigo_corto("Solicitud de información No 031-2026"), "SI-031")

    def test_etapa_del_ultimo_hito_alcanzado(self):
        h = [("2026-11-20 10:00", "plazo_para_presentar_1"), ("2026-11-25 23:59", "expedici_n_y_publicaci_n"),
             ("2026-12-04 10:00", "presentaci_n_de_ofertas"), ("2027-06-01 00:00", "plazo_de_ejecuci_n_del")]
        self.assertEqual(reglas.etapa_en(h, dt.datetime(2026, 11, 24, 12)), "Pre-pliego")
        self.assertEqual(reglas.etapa_en(h, dt.datetime(2026, 11, 26, 9)), "Pliego definitivo")
        self.assertEqual(reglas.etapa_en(h, dt.datetime(2026, 12, 5, 9)), "Cronograma cumplido")


class Ciclo(unittest.TestCase):
    def test_dos_actualizaciones_novedades_y_gestion(self):
        with tempfile.TemporaryDirectory() as d:
            cfg = config.cargar(pathlib.Path(d))
            A = almacen.Almacen(almacen.abrir(cfg.base))
            procesos = [fila("CO1.REQ.1", "CO1.BDOS.9", ultpub="2026-09-21"),
                        fila("CO1.REQ.2", "CO1.BDOS.9", ultpub="2026-10-27")]
            crono = [{"id_del_proceso": "CO1.REQ.1", "presentaci_n_de_ofertas": "2026-10-26 15:00:00.0"},
                     {"id_del_proceso": "CO1.REQ.2", "presentaci_n_de_ofertas": "2026-12-01 15:00:00.0",
                      "publicaci_n_acto": "2026-12-10 01:00:00.0"}]
            t1 = dt.datetime(2026, 11, 24, 8)
            e1 = A.iniciar(t1, "x")
            self.assertEqual(A.guardar(e1, t1, procesos, crono, [], []), 0)   # línea base
            A.terminar(e1, t1, "completa")
            cr = {h["campo"]: h for h in A.cronograma("CO1.BDOS.9")}
            self.assertEqual(cr["presentaci_n_de_ofertas"]["fecha_hora"], "2026-12-01 10:00")
            self.assertIn("2026-10-26 10:00", cr["presentaci_n_de_ofertas"]["anteriores"])
            excel.construir(cfg.excel, A, e1, t1, cfg, [])
            wb = openpyxl.load_workbook(cfg.excel)
            ws = wb["PROCESOS_EN_CURSO"]
            cab = [c.value for c in ws[4]]
            ws.cell(row=5, column=cab.index("Responsable") + 1, value="Ana")
            wb.save(cfg.excel)
            t2 = dt.datetime(2026, 11, 25, 8)
            e2 = A.iniciar(t2, "x")
            for k, campo, fh, vals in excel.leer_gestion(cfg.excel):
                A.registrar_gestion(e2, t2, k, campo, fh, **vals)
            crono[1]["presentaci_n_de_ofertas"] = "2026-12-03 15:00:00.0"          # adenda
            procesos2 = procesos + [fila("CO1.REQ.3", "CO1.BDOS.8", ref="FDLT-SAMC-441-2026 (1)")]
            crono2 = crono + [{"id_del_proceso": "CO1.REQ.3", "presentaci_n_de_ofertas": "2026-12-11 15:00:00.0"}]
            n = A.guardar(e2, t2, procesos2, crono2, [], [])
            tipos = sorted(x["tipo"] for x in A.novedades(e2))
            self.assertEqual(tipos, ["Cambio de fecha", "Proceso nuevo en seguimiento"])
            self.assertEqual(n, 2)
            A.terminar(e2, t2, "completa")
            excel.construir(cfg.excel, A, e2, t2, cfg, [])
            wb = openpyxl.load_workbook(cfg.excel)
            ws = wb["PROCESOS_EN_CURSO"]
            cab = [c.value for c in ws[4]]
            responsables = [r[cab.index("Responsable")] for r in ws.iter_rows(min_row=5, values_only=True)]
            self.assertIn("Ana", responsables)
            t3 = dt.datetime(2026, 11, 26, 8)
            e3 = A.iniciar(t3, "x")
            procesos3 = [dict(p, estado_del_procedimiento="Seleccionado") if p["id_del_portafolio"] == "CO1.BDOS.9"
                         else p for p in procesos2]
            A.guardar(e3, t3, procesos3, crono2, [], [])
            self.assertEqual([x["tipo"] for x in A.novedades(e3)], ["Salió del seguimiento"])
            self.assertEqual(sqlite3.connect(cfg.base).execute(
                "SELECT en_seguimiento FROM proceso WHERE id_portafolio='CO1.BDOS.9'").fetchone()[0], 0)


if __name__ == "__main__":
    unittest.main()
