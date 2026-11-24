# -*- coding: utf-8 -*-
"""Actualiza el seguimiento: consulta SECOP II, guarda en la base local y rehace el Excel.

    cronogramas.exe                  (o: python -m cronogramas)
    cronogramas.exe --sin-abrir      no abre el Excel al terminar
    cronogramas.exe --sin-consultar  rehace el Excel con lo guardado (sin internet)
    cronogramas.exe --carpeta RUTA   usa otra carpeta de trabajo
"""
import argparse
import gzip
import json
import os
import shutil
import subprocess
import sys
import time

from . import __version__, almacen, config, excel, reglas, secop


def decir(t=""):
    print(t, flush=True)


def excel_cerrado(ruta):
    if not ruta.exists():
        return True
    try:
        with open(ruta, "r+b"):
            return True
    except PermissionError:
        return False


def guardar_consulta(cfg, momento, **partes):
    """Copia comprimida de lo que respondió SECOP II (para poder revisar qué se consultó)."""
    d = cfg.datos / "consultas" / momento.strftime("%Y-%m-%d_%H%M%S")
    d.mkdir(parents=True, exist_ok=True)
    for nombre, filas in partes.items():
        with gzip.open(d / f"{nombre}.json.gz", "wt", encoding="utf-8") as f:
            json.dump(filas, f, ensure_ascii=False)


def abrir_archivo(ruta):
    try:
        if sys.platform.startswith("win"):
            os.startfile(ruta)  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.run(["open", str(ruta)], check=False)
    except OSError:
        pass


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(prog="cronogramas", description="Seguimiento de cronogramas de SECOP II")
    ap.add_argument("--carpeta", help="carpeta de trabajo (por defecto, la del programa)")
    ap.add_argument("--sin-abrir", action="store_true", help="no abrir el Excel al terminar")
    ap.add_argument("--sin-consultar", action="store_true", help="rehacer el Excel con lo guardado, sin internet")
    ap.add_argument("--version", action="version", version=__version__)
    args = ap.parse_args(argv)
    cfg = config.cargar(args.carpeta)
    if args.sin_abrir:
        cfg.abrir_excel = False

    decir(f"\n  Cronogramas de procesos · {cfg.nombre_mostrar}  (v{__version__})")
    decir(f"  Carpeta: {cfg.carpeta}\n")
    if not excel_cerrado(cfg.excel):
        decir(f"  ✗ «{cfg.excel.name}» está abierto. Ciérrelo y vuelva a ejecutar.")
        return 2

    db = almacen.abrir(cfg.base)
    A = almacen.Almacen(db)
    momento = reglas.ahora()
    t0 = time.time()
    eid = A.iniciar(momento, __version__)
    avisos = []
    try:
        anotados = excel.leer_gestion(cfg.excel)
        nuevos = sum(A.registrar_gestion(eid, momento, k, campo, fh, **vals) for k, campo, fh, vals in anotados)
        db.commit()
        if anotados:
            decir(f"  1. Anotaciones del equipo leídas del Excel anterior: {nuevos} cambios guardados")
        cuentas = dict(filas_procesos=None, filas_cronograma=None, adendas=None, novedades=0)
        if args.sin_consultar:
            decir("  2. Sin consultar SECOP II: se usa lo guardado")
        else:
            cli = secop.Cliente(cfg, avisar=decir)
            decir("  2. Consultando SECOP II (datos.gov.co)…")
            abiertos = cli.procesos_abiertos()
            exps = {reglas.expediente(f) for f in abiertos} | set(A.expedientes_en_seguimiento())
            decir(f"     procesos abiertos de la entidad: {len(abiertos)} filas · {len(exps)} expedientes a revisar")
            procesos = cli.procesos_de([k for k in exps if k.startswith("CO1.BDOS")])
            vistos = {f["id_del_proceso"] for f in procesos}
            procesos += [f for f in abiertos if f["id_del_proceso"] not in vistos]
            crono = cli.cronogramas([f["id_del_proceso"] for f in procesos])
            decir(f"     cronogramas: {len(crono)}")
            # adendas solo de lo que se sigue y tuvo movimiento en el último año (la consulta es lenta)
            grupos = {}
            for f in procesos:
                grupos.setdefault(reglas.expediente(f), []).append(f)
            limite = (momento.date().replace(year=momento.year - 1)).isoformat()
            con_movimiento = [k for k, fs in grupos.items() if k.startswith("CO1.BDOS") and reglas.en_seguimiento(fs)[0]
                              and max((f.get("fecha_de_ultima_publicaci") or "") for f in fs)[:10] >= limite]
            ads = cli.adendas(con_movimiento)
            decir(f"     documentos de adenda: {len(ads)}")
            guardar_consulta(cfg, momento, procesos=procesos, cronograma=crono, adendas=ads)
            nov = A.guardar(eid, momento, procesos, crono, ads, avisos)
            cuentas = dict(filas_procesos=len(procesos), filas_cronograma=len(crono), adendas=len(ads),
                           novedades=nov, consultas=cli.consultas)
            decir(f"  3. Guardado en la base local · novedades: {nov}")
        en_seg = len(A.expedientes_en_seguimiento())
        A.terminar(eid, reglas.ahora(), "completa", en_seguimiento=en_seg, avisos=avisos, **cuentas)
        res = excel.construir(cfg.excel, A, eid, momento, cfg, avisos)
        cfg.historial.mkdir(parents=True, exist_ok=True)
        copia = cfg.historial / f"Cronogramas - {momento:%Y-%m-%d %H%M}.xlsx"
        shutil.copy2(cfg.excel, copia)
        decir(f"  4. Excel listo: {cfg.excel.name}  (copia en historial/)\n")
        decir(f"     Procesos en curso: {res['en_curso']}   Hitos hoy: {res['hoy']} "
              f"({res['hoy_entidad']} de la Entidad)   Próximos 7 días: {res['semana']}")
        decir(f"     Novedades: {res['novedades']}   Puntos para revisar: {res['revisar']}")
        for a in avisos:
            decir(f"     ! {a}")
        decir(f"\n  Terminado en {time.time() - t0:.0f} s.")
    except secop.ErrorConsulta as e:
        A.terminar(eid, reglas.ahora(), "falló", avisos=[str(e)])
        decir(f"\n  ✗ {e}\n    No se cambió el Excel. Revise la conexión a internet y vuelva a intentar.")
        return 1
    except Exception as e:  # noqa: BLE001
        A.terminar(eid, reglas.ahora(), "falló", avisos=[f"{type(e).__name__}: {e}"])
        raise
    finally:
        db.close()
    if cfg.abrir_excel:
        abrir_archivo(cfg.excel)
    return 0
