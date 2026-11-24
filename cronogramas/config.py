# -*- coding: utf-8 -*-
"""Configuración: valores por defecto, que se pueden cambiar en config.ini (junto al programa)."""
import configparser
import pathlib
import sys
from dataclasses import dataclass, field

ESTADOS_ABIERTOS = ("Publicado", "Abierto", "Evaluación")      # publicado (prepliego) → antes de adjudicar
ESTADOS_CERRADOS = {"Seleccionado", "Adjudicado", "Cancelado", "Desierto", "Terminado", "Cerrado"}


@dataclass
class Config:
    carpeta: pathlib.Path
    codigo_entidad: str = "702235227"
    nombres_entidad: list = field(default_factory=lambda: [
        "ALCALDIA LOCAL DE TUNJUELITO", "BOGOTÁ D.C. - ALCALDÍA LOCAL DE TUNJUELITO"])
    nombre_mostrar: str = "Fondo de Desarrollo Local de Tunjuelito"
    dias_abiertos_sin_fechas: int = 90
    datasets_archivos: list = field(default_factory=lambda: ["dmgg-8hin", "h4m3-b4e2"])
    token_app: str = ""                       # opcional: token de aplicación de datos.gov.co
    abrir_excel: bool = True

    @property
    def datos(self):
        return self.carpeta / "datos"

    @property
    def base(self):
        return self.datos / "cronogramas.db"

    @property
    def excel(self):
        return self.carpeta / "Seguimiento de cronogramas.xlsx"

    @property
    def historial(self):
        return self.carpeta / "historial"


def carpeta_programa():
    """Junto al .exe cuando está empaquetado; si no, la carpeta desde donde se ejecuta."""
    if getattr(sys, "frozen", False):
        return pathlib.Path(sys.executable).resolve().parent
    return pathlib.Path.cwd()


def cargar(carpeta=None):
    cfg = Config(carpeta=pathlib.Path(carpeta) if carpeta else carpeta_programa())
    ini = cfg.carpeta / "config.ini"
    if ini.exists():
        p = configparser.ConfigParser()
        p.read(ini, encoding="utf-8")
        s = p["cronogramas"] if p.has_section("cronogramas") else {}
        cfg.codigo_entidad = s.get("codigo_entidad", cfg.codigo_entidad).strip()
        if s.get("nombres_entidad"):
            cfg.nombres_entidad = [x.strip() for x in s["nombres_entidad"].split("|") if x.strip()]
        cfg.nombre_mostrar = s.get("nombre_mostrar", cfg.nombre_mostrar).strip()
        cfg.dias_abiertos_sin_fechas = int(s.get("dias_abiertos_sin_fechas", cfg.dias_abiertos_sin_fechas))
        if s.get("datasets_archivos"):
            cfg.datasets_archivos = [x.strip() for x in s["datasets_archivos"].split(",") if x.strip()]
        cfg.token_app = s.get("token_app", "").strip()
        cfg.abrir_excel = s.get("abrir_excel", "si").strip().lower() in ("si", "sí", "true", "1")
    return cfg
