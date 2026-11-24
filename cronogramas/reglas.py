# -*- coding: utf-8 -*-
"""Reglas verificadas contra adendas reales (ver README, «Cómo se validó»).

1. Un proceso (expediente CO1.BDOS) tiene VARIAS filas CO1.REQ, una por fase. Las filas no traen el
   mismo cronograma: rige la de publicación más reciente que tenga cronograma; las otras guardan las
   fechas de antes (sirven para saber la fecha original de un hito que cambió).
2. El cronograma (f3ev-294c) viene en hora UTC; la ficha del proceso (p6dx-8zbt), en hora local.
   Todo se muestra en hora de Bogotá (UTC-5, sin horario de verano).
3. Seguimiento = abierto y sin adjudicar: estado Publicado, Abierto o Evaluación; adjudicado ≠ «Si»;
   ninguna fase del expediente en Seleccionado, Adjudicado, Cancelado, Desierto, Terminado o Cerrado.
"""
import datetime as dt

from .catalogo import CUMPLIDO, EJECUCION, ETAPAS, HITOS
from .config import ESTADOS_ABIERTOS, ESTADOS_CERRADOS

BOGOTA = dt.timezone(dt.timedelta(hours=-5))


def ahora():
    return dt.datetime.now(BOGOTA).replace(tzinfo=None)


def bogota(v):
    """'2026-11-11 15:00:00.0000000' (UTC) → '2026-11-11 10:00'."""
    if not v:
        return None
    try:
        t = dt.datetime.fromisoformat(str(v)[:19].replace("T", " "))
    except ValueError:
        return None
    return (t - dt.timedelta(hours=5)).strftime("%Y-%m-%d %H:%M")


def local(v):
    """La ficha del proceso ya viene en hora local: '2026-11-13T14:00:00.000' → '2026-11-13 14:00'."""
    return str(v)[:16].replace("T", " ") if v else None


def numero_req(fila):
    cola = fila["id_del_proceso"].rsplit(".", 1)[-1]
    return int(cola) if cola.isdigit() else 0


def expediente(fila):
    return fila.get("id_del_portafolio") or fila["id_del_proceso"]


def fila_vigente(filas, con_cronograma):
    return max(filas, key=lambda f: (f["id_del_proceso"] in con_cronograma,
                                     f.get("fecha_de_ultima_publicaci") or "", numero_req(f)))


def en_seguimiento(filas):
    """(bool, motivo) para un expediente con todas sus filas."""
    estados = {f.get("estado_del_procedimiento") for f in filas}
    cerrados = estados & ESTADOS_CERRADOS
    if cerrados:
        return False, " / ".join(sorted(cerrados))
    if any((f.get("adjudicado") or "").strip().lower() in ("si", "sí") for f in filas):
        return False, "Adjudicado"
    if not estados & set(ESTADOS_ABIERTOS):
        return False, " / ".join(sorted(e for e in estados if e)) or "sin estado"
    return True, ""


def codigo_corto(referencia):
    """'FDLT-SAMC-426-2026 (166972)' → 'SAMC-426'; 'FDLT-LP 429-2026 (167708)' → 'LP-429'."""
    import re
    r = (referencia or "").upper().replace("FLDT", "FDLT")
    m = re.search(r"FDLT[\s\-]*([A-Z]{2,5})[\s\-\.]*0*(\d{1,4})", r)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):03d}" if len(m.group(2)) < 3 else f"{m.group(1)}-{m.group(2)}"
    m = re.search(r"SOLICITUD DE INFORMACI[OÓ]N\s*N[OÚ°º\.]*\s*0*(\d+)", r)
    if m:
        return f"SI-{int(m.group(1)):03d}"
    return (referencia or "")[:20]


def etapa_en(hitos, momento):
    """La etapa del último hito alcanzado a ese momento. hitos: [(fecha_hora, campo)].
    Si ya pasaron todos los hitos del proceso (sin contar el fin del plazo de ejecución): cumplido."""
    propios = [(f, c) for f, c in hitos if HITOS.get(c, ("", "Otro"))[1] != EJECUCION]
    if not propios:
        return ETAPAS[0]
    m = momento.strftime("%Y-%m-%d %H:%M")
    if all(f <= m for f, _ in propios):
        return CUMPLIDO
    pasados = [(f, HITOS.get(c, ("", ETAPAS[0], "", 0))[3], c) for f, c in propios if f <= m]
    if not pasados:
        return ETAPAS[0]
    _, _, c = max(pasados)
    etapa = HITOS.get(c, ("", ETAPAS[0]))[1]
    return etapa if etapa in ETAPAS else ETAPAS[0]
