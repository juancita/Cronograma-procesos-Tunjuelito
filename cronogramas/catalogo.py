# -*- coding: utf-8 -*-
"""Hitos del cronograma de SECOP II (conjunto f3ev-294c «SECOPII - Fechas Cronograma»).

campo → (nombre oficial, etapa, quién actúa, orden en el flujo, código del tablero).
El nombre es la descripción que publica SECOP II para cada columna. Si SECOP agrega una columna
nueva, el programa la muestra con su nombre técnico y lo avisa en REGISTRO_EJECUCIONES.

Etapas (las del tablero del equipo): Pre-pliego · Pliego definitivo · Ofertas y evaluación ·
Adjudicación · Contrato; «Ejecución» es el fin del plazo del contrato y no cuenta como hito del proceso.
"""
ETAPAS = ("Pre-pliego", "Pliego definitivo", "Ofertas y evaluación", "Adjudicación", "Contrato")
CUMPLIDO = "Cronograma cumplido"
EJECUCION = "Ejecución"
NO_HITO = {"id_del_proceso", "proceso_de_compra", "modalidad_de_contratacion"}
# Fechas de registro del proceso: no son plazos, no se cuentan como «cambio de fecha».
SIN_FECHA_ANTERIOR = {"fecha_de_creaci_n", "fecha_de_publicaci_n", "fecha_de_publicaci_n_del_2"}
CODIGOS = {"AP": "acto de apertura o invitación", "R": "audiencia de riesgos", "C": "cierre de ofertas",
           "IE": "informe de evaluación", "AA": "audiencia de adjudicación", "A": "adjudicación",
           "F": "firma del contrato", "G": "garantías", "•": "otro hito"}

HITOS = {
    "fecha_de_creaci_n": ("Fecha de creación de Proceso de Compra", "Pre-pliego", "Entidad", 10, ""),
    "fecha_de_publicaci_n": ("Fecha de publicacion del proceso", "Pre-pliego", "Entidad", 20, ""),
    "fecha_de_publicaci_n_del_2": ("Fecha de publicación del proceso", "Pre-pliego", "Entidad", 21, ""),
    "publicaci_n_de_estudios": ("Publicación de estudios previos", "Pre-pliego", "Entidad", 30, ""),
    "publicaci_n_estudios_previos": ("Publicación Estudios Previos", "Pre-pliego", "Entidad", 31, ""),
    "publicaci_n_estudios_y": ("Publicación Estudios y Documentos Previos", "Pre-pliego", "Entidad", 32, ""),
    "publicaci_n_de_estudios_1": ("Publicación de estudios previos y proyecto de pliego de condiciones", "Pre-pliego", "Entidad", 33, ""),
    "fecha_de_publicaci_n_del_4": ("Fecha de publicación del proyecto de pliego de condiciones", "Pre-pliego", "Entidad", 34, ""),
    "fecha_de_publicaci_n_del_3": ("Fecha de publicación del proyecto de aviso de convocatoria para precalificación", "Pre-pliego", "Entidad", 35, ""),
    "publicaci_n_aviso_convocatoria": ("Publicación Aviso Convocatoria para pre-calificación", "Pre-pliego", "Entidad", 36, ""),
    "publicaci_n_del_aviso_de": ("Publicación del aviso de convocatoria pública", "Pre-pliego", "Entidad", 37, ""),
    "publicaci_n_del_aviso_art": ("Publicación del Aviso (artículo 30 de la Ley 80 de 1993)", "Pre-pliego", "Entidad", 38, ""),
    "plazo_para_presentar_1": ("Plazo para presentar observaciones al proyecto de Pliego de Condiciones", "Pre-pliego", "Interesados", 40, ""),
    "presentaci_n_de_observaciones_3": ("Presentación de Observaciones al Pliego de Condiciones", "Pre-pliego", "Interesados", 41, ""),
    "respuesta_a_las_observaciones_1": ("Respuesta a las observaciones a los Pliegos de Condiciones", "Pre-pliego", "Entidad", 44, ""),
    "plazo_para_manifestaci_n_1": ("Plazo para manifestación de interés de limitar la convocatoria a MiPymes", "Pre-pliego", "Interesados", 45, ""),
    "publicaci_n_del_aviso_de_1": ("Publicación del aviso de limitación a MiPymes o si podrá participar cualquier interesado", "Pre-pliego", "Entidad", 46, ""),
    "publicaci_n_de_la_lista_de": ("Publicación de la lista de precalificados", "Pre-pliego", "Entidad", 47, ""),
    "publicaci_n_del_informe_de_1": ("Publicación del informe de precalificación", "Pre-pliego", "Entidad", 48, ""),
    "expedici_n_y_publicaci_n": ("Fecha de Expedición y publicación acto administrativo de apertura del proceso de selección", "Pliego definitivo", "Entidad", 50, "AP"),
    "publicaci_n_de_la_invitaci": ("Publicación de la invitación", "Pliego definitivo", "Entidad", 51, "AP"),
    "fecha_prevista_de_publicaci": ("Fecha prevista de publicación de los pliegos de condiciones definitivos", "Pliego definitivo", "Entidad", 52, ""),
    "fecha_prevista_de_publicaci_3": ("Fecha Prevista de Publicación del Pliego de Condiciones Definitivo y Demonstración de Interés", "Pliego definitivo", "Entidad", 53, ""),
    "fecha_de_publicaci_n_del_1": ("Fecha de publicación del pliego de condiciones definitivo y demostración de interés", "Pliego definitivo", "Entidad", 54, ""),
    "fecha_de_publicaci_n_del": ("Fecha de publicación del pliego de condiciones definitivo", "Pliego definitivo", "Entidad", 55, ""),
    "plazo_para_manifestaci_n": ("Plazo para manifestación de Interés", "Pliego definitivo", "Interesados", 56, ""),
    "t_rmino_para_valorar_las": ("Término para valorar las manifestaciones de interés", "Pliego definitivo", "Entidad", 57, ""),
    "realizaci_n_del_sorteo": ("Realización del sorteo", "Pliego definitivo", "Entidad", 58, ""),
    "fecha_audiencia_de_asignaci": ("Fecha de Audiencia de asignación de Riesgos", "Pliego definitivo", "Entidad e interesados", 59, "R"),
    "plazo_para_solicitar": ("Plazo para solicitar aclaraciones", "Pliego definitivo", "Interesados", 60, ""),
    "presentaci_n_de_observaciones": ("Presentación de Observaciones a los Pliegos de Condiciones definitivos", "Pliego definitivo", "Interesados", 61, ""),
    "respuesta_a_las_observaciones_3": ("Respuesta a las observaciones al pliego de condiciones definitivo", "Pliego definitivo", "Entidad", 62, ""),
    "plazo_m_ximo_para_expedir": ("Plazo máximo para expedir adendas", "Pliego definitivo", "Entidad", 63, ""),
    "fecha_prevista_para_registro": ("Fecha prevista para registro de ofertas", "Pliego definitivo", "Interesados", 64, ""),
    "fecha_l_mite_para_responder": ("Fecha límite para responder la solicitud de información", "Ofertas y evaluación", "Interesados", 65, "C"),
    "presentaci_n_de_ofertas": ("Presentación de Ofertas", "Ofertas y evaluación", "Interesados", 66, "C"),
    "plazo_de_validez_de_las": ("Plazo de validez de las ofertas", "Ofertas y evaluación", "Interesados", 67, ""),
    "fecha_l_mite_para_la_apertura": ("Fecha límite para la apertura de sobres", "Ofertas y evaluación", "Entidad", 68, ""),
    "fecha_de_apertura_de_ofertas": ("Fecha de Apertura de Ofertas", "Ofertas y evaluación", "Entidad", 69, ""),
    "informe_de_presentaci_n_de": ("Informe de presentación de ofertas", "Ofertas y evaluación", "Entidad", 70, ""),
    "fecha_de_apertura_habilitantes": ("Fecha de Apertura de sobre de requisitos habilitantes y técnicos", "Ofertas y evaluación", "Entidad", 71, ""),
    "publicaci_n_del_informe_de": ("Publicación del informe de evaluación de las Ofertas", "Ofertas y evaluación", "Entidad", 72, "IE"),
    "publicaci_n_del_informe_de_2": ("Publicación del informe de verificación o evaluación", "Ofertas y evaluación", "Entidad", 73, "IE"),
    "presentaci_n_de_observaciones_1": ("Presentación de observaciones al informe de verificación o evaluación", "Ofertas y evaluación", "Interesados", 74, ""),
    "presentaci_n_de_observaciones_2": ("Presentación de observaciones al informe evaluación de las ofertas", "Ofertas y evaluación", "Interesados", 75, ""),
    "fecha_apertura_de_sobres": ("Fecha de Apertura del sobre económico", "Ofertas y evaluación", "Entidad", 76, ""),
    "fecha_apertura_del_sobre": ("Fecha Apertura del sobre económico", "Ofertas y evaluación", "Entidad", 77, ""),
    "fecha_audiencia_de_adjudicaci": ("Fecha de Audiencia de Adjudicación", "Adjudicación", "Entidad", 78, "AA"),
    "publicaci_n_acto": ("Publicación Acto Administrativo de adjudicación o de Declaratoria de Desierto", "Adjudicación", "Entidad", 79, "A"),
    "fecha_de_aceptaci_n_de_ofertas": ("Fecha de Aceptación de ofertas", "Adjudicación", "Entidad", 80, "A"),
    "firma_del_contrato": ("Firma del Contrato", "Contrato", "Entidad y adjudicatario", 82, "F"),
    "fecha_de_firma_del_contrato": ("Fecha de Firma del Contrato", "Contrato", "Entidad y adjudicatario", 83, "F"),
    "fecha_entrega_garant_a": ("Fecha de Entrega de las garantías de ejecución del contrato", "Contrato", "Contratista", 84, "G"),
    "fecha_entrega_garant_a_o": ("Fecha de Entrega de las garantías o pólizas de ejecución del contrato", "Contrato", "Contratista", 85, "G"),
    "fecha_aprobaci_n_garant_as": ("Fecha de Aprobación de las garantías de ejecución del contrato", "Contrato", "Entidad", 86, "G"),
    "fecha_aprobaci_n_garant_as_1": ("Fecha de Aprobación de las garantías o pólizas e inicio de ejecución del contrato", "Contrato", "Entidad", 87, "G"),
    "fecha_de_inicio_de_ejecuci": ("Fecha de inicio de ejecución del contrato", "Contrato", "Entidad y contratista", 88, ""),
    "plazo_de_ejecuci_n_del": ("Plazo de ejecución del contrato", "Ejecución", "Contratista", 90, ""),
}
