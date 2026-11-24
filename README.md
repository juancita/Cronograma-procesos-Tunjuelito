# Cronogramas de procesos · SECOP II · FDL Tunjuelito

Seguimiento diario de los procesos de selección del Fondo de Desarrollo Local de Tunjuelito, desde que se publica el proyecto de pliego hasta antes de la adjudicación. El programa consulta los **datos abiertos de SECOP II** y arma un Excel con:
- los hitos que vienen;
- un tablero día a día;
- la etapa de cada proceso;
- las fechas que cambiaron;
- las adendas;
- lo que el equipo ha gestionado.

Solo **lee** información pública de [datos.gov.co](https://www.datos.gov.co). No necesita usuario ni contraseña.

## Descargar y usar (Windows)

1. Descargar **[Cronogramas-SECOP-Tunjuelito.zip](../../releases/latest/download/Cronogramas-SECOP-Tunjuelito.zip)** (última versión, en *Releases*).
2. Descomprimir la carpeta donde se vaya a dejar, por ejemplo en *Documentos*. No ejecutarlo desde dentro del ZIP.
3. Doble clic en **`Actualizar cronogramas.bat`**. Tarda alrededor de un minuto y al final abre **`Seguimiento de cronogramas.xlsx`**.
4. Para actualizar: cerrar el Excel y volver a dar doble clic en el `.bat`.

No hay que instalar nada: el ZIP trae el programa compilado. Si Windows muestra «Windows protegió su PC», elegir *Más información → Ejecutar de todas formas*. El programa no está firmado digitalmente.

### Lo que anota el equipo se conserva

Las columnas con encabezado **verde** las llena el equipo:
- **Nombre corto:** el que aparece en el tablero.
- **Responsable**, **Estado de gestión** y **Observaciones de gestión**.
- **Gestionado** y **Observación del hito**, en cada hito.

Cada vez que se actualiza, el programa lee esas columnas del Excel anterior, las guarda en la base local (con su historia) y las vuelve a poner en el Excel nuevo.

### Carpetas que crea (no borrarlas)

| Carpeta | Qué guarda |
|---|---|
| `datos/cronogramas.db` | La base local: procesos, historial de fechas de cada hito, novedades, adendas, gestión del equipo y registro de cada actualización |
| `datos/consultas/` | Copia comprimida de lo que respondió SECOP II en cada actualización |
| `historial/` | Una copia del Excel por cada actualización |

## Hojas del Excel

| Hoja | Para qué |
|---|---|
| **RESUMEN** | Dónde está cada proceso hoy (por etapa), presupuesto en juego, hitos de hoy y de la semana, cómo se usa |
| **NOVEDADES** | Qué cambió desde la actualización anterior: procesos nuevos, fechas cambiadas, adendas, procesos que salen (adjudicados, cancelados, desiertos) con su adjudicatario |
| **PROXIMOS_HITOS** | Cada hito pendiente en orden de fecha y hora, con alerta por cercanía (hoy, 1-3 días, 4-7, 8-15) y quién actúa. En rojo, lo que le toca a la Entidad |
| **TABLERO_DIARIO** | Procesos × días hábiles: código del hito (AP, R, C, IE, AA, A, F, G) y color de la etapa |
| **PROCESOS_EN_CURSO** | Un proceso por fila: etapa hoy, próximo hito, apertura, cierre, informe, adjudicación, firma, adendas, fechas cambiadas |
| **EVOLUCION** | Cuántos procesos hay en cada etapa día a día, con gráfico |
| **PUNTOS_PARA_REVISAR** | Avisos automáticos: hitos en orden atípico, falta de fecha de adjudicación, hitos de fin de año o en fin de semana, cronogramas muy modificados, cierre distinto entre cronograma y ficha |
| **CAMBIOS_DE_FECHA** | Fecha anterior → fecha vigente de cada hito que cambió |
| **CRONOGRAMA_COMPLETO** | Todos los hitos de cada proceso, pasados y pendientes |
| **ABIERTOS_SIN_FECHAS** | Abiertos sin adjudicar cuyo cronograma ya terminó (posible trámite pendiente en la plataforma) |
| **ESTUDIOS_DE_MERCADO** | Solicitudes de información a proveedores, aparte porque no son procesos de selección |
| **ADENDAS** | Adendas publicadas, con enlace de descarga |
| **REGISTRO_PROCESOS** | Todo proceso que ha estado en seguimiento: cuándo entró, cuándo salió y por qué |
| **HISTORIAL_GESTION** | Cada anotación del equipo con su fecha |
| **REGISTRO_EJECUCIONES** | Cada actualización: qué se consultó y qué se encontró |

## Cómo funciona

| Conjunto de datos abiertos | Qué aporta |
|---|---|
| `p6dx-8zbt` · SECOP II - Procesos de Contratación | Procesos de la entidad (por código y por nombre), su estado, fase, presupuesto y adjudicación |
| `f3ev-294c` · SECOPII - Fechas Cronograma | Los hitos de cada proceso (≈60 posibles según la modalidad) |
| SECOP II - Archivos Descarga Historico (año actual y anterior) | Las adendas publicadas |

### Reglas validadas

Estas reglas se verificaron contra adendas reales, por ejemplo la adenda 2 de FDLT-LP-379-2026:

1. **Un proceso tiene varias filas, una por fase** (base, manifestación de interés, presentación de oferta, fase de selección), y **no traen el mismo cronograma**. Rige la de publicación más reciente. Las otras conservan las fechas de antes de una adenda o ajuste; de ahí sale «Fecha antes del cambio».
2. **El cronograma viene en hora UTC**; el programa lo muestra en hora de Bogotá (−5 h).
3. **Entra al seguimiento** un proceso con estado *Publicado*, *Abierto* o *Evaluación*, no adjudicado y sin ninguna fase en *Seleccionado*, *Adjudicado*, *Cancelado*, *Desierto*, *Terminado* o *Cerrado*. Los borradores no se publican, así que no entran.
4. **Etapa hoy** es la del último hito ya alcanzado:
   - *Pre-pliego*: proyecto de pliego publicado y sus observaciones.
   - *Pliego definitivo*: acto de apertura, audiencias, observaciones y adendas.
   - *Ofertas y evaluación*: cierre, apertura, informe y observaciones.
   - *Adjudicación*.
   - *Contrato*: firma y garantías.
   - *Cronograma cumplido*.

### Límites conocidos

- Los datos abiertos se actualizan una vez al día; una adenda del mismo día puede no estar todavía. **Antes de actuar, confirmar en la página del proceso** (enlace en cada fila).
- El cronograma publicado no trae todos los hitos del pliego (por ejemplo, el informe final de evaluación).
- Una fecha puede cambiar sin una adenda publicada con ese nombre.

## Configuración (`config.ini`)

| Clave | Para qué |
|---|---|
| `codigo_entidad`, `nombres_entidad` | La entidad cuyos procesos se traen |
| `dias_abiertos_sin_fechas` | Ventana de la hoja ABIERTOS_SIN_FECHAS |
| `datasets_archivos` | Respaldo si el catálogo de datos abiertos no responde |
| `token_app` | Opcional, solo si la API empieza a limitar |
| `abrir_excel` | Abrir el Excel al terminar |

Opciones del programa:
- `cronogramas.exe --sin-consultar` rehace el Excel con lo guardado, sin internet.
- `--sin-abrir` no abre el Excel al terminar.
- `--carpeta RUTA` usa otra carpeta de trabajo.

## Para desarrollar

```bash
pip install -r requirements.txt
python -m unittest discover -s tests -v      # pruebas sin internet
python -m cronogramas --carpeta prueba        # corrida real
```

**Publicar una versión:**
1. Subir `__version__` en `cronogramas/__init__.py` y anotar en `CAMBIOS.md`.
2. Crear la etiqueta: `git tag v1.0.1 && git push origin v1.0.1`.

GitHub Actions (`.github/workflows/empaquetar.yml`) corre las pruebas en Windows, compila `cronogramas.exe` con PyInstaller, hace una prueba real contra SECOP II y publica el ZIP en *Releases*. Cada push a `main` deja además el ZIP como artefacto del flujo.

**Si la red de la oficina inspecciona HTTPS**, el programa usa los certificados de Windows (`truststore`) y el proxy configurado en Windows.
