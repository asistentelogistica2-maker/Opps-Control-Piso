# Proyecto: Generador de Órdenes de Producción (OPP)

## Resumen
Aplicación web en Python/Flask que automatiza la generación de OPPs a partir de un Excel de entrada, produciendo dos archivos de salida: uno para carga en ERP (Generic/Siesa) y otro para impresión de stickers (Bartender).

---

## Stack tecnológico
| Capa | Tecnología |
|---|---|
| Backend | Python 3.14 + Flask 3.1.3 |
| Procesamiento Excel | openpyxl 3.1.x |
| Servidor producción | Gunicorn 26.0.0 |
| Frontend | Bootstrap 5.3 + Bootstrap Icons + DM Sans (Google Fonts) |
| Deploy | Render (plan gratuito) |
| Repositorio | GitHub |

---

## Cuentas y URLs
| Recurso | Valor |
|---|---|
| GitHub usuario | asistentelogistica2-maker |
| Repositorio | https://github.com/asistentelogistica2-maker/Opps-Control-Piso |
| App en producción | https://opps-control-piso.onrender.com |
| Email cuenta | asistente.logistica2@gmail.com |

---

## Estructura de archivos
```
Proyecto Control Piso OPP/
├── app.py                        ← Flask: rutas y lógica web
├── Procfile                      ← Render: web: gunicorn app:app
├── requirements.txt              ← Dependencias con versiones exactas
├── .env                          ← Variables locales (no se sube a GitHub)
├── .gitignore
├── README.md
│
├── logic/
│   ├── __init__.py
│   ├── generator.py              ← Motor de generación de OPPs
│   └── excel_io.py               ← Lectura/escritura de Excel (BytesIO)
│
├── templates/
│   ├── base.html                 ← Layout base: header sticky, design system Inter Doors
│   ├── index.html                ← Página principal (carga + opciones)
│   ├── resultados.html           ← Vista previa + descarga de archivos
│   └── estructura.html           ← CRUD de estructura productiva
│
├── static/
│   └── img/
│       ├── favicon.png           ← Favicon de la app
│       ├── LOGO_ID.png           ← Logo Inter Doors (navbar)
│       ├── Diagrama de Proceso Control Piso.png
│       ├── Preview Generacion de OPP.png
│       └── Preview Generacion de OPP 2.png
│
├── config/
│   └── estructura.json           ← Mapa referencia → procesos (persistente en git)
│
└── data/
    └── opp_counter.json          ← Contador secuencial de OPPs
```

---

## Rutas de la aplicación
| Método | Ruta | Descripción |
|---|---|---|
| GET | `/` | Página principal con formulario |
| POST | `/generar` | Procesa Excel y genera OPPs |
| GET | `/descargar/<token>/<tipo>` | Descarga Excel ERP o Stickers |
| GET | `/plantilla` | Descarga plantilla de entrada |
| GET | `/estructura` | Ver/gestionar estructura productiva |
| POST | `/estructura/guardar` | Crear o editar una referencia |
| POST | `/estructura/eliminar` | Eliminar una referencia |

---

## Lógica de negocio

### Flujo principal
1. Usuario selecciona **Tipo de OPP** (Stock / RQ / SP)
2. Sube Excel con columnas: `Cliente, Referencia, Cantidad, Notas del ítem, Notas generales`
3. El sistema busca cada referencia en `estructura.json`
4. Por cada proceso en la ruta de producción → genera una OPP con número secuencial
5. Solo el **último proceso** genera stickers (cantidad = número de piezas del ítem)
6. Genera dos Excel en memoria (BytesIO) y los ofrece para descarga

### Regla de stickers por tipo de OPP
- **Stock**: los stickers ya existen como imágenes en Google Drive (pre-fabricados). La opción de generar stickers se **deshabilita automáticamente** en el formulario cuando se selecciona Stock. Pendiente: conectar link de carpeta Drive.
- **RQ / SP**: se generan stickers normalmente en Excel para Bartender.
- Cantidad de stickers = cantidad del ítem. Formato de pieza: `1/5`, `2/5`, `3/5`...

### Estructura productiva
Almacenada en `config/estructura.json`:
```json
{
  "REF001": {
    "descripcion": "Nombre del producto",
    "procesos": ["Corte", "Costura", "Terminado", "Empaque"]
  }
}
```
> ⚠️ Cambios hechos desde la web se pierden al redesplegar. Para persistirlos: descargar el JSON, reemplazar `config/estructura.json` en el repo y hacer push.

---

## Tipos de OPP disponibles
| Tipo | Descripción | Stickers |
|---|---|---|
| Stock | Producción para inventario | Desde Google Drive (imágenes pre-hechas) |
| RQ | Requisición de cliente | Se generan en Excel (Bartender) |
| SP | Servicio / Proyecto | Se generan en Excel (Bartender) |

---

## Salidas generadas

### Excel ERP (`OPP_ERP.xlsx`) — formato Siesa/Generic

> **Esquema maestro-detalle (header-detail):** relación 1:N entre hojas.
> - **Hoja 1 – Cabecera (maestro):** una fila por pedido. `ID_Orden` es clave primaria.
> - **Hoja 2 – Líneas (detalle):** una fila por ítem. `ID_Orden` referencia al maestro; `ID_Linea` es consecutivo propio de cada ítem (clave primaria del detalle).

**Hoja 1: Documentos (maestro)**
| ID_Orden | CONSECUTIVO DCTO | FECHA AAAAMMDD | PLANIFICADOR | REF1 | REF2 | REF3 | NOTAS |
| Clave primaria | OPP número | Fecha hoy | *(pendiente)* | Cliente | *(pendiente)* | *(pendiente)* | Notas generales |

**Hoja 2: Items (detalle)**
| ID_Orden | ID_Linea | NUMERO DCTO | REGISTRO MVTO | REFERENCIA | EXT1 | EXT2 | U.M | CANT PLANEADA | FECHA INICIO | FECHA TERMINACION | METODO LISTA | METODO RUTA | MEDIDA REAL | BODEGA |
| FK → maestro | Consecutivo ítem | OPP número | *(pendiente)* | Referencia | *(pendiente)* | *(pendiente)* | *(pendiente)* | Cantidad | *(pendiente)* | *(pendiente)* | *(pendiente)* | Proceso | Notas ítem | *(pendiente)* |

> Las columnas marcadas *(pendiente)* quedan vacías hasta confirmar el mapeo con el ERP.

### Excel Stickers (`OPP_Stickers.xlsx`) — solo para RQ y SP
| Cliente | Numero de documento | Medida real | Numero de pieza | Cantidad |

### PDF Stickers (`stickers.pdf`) — reemplaza Excel
El archivo Excel de stickers ha sido reemplazado por generación en PDF con especificaciones precisas:

**Características**:
- **Dimensiones exactas**: 100mm × 25mm (una página por sticker)
- **Layout profesional**: 4 zonas (Logo | Info | Pieza | QR)
  - Logo (0-21mm): ISOTIPO-INTERDOORS-POSITIVO.png centrado + divisor gris
  - Info (21-66mm): Cliente, doc, medida con divisores y acento negro
  - Pieza (66-80mm): Número de pieza (ej: 1/15) centrado
  - QR (80-100mm): Código automático `Interdoors|cliente|doc|medida|pieza`
- **Generación en tiempo real**: Desde Flask sin escritura a disco
- **Descarga directa**: Botón "Descargar Stickers" en resultados

**Integración en app.py** (~línea 58-74):
1. Se reciben `sticker_rows` de `generate_opps()`
2. Se convierten al formato requerido (cliente, doc, medida, pieza)
3. Se llama `export_stickers_from_orders(stickers_data, temp_pdf)`
4. El PDF se cachea en memoria (_cache) para descarga

```python
# Convertir formato
stickers_data = [{
    'cliente': row['Cliente'],
    'doc': str(row['Numero de documento']),
    'medida': row['Medida real'],
    'pieza': row['Numero de pieza']
} for row in sticker_rows]

# Generar PDF
export_stickers_from_orders(stickers_data, temp_pdf)
```

---

## Design System — Inter Doors
Implementado en `templates/base.html` vía CSS custom properties:

```css
--ink: #2A2927        /* texto principal */
--ink-2: #4A4946      /* texto secundario */
--ink-3: #6E6D6A      /* captions */
--line: #DEDDD9       /* bordes */
--bg: #ECEAE5         /* fondo página */
--surface-2: #E8E6E1  /* header, card headers */
--surface: #FFFFFF    /* cards */
--yellow: #F3C615     /* amarillo primario (CTAs, activo) */
--yellow-mid: #C9A500 /* amarillo texto sobre fondo claro */
--yellow-pale: #FBF5D6/* fondo hover sutil */
--teal: #0C8E82       /* acento verde (stickers) */
--navy: #323957       /* acento azul (ERP) */
```

### Header (sticky, 64px)
- **Izquierda**: Logo `LOGO_ID.png`
- **Centro**: Título de página en uppercase gris
- **Derecha**: Tabs con subrayado amarillo en activo (Inicio · Estructura)
- Fondo: `surface-2` (#E8E6E1), sombra inferior

### Comportamiento UI notable
- Tipo OPP **Stock** → toggle Stickers se apaga y bloquea automáticamente (JS)
- Botón "Descargar plantilla" vive dentro del bloque Tipo de OPP
- Estructura Productiva solo accesible desde el header (no en inicio)
- Tab Stickers en resultados: solo conteo + botón descarga (sin tabla detallada)

---

## Flujo de trabajo (desarrollo continuo)
```bash
# 1. Correr localmente
venv\Scripts\python app.py
# Abrir: http://127.0.0.1:5000

# 2. Cuando los cambios están listos
git add .
git commit -m "descripción del cambio"
git push
# Render detecta el push y redespliega automáticamente (~2 min)
```

---

## Estado actual del proyecto
- [x] Estructura base Flask + Render configurada
- [x] Lógica de generación de OPPs (generator.py)
- [x] Lectura/escritura Excel en memoria (sin archivos locales)
- [x] UI con design system Inter Doors (DM Sans, tokens corporativos)
- [x] Header sticky profesional (logo + título centrado + tabs)
- [x] Selector de tipo de OPP (Stock / RQ / SP) con hover amarillo
- [x] Stock deshabilita stickers automáticamente
- [x] Plantilla Excel dentro del bloque Tipo de OPP
- [x] CRUD de estructura productiva desde el header
- [x] Excel ERP con 2 hojas (Documentos + Items) — formato Siesa
- [x] Descarga directa de Excel generados
- [x] Favicon y logo en static/img/
- [x] **Stickers en PDF** (reemplazó Excel) - 100mm×25mm, QR automático, multi-zona
- [ ] Mapeo completo de columnas ERP (pendiente confirmar con ERP)
- [ ] Link Google Drive stickers Stock (pendiente tener imágenes)
- [ ] Mostrar imagen sticker desde Drive en resultados (Stock)
- [ ] Persistencia de estructura productiva (base de datos o env var)
- [ ] Autenticación / login
- [ ] Historial de OPPs generadas

---

## Notas importantes
- **Render plan gratuito:** hiberna el servicio tras 15 min sin uso. Primera carga puede demorar ~30 seg.
- **Archivos generados:** se guardan en memoria. Si el servidor reinicia, los links de descarga dejan de funcionar (hay que generar de nuevo).
- **Contador OPP:** `data/opp_counter.json` se resetea con cada nuevo deploy en Render.
- **Puerto local:** Flask corre en `5000`. Si hay conflicto de puerto, matar procesos con `netstat -ano | findstr :5000`.
- **Columnas ERP pendientes:** PLANIFICADOR, REF2, REF3, REGISTRO MVTO, EXT1, EXT2, U.M, FECHA INICIO, FECHA TERMINACION, METODO LISTA, BODEGA — quedan vacías hasta confirmar mapeo.
