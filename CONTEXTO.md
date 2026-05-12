# Proyecto: Generador de Órdenes de Producción (OPP)

## Resumen
Aplicación web en Python/Flask que automatiza la generación de OPPs a partir de un Excel de entrada, produciendo el archivo `OPP's_jumbo.xlsx` para carga en ERP (Generic/Siesa). Actualmente operativo para el tipo **Stock**; tipos RQ y SP están pendientes de implementación.

---

## Stack tecnológico
| Capa | Tecnología |
|---|---|
| Backend | Python 3.14 + Flask 3.1.3 |
| Procesamiento Excel | openpyxl 3.1.x |
| Base de datos | Firebase Realtime Database (proyecto: picking-d3107) |
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
| Firebase DB | https://picking-d3107-default-rtdb.firebaseio.com |
| Email cuenta | asistente.logistica2@gmail.com |

---

## Estructura de archivos
```
Proyecto Control Piso OPP/
├── app.py                        ← Flask: rutas y lógica web
├── Procfile                      ← Render: web: gunicorn app:app
├── requirements.txt              ← Dependencias con versiones exactas
├── firebase-credentials.json     ← Credenciales Firebase LOCAL (no sube a GitHub)
├── .env                          ← Variables locales (no sube a GitHub)
├── .gitignore
│
├── logic/
│   ├── __init__.py
│   ├── generator.py              ← Motor de generación + load_referencias_stock()
│   ├── excel_io.py               ← Lectura/escritura Excel, templates, lookup referencias
│   └── firebase_db.py            ← Conexión Firebase RTDB (load/save estructura y referencias)
│
├── templates/
│   ├── base.html                 ← Layout base: header sticky, design system Inter Doors
│   ├── index.html                ← Página principal (tipo OPP + carga archivo)
│   ├── resultados.html           ← Vista previa tabla OPPs + botón descarga jumbo
│   └── estructura.html           ← Gestión de Referencias Productivas
│
├── static/
│   └── img/
│       ├── favicon.png
│       ├── LOGO_ID.png           ← Logo Inter Doors (navbar)
│       └── ISOTIPO-INTERDOORS-POSITIVO.png
│
└── config/
    └── estructura.json           ← Fallback local (no usado en producción, Firebase es fuente)
```

---

## Credenciales Firebase
- **Producción (Render):** variable de entorno `FIREBASE_CREDENTIALS` con el JSON completo del service account.
- **Local:** archivo `firebase-credentials.json` en la raíz del proyecto (en `.gitignore`).
- El código detecta automáticamente cuál usar (`firebase_db.py → _init()`).

---

## Rutas de la aplicación
| Método | Ruta | Descripción |
|---|---|---|
| GET | `/` | Página principal con formulario |
| POST | `/generar` | Procesa Excel y genera OPPs |
| GET | `/descargar/<token>/jumbo` | Descarga `OPP's_jumbo.xlsx` |
| GET | `/plantilla` | Descarga plantilla de entrada vacía (Fecha Programación / Referencia / Color / Cantidad) |
| GET | `/estructura` | Ver/gestionar Referencias Productivas |
| GET | `/referencias/plantilla` | Descarga template unificado de 14 columnas |
| POST | `/referencias/importar` | Importa referencias desde Excel a Firebase |

---

## Lógica de negocio — Stock

### Flujo principal
1. Usuario selecciona **Stock** (RQ y SP deshabilitados — próximamente)
2. Sube Excel con columnas: `Fecha · Referencia · Color · Cantidad`
3. El sistema busca cada fila en Firebase por clave `REFERENCIA|COLOR`
4. Si col **Proceso 1** (H) tiene letra Y col **Proceso 2** (J) también → genera **2 OPPs**
5. Si solo col **Proceso 1** tiene letra → genera **1 OPP**
6. El consecutivo de OPPs arranca desde 1 en cada generación (el ERP asigna el número definitivo)
7. Genera `OPP's_jumbo.xlsx` en memoria y lo ofrece para descarga

### Mapeo `OPP's_jumbo.xlsx`

**Hoja Documentos** (una fila por OPP):
| Columna | Valor |
|---|---|
| CONSECUTIVO DCTO | Número OPP (1, 2, 3...) |
| FECHA AAAAMMDD | Fecha del día en que se genera el archivo |
| PLANIFICADOR | `71364487` (fijo) |
| REF1 | Letra Proceso 1 (col H) si es OPP1 / Letra Proceso 2 (col J) si es OPP2 |
| REF2 | Col J (REF2) — igual en ambas OPPs |
| REF3 | Vacío |
| NOTAS | Notas Proceso 1 (col M) si es OPP1 / Notas Proceso 2 (col N) si es OPP2 |

**Hoja Items** (una fila por OPP):
| Columna | Valor |
|---|---|
| NUMERO DCTO | Mismo número OPP |
| REGISTRO MVTO | Consecutivo de ítems por documento (1, 2, 3...) |
| REFERENCIA | Col B (Referencia PRD) si OPP1 con 2 OPPs / Col A en los demás casos |
| EXT1 | Col E (Color #) |
| EXT2 | Col F (Medida) |
| U.M | Col G (Unidad de medida) |
| CANT PLANEADA | Cantidad del input |
| FECHA INICIO | Fecha del input si OPP1 / Fecha input + 2 días si OPP2 |
| FECHA TERMINACION | Igual que FECHA INICIO según OPP |
| METODO LISTA | `0001` (texto) |
| METODO RUTA | `0001` (texto) |
| MEDIDA REAL | Vacío |
| BODEGA | `80123` si genera 1 OPP / OPP1: `80106`, OPP2: `80123` si genera 2 OPPs |

---

## Referencias Productivas (Firebase)

### Estructura en Firebase
Nodo `/estructura`. Entradas con `|` en la clave son referencias productivas:
```
"PUDT0260|CEDRO SIL": {
  "referencia_a": "PUDT0260",
  "referencia_b": "PUDC0260",
  "descripcion": "PUERTA DEKO 60x200 cm",
  "color": "CEDRO SIL",
  "color_num": 284,
  "medida": "STD",
  "um": "UNI",
  "ref1": "A",
  "nombre_proceso1": "Arborit",
  "ref2_i": "E",
  "nombre_proceso2": "Empaque",
  "ref2_j": "STOCK",
  "notas1": "TODAS VAN A 2 MTS...",
  "notas2": "TODAS VAN A 2 MTS..."
}
```

### Descarga de plantilla con datos actuales
La ruta `/referencias/plantilla` genera el Excel de 14 columnas **con todos los registros actuales de Firebase**, no vacío. Permite descargar, editar y re-importar.

### Template Excel unificado (14 columnas)
| Col | Nombre | Descripción |
|---|---|---|
| A | Referencia | Código referencia (ej: PUDT0260) |
| B | Referencia PRD | Código producción (ej: PUDC0260) |
| C | Descripción | Nombre del producto |
| D | Color | Nombre del color |
| E | Color # | Código numérico del color |
| F | Medida | Medida estándar |
| G | U.M | Unidad de medida |
| H | Proceso 1 | Letra del primer proceso (A, M, E...) |
| I | Nombre Proceso 1 | Nombre completo (informativo) |
| J | Proceso 2 | Letra del segundo proceso (vacío si solo 1 OPP) |
| K | Nombre Proceso 2 | Nombre completo (informativo) |
| L | REF2 | Valor columna REF2 del ERP |
| M | Notas Proceso 1 | Instrucciones OPP1 |
| N | Notas Proceso 2 | Instrucciones OPP2 |

> Las claves en Firebase se sanitizan: caracteres `$ # [ ] / .` se reemplazan por `-`.

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
--navy: #323957       /* acento azul */
--teal: #0C8E82       /* acento verde */
```

### Header (sticky, 64px)
- **Izquierda**: Logo `LOGO_ID.png`
- **Centro**: Título de página en uppercase gris
- **Derecha**: Tabs con subrayado amarillo en activo (Inicio · Referencias)

---

## Lectura del Excel de entrada

- El encabezado de la columna de fecha se normaliza (sin tildes, minúsculas) antes de buscarlo en el mapping, por lo que `"Fecha"`, `"Fecha Programación"` y `"Fecha Programacion"` son equivalentes.
- Si la fecha viene como objeto `datetime`/`date` de openpyxl se usa directamente.
- Si viene como texto se intenta parsear con estos formatos en orden: `YYYY-MM-DD`, `DD/MM/YYYY`, `DD-MM-YYYY`, `YYYY/MM/DD`, `YYYYMMDD`. Si ninguno funciona se usa `date.today()` como último recurso.

---

## Flujo de trabajo (desarrollo continuo)
```bash
# 1. Correr localmente (requiere firebase-credentials.json en raíz)
venv\Scripts\python app.py
# Abrir: http://127.0.0.1:5000

# 2. Subir cambios
git add <archivos>
git commit -m "descripción"
git push
# Render detecta el push y redespliega (~2 min)
```

---

## Estado actual del proyecto
- [x] Estructura base Flask + Render configurada
- [x] Firebase Realtime Database como persistencia (reemplazó JSON locales)
- [x] Lógica Stock: lookup por Referencia+Color, genera 1 o 2 OPPs según columnas H/J
- [x] Salida `OPP's_jumbo.xlsx` con mapeo completo de columnas ERP
- [x] Consecutivo OPP inicia desde 1 por generación (ERP asigna número real)
- [x] UI con design system Inter Doors
- [x] Selector tipo OPP (Stock activo / RQ y SP próximamente)
- [x] Plantilla entrada vacía con encabezado "Fecha Programación" · Referencia · Color · Cantidad
- [x] Referencias Productivas: template 14 cols, importación masiva a Firebase
- [x] Tabla referencias muestra nombre completo de procesos
- [x] Planificador actualizado a `71364487`
- [x] Campo BODEGA: `80123` si 1 OPP; OPP1 `80106` / OPP2 `80123` si 2 OPPs
- [x] Plantilla de referencias descarga con datos actuales de Firebase
- [x] Tabla referencias: paginación (8/página), búsqueda y filtros (Proceso / Color / Tipo)
- [x] Badges de proceso con color por nombre; badge de referencia oscuro con punto amarillo
- [x] Panel importación referencias con zona drag-and-drop
- [ ] Implementar lógica RQ
- [ ] Implementar lógica SP
- [ ] Historial de OPPs generadas

---

## Notas importantes
- **Render plan gratuito:** hiberna el servicio tras 15 min sin uso. Primera carga puede demorar ~30 seg.
- **Archivos generados:** se guardan en memoria (`_cache`). Si el servidor reinicia, los links de descarga dejan de funcionar (hay que generar de nuevo).
- **Firebase gratuito (Render):** la base de datos de Render expira en 90 días. Los datos están en Firebase que es permanente.
- **Puerto local:** Flask corre en `5000`. Si hay conflicto: `netstat -ano | findstr :5000`.
