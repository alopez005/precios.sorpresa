# ⚡ PriceSync 1.1 — Actualizador de Precios

**Sistema de sincronización de precios para comercios con tienda física y e-commerce en Tienda Nube.**

PriceSync compara tu catálogo interno contra listas de proveedores (archivos Excel o tiendas web), detecta variaciones de precio, y genera los archivos listos para importar tanto a tu sistema de gestión como a Tienda Nube.

🔗 **Demo en vivo:** [alopez005.github.io/precios.sorpresa](https://alopez005.github.io/precios.sorpresa)

---

## Características principales

### 📁 Paso 1 — Carga de archivos
- Importa tu **archivo maestro** (catálogo actual) y la **lista del proveedor** (nuevos precios) en formato `.xls` / `.xlsx` / `.csv`.
- Detección automática de columnas (EAN, descripción, precio, código interno).
- Detección automática de fila de encabezados para archivos con formato irregular.
- Soporte para múltiples hojas en un mismo libro Excel.
- Selector de modo: **Juguetes/General** (precio de costo + IVA + margen) o **Libros** (PVP directo del proveedor).

### 🔍 Paso 2 — Revisión y matching
- **Match por EAN** (código de barras): prioridad máxima, matching exacto.
- **Match por descripción**: normalización Unicode, comparación exacta y compacta (sin espacios ni caracteres especiales).
- Dashboard con estadísticas: productos que subieron, bajaron, sin cambio, y sin match.
- Filtros interactivos por estado y barra de búsqueda.
- Configuración de margen de ganancia (sobre costo o sobre venta) con recálculo en vivo.
- Selección individual o masiva de productos a exportar.
- Herramienta de **diagnóstico** para inspeccionar cómo el sistema lee los códigos de cada archivo.

### 📊 Paso 3 — Exportación
- **Excel Interno**: genera un `.xlsx` con los precios actualizados, columnas de auditoría (precio anterior, nuevo, variación %, estado, tipo de match, fecha).
- **CSV para Tienda Nube**: sube tu Excel de gestión actualizado + el CSV actual de Tienda Nube, y PriceSync genera un CSV listo para importar con solo los precios modificados. Vista previa antes de descargar.
- Modo **Tienda Nube Directo** (Panel 4): flujo acelerado sin pasar por los pasos 1-2, ideal para actualizaciones recurrentes.

### 🌐 Paso 5 — Proveedores Web (Scraping)
- Registrá proveedores que venden online (tiendas Tiendanube) con nombre y URL de categoría.
- **Scraping automático** con múltiples estrategias de extracción:
  - `googleItems` (array JS embebido en tiendas Tiendanube)
  - Atributos `js-item-name` + `data-product-price` del DOM
  - JSON-LD estructurado (`schema.org/Product`)
  - Detección de links a productos con precio cercano
- **Paginación automática**: detecta `LS.productsCount` y scrapea todas las páginas.
- Uso de proxies CORS (`codetabs`, `allorigins`) para acceso cross-origin.
- Detección de bloqueo Cloudflare con fallback automático.
- **Matching inteligente** contra tu catálogo:
  - Fase 1: filtra por marca del proveedor (threshold 50%).
  - Fase 2: búsqueda general con threshold más estricto (70%).
  - Matching por nombre exacto, contención y token overlap.
- Los proveedores se persisten en `localStorage` del navegador.
- Resultados exportables a Excel interno o directamente a CSV de Tienda Nube.
- Opción de enviar resultados al flujo principal (Paso 2) para continuar con la exportación estándar.

---

## Stack técnico

| Componente | Tecnología |
|---|---|
| Frontend | HTML5, CSS3 (custom properties, responsive), Vanilla JS |
| Excel I/O | [SheetJS (xlsx)](https://sheetjs.com/) v0.18.5 vía CDN |
| Tipografía | DM Sans + Space Mono (Google Fonts) |
| Deploy | GitHub Pages (archivo estático, zero dependencies) |
| Almacenamiento | `localStorage` para proveedores web |

**Sin backend, sin frameworks, sin build step.** Todo corre en el navegador del usuario. Los archivos nunca salen de la máquina.

---

## Estructura del proyecto

```
precios.sorpresa/
├── index.html      # Estructura HTML (5 paneles + modal de diagnóstico)
├── app.js          # Lógica completa (~1900 líneas)
│   ├── Toast & Confirm dialogs
│   ├── Normalización (EAN, descripciones, precios)
│   ├── Similaridad (Levenshtein, token overlap)
│   ├── Lectura de Excel/CSV
│   ├── Motor de matching (EAN → desc exacta → compacta)
│   ├── Renderizado de tablas y estadísticas
│   ├── Exportación Excel interno
│   ├── Flujo Tienda Nube (2 archivos)
│   ├── Scraping de proveedores web (Panel 5)
│   └── Matching web → catálogo (brand-first + fallback)
├── styles.css      # Estilos (~1000 líneas, dark theme, responsive)
└── README.md
```

---

## Uso rápido

### Flujo estándar (Pasos 1 → 2 → 3)

1. Seleccioná el modo (**Juguetes** o **Libros**).
2. Cargá tu archivo maestro y la lista del proveedor.
3. Ajustá las columnas si la detección automática no fue precisa.
4. Clic en **Procesar archivos**.
5. Revisá los resultados, ajustá márgenes si es necesario.
6. Exportá el Excel interno y/o generá el CSV para Tienda Nube.

### Flujo directo Tienda Nube (Panel 4)

1. Clic en **🛒 Tienda Nube directo**.
2. Subí tu Excel de gestión (con PVP actualizado) + el CSV de Tienda Nube.
3. PriceSync compara por EAN y genera el CSV con los precios nuevos.

### Flujo Proveedores Web (Panel 5)

1. Clic en **🌐 Proveedores Web**.
2. Registrá proveedores con su URL de catálogo.
3. Clic en **Iniciar sincronización** — PriceSync scrapea los precios.
4. Revisá los resultados y exportá directo a Excel o Tienda Nube.

---

## Configuración de IVA (modo Juguetes)

- **Neto (sin IVA):** el precio del proveedor se multiplica por `(1 + alícuota)`.
- **Con IVA incluido:** se usa el precio tal cual.
- Alícuota configurable (default: 21%).
- En modo Libros el IVA no aplica (PVP directo).

---

## Normalización de datos

PriceSync aplica normalización agresiva para maximizar el matching:

- **EAN/códigos**: elimina espacios, guiones, puntos, notación científica. Convierte `7.16E+12` → `7160000000000`.
- **Descripciones**: lowercase, strip acentos (NFD), elimina caracteres especiales, normaliza espacios.
- **Precios**: parsea formatos argentinos (`1.234,56`), americanos (`1,234.56`), y mixtos.

---

## Privacidad

Todos los archivos se procesan 100% en el navegador. Nada se sube a ningún servidor. Los proveedores web se almacenan en `localStorage` del navegador local.

---

## Autor

Desarrollado por **Andrés López** — Esperanza, Santa Fe, Argentina.

Herramienta interna para [Sorpresa Online](https://sorpresaonline.ar).
