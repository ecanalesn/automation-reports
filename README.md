# 📊 Report Automation Toolkit: Automatización de Informes de Negocio

![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776ab)
![pandas](https://img.shields.io/badge/pandas-150458)
![openpyxl](https://img.shields.io/badge/openpyxl-Excel-217346)
![PHP](https://img.shields.io/badge/PHP-WooCommerce-777bb4)

> ☕ Parte de la lógica está migrada a Java, con Maven, tests y CI, en [Automation Reports Java](https://github.com/ecanalesn/automation-reports-java).

## 📋 Descripción del Proyecto
Report Automation Toolkit es una colección de scripts en Python que convierten los exports de pedidos de una tienda online (WooCommerce) en informes de Excel con formato, listos para el equipo. Cada script resuelve un caso real: el informe de ventas semanal, la previsión de ingresos de productos a plazos, la detección de impagos y la exportación incremental de las conversaciones de un chat. Se completa con dos snippets de WooCommerce en PHP.

> [!NOTE]
> Es la versión genérica y anonimizada de herramientas que se usaban **en producción** para sustituir tareas manuales recurrentes: consolidar exports de varias tiendas y divisas, clasificar productos y preparar cada semana los informes para el equipo.

## ✨ Características Principales
- 🔗 **Consolidación de Datos**: une los exports de varias tiendas y divisas en un único conjunto, convertido a la divisa base.
- 🏷️ **Clasificación por Reglas**: asigna cada producto a una categoría con expresiones regulares compiladas, tras normalizar el texto.
- 📄 **Excel con Formato**: hojas por categoría, estilos copiados desde una plantilla, rellenos condicionales, hipervínculos y validación de datos.
- 📅 **Previsión de Ingresos**: proyecta mes a mes los plazos pendientes de los productos a plazos sobre un horizonte de fechas.
- ⚠️ **Detección de Impagos**: separa los impagos ya reportados de los recién detectados y marca los importes anómalos (por ejemplo, un plazo cobrado en otra divisa).
- 🔄 **Exportación Incremental y Reanudable**: al relanzarla solo reexporta las conversaciones nuevas o con mensajes nuevos; las demás se reutilizan.
- ⚡ **Descargas Concurrentes**: llamadas a la API en paralelo con `ThreadPoolExecutor`.
- 🧩 **Snippets de WooCommerce**: comisión de afiliado por tramos y campo de checkout condicional por país.

### Scripts incluidos

| Script | Entrada | Salida |
|---|---|---|
| [`sales_report.py`](sales_report/sales_report.py) | Exports de pedidos (.xlsx) + plantilla | Informe semanal de ventas con una hoja por categoría |
| [`forecast_report.py`](forecast_report/forecast_report.py) | Exports de pedidos (.xlsx) | Previsión de ingresos mes a mes |
| [`arrears_report.py`](arrears_report/arrears_report.py) | Exports de pedidos + impagos conocidos (opcional) | Impagos reportados y nuevos, resaltados en colores distintos |
| [`chat_export.py`](chat_export/chat_export.py) | API REST del proveedor de chat | Índice en Excel con hipervínculos + un PDF y los adjuntos de cada conversación |

## 🛠️ Tecnologías Utilizadas
- **Lenguaje**: Python 3.10+
- **Análisis de Datos**: pandas
- **Generación de Excel**: openpyxl (estilos, rellenos, hipervínculos y validación de datos)
- **Peticiones HTTP**: requests
- **Concurrencia**: `concurrent.futures` (ThreadPoolExecutor)
- **Normalización de Texto y Reglas**: unicodedata y expresiones regulares (`re`)
- **Snippets Web**: PHP sobre WooCommerce y AffiliateWP
- **Control de versiones**: Git

## 📦 Requisitos Previos
- Python 3.10 o superior
- pip
- Git

## 📁 Estructura del Proyecto
```text
automation-reports/
├── sales_report/
│   └── sales_report.py                 # Informe de ventas semanal por categoría (Excel sobre plantilla)
├── forecast_report/
│   └── forecast_report.py              # Previsión de ingresos de productos a plazos, mes a mes
├── arrears_report/
│   └── arrears_report.py               # Impagos: reportados vs. detectados + importes anómalos
├── chat_export/
│   └── chat_export.py                  # Exportador incremental y reanudable de conversaciones (índice + PDFs)
├── wordpress_snippets/
│   ├── tiered_commission.php           # Comisión de afiliado por tramos (AffiliateWP + WooCommerce)
│   └── conditional_checkout_field.php  # Campo de checkout condicional por país (WooCommerce)
├── requirements.txt                    # Dependencias (pandas, openpyxl, requests)
├── .gitignore                          # Ignora datos reales (.xlsx, .csv, .pdf, .env)
└── README.md
```

## 🚀 Instalación y Ejecución

### 1. Clonar el repositorio
```bash
git clone https://github.com/ecanalesn/automation-reports.git
cd automation-reports
```

### 2. Instalar las dependencias
```bash
pip install -r requirements.txt
```

### 3. Generar el informe de ventas
```bash
python sales_report/sales_report.py orders_a.xlsx orders_b.xlsx \
    --template template.xlsx --output weekly_report.xlsx
```

### 4. Generar la previsión de ingresos
```bash
python forecast_report/forecast_report.py orders_a.xlsx orders_b.xlsx \
    --output forecast.xlsx --today 2026-06-10 --end 2027-06-30
```

### 5. Generar el informe de impagos
```bash
python arrears_report/arrears_report.py orders_a.xlsx orders_b.xlsx \
    --known known_arrears.xlsx --output arrears_report.xlsx
```

### 6. Exportar las conversaciones (modo incremental)
```bash
export CHAT_API_KEY=tu_token
python chat_export/chat_export.py --out ./export
```

Cada script es una CLI independiente. El docstring del principio de cada archivo explica todas sus opciones.

### Opciones de la línea de comandos

| Script | Opción | Descripción | Por defecto |
|---|---|---|---|
| `sales_report` | `orders…` | Uno o más exports de pedidos (.xlsx) | (obligatorio) |
| | `--template` | Libro de plantilla | `template.xlsx` |
| | `--output` | Excel de salida | `weekly_report.xlsx` |
| `forecast_report` | `orders…` | Uno o más exports de pedidos (.xlsx) | (obligatorio) |
| | `--today` | Fecha de referencia (AAAA-MM-DD) | hoy |
| | `--end` | Horizonte de la previsión (AAAA-MM-DD) | el de la configuración |
| | `--output` | Excel de salida | `forecast.xlsx` |
| `arrears_report` | `orders…` | Uno o más exports de pedidos (.xlsx) | (obligatorio) |
| | `--known` | Archivo de impagos ya conocidos (.xlsx) | (opcional) |
| | `--output` | Excel de salida | `arrears_report.xlsx` |
| `chat_export` | `--out` | Carpeta de salida | `./export` |

## 🔧 Configuración
- **Taxonomía de productos**: cada script tiene una sección de `CONFIGURACIÓN` al principio con las reglas de categorías (expresiones regulares), las columnas de salida y los nombres de hoja. Los valores de ejemplo se sustituyen por los propios.
- **Plantilla de Excel**: `sales_report` copia los estilos de una plantilla que hay que proporcionar (la fila de cabecera y la fila modelo son configurables).
- **Divisas**: los scripts con importes usan un diccionario de tasas de cambio a la divisa base, editable en la configuración.
- **API de chat**: `chat_export` lee la clave de la variable de entorno `CHAT_API_KEY` (y, opcionalmente, la URL base de `CHAT_API_BASE`). Los endpoints y los campos son marcadores que se adaptan al proveedor real.

## 📐 Cómo Funciona
- **Ventas**: normaliza el nombre del producto (minúsculas, sin acentos), le aplica las reglas de categoría en orden y escribe cada pedido en la hoja de su categoría, copiando el formato de la fila modelo de la plantilla.
- **Previsión**: lee el número de plazos del nombre del producto, cuenta los plazos ya cobrados por cliente y reparte los pendientes mes a mes hasta la fecha de horizonte.
- **Impagos**: agrupa los pedidos por cliente y producto, calcula cuántos plazos mensuales deberían haber vencido desde la primera compra (con un margen de tolerancia en días) y los compara con los cobrados. Después los cruza con la lista manual para separar lo ya controlado de lo nuevo.
- **Exportación de chat**: compara el listado actual de la API con el índice de la ejecución anterior y solo descarga lo nuevo o modificado, en paralelo.

## 🗺️ Posibles Mejoras
- Tests automáticos con pytest y datos de ejemplo ficticios (ya hechos en la [versión de Java](https://github.com/ecanalesn/automation-reports-java)).
- Integración continua con GitHub Actions.
- Configuración externa (YAML) en lugar de constantes al principio de cada script.
- Programar la ejecución semanal y enviar el informe por correo automáticamente.

> [!NOTE]
> Es código de muestra: los productos, categorías, SKU, tasas de cambio y nombres de archivo son marcadores de ejemplo. No contiene datos ni lógica de negocio de terceros; el diseño y la implementación son propios.

**Desarrollado por**: Estefanía Canales
**Fecha**: 12/08/2026


**Desarrollado por**: Estefanía Canales
**Fecha**: 12/08/2026
