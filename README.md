# 📊 Report Automation Toolkit - Automatización de Informes de Negocio

> Colección de scripts en Python para automatizar informes de negocio recurrentes a partir de exports de e-commerce (WooCommerce), más algunos snippets de WooCommerce/PHP. Versión genérica y anonimizada de herramientas desarrolladas en producción.

---

## 📋 Descripción del Proyecto

**Report Automation Toolkit** es un conjunto de utilidades que transforman exports tabulares desordenados de una tienda online en informes de Excel con formato, listos para el equipo. Cada script resuelve un caso real: informes de ventas semanales, previsión de ingresos de productos a plazos, detección de impagos y exportación incremental de conversaciones de un chat.

Todo el código está publicado como **muestra técnica**: los nombres de productos, categorías, SKU, tasas de cambio y nombres de archivo son marcadores de ejemplo. No contiene datos ni lógica de negocio de terceros; el valor está en la estructura y las técnicas, que son reutilizables en cualquier proyecto.

### ✨ Características Principales

- **Consolidación de datos** de varias tiendas y divisas en un único conjunto
- **Clasificación por reglas** de productos mediante expresiones regulares compiladas
- **Generación de Excel con formato**: hojas por categoría, copia de estilos desde plantilla, rellenos condicionales, hipervínculos y validación de datos
- **Previsión de ingresos** de productos a plazos, proyectados mes a mes sobre un horizonte de fechas
- **Detección de impagos** con separación entre reportados y recién detectados, y marcado de anomalías de divisa
- **Exportación incremental y reanudable** de conversaciones (solo reprocesa lo nuevo o modificado)
- **Descargas concurrentes** con `ThreadPoolExecutor` para las llamadas a la API
- **Snippets de WooCommerce/PHP**: comisión de afiliado por tramos y campo de checkout condicional por país

## 🛠️ Tecnologías Utilizadas

- **Lenguaje**: Python 3
- **Análisis de datos**: pandas
- **Generación de Excel**: openpyxl (estilos, rellenos, hipervínculos, validación de datos)
- **Peticiones HTTP**: requests
- **Concurrencia**: `concurrent.futures` (ThreadPoolExecutor)
- **Normalización de texto y reglas**: unicodedata, expresiones regulares (re)
- **Snippets web**: PHP sobre WooCommerce y AffiliateWP

## 📦 Requisitos Previos

- Python 3.10 o superior
- pip
- Git

## 🚀 Instalación y Ejecución

### 1. Clonar el repositorio
```bash
git clone [URL_DEL_REPOSITORIO]
cd automation-reports
```

### 2. Instalar dependencias
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

### 6. Exportar conversaciones (modo incremental)
```bash
export CHAT_API_KEY=tu_token
python chat_export/chat_export.py --out ./export
```

> Cada script es una CLI independiente. Consulta el docstring del principio de cada archivo para ver todas sus opciones.

## 📁 Estructura del Proyecto

```
automation-reports/
├── sales_report/
│   └── sales_report.py         # Informe de ventas semanal por categoría (Excel sobre plantilla)
├── forecast_report/
│   └── forecast_report.py      # Previsión de ingresos de productos a plazos, mes a mes
├── arrears_report/
│   └── arrears_report.py       # Detección de impagos: reportados vs. detectados + anomalías
├── chat_export/
│   └── chat_export.py          # Exportador incremental y reanudable de conversaciones (índice + PDFs)
├── wordpress_snippets/
│   ├── tiered_commission.php        # Comisión de afiliado por tramos (AffiliateWP + WooCommerce)
│   └── conditional_checkout_field.php  # Campo de checkout condicional por país (WooCommerce)
├── requirements.txt            # Dependencias (pandas, openpyxl, requests)
├── .gitignore                  # Ignora datos reales (.xlsx, .csv, .pdf, .env)
└── README.md
```

## 🔧 Configuración

- **Taxonomía de productos**: cada script tiene una sección de `CONFIGURACIÓN` al principio con las reglas de categorías (expresiones regulares), columnas de salida y nombres de hoja. Sustituye los marcadores de ejemplo por los tuyos.
- **Plantilla de Excel**: `sales_report` copia estilos de una plantilla que debes proporcionar (fila de cabecera y fila modelo configurables).
- **Divisas**: los scripts con importes usan un diccionario de tasas de cambio a la divisa base, editable en la configuración.
- **API de chat**: `chat_export` lee la clave de la variable de entorno `CHAT_API_KEY`; los endpoints y campos son marcadores que se adaptan al proveedor real.

## 📝 Nota

Es código de muestra/referencia. Espera columnas de entrada y una plantilla que tú proporciones; la taxonomía de ejemplo está pensada para que la sustituyas por la tuya.

**Autoría**: diseño e implementación propios. Publicado como muestra técnica en versión genérica, sin datos ni lógica de negocio de terceros.

---

**Desarrollado por**: Estefanía Canales
**Fecha**: 12/08/2026
