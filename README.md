# Generador de Catálogo PDF - Distribuidora Bellingieri

Este proyecto es un script en Python diseñado para automatizar la creación de catálogos comerciales en formato PDF. Lee la información de los productos desde un archivo Excel (`catalogo.xlsx`), inserta las imágenes correspondientes desde una carpeta local, y genera un catálogo con un diseño elegante y personalizado (marcos dorados, cabecera negra, 4 productos por página). Al finalizar, fusiona automáticamente una portada en formato PDF al inicio del catálogo.

## 🛠️ Requisitos Previos

Para ejecutar este proyecto, necesitas tener instalado **Python** en tu sistema. Se recomienda usar Python 3.10 o superior.

### Instalación de Python (Linux / Ubuntu)
Si estás en Ubuntu o Debian, puedes instalar Python y las herramientas necesarias (incluyendo el gestor de entornos virtuales) abriendo una terminal y ejecutando:

```bash
sudo apt update
sudo apt install python3 python3-pip python3-venv

mi_catalogo/
├── catalogo.xlsx          # Base de datos de los productos (Excel)
├── generar_catalogo.py    # El script principal de Python
├── logo.png               # Logo de la empresa (idealmente con fondo transparente)
├── portada.pdf            # PDF con el diseño de la portada oficial
├── README.md              # Este archivo de instrucciones
├── .gitignore             # Archivo para ignorar carpetas pesadas en Git
└── imagenes/              # Carpeta que contiene todas las fotos de los productos
    ├── img-0001.jpg
    ├── img-0002.jpg
    └── ...


cd ruta/a/tu/carpeta/mi_catalogo
python3 -m venv .venv
source .venv/bin/activate

pip install pandas openpyxl reportlab pillow pypdf

python generar_catalogo.py