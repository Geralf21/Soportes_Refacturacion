# Escáner y Procesador de PDFs

Este script procesa archivos PDF desde un directorio de entrada, extrae información mediante PyMuPDF (`fitz`) y expresiones regulares, y cruza los datos con una base de datos en Excel utilizando coincidencias de texto (`difflib`).

## Requisitos

El script requiere bibliotecas externas. Instálalas ejecutando el siguiente comando:

```bash
pip install pandas PyMuPDF
```

*(Nota: `os`, `sys`, `re`, `difflib` y `warnings` están incluidas en la biblioteca estándar de Python y no requieren instalación).*

## Estructura de Directorios

Para que el script funcione correctamente mediante sus rutas relativas, el directorio de trabajo debe mantener esta estructura:

*   **`Entrada_PDF/`**: Directorio donde debes colocar los documentos PDF a escanear.
*   **`Salida_PDF/`**: Directorio de destino donde el script guardará los resultados procesados.
*   **`BaseDeDatos/`**: Directorio que debe contener el archivo Excel de referencia llamado `Base_0BS.xlsx`.

## Uso

1. Verifica que el archivo Excel de instituciones esté ubicado en `BaseDeDatos/Base_0BS.xlsx`.
2. Coloca los archivos PDF a analizar dentro de la carpeta `Entrada_PDF`.
3. Ejecuta el script principal desde tu terminal:

```bash
python Escaneo_PDF.py
```