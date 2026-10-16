import os
import sys
import fitz 
import re
import pandas as pd
from difflib import SequenceMatcher
import warnings

# Ignorar advertencias de pandas
warnings.filterwarnings("ignore")

# ==========================================
# 1. CONFIGURACIÓN DE RUTAS
# ==========================================
if getattr(sys, 'frozen', False):
    ruta_base = os.path.dirname(sys.executable)
else:
    ruta_base = os.path.dirname(os.path.abspath(__file__))

CARPETA_ENTRADA_PDF = os.path.join(ruta_base, 'Entrada_PDFs')
CARPETA_SALIDA_PDF = os.path.join(ruta_base, 'Salida_PDFs')
RUTA_DB = os.path.join(ruta_base, 'BasedeDatos', 'Base_OBS.xlsx')

# ==========================================
# 2. CARGAR DICCIONARIO DESDE EXCEL
# ==========================================
def cargar_instituciones_db():
    print("-" * 50)
    print(" CARGANDO BASE DE DATOS DE INSTITUCIONES...")
    
    if not os.path.exists(RUTA_DB):
        print(f"❌ Error: No se encuentra el archivo en: {RUTA_DB}")
        return None

    try:
        # Forzamos todo a string para evitar que los números pierdan ceros o guiones
        df_db = pd.read_excel(RUTA_DB, engine='openpyxl', dtype=str)
        df_db = df_db.fillna('')
        
        instituciones = []
        for _, row in df_db.iterrows():
            instituciones.append({
                'documento': str(row.get('Numero de documento', '')).strip().upper(),
                'codigo_cliente': str(row.get('CodigoCliente', '')).strip().upper(),
                'razon_social': str(row.get('Razon Social', '')).strip().upper(),
                'abreviado': str(row.get('ABREVIADOS', 'DESC')).strip().upper() 
            })
            
        print(f"✅ Se cargaron {len(instituciones)} instituciones exitosamente.")
        print("-" * 50)
        return instituciones

    except Exception as e:
        print(f"❌ Error leyendo la base de datos: {e}")
        return None

# ==========================================
# 3. LÓGICA PRINCIPAL (PROCESAR PDFS)
# ==========================================
def procesar_facturas_pdf():
    # 1. Preparar carpetas con manejo de errores
    try:
        if not os.path.exists(CARPETA_ENTRADA_PDF): 
            os.makedirs(CARPETA_ENTRADA_PDF)
        if not os.path.exists(CARPETA_SALIDA_PDF): 
            os.makedirs(CARPETA_SALIDA_PDF)
    except Exception as e:
        print(f"❌ Error del sistema al intentar acceder a las carpetas: {e}")
        return

    # 2. Cargar la DB
    instituciones_db = cargar_instituciones_db()
    if not instituciones_db:
        return

    # --- OPTIMIZACIÓN DE VELOCIDAD Y FILTROS ---
    REGEX_DESCRIPCION = re.compile(r"REFACTURACION\s+POR\s+INCREMENTO\s+DE\s+VALORES\s+SOBRE\s+FC")
    
    REGEX_NUMERO = re.compile(r"\b(\d{4,5}-\d{8})\b")
    
    REGEX_LETRA = re.compile(r"\n\s*(A|B|C)\s*\n")

    try:
        archivos_pdf = [f for f in os.listdir(CARPETA_ENTRADA_PDF) if f.lower().endswith('.pdf')]
    except FileNotFoundError:
        print(f"❌ Error: No se pudo leer la carpeta {CARPETA_ENTRADA_PDF}.")
        return
    
    print(f"\n--- Se encontraron {len(archivos_pdf)} archivos PDF en '{CARPETA_ENTRADA_PDF}' ---")

    for archivo in archivos_pdf:
        print(f"\nProcesando documento: {archivo}")
        ruta_pdf_origen = os.path.join(CARPETA_ENTRADA_PDF, archivo)

        try:
            doc = fitz.open(ruta_pdf_origen)
            
            for num_pag in range(len(doc)):
                pagina = doc.load_page(num_pag)
                texto_pagina = pagina.get_text().upper()   # type: ignore
                
                # --- FILTRO POR DESCRIPCIÓN ---
                if not REGEX_DESCRIPCION.search(texto_pagina):
                    print(f"   -> ⏭️ Pág {num_pag + 1}: Ignorada (no es refacturación).")
                    continue 
                
                # --- A. BUSCAR NÚMERO DE FACTURA Y LETRA ---
                match_numero = REGEX_NUMERO.search(texto_pagina)
                
                if not match_numero:
                    print(f"   -> ⚠️ Pág {num_pag + 1}: No se encontró número de factura válido con 'N°'.")
                    continue
                
                numero_base = match_numero.group(1)
                letra_factura = "X" 
                match_letra = REGEX_LETRA.search(texto_pagina[:800])
                
                if match_letra:
                    letra_factura = match_letra.group(1)
                else:
                    if "01" in texto_pagina[:800] and "COD" in texto_pagina[:800]:
                        letra_factura = "A"
                    elif "06" in texto_pagina[:800] or "\n06\n" in texto_pagina[:800]:
                        letra_factura = "B"
                    elif "11" in texto_pagina[:800] or "\n11\n" in texto_pagina[:800]:
                        letra_factura = "C"
                    
                nro_factura = f"FC-{letra_factura}-{numero_base}"

                # --- B. BUSCAR INSTITUCIÓN EN EL TEXTO ---
                abreviatura_encontrada = "DESC" 
                texto_sin_formatos = texto_pagina.replace('-', '').replace(' ', '')
                lineas_pdf = texto_pagina.split('\n')

                for inst in instituciones_db:
                    cuit_db = inst['documento']
                    cod_cliente_db = inst['codigo_cliente']
                    razon = inst['razon_social']
                    
                    cuit_encontrado = False
                    razon_encontrada = False
                    codigo_encontrado = False

                    # 1. Filtro fuerte: Código de Cliente
                    if cod_cliente_db and len(cod_cliente_db) > 3 and cod_cliente_db in texto_pagina:
                        codigo_encontrado = True

                    # 2. Filtro de CUIT
                    if cuit_db:
                        doc_limpio = cuit_db.replace('-', '').replace(' ', '')
                        if len(doc_limpio) > 5 and doc_limpio in texto_sin_formatos:
                            cuit_encontrado = True

                    # 3. Filtro de Razón Social
                    if razon:
                        for linea in lineas_pdf:
                            linea_limpia = linea.strip()
                            if len(linea_limpia) > 5: 
                                similitud = SequenceMatcher(None, razon, linea_limpia).ratio()
                                if similitud >= 0.80:
                                    razon_encontrada = True
                                    break 

                    # Lógica estricta de validación
                    if codigo_encontrado or (cuit_encontrado and razon_encontrada):
                        abreviatura_encontrada = inst['abreviado']
                        break 

                # --- C. CREAR CARPETA Y GUARDAR EL NUEVO PDF ---
                abreviatura_segura = re.sub(r'[\\/*?:"<>|]', '_', abreviatura_encontrada)
                
                carpeta_institucion = os.path.join(CARPETA_SALIDA_PDF, abreviatura_segura)
                if not os.path.exists(carpeta_institucion):
                    os.makedirs(carpeta_institucion)

                nombre_nuevo_pdf = f"{nro_factura}.pdf"
                ruta_pdf_salida = os.path.join(carpeta_institucion, nombre_nuevo_pdf)

                nuevo_doc = fitz.open()
                nuevo_doc.insert_pdf(doc, from_page=num_pag, to_page=num_pag)
                nuevo_doc.save(ruta_pdf_salida)
                nuevo_doc.close()

                print(f"   -> ✅ Pág {num_pag + 1} guardada en '{abreviatura_segura}' como: {nombre_nuevo_pdf}")

            doc.close()

        except Exception as e:
            print(f"❌ Error procesando el PDF {archivo}: {e}")

    print("\n" + "="*50)
    print(" ¡Proceso Finalizado! ")
    print("="*50)

if __name__ == "__main__":
    os.system('cls' if os.name == 'nt' else 'clear') 
    procesar_facturas_pdf()
