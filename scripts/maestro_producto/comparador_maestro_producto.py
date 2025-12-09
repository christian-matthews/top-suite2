#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Comparador de Maestro Producto (Test1 vs Copia de FINAL)
Convierte ambos archivos a formato SAP y compara para detectar:
- Líneas nuevas (SKU_HIJO que están en FINAL pero no en Test1)
- Líneas modificadas (SKU_HIJO con atributos diferentes)
"""

import pandas as pd
import numpy as np
import os
import sys
import shutil
from datetime import datetime
from pathlib import Path
from typing import Tuple, Dict, List, Optional
from collections import OrderedDict
import time

# Importar funciones del procesador_excel
from procesador_excel import (
    leer_instrucciones_tgt,
    leer_datos_base,
    procesar_columna,
    generar_tabla_tgt
)


def normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normaliza los nombres de columnas: elimina espacios y estandariza.
    """
    df = df.copy()
    # Normalizar SKU_HIJO (puede venir como "SKU HIJO" o "SKU_HIJO")
    if "SKU HIJO" in df.columns:
        df.rename(columns={"SKU HIJO": "SKU_HIJO"}, inplace=True)
    df.columns = df.columns.str.strip()
    return df


def convertir_a_sap(archivo_excel: str, output_dir: str, nombre_salida: str) -> Optional[str]:
    """
    Convierte un archivo de maestro producto a formato SAP.
    
    Args:
        archivo_excel: Ruta al archivo Excel
        output_dir: Directorio donde guardar el resultado
        nombre_salida: Nombre del archivo de salida
        
    Returns:
        Ruta del archivo generado o None si hay error
    """
    try:
        print(f"\n🔄 Convirtiendo {os.path.basename(archivo_excel)} a formato SAP...")
        
        # Verificar si tiene hoja TGT (como Test1)
        xl_file = pd.ExcelFile(archivo_excel)
        
        if "TGT" in xl_file.sheet_names and "BASE GS1 (2)" in xl_file.sheet_names:
            # Usar el procesador_excel directamente
            print("   ✓ Archivo tiene estructura TGT, usando procesador_excel...")
            archivo_salida = os.path.join(output_dir, nombre_salida)
            
            # Leer instrucciones TGT
            instrucciones = leer_instrucciones_tgt(archivo_excel)
            if not instrucciones:
                return None
            
            # Leer datos base
            df_base = leer_datos_base(archivo_excel)
            if df_base is None:
                return None
            
            # Procesar según instrucciones
            tabla_final = OrderedDict()
            tablas_auxiliares = {}
            num_filas = len(df_base)
            
            for col_idx in sorted(instrucciones.keys()):
                inst = instrucciones[col_idx]
                valores, tabla_aux = procesar_columna(
                    df_base,
                    inst['regla_llenado'],
                    inst['nombre_campo'],
                    inst['generar_auxiliar'],
                    num_filas
                )
                tabla_final[inst['nombre_campo']] = valores
                if tabla_aux is not None:
                    tablas_auxiliares[inst['descripcion']] = tabla_aux
            
            # Crear DataFrame final
            nombres_campos_ordenados = [instrucciones[col_idx]['nombre_campo'] 
                                       for col_idx in sorted(instrucciones.keys())]
            df_final = pd.DataFrame(tabla_final, columns=nombres_campos_ordenados)
            
            # Guardar
            with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
                df_final.to_excel(writer, sheet_name='TGT_FINAL', index=False)
                for nombre_hoja, tabla_aux in tablas_auxiliares.items():
                    nombre_hoja_corto = nombre_hoja[:31] if len(nombre_hoja) > 31 else nombre_hoja
                    tabla_aux.to_excel(writer, sheet_name=nombre_hoja_corto, index=False)
            
            print(f"   ✓ Conversión completada: {nombre_salida}")
            return archivo_salida
            
        else:
            # Archivo sin estructura TGT (como Copia de FINAL)
            # Necesitamos usar las instrucciones de Test1 para convertir
            print("   ⚠ Archivo sin estructura TGT, necesitamos usar instrucciones de referencia...")
            return None
            
    except Exception as e:
        print(f"   ✗ Error al convertir: {str(e)}")
        import traceback
        traceback.print_exc()
        return None


def convertir_final_con_instrucciones_test1(
    archivo_final: str,
    archivo_test1: str,
    output_dir: str,
    nombre_salida: str
) -> Optional[str]:
    """
    Convierte Copia de FINAL a formato SAP usando las instrucciones de Test1.
    """
    try:
        print(f"\n🔄 Convirtiendo {os.path.basename(archivo_final)} usando instrucciones de Test1...")
        
        # Leer instrucciones de Test1
        instrucciones = leer_instrucciones_tgt(archivo_test1)
        if not instrucciones:
            print("   ✗ No se pudieron leer instrucciones de Test1")
            return None
        
        # Leer datos de FINAL (hoja Hoja1)
        df_base = pd.read_excel(archivo_final, sheet_name="Hoja1")
        df_base = normalizar_columnas(df_base)
        print(f"   ✓ Datos cargados: {len(df_base)} filas")
        
        # Procesar según instrucciones de Test1
        tabla_final = OrderedDict()
        tablas_auxiliares = {}
        num_filas = len(df_base)
        
        for col_idx in sorted(instrucciones.keys()):
            inst = instrucciones[col_idx]
            regla = inst['regla_llenado']
            
            # Si la regla es un nombre de columna, verificar si existe en FINAL
            if regla not in ["correlativo"] and not regla.startswith("VALOR :"):
                # Buscar columna equivalente (puede tener nombre ligeramente diferente)
                columna_encontrada = None
                for col in df_base.columns:
                    if col.strip().upper() == regla.strip().upper():
                        columna_encontrada = col
                        break
                
                if columna_encontrada:
                    regla = columna_encontrada
                else:
                    print(f"   ⚠ Columna '{regla}' no encontrada en FINAL, usando vacío")
            
            valores, tabla_aux = procesar_columna(
                df_base,
                regla,
                inst['nombre_campo'],
                inst['generar_auxiliar'],
                num_filas
            )
            tabla_final[inst['nombre_campo']] = valores
            if tabla_aux is not None:
                tablas_auxiliares[inst['descripcion']] = tabla_aux
        
        # Crear DataFrame final
        nombres_campos_ordenados = [instrucciones[col_idx]['nombre_campo'] 
                                   for col_idx in sorted(instrucciones.keys())]
        df_final = pd.DataFrame(tabla_final, columns=nombres_campos_ordenados)
        
        # Guardar
        archivo_salida = os.path.join(output_dir, nombre_salida)
        with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='TGT_FINAL', index=False)
            for nombre_hoja, tabla_aux in tablas_auxiliares.items():
                nombre_hoja_corto = nombre_hoja[:31] if len(nombre_hoja) > 31 else nombre_hoja
                tabla_aux.to_excel(writer, sheet_name=nombre_hoja_corto, index=False)
        
        print(f"   ✓ Conversión completada: {nombre_salida}")
        return archivo_salida
        
    except Exception as e:
        print(f"   ✗ Error al convertir: {str(e)}")
        import traceback
        traceback.print_exc()
        return None


def normalizar_sku_hijo(value) -> str:
    """
    Normaliza SKU_HIJO para comparación consistente.
    """
    if pd.isna(value):
        return ""
    
    if isinstance(value, (int, float)):
        return str(int(value))
    
    str_value = str(value).strip()
    if str_value.endswith('.0'):
        str_value = str_value[:-2]
    
    return str_value


def normalizar_valor_comparacion(value) -> str:
    """
    Normaliza un valor para comparación.
    """
    if pd.isna(value):
        return ""
    
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)
    
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    
    return str(value).strip().lower()


def formatear_valor_display(value) -> str:
    """
    Formatea un valor para mostrar en el reporte.
    """
    if pd.isna(value):
        return "(vacío)"
    
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return f"{value:.10f}".rstrip('0').rstrip('.')
    
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    
    return str(value)


def comparar_archivos_maestro_producto(
    archivo_test1: str,
    archivo_final: str,
    key_column: str = "SKU_HIJO"
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Compara dos archivos en formato maestro producto y detecta NUEVOS y MODIFICADOS.
    
    Returns:
        Tupla (df_nuevos, df_modificados)
    """
    print(f"\n🔍 Comparando archivos en formato maestro producto...")
    
    # Cargar archivos
    # Test1 tiene hoja "BASE GS1 (2)"
    try:
        df_test1 = pd.read_excel(archivo_test1, sheet_name="BASE GS1 (2)")
    except:
        # Intentar con primera hoja
        df_test1 = pd.read_excel(archivo_test1, sheet_name=0)
    
    # FINAL tiene hoja "Hoja1"
    try:
        df_final = pd.read_excel(archivo_final, sheet_name="Hoja1")
    except:
        df_final = pd.read_excel(archivo_final, sheet_name=0)
    
    # Normalizar columnas
    df_test1 = normalizar_columnas(df_test1)
    df_final = normalizar_columnas(df_final)
    
    # Buscar columna clave (SKU_HIJO o SKU HIJO)
    key_column_found = None
    if key_column in df_test1.columns:
        key_column_found = key_column
    else:
        # Buscar variantes
        posibles_claves = [col for col in df_test1.columns 
                          if "sku" in col.lower() and "hijo" in col.lower()]
        if posibles_claves:
            key_column_found = posibles_claves[0]
    
    if not key_column_found:
        print("   ✗ No se encontró columna SKU_HIJO para comparación")
        return pd.DataFrame(), pd.DataFrame()
    
    key_column = key_column_found
    print(f"   ✓ Usando columna clave: {key_column}")
    
    # Filtrar registros válidos (no NaN en columna clave)
    df_test1_validos = df_test1[df_test1[key_column].notna()].copy()
    df_final_validos = df_final[df_final[key_column].notna()].copy()
    
    print(f"   ✓ Test1: {len(df_test1_validos)} registros válidos")
    print(f"   ✓ FINAL: {len(df_final_validos)} registros válidos")
    
    # Normalizar columna clave para comparación
    df_test1_validos[key_column + "_norm"] = df_test1_validos[key_column].apply(normalizar_sku_hijo)
    df_final_validos[key_column + "_norm"] = df_final_validos[key_column].apply(normalizar_sku_hijo)
    
    # Crear sets para comparación
    sku_test1 = set(df_test1_validos[key_column + "_norm"])
    sku_final = set(df_final_validos[key_column + "_norm"])
    
    # NUEVOS: están en FINAL pero no en Test1
    nuevos_sku = sku_final - sku_test1
    df_nuevos = df_final_validos[df_final_validos[key_column + "_norm"].isin(nuevos_sku)].copy()
    if key_column + "_norm" in df_nuevos.columns:
        df_nuevos = df_nuevos.drop(columns=[key_column + "_norm"])
    
    print(f"   ✓ NUEVOS encontrados: {len(df_nuevos)}")
    
    # MODIFICADOS: están en ambos pero con diferencias
    comunes_sku = sku_test1 & sku_final
    modificados_list = []
    
    print(f"   🔍 Analizando {len(comunes_sku)} registros comunes para detectar modificaciones...")
    
    for sku_norm in comunes_sku:
        row_test1 = df_test1_validos[df_test1_validos[key_column + "_norm"] == sku_norm].iloc[0]
        row_final = df_final_validos[df_final_validos[key_column + "_norm"] == sku_norm].iloc[0]
        
        sku_original = row_final[key_column]
        
        # Comparar todas las columnas excepto la clave y la temporal
        columns_to_compare = [col for col in df_test1_validos.columns 
                             if col != key_column and col != key_column + "_norm"]
        
        for col in columns_to_compare:
            if col not in df_final_validos.columns:
                continue
            
            val_test1 = row_test1[col]
            val_final = row_final[col]
            
            norm_test1 = normalizar_valor_comparacion(val_test1)
            norm_final = normalizar_valor_comparacion(val_final)
            
            if norm_test1 != norm_final:
                tipo_cambio = "texto"
                if isinstance(val_test1, (int, float)) or isinstance(val_final, (int, float)):
                    tipo_cambio = "numérico"
                elif isinstance(val_test1, datetime) or isinstance(val_final, datetime):
                    tipo_cambio = "fecha"
                
                modificados_list.append({
                    key_column: sku_original,
                    "COLUMNA": col,
                    "VALOR_TEST1": formatear_valor_display(val_test1),
                    "VALOR_FINAL": formatear_valor_display(val_final),
                    "ESTADO": "MODIFICADO",
                    "TIPO_CAMBIO": tipo_cambio
                })
    
    df_modificados = pd.DataFrame(modificados_list)
    print(f"   ✓ MODIFICADOS encontrados: {df_modificados[key_column].nunique() if not df_modificados.empty else 0} SKU únicos, {len(df_modificados)} cambios")
    
    return df_nuevos, df_modificados


def generar_reporte_excel(
    df_nuevos: pd.DataFrame,
    df_modificados: pd.DataFrame,
    timestamp: str,
    output_dir: str
) -> str:
    """
    Genera archivo Excel con hojas NUEVOS y MODIFICADOS.
    """
    filename = f"COMPARACION_MAESTRO_PRODUCTO_{timestamp}.xlsx"
    filepath = os.path.join(output_dir, filename)
    
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        # Hoja NUEVOS
        if not df_nuevos.empty:
            df_nuevos.to_excel(writer, sheet_name="NUEVOS", index=False)
        else:
            pd.DataFrame(columns=df_nuevos.columns if not df_nuevos.empty else ["SKU_HIJO"]).to_excel(
                writer, sheet_name="NUEVOS", index=False
            )
        
        # Hoja MODIFICADOS
        if not df_modificados.empty:
            df_modificados.to_excel(writer, sheet_name="MODIFICADOS", index=False)
        else:
            pd.DataFrame(columns=["SKU_HIJO", "COLUMNA", "VALOR_TEST1", "VALOR_FINAL", "ESTADO", "TIPO_CAMBIO"]).to_excel(
                writer, sheet_name="MODIFICADOS", index=False
            )
    
    return filepath


def crear_carpeta_timestamp(base_dir: str) -> str:
    """
    Crea una carpeta con timestamp para guardar los resultados.
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    carpeta = os.path.join(base_dir, f"PROCESAMIENTO_{timestamp}")
    os.makedirs(carpeta, exist_ok=True)
    return carpeta, timestamp


def main():
    """
    Función principal.
    """
    print("=" * 80)
    print("COMPARADOR DE MAESTRO PRODUCTO (Test1 vs Copia de FINAL)")
    print("=" * 80)
    
    # Directorio base
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Rutas de archivos
    archivo_test1 = os.path.join(base_dir, "Test1.xlsx")
    archivo_final = os.path.join(base_dir, "Copia de FINAL.xlsx")
    archivo_base_sap = os.path.join(base_dir, "resultado_Test1_20250806_232635.xlsx")
    
    # Verificar que existen los archivos
    if not os.path.exists(archivo_test1):
        print(f"✗ Error: No se encuentra {archivo_test1}")
        sys.exit(1)
    
    if not os.path.exists(archivo_final):
        print(f"✗ Error: No se encuentra {archivo_final}")
        sys.exit(1)
    
    # Crear carpeta con timestamp
    carpeta_procesamiento, timestamp = crear_carpeta_timestamp(base_dir)
    print(f"\n📁 Carpeta de procesamiento: {carpeta_procesamiento}")
    
    try:
        # 1. Comparar archivos directamente en formato maestro producto
        print("\n" + "=" * 80)
        print("PASO 1: Comparar archivos en formato maestro producto")
        print("=" * 80)
        
        df_nuevos, df_modificados = comparar_archivos_maestro_producto(
            archivo_test1,
            archivo_final
        )
        
        # 2. Generar reporte Excel
        print("\n" + "=" * 80)
        print("PASO 2: Generar reporte Excel")
        print("=" * 80)
        
        archivo_reporte = generar_reporte_excel(
            df_nuevos,
            df_modificados,
            timestamp,
            carpeta_procesamiento
        )
        print(f"✓ Reporte generado: {os.path.basename(archivo_reporte)}")
        
        # 3. Copiar archivos originales a la carpeta
        print("\n" + "=" * 80)
        print("PASO 3: Guardar archivos procesados")
        print("=" * 80)
        
        shutil.copy2(archivo_test1, os.path.join(carpeta_procesamiento, "Test1.xlsx"))
        shutil.copy2(archivo_final, os.path.join(carpeta_procesamiento, "Copia de FINAL.xlsx"))
        print("✓ Archivos originales copiados")
        
        # Resumen final
        print("\n" + "=" * 80)
        print("RESUMEN FINAL")
        print("=" * 80)
        print(f"📁 Carpeta de resultados: {carpeta_procesamiento}")
        print(f"📊 Registros NUEVOS: {len(df_nuevos)}")
        print(f"📊 Registros MODIFICADOS: {df_modificados['SKU_HIJO'].nunique() if not df_modificados.empty else 0} SKU únicos")
        print(f"📊 Total de cambios detectados: {len(df_modificados)}")
        print(f"📄 Archivo de reporte: {os.path.basename(archivo_reporte)}")
        print("=" * 80)
        print("✓ Proceso completado exitosamente")
        
    except Exception as e:
        print(f"\n✗ Error inesperado: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

