#!/usr/bin/env python3
"""
Script para procesar automáticamente datos Excel según instrucciones en hoja TGT
Autor: Asistente IA
Descripción: Lee instrucciones de TGT y procesa datos de BASE GS1 (2)
"""

import pandas as pd
import numpy as np
from pathlib import Path
import sys
from datetime import datetime
import time
from collections import OrderedDict

def analizar_estructura_archivo(archivo_excel):
    """
    Analiza la estructura del archivo Excel y muestra información sobre las hojas
    """
    try:
        # Leer todas las hojas del archivo
        xl_file = pd.ExcelFile(archivo_excel)
        print(f"📊 Análisis del archivo: {archivo_excel}")
        print(f"🗂️  Hojas disponibles: {xl_file.sheet_names}")
        
        for hoja in xl_file.sheet_names:
            df = pd.read_excel(archivo_excel, sheet_name=hoja, header=None)
            print(f"\n📋 Hoja '{hoja}':")
            print(f"   - Dimensiones: {df.shape[0]} filas x {df.shape[1]} columnas")
            print(f"   - Primeras 3 filas:")
            print(df.head(3).to_string(index=False))
            print("-" * 50)
        
        return xl_file.sheet_names
        
    except Exception as e:
        print(f"❌ Error al analizar el archivo: {e}")
        return None

def leer_instrucciones_tgt(archivo_excel, hoja_tgt="TGT"):
    """
    Lee las instrucciones de la hoja TGT (primeras 4 filas)
    Retorna diccionario con las instrucciones parseadas
    """
    try:
        # Leer las primeras 4 filas de la hoja TGT
        df_instrucciones = pd.read_excel(archivo_excel, sheet_name=hoja_tgt, header=None, nrows=4)
        
        instrucciones = {}
        num_columnas = df_instrucciones.shape[1]
        
        for col_idx in range(num_columnas):
            # Extraer información de cada columna
            regla_llenado = df_instrucciones.iloc[0, col_idx] if not pd.isna(df_instrucciones.iloc[0, col_idx]) else ""
            generar_auxiliar = df_instrucciones.iloc[1, col_idx] if not pd.isna(df_instrucciones.iloc[1, col_idx]) else ""
            nombre_campo = df_instrucciones.iloc[2, col_idx] if not pd.isna(df_instrucciones.iloc[2, col_idx]) else ""
            descripcion = df_instrucciones.iloc[3, col_idx] if not pd.isna(df_instrucciones.iloc[3, col_idx]) else ""
            
            # PROCESAR TODAS LAS COLUMNAS, respetando el orden de la hoja TGT
            # Si la fila 1 está vacía, usar "no" como valor por defecto
            if generar_auxiliar == "":
                generar_auxiliar = "no"
                
            instrucciones[col_idx] = {
                'regla_llenado': str(regla_llenado),
                'generar_auxiliar': str(generar_auxiliar).lower() == 'si',
                'nombre_campo': str(nombre_campo),
                'descripcion': str(descripcion),
                'orden_original': col_idx  # Mantener el orden original
            }
        
        return instrucciones
        
    except Exception as e:
        print(f"❌ Error al leer instrucciones TGT: {e}")
        return None

def leer_datos_base(archivo_excel, hoja_base="BASE GS1 (2)", limitar_filas=None):
    """
    Lee los datos de la hoja base
    """
    try:
        # Leer la hoja base con encabezados
        if limitar_filas:
            print(f"🔢 Limitando procesamiento a las primeras {limitar_filas} filas...")
            df_base = pd.read_excel(archivo_excel, sheet_name=hoja_base, nrows=limitar_filas)
        else:
            df_base = pd.read_excel(archivo_excel, sheet_name=hoja_base)
            
        print(f"📊 Datos base cargados: {df_base.shape[0]} filas x {df_base.shape[1]} columnas")
        print(f"🏷️  Columnas disponibles: {list(df_base.columns)}")
        
        return df_base
        
    except Exception as e:
        print(f"❌ Error al leer datos base: {e}")
        return None

def mostrar_progreso(actual, total, prefijo="Progreso"):
    """
    Muestra una barra de progreso simple
    """
    porcentaje = (actual / total) * 100
    barra_longitud = 30
    progreso = int(barra_longitud * actual / total)
    barra = "█" * progreso + "░" * (barra_longitud - progreso)
    print(f"\r{prefijo}: [{barra}] {porcentaje:.1f}% ({actual}/{total})", end="", flush=True)

def procesar_columna(df_base, regla, nombre_campo, generar_auxiliar, num_filas):
    """
    Procesa una columna según la regla especificada
    """
    # TABLA FIJA DE GRUPO DE ARTÍCULOS - NO SE PUEDE MODIFICAR
    TABLA_GRUPO_ARTICULOS = {
        'ROPA INTERIOR': 100,
        'LOUNGEWEAR': 101,
        'ACCESORIOS': 102,
        'ACTIVE': 103,
        'CALCETINES': 104,
        'APPAREL': 105
    }
    
    if regla.lower() == "correlativo":
        # Generar números correlativos
        print(f"      ⚡ Generando {num_filas:,} números correlativos...")
        return list(range(1, num_filas + 1)), None
        
    elif regla.startswith("VALOR :"):
        # Extraer valor constante de dentro de las comillas
        if '"' in regla:
            # Buscar contenido entre comillas dobles
            inicio = regla.find('"') + 1
            fin = regla.rfind('"')
            valor_constante = regla[inicio:fin] if inicio <= fin else ""
        elif "'" in regla:
            # Buscar contenido entre comillas simples
            inicio = regla.find("'") + 1
            fin = regla.rfind("'")
            valor_constante = regla[inicio:fin] if inicio <= fin else ""
        else:
            # Si no hay comillas, tomar lo que está después de ":"
            valor_constante = regla.split(":")[1].strip()
        
        print(f"      ⚡ Llenando {num_filas:,} celdas con valor constante: '{valor_constante}'")
        return [valor_constante] * num_filas, None
        
    else:
        # La regla es el nombre de una columna en la base
        if regla in df_base.columns:
            print(f"      ⚡ Copiando datos de columna '{regla}'...")
            valores = df_base[regla].fillna("").tolist()
            
            if generar_auxiliar:
                print(f"      🔍 Buscando valores únicos en {num_filas:,} registros...")
                
                # CONDICIÓN ESPECIAL PARA GRUPO DE ARTÍCULOS - USAR TABLA FIJA
                if regla == "Grupo / Clase":
                    print(f"      🔒 Usando tabla fija para Grupo de Artículos...")
                    # Usar la tabla fija predefinida
                    valores_unicos = list(TABLA_GRUPO_ARTICULOS.keys())
                    codigos_tabla = list(TABLA_GRUPO_ARTICULOS.values())
                    
                    tabla_auxiliar = pd.DataFrame({
                        'ID': codigos_tabla,
                        'VALOR': valores_unicos
                    })
                    
                    print(f"      🗂️  Usando mapeo fijo de {len(valores_unicos)} grupos predefinidos...")
                    # Usar el mapeo fijo
                    mapeo = TABLA_GRUPO_ARTICULOS
                else:
                    # Crear tabla auxiliar con valores únicos (comportamiento normal)
                    valores_unicos = df_base[regla].dropna().unique()
                    tabla_auxiliar = pd.DataFrame({
                        'ID': range(1, len(valores_unicos) + 1),
                        'VALOR': valores_unicos
                    })
                    
                    print(f"      🗂️  Creando mapeo de {len(valores_unicos)} valores únicos...")
                    # Crear mapeo de valor a ID
                    mapeo = dict(zip(valores_unicos, range(1, len(valores_unicos) + 1)))
                
                print(f"      🔄 Reemplazando valores por IDs...")
                # Reemplazar valores por IDs (dejar vacío si no hay valor en lugar de 0)
                valores_con_id = [mapeo.get(val, "") if pd.notna(val) and val != "" else "" for val in df_base[regla]]
                
                return valores_con_id, tabla_auxiliar
            else:
                return valores, None
        else:
            print(f"      ⚠️  Columna '{regla}' no encontrada - llenando con valores vacíos")
            return [""] * num_filas, None

def generar_nombre_archivo_salida(archivo_entrada):
    """
    Genera un nombre único para el archivo de salida usando timestamp
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    archivo_base = Path(archivo_entrada).stem
    directorio = Path(archivo_entrada).parent
    
    return str(directorio / f"resultado_{archivo_base}_{timestamp}.xlsx")

def generar_tabla_tgt(archivo_excel, archivo_salida=None):
    """
    Función principal que genera la tabla TGT según las instrucciones
    """
    print("🚀 Iniciando procesamiento...")
    
    # Analizar estructura del archivo
    hojas = analizar_estructura_archivo(archivo_excel)
    if not hojas:
        return False
    
    # Verificar que existen las hojas necesarias
    if "TGT" not in hojas:
        print("❌ Error: No se encontró la hoja 'TGT'")
        return False
    
    if "BASE GS1 (2)" not in hojas:
        print("❌ Error: No se encontró la hoja 'BASE GS1 (2)'")
        return False
    
    # Leer instrucciones
    print("\n📋 Leyendo instrucciones de la hoja TGT...")
    instrucciones = leer_instrucciones_tgt(archivo_excel)
    if not instrucciones:
        return False
    
    print("📝 Instrucciones encontradas:")
    for col_idx, inst in instrucciones.items():
        print(f"   Columna {col_idx}: {inst['regla_llenado']} | Auxiliar: {inst['generar_auxiliar']} | Campo: {inst['nombre_campo']}")
    
    # Leer datos base (limitado a 10714 filas - exactamente los registros con datos)
    print("\n📊 Leyendo datos de la hoja base...")
    df_base = leer_datos_base(archivo_excel, limitar_filas=10714)
    if df_base is None:
        return False
    
    # Procesar cada columna según las instrucciones EN ORDEN
    print(f"\n⚙️  Procesando {len(instrucciones)} columnas...")
    tabla_final = OrderedDict()
    tablas_auxiliares = {}
    num_filas = len(df_base)
    
    total_columnas = len(instrucciones)
    columna_actual = 0
    
    # PROCESAR EN EL ORDEN CORRECTO (por índice de columna)
    for col_idx in sorted(instrucciones.keys()):
        inst = instrucciones[col_idx]
        columna_actual += 1
        print(f"\n   📊 [{columna_actual}/{total_columnas}] Procesando columna {col_idx}: {inst['nombre_campo']}")
        mostrar_progreso(columna_actual - 1, total_columnas, "Progreso general")
        print()  # Nueva línea después de la barra de progreso
        
        inicio_tiempo = time.time()
        valores, tabla_aux = procesar_columna(
            df_base, 
            inst['regla_llenado'], 
            inst['nombre_campo'], 
            inst['generar_auxiliar'], 
            num_filas
        )
        tiempo_transcurrido = time.time() - inicio_tiempo
        
        tabla_final[inst['nombre_campo']] = valores
        
        if tabla_aux is not None:
            tablas_auxiliares[inst['descripcion']] = tabla_aux
            print(f"     ✅ Tabla auxiliar generada: {inst['descripcion']} ({len(tabla_aux)} valores únicos)")
        
        print(f"     ⏱️  Tiempo: {tiempo_transcurrido:.2f} segundos")
    
    # Mostrar progreso final
    print()
    mostrar_progreso(total_columnas, total_columnas, "Progreso general")
    print("\n")
    
    # Crear DataFrame final con encabezados de TGT
    print("🔧 Construyendo DataFrame final...")
    
    # Preparar nombres de campos (fila 3 de TGT) como encabezados
    nombres_campos_ordenados = []
    descripciones_ordenadas = []
    for col_idx in sorted(instrucciones.keys()):
        nombres_campos_ordenados.append(instrucciones[col_idx]['nombre_campo'])
        descripciones_ordenadas.append(instrucciones[col_idx]['descripcion'])
    
    # Crear DataFrame con los nombres de campos como columnas
    df_final = pd.DataFrame(tabla_final, columns=nombres_campos_ordenados)
    
    # Agregar fila de descripciones (fila 4 de TGT) como segunda fila
    print("📝 Agregando encabezados: fila 3 y 4 de TGT...")
    fila_descripciones = pd.DataFrame([descripciones_ordenadas], columns=df_final.columns)
    
    # Insertar la fila de descripciones como segunda fila (índice 0)
    df_final = pd.concat([fila_descripciones, df_final], ignore_index=True)
    
    # Generar archivo de salida con nombre único
    if archivo_salida is None:
        archivo_salida = generar_nombre_archivo_salida(archivo_excel)
    
    print(f"\n💾 Guardando resultados en: {archivo_salida}")
    print("📝 Esto puede tomar varios minutos para archivos grandes...")
    
    inicio_guardado = time.time()
    
    with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
        # Escribir tabla principal
        print("   📊 Escribiendo hoja principal TGT_FINAL...")
        df_final.to_excel(writer, sheet_name='TGT_FINAL', index=False)
        print(f"   ✅ Hoja 'TGT_FINAL' creada ({df_final.shape[0]:,} filas x {df_final.shape[1]} columnas)")
        print(f"      📝 Fila 1: Nombres de campos (TGT fila 3)")
        print(f"      📝 Fila 2: Descripciones de campos (TGT fila 4)")
        print(f"      📝 Filas 3-{df_final.shape[0]}: Datos procesados ({df_final.shape[0]-2:,} registros)")
        
        # Escribir tablas auxiliares
        total_auxiliares = len(tablas_auxiliares)
        auxiliar_actual = 0
        
        for nombre_hoja, tabla_aux in tablas_auxiliares.items():
            auxiliar_actual += 1
            # Truncar nombre de hoja si es muy largo (Excel tiene límite de 31 caracteres)
            nombre_hoja_corto = nombre_hoja[:31] if len(nombre_hoja) > 31 else nombre_hoja
            
            print(f"   📋 [{auxiliar_actual}/{total_auxiliares}] Escribiendo hoja auxiliar '{nombre_hoja_corto}'...")
            tabla_aux.to_excel(writer, sheet_name=nombre_hoja_corto, index=False)
            print(f"   ✅ Hoja auxiliar '{nombre_hoja_corto}' creada ({tabla_aux.shape[0]} valores únicos)")
            
            mostrar_progreso(auxiliar_actual, total_auxiliares, "Guardando auxiliares")
        
        print()  # Nueva línea después de la barra de progreso final
    
    tiempo_guardado = time.time() - inicio_guardado
    print(f"   ⏱️  Tiempo de guardado: {tiempo_guardado:.2f} segundos")
    
    print(f"\n🎉 ¡Procesamiento completado exitosamente!")
    print(f"📁 Archivo generado: {archivo_salida}")
    
    return True

def main():
    """
    Función principal
    """
    archivo_entrada = "/Users/christianmatthews/Library/Mobile Documents/com~apple~CloudDocs/CURSOR/TOP/MAESTRO PRODUCTO/Test1.xlsx"
    
    if not Path(archivo_entrada).exists():
        print(f"❌ Error: El archivo {archivo_entrada} no existe")
        return
    
    print("=" * 60)
    print("🔧 PROCESADOR AUTOMÁTICO DE DATOS EXCEL")
    print("=" * 60)
    
    exito = generar_tabla_tgt(archivo_entrada)
    
    if exito:
        print("\n✅ Proceso completado con éxito")
    else:
        print("\n❌ El proceso falló")

if __name__ == "__main__":
    main()