#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PROMPT2 - Procesamiento MP KEY vs VENTAS
Versión final lista para producción + modo testing incluido
"""

import pandas as pd
import sys
import os
import shutil
from pathlib import Path
from typing import Tuple, Dict, Optional
from datetime import datetime

# Configuración
MODO_TESTING = False  # Cambiar a True para modo testing
ARCHIVO_MP_KEY = "MP KEY.xlsx"
ARCHIVO_VENTAS = "Ventas JUL-AGO-SEP-OCT.xlsx"
ARCHIVO_SALIDA_CARGAR = "TX_Carga.xlsx"
ARCHIVO_SALIDA_PENDIENTES = "TX_Pendientes.xlsx"


def limpiar_carpeta_origen():
    """Mueve archivos de procesos anteriores a carpetas organizadas"""
    print("\n🧹 Limpiando carpeta de origen...")
    
    # Archivos que deben quedarse (origen)
    archivos_origen = [ARCHIVO_MP_KEY, ARCHIVO_VENTAS]
    
    # Patrones de archivos a mover
    patrones_procesos = [
        'Proceso_*',
        'TX_Carga_*.xlsx',
        'TX_Pendientes_*.xlsx',
        'transacciones por cargar F1_*.xlsx',
        'transacciones pendientes F1_*.xlsx'
    ]
    
    archivos_movidos = 0
    
    # Buscar y mover archivos de procesos anteriores
    for patron in patrones_procesos:
        archivos = list(Path('.').glob(patron))
        
        for archivo in archivos:
            # No mover archivos de origen
            if archivo.name in archivos_origen:
                continue
            
            # Si es una carpeta Proceso_, moverla a Archivados/
            if archivo.is_dir() and archivo.name.startswith('Proceso_'):
                carpeta_archivo = Path('Archivados')
                carpeta_archivo.mkdir(exist_ok=True)
                destino = carpeta_archivo / archivo.name
                
                if not destino.exists():
                    try:
                        shutil.move(str(archivo), str(destino))
                        print(f"   ✅ Carpeta {archivo.name} movida a Archivados/")
                        archivos_movidos += 1
                    except Exception as e:
                        print(f"   ⚠️  Error al mover {archivo.name}: {e}")
            
            # Si es un archivo de resultado, moverlo a Archivados/Resultados/
            elif archivo.is_file():
                carpeta_resultados = Path('Archivados') / 'Resultados'
                carpeta_resultados.mkdir(parents=True, exist_ok=True)
                destino = carpeta_resultados / archivo.name
                
                if not destino.exists():
                    try:
                        shutil.move(str(archivo), str(destino))
                        print(f"   ✅ {archivo.name} movido a Archivados/Resultados/")
                        archivos_movidos += 1
                    except Exception as e:
                        print(f"   ⚠️  Error al mover {archivo.name}: {e}")
    
    if archivos_movidos > 0:
        print(f"   📦 {archivos_movidos} archivos/carpetas movidos a Archivados/")
    else:
        print("   ✅ Carpeta de origen ya está limpia")


def crear_carpeta_proceso():
    """Crea una carpeta para el proceso con timestamp"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    carpeta = f"Proceso_{timestamp}"
    
    # Crear carpeta si no existe
    Path(carpeta).mkdir(exist_ok=True)
    
    return carpeta, timestamp


def copiar_archivos_origen(carpeta: str):
    """Copia los archivos de origen a la carpeta del proceso"""
    print(f"\n📁 Copiando archivos de origen a {carpeta}/...")
    
    archivos_origen = [ARCHIVO_MP_KEY, ARCHIVO_VENTAS]
    
    for archivo in archivos_origen:
        if os.path.exists(archivo):
            try:
                destino = os.path.join(carpeta, archivo)
                shutil.copy2(archivo, destino)
                print(f"   ✅ {archivo} copiado")
            except Exception as e:
                print(f"   ⚠️  Error al copiar {archivo}: {e}")
        else:
            print(f"   ⚠️  {archivo} no encontrado")


def generar_nombres_con_timestamp(carpeta: str, timestamp: str):
    """Genera nombres de archivo con timestamp dentro de la carpeta"""
    base_cargar = ARCHIVO_SALIDA_CARGAR.replace('.xlsx', '')
    base_pendientes = ARCHIVO_SALIDA_PENDIENTES.replace('.xlsx', '')
    
    archivo_cargar = os.path.join(carpeta, f"{base_cargar}_{timestamp}.xlsx")
    archivo_pendientes = os.path.join(carpeta, f"{base_pendientes}_{timestamp}.xlsx")
    
    return archivo_cargar, archivo_pendientes


def detectar_columna_clave_p(df: pd.DataFrame, posibles_nombres: list) -> Optional[str]:
    """Detecta automáticamente la columna CLAVE_P"""
    df_cols_upper = [str(col).upper().strip() for col in df.columns]
    
    # Priorizar KEY_MS sobre otras columnas
    if 'KEY_MS' in df.columns:
        return 'KEY_MS'
    
    # Buscar coincidencias exactas primero
    for nombre in posibles_nombres:
        nombre_upper = nombre.upper().strip()
        if nombre_upper in df_cols_upper:
            idx = df_cols_upper.index(nombre_upper)
            return df.columns[idx]
    
    # Buscar coincidencias parciales (pero más estrictas)
    for nombre in posibles_nombres:
        nombre_upper = nombre.upper().strip()
        for idx, col_upper in enumerate(df_cols_upper):
            # Coincidencia exacta o que el nombre esté completo en la columna
            if nombre_upper == col_upper or (nombre_upper in col_upper and len(nombre_upper) >= 3):
                # Evitar KEY_ECLOUD si hay otras opciones
                if 'ECLOUD' not in col_upper or len(posibles_nombres) == 1:
                    return df.columns[idx]
    
    # Búsqueda más flexible (evitar KEY_ECLOUD y columnas de una sola letra)
    for col in df.columns:
        col_str = str(col).upper().strip()
        if len(col_str) > 1:  # Evitar columnas de una sola letra
            if any(keyword in col_str for keyword in ['CLAVE', 'KEY']):
                if 'ECLOUD' not in col_str:
                    if 'P' in col_str or 'PRODUCTO' in col_str or 'MS' in col_str or col_str == 'KEY':
                        return col
    
    return None


def detectar_columna_no_sap(df: pd.DataFrame) -> Optional[str]:
    """Detecta automáticamente la columna NO_SAP"""
    df_cols_upper = [str(col).upper().strip() for col in df.columns]
    
    # Buscar primero SKU_HIJO que es común en nuevos formatos
    if 'SKU_HIJO' in df.columns:
        return 'SKU_HIJO'
    
    posibles = ['NO_SAP', 'NUMERO DE ARTICULO', 'NUMERO ARTICULO', 'NÚMERO DE ARTÍCULO', 
                'ARTICULO', 'ARTÍCULO', 'CODIGO', 'CÓDIGO', 'SKU_HIJO', 'SKU PADRE']
    
    for nombre in posibles:
        nombre_upper = nombre.upper().strip()
        # Buscar coincidencia exacta primero
        if nombre_upper in df_cols_upper:
            idx = df_cols_upper.index(nombre_upper)
            return df.columns[idx]
        # Luego buscar coincidencias parciales (pero evitar columnas de una sola letra)
        for idx, col_upper in enumerate(df_cols_upper):
            if len(col_upper) > 2:  # Evitar columnas de una o dos letras
                if nombre_upper in col_upper or col_upper in nombre_upper:
                    return df.columns[idx]
    
    return None


def detectar_columna_numero(df: pd.DataFrame) -> Optional[str]:
    """Detecta automáticamente la columna Numero"""
    df_cols_upper = [str(col).upper().strip() for col in df.columns]
    
    posibles = ['NUMERO', 'NÚMERO', 'NUM', 'NRO', 'DOCUMENTO', 'FACTURA']
    
    for nombre in posibles:
        nombre_upper = nombre.upper().strip()
        for idx, col_upper in enumerate(df_cols_upper):
            if nombre_upper == col_upper or nombre_upper in col_upper:
                return df.columns[idx]
    
    return None


def cargar_mp_key(archivo: str) -> Tuple[pd.DataFrame, str, str]:
    """Carga y valida el archivo MP KEY"""
    print(f"📂 Cargando {archivo}...")
    
    # Intentar diferentes configuraciones de lectura
    df = None
    header_row = None
    
    # Buscar la fila de encabezado
    for i in range(10):
        try:
            df_test = pd.read_excel(archivo, header=i, nrows=5)
            if len(df_test.columns) > 3:  # Archivo tiene suficientes columnas
                df = pd.read_excel(archivo, header=i)
                header_row = i
                break
        except:
            continue
    
    if df is None:
        df = pd.read_excel(archivo)
    
    # Detectar columnas relevantes
    col_clave_p = detectar_columna_clave_p(df, ['CLAVE_P', 'KEY', 'KEY_MS', 'CLAVE PRODUCTO'])
    col_no_sap = detectar_columna_no_sap(df)
    
    if col_clave_p is None:
        raise ValueError(f"❌ No se encontró la columna CLAVE_P en {archivo}. Columnas disponibles: {df.columns.tolist()}")
    
    # NO_SAP es opcional
    if col_no_sap is None:
        print(f"⚠️  NO_SAP no encontrado en {archivo}, continuando solo con CLAVE_P")
        df = df[[col_clave_p]].copy()
        df.columns = ['CLAVE_P']
    else:
        print(f"✅ Columnas detectadas: CLAVE_P='{col_clave_p}', NO_SAP='{col_no_sap}'")
        df = df[[col_clave_p, col_no_sap]].copy()
        df.columns = ['CLAVE_P', 'NO_SAP']
    
    # Limpiar datos
    df = df.dropna(subset=['CLAVE_P'])
    df['CLAVE_P'] = df['CLAVE_P'].astype(str).str.strip()
    df = df[df['CLAVE_P'] != '']
    
    print(f"✅ MP KEY cargado: {len(df)} registros únicos")
    
    if MODO_TESTING:
        print(f"🔍 TESTING - Primeras 5 CLAVE_P: {df['CLAVE_P'].head().tolist()}")
    
    return df, col_clave_p, col_no_sap


def cargar_ventas(archivo: str) -> Tuple[pd.DataFrame, str, str]:
    """Carga y valida el archivo VENTAS"""
    print(f"📂 Cargando {archivo}...")
    
    # Intentar diferentes configuraciones de lectura
    df = None
    header_row = None
    col_clave_p = None
    col_numero = None
    
    # Buscar la fila de encabezado que contenga las columnas necesarias
    for i in range(15):
        try:
            df_test = pd.read_excel(archivo, header=i, nrows=10)
            if len(df_test.columns) > 3:  # Archivo tiene suficientes columnas
                # Intentar detectar columnas en este encabezado
                temp_clave_p = detectar_columna_clave_p(df_test, ['CLAVE_P', 'KEY', 'KEY_MS', 'CLAVE PRODUCTO'])
                temp_numero = detectar_columna_numero(df_test)
                
                if temp_clave_p is not None and temp_numero is not None:
                    # Encontramos el encabezado correcto
                    df = pd.read_excel(archivo, header=i)
                    col_clave_p = temp_clave_p
                    col_numero = temp_numero
                    header_row = i
                    break
        except Exception as e:
            if MODO_TESTING:
                print(f"🔍 TESTING - Error en fila {i}: {e}")
            continue
    
    if df is None:
        # Último intento: leer sin encabezado y buscar manualmente
        df_raw = pd.read_excel(archivo, header=None)
        # Buscar fila que contenga 'KEY_MS' o 'Numero'
        for i in range(min(15, len(df_raw))):
            row_values = [str(val).upper() for val in df_raw.iloc[i].values if pd.notna(val)]
            if 'KEY_MS' in row_values or 'NUMERO' in row_values:
                df = pd.read_excel(archivo, header=i)
                col_clave_p = detectar_columna_clave_p(df, ['CLAVE_P', 'KEY', 'KEY_MS', 'CLAVE PRODUCTO'])
                col_numero = detectar_columna_numero(df)
                break
        
        if df is None:
            df = pd.read_excel(archivo)
            col_clave_p = detectar_columna_clave_p(df, ['CLAVE_P', 'KEY', 'KEY_MS', 'CLAVE PRODUCTO'])
            col_numero = detectar_columna_numero(df)
    
    if col_clave_p is None:
        raise ValueError(f"❌ No se encontró la columna CLAVE_P en {archivo}. Columnas disponibles: {df.columns.tolist()}")
    
    if col_numero is None:
        raise ValueError(f"❌ No se encontró la columna Numero en {archivo}. Columnas disponibles: {df.columns.tolist()}")
    
    print(f"✅ Columnas detectadas: CLAVE_P='{col_clave_p}', Numero='{col_numero}'")
    
    # Limpiar datos
    df = df.copy()
    df['CLAVE_P'] = df[col_clave_p].astype(str).str.strip()
    df['Numero'] = df[col_numero].astype(str).str.strip()
    
    # Eliminar filas sin CLAVE_P o Numero
    df = df.dropna(subset=['CLAVE_P', 'Numero'])
    df = df[(df['CLAVE_P'] != '') & (df['Numero'] != '')]
    
    print(f"✅ VENTAS cargado: {len(df)} transacciones")
    
    if MODO_TESTING:
        print(f"🔍 TESTING - Primeras 5 CLAVE_P: {df['CLAVE_P'].head().tolist()}")
        print(f"🔍 TESTING - Primeras 5 Numero: {df['Numero'].head().tolist()}")
    
    return df, col_clave_p, col_numero


def normalizar_clave(clave: str) -> str:
    """Normaliza una clave para comparación (elimina solo espacios, convierte a mayúsculas)"""
    if pd.isna(clave):
        return ''
    return str(clave).upper().strip().replace(' ', '')


def procesar_transacciones(df_ventas: pd.DataFrame, df_mp_key: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Procesa las transacciones según las reglas de negocio"""
    print("\n🔄 Procesando transacciones...")
    
    # Guardar una copia del original para validaciones
    df_ventas_original = df_ventas.copy()
    
    # Normalizar claves del maestro
    df_mp_key['CLAVE_P_NORM'] = df_mp_key['CLAVE_P'].apply(normalizar_clave)
    claves_validas = set(df_mp_key['CLAVE_P_NORM'].unique())
    print(f"📊 Total de CLAVE_P únicos en MP KEY: {len(claves_validas)}")
    
    # Normalizar claves de ventas
    df_ventas['CLAVE_P_NORM'] = df_ventas['CLAVE_P'].apply(normalizar_clave)
    
    # Marcar cada línea de VENTAS
    df_ventas['STATUS'] = df_ventas['CLAVE_P_NORM'].apply(
        lambda x: 'OK' if x in claves_validas else 'NO_ENCONTRADO'
    )
    
    # Estadísticas
    total_ok = (df_ventas['STATUS'] == 'OK').sum()
    total_no_encontrado = (df_ventas['STATUS'] == 'NO_ENCONTRADO').sum()
    print(f"📊 Líneas OK: {total_ok}")
    print(f"📊 Líneas NO_ENCONTRADO: {total_no_encontrado}")
    
    # Agrupar por Numero y clasificar
    print("\n📦 Agrupando por Numero...")
    
    grupos_status = df_ventas.groupby('Numero')['STATUS'].apply(
        lambda x: 'OK' if all(s == 'OK' for s in x) else 'PENDIENTE'
    ).to_dict()
    
    df_ventas['GRUPO_STATUS'] = df_ventas['Numero'].map(grupos_status)
    
    # Separar en dos grupos
    df_cargar = df_ventas[df_ventas['GRUPO_STATUS'] == 'OK'].copy()
    df_pendientes = df_ventas[df_ventas['GRUPO_STATUS'] == 'PENDIENTE'].copy()
    
    # Agregar columna ENC a pendientes antes de eliminar STATUS
    if len(df_pendientes) > 0:
        df_pendientes['ENC'] = df_pendientes['STATUS'].apply(lambda x: 1 if x == 'OK' else 0)
    
    # Crear listado de códigos no encontrados (únicos)
    codigos_no_encontrados = df_pendientes[df_pendientes['STATUS'] == 'NO_ENCONTRADO']['CLAVE_P'].unique()
    df_codigos_no_encontrados = pd.DataFrame({
        'CLAVE_P': codigos_no_encontrados
    }).sort_values('CLAVE_P').reset_index(drop=True)
    
    # Agregar SAP_ID a TX_Carga (hacer merge con MP KEY)
    if len(df_cargar) > 0:
        # Crear diccionario de CLAVE_P_NORM -> NO_SAP (SAP_ID)
        if 'NO_SAP' in df_mp_key.columns:
            # Crear mapeo de CLAVE_P normalizado a NO_SAP
            df_mp_key_mapeo = df_mp_key[['CLAVE_P_NORM', 'NO_SAP']].copy()
            df_mp_key_mapeo = df_mp_key_mapeo.drop_duplicates(subset=['CLAVE_P_NORM'])
            mapeo_sap = dict(zip(df_mp_key_mapeo['CLAVE_P_NORM'], df_mp_key_mapeo['NO_SAP']))
            
            # Agregar SAP_ID a df_cargar
            df_cargar['SAP_ID'] = df_cargar['CLAVE_P_NORM'].map(mapeo_sap)
            
            # Reordenar columnas: mover SAP_ID después de CLAVE_P
            cols = df_cargar.columns.tolist()
            if 'CLAVE_P' in cols and 'SAP_ID' in cols:
                # Encontrar posición de CLAVE_P
                idx_clave_p = cols.index('CLAVE_P')
                # Remover SAP_ID de su posición actual
                cols.remove('SAP_ID')
                # Insertar SAP_ID después de CLAVE_P
                cols.insert(idx_clave_p + 1, 'SAP_ID')
                df_cargar = df_cargar[cols]
        else:
            print("⚠️  NO_SAP no disponible en MP KEY, no se agregará SAP_ID")
            df_cargar['SAP_ID'] = None
    
    # Eliminar columnas auxiliares
    df_cargar = df_cargar.drop(columns=['STATUS', 'GRUPO_STATUS', 'CLAVE_P_NORM'], errors='ignore')
    df_pendientes = df_pendientes.drop(columns=['STATUS', 'GRUPO_STATUS', 'CLAVE_P_NORM'], errors='ignore')
    
    # Estadísticas finales
    grupos_ok = len([s for s in grupos_status.values() if s == 'OK'])
    grupos_pendientes = len([s for s in grupos_status.values() if s == 'PENDIENTE'])
    
    print(f"\n📊 RESUMEN:")
    print(f"   • Total de transacciones: {len(df_ventas)}")
    print(f"   • Total de grupos (Numero): {len(grupos_status)}")
    print(f"   • Grupos OK: {grupos_ok}")
    print(f"   • Grupos pendientes: {grupos_pendientes}")
    print(f"   • Transacciones por cargar: {len(df_cargar)}")
    print(f"   • Transacciones pendientes: {len(df_pendientes)}")
    print(f"   • % coincidencias CLAVE_P: {(total_ok/len(df_ventas)*100):.2f}%")
    
    return df_cargar, df_pendientes, df_ventas_original, df_codigos_no_encontrados, df_mp_key


def validar_resultados(df_original: pd.DataFrame, df_cargar: pd.DataFrame, df_pendientes: pd.DataFrame):
    """Valida que los resultados sean correctos"""
    print("\n🔍 Validando resultados...")
    
    errores = []
    advertencias = []
    
    # 1. Validar que no hay números duplicados entre archivos
    if len(df_cargar) > 0 and len(df_pendientes) > 0:
        numeros_cargar = set(df_cargar['Numero'].astype(str).unique())
        numeros_pendientes = set(df_pendientes['Numero'].astype(str).unique())
        duplicados = numeros_cargar.intersection(numeros_pendientes)
        
        if len(duplicados) > 0:
            errores.append(f"❌ ERROR: Se encontraron {len(duplicados)} números duplicados entre archivos")
            if MODO_TESTING:
                print(f"   Ejemplos de duplicados: {list(duplicados)[:10]}")
        else:
            print("✅ Validación 1: No hay números duplicados entre archivos")
    else:
        print("⚠️  Validación 1: No se puede validar (uno de los archivos está vacío)")
    
    # 2. Validar total de líneas
    total_original = len(df_original)
    total_cargar = len(df_cargar)
    total_pendientes = len(df_pendientes)
    total_salida = total_cargar + total_pendientes
    
    if total_original != total_salida:
        errores.append(f"❌ ERROR: Total de líneas no cuadra. Original: {total_original}, Suma salidas: {total_salida}, Diferencia: {abs(total_original - total_salida)}")
    else:
        print(f"✅ Validación 2: Total de líneas cuadra ({total_original} = {total_cargar} + {total_pendientes})")
    
    # 3. Validar total de números únicos
    numeros_original = set(df_original['Numero'].astype(str).unique())
    numeros_cargar_unicos = set(df_cargar['Numero'].astype(str).unique()) if len(df_cargar) > 0 else set()
    numeros_pendientes_unicos = set(df_pendientes['Numero'].astype(str).unique()) if len(df_pendientes) > 0 else set()
    numeros_salida_unicos = numeros_cargar_unicos.union(numeros_pendientes_unicos)
    
    if len(numeros_original) != len(numeros_salida_unicos):
        errores.append(f"❌ ERROR: Total de números únicos no cuadra. Original: {len(numeros_original)}, Salida: {len(numeros_salida_unicos)}, Diferencia: {abs(len(numeros_original) - len(numeros_salida_unicos))}")
        # Verificar si hay números faltantes o extra
        faltantes = numeros_original - numeros_salida_unicos
        extras = numeros_salida_unicos - numeros_original
        if len(faltantes) > 0:
            errores.append(f"   Números en original pero no en salida: {len(faltantes)}")
        if len(extras) > 0:
            errores.append(f"   Números en salida pero no en original: {len(extras)}")
    else:
        print(f"✅ Validación 3: Total de números únicos cuadra (Original: {len(numeros_original)}, Cargar: {len(numeros_cargar_unicos)}, Pendientes: {len(numeros_pendientes_unicos)}, Total salida: {len(numeros_salida_unicos)})")
    
    # 4. Validar TotalLinea (si existe la columna)
    if 'TotalLinea' in df_original.columns:
        try:
            total_linea_original = df_original['TotalLinea'].sum()
            total_linea_cargar = df_cargar['TotalLinea'].sum() if len(df_cargar) > 0 else 0
            total_linea_pendientes = df_pendientes['TotalLinea'].sum() if len(df_pendientes) > 0 else 0
            total_linea_salida = total_linea_cargar + total_linea_pendientes
            
            diferencia = abs(total_linea_original - total_linea_salida)
            if diferencia > 0.01:  # Tolerancia para errores de punto flotante
                errores.append(f"❌ ERROR: TotalLinea no cuadra. Original: {total_linea_original:.2f}, Salida: {total_linea_salida:.2f}, Diferencia: {diferencia:.2f}")
            else:
                print(f"✅ Validación 4: TotalLinea cuadra ({total_linea_original:.2f} = {total_linea_cargar:.2f} + {total_linea_pendientes:.2f})")
        except Exception as e:
            advertencias.append(f"⚠️  No se pudo validar TotalLinea: {e}")
    else:
        advertencias.append("⚠️  Columna 'TotalLinea' no encontrada en archivo original")
    
    # Mostrar resultados
    print("\n" + "=" * 60)
    if len(errores) > 0:
        print("❌ ERRORES ENCONTRADOS:")
        for error in errores:
            print(f"   {error}")
        return False
    else:
        print("✅ Todas las validaciones pasaron correctamente")
        if len(advertencias) > 0:
            print("\n⚠️  ADVERTENCIAS:")
            for adv in advertencias:
                print(f"   {adv}")
        return True
    print("=" * 60)


def guardar_archivos(df_cargar: pd.DataFrame, df_pendientes: pd.DataFrame, df_codigos_no_encontrados: pd.DataFrame, carpeta: str, timestamp: str):
    """Guarda los archivos de salida con timestamp en la carpeta del proceso"""
    print(f"\n💾 Guardando archivos de salida en {carpeta}/...")
    
    # Generar nombres con timestamp dentro de la carpeta
    archivo_cargar, archivo_pendientes = generar_nombres_con_timestamp(carpeta, timestamp)
    
    # Guardar transacciones por cargar
    if len(df_cargar) > 0:
        try:
            print(f"   Guardando {archivo_cargar}...")
            with pd.ExcelWriter(archivo_cargar, engine='openpyxl', mode='w') as writer:
                df_cargar.to_excel(writer, sheet_name='Sheet1', index=False)
            print(f"✅ {archivo_cargar} guardado ({len(df_cargar)} transacciones)")
        except Exception as e:
            print(f"⚠️  Error al guardar {archivo_cargar}: {e}")
    else:
        print(f"⚠️  No hay transacciones para cargar")
    
    # Guardar transacciones pendientes (con hoja 1 y hoja 2)
    if len(df_pendientes) > 0:
        try:
            print(f"   Guardando {archivo_pendientes}...")
            with pd.ExcelWriter(archivo_pendientes, engine='openpyxl', mode='w') as writer:
                # Hoja 1: Transacciones pendientes con columna ENC
                df_pendientes.to_excel(writer, sheet_name='Sheet1', index=False)
                
                # Hoja 2: Listado de códigos no encontrados
                if len(df_codigos_no_encontrados) > 0:
                    df_codigos_no_encontrados.to_excel(writer, sheet_name='Sheet2', index=False)
                    print(f"      Hoja 2: {len(df_codigos_no_encontrados)} códigos no encontrados")
                else:
                    # Crear hoja vacía si no hay códigos no encontrados
                    pd.DataFrame(columns=['CLAVE_P']).to_excel(writer, sheet_name='Sheet2', index=False)
                    print(f"      Hoja 2: Sin códigos no encontrados")
            
            print(f"✅ {archivo_pendientes} guardado ({len(df_pendientes)} transacciones)")
            print(f"   - Hoja 1: Transacciones pendientes con columna ENC")
            print(f"   - Hoja 2: Listado de códigos no encontrados")
        except Exception as e:
            print(f"⚠️  Error al guardar {archivo_pendientes}: {e}")
            print(f"   Intentando guardar en formato CSV como alternativa...")
            try:
                csv_file = archivo_pendientes.replace('.xlsx', '.csv')
                df_pendientes.to_csv(csv_file, index=False, encoding='utf-8-sig')
                print(f"✅ Guardado como CSV: {csv_file}")
            except Exception as e2:
                print(f"❌ Error también al guardar CSV: {e2}")
    else:
        print(f"⚠️  No hay transacciones pendientes")
    
    return archivo_cargar, archivo_pendientes


def analizar_y_guardar_resumen(df_ventas: pd.DataFrame, df_mp_key: pd.DataFrame, df_pendientes: pd.DataFrame, 
                                df_codigos_no_encontrados: pd.DataFrame, carpeta: str, timestamp: str):
    """Analiza las transacciones pendientes y guarda el resumen en Excel"""
    
    # Calcular métricas
    total_transacciones = len(df_ventas)
    total_codigos_unicos_ventas = df_ventas['CLAVE_P'].nunique()
    total_documentos_unicos = df_ventas['Numero'].nunique()
    
    # Normalizar claves del MP KEY
    df_mp_key['CLAVE_P_NORM'] = df_mp_key['CLAVE_P'].apply(normalizar_clave)
    total_codigos_validos = df_mp_key['CLAVE_P_NORM'].nunique()
    
    # Métricas de pendientes
    transacciones_afectadas = len(df_pendientes)
    documentos_afectados = df_pendientes['Numero'].nunique()
    codigos_faltantes = len(df_codigos_no_encontrados)
    
    # Contar líneas sin código (ENC=0) y con código (ENC=1)
    if 'ENC' in df_pendientes.columns:
        lineas_sin_codigo = (df_pendientes['ENC'] == 0).sum()
        lineas_con_codigo = (df_pendientes['ENC'] == 1).sum()
    else:
        # Calcular manualmente si no hay columna ENC
        claves_validas = set(df_mp_key['CLAVE_P_NORM'].unique())
        df_pendientes['CLAVE_P_NORM'] = df_pendientes['CLAVE_P'].apply(normalizar_clave)
        lineas_sin_codigo = (~df_pendientes['CLAVE_P_NORM'].isin(claves_validas)).sum()
        lineas_con_codigo = df_pendientes['CLAVE_P_NORM'].isin(claves_validas).sum()
    
    # Calcular porcentajes
    porcentaje_codigos_faltantes = (codigos_faltantes / total_codigos_unicos_ventas * 100) if total_codigos_unicos_ventas > 0 else 0
    porcentaje_sin_codigo = (lineas_sin_codigo / total_transacciones * 100) if total_transacciones > 0 else 0
    porcentaje_trans_afectadas = (transacciones_afectadas / total_transacciones * 100) if total_transacciones > 0 else 0
    porcentaje_doc_afectados = (documentos_afectados / total_documentos_unicos * 100) if total_documentos_unicos > 0 else 0
    
    # Crear DataFrame con resumen
    resumen_data = {
        'Métrica': [
            '1. CÓDIGOS FALTANTES',
            '   Total códigos válidos en MP KEY',
            '   Total códigos únicos en ventas',
            '   Códigos faltantes (no encontrados)',
            '   Porcentaje faltante (%)',
            '',
            '2. TRANSACCIONES AFECTADAS',
            '   Total de transacciones',
            '   Líneas sin código válido',
            '   Porcentaje sin código (%)',
            '   Líneas afectadas (total)',
            '   Porcentaje afectadas (%)',
            '   Líneas con código válido (en docs afectados)',
            '',
            '3. DOCUMENTOS/FACTURAS AFECTADAS',
            '   Total de documentos únicos',
            '   Documentos afectados',
            '   Porcentaje afectado (%)'
        ],
        'Valor': [
            '',
            f'{total_codigos_validos:,}',
            f'{total_codigos_unicos_ventas:,}',
            f'{codigos_faltantes:,}',
            f'{porcentaje_codigos_faltantes:.2f}%',
            '',
            '',
            f'{total_transacciones:,}',
            f'{lineas_sin_codigo:,}',
            f'{porcentaje_sin_codigo:.2f}%',
            f'{transacciones_afectadas:,}',
            f'{porcentaje_trans_afectadas:.2f}%',
            f'{lineas_con_codigo:,}',
            '',
            '',
            f'{total_documentos_unicos:,}',
            f'{documentos_afectados:,}',
            f'{porcentaje_doc_afectados:.2f}%'
        ]
    }
    
    df_resumen = pd.DataFrame(resumen_data)
    
    # Guardar resumen en Excel
    archivo_resumen = os.path.join(carpeta, f"Resumen_Analisis_{timestamp}.xlsx")
    
    try:
        with pd.ExcelWriter(archivo_resumen, engine='openpyxl') as writer:
            # Hoja 1: Resumen ejecutivo
            df_resumen.to_excel(writer, sheet_name='Resumen', index=False)
            
            # Hoja 2: Detalle de códigos faltantes
            if len(df_codigos_no_encontrados) > 0:
                df_codigos_no_encontrados.to_excel(writer, sheet_name='Codigos_Faltantes', index=False)
            
            # Hoja 3: Estadísticas por documento afectado
            if len(df_pendientes) > 0:
                stats_docs = df_pendientes.groupby('Numero').agg({
                    'CLAVE_P': 'count',
                    'ENC': lambda x: (x == 0).sum() if 'ENC' in df_pendientes.columns else 0
                }).reset_index()
                stats_docs.columns = ['Numero', 'Total_Lineas', 'Lineas_Sin_Codigo']
                stats_docs = stats_docs.sort_values('Lineas_Sin_Codigo', ascending=False)
                stats_docs.to_excel(writer, sheet_name='Docs_Afectados', index=False)
        
        print(f"✅ Resumen guardado: {os.path.basename(archivo_resumen)}")
        print(f"\n📊 RESUMEN DEL ANÁLISIS:")
        print(f"   • Códigos faltantes: {codigos_faltantes:,} ({porcentaje_codigos_faltantes:.2f}%)")
        print(f"   • Transacciones afectadas: {transacciones_afectadas:,} ({porcentaje_trans_afectadas:.2f}%)")
        print(f"   • Documentos afectados: {documentos_afectados:,} ({porcentaje_doc_afectados:.2f}%)")
        
    except Exception as e:
        print(f"⚠️  Error al guardar resumen: {e}")


def main():
    """Función principal"""
    print("=" * 60)
    print("🚀 PROMPT2 - Procesamiento MP KEY vs VENTAS")
    print("=" * 60)
    
    if MODO_TESTING:
        print("🧪 MODO TESTING ACTIVADO\n")
    
    try:
        # 0. Limpiar carpeta de origen (mover archivos anteriores)
        limpiar_carpeta_origen()
        
        # 1. Crear carpeta para el proceso
        carpeta, timestamp = crear_carpeta_proceso()
        print(f"\n📁 Carpeta del proceso: {carpeta}/")
        
        # 2. Cargar archivos
        df_mp_key, _, _ = cargar_mp_key(ARCHIVO_MP_KEY)
        df_ventas, _, _ = cargar_ventas(ARCHIVO_VENTAS)
        
        # 3. Copiar archivos de origen a la carpeta
        copiar_archivos_origen(carpeta)
        
        # 4. Procesar transacciones
        df_cargar, df_pendientes, df_ventas_original, df_codigos_no_encontrados, df_mp_key_procesado = procesar_transacciones(df_ventas, df_mp_key)
        
        # 5. Validar resultados
        validacion_ok = validar_resultados(df_ventas_original, df_cargar, df_pendientes)
        
        # 6. Guardar archivos en la carpeta del proceso
        archivo_cargar, archivo_pendientes = guardar_archivos(df_cargar, df_pendientes, df_codigos_no_encontrados, carpeta, timestamp)
        
        # 7. Analizar pendientes y guardar resumen
        if len(df_pendientes) > 0:
            print("\n" + "=" * 60)
            print("📊 Analizando transacciones pendientes...")
            print("=" * 60)
            analizar_y_guardar_resumen(df_ventas_original, df_mp_key, df_pendientes, df_codigos_no_encontrados, carpeta, timestamp)
        
        print("\n" + "=" * 60)
        if validacion_ok:
            print("✅ Proceso completado exitosamente")
            print(f"📁 Carpeta del proceso: {carpeta}/")
            print(f"📄 Archivos en la carpeta:")
            print(f"   • {ARCHIVO_MP_KEY} (origen)")
            print(f"   • {ARCHIVO_VENTAS} (origen)")
            print(f"   • {os.path.basename(archivo_cargar)}")
            print(f"   • {os.path.basename(archivo_pendientes)}")
            if len(df_pendientes) > 0:
                print(f"   • Resumen_Analisis_{timestamp}.xlsx")
        else:
            print("⚠️  Proceso completado con errores de validación")
            print("   Revisa los errores mostrados arriba")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        if MODO_TESTING:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

