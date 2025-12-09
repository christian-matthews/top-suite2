#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Comparador de Planillas Excel (BASE vs FINAL)
Compara dos archivos Excel y genera reportes de diferencias.
"""

import pandas as pd
import numpy as np
import argparse
import os
import sys
import shutil
from datetime import datetime
from typing import Tuple, Dict, List, Optional


def load_excel(path: str, sheet_name: str, fallback_sheet: str = None) -> pd.DataFrame:
    """
    Carga un archivo Excel y devuelve un DataFrame.
    
    Args:
        path: Ruta al archivo Excel
        sheet_name: Nombre de la hoja a leer
        fallback_sheet: Nombre alternativo de hoja si la principal no existe
        
    Returns:
        DataFrame con los datos de la hoja
        
    Raises:
        FileNotFoundError: Si el archivo no existe
        ValueError: Si la hoja no existe
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"El archivo no existe: {path}")
    
    try:
        # Leer el archivo Excel
        excel_file = pd.ExcelFile(path)
        
        # Intentar usar la hoja principal, si no existe usar fallback
        actual_sheet = sheet_name
        if sheet_name not in excel_file.sheet_names:
            if fallback_sheet and fallback_sheet in excel_file.sheet_names:
                actual_sheet = fallback_sheet
                print(f"Advertencia: La hoja '{sheet_name}' no existe, usando '{fallback_sheet}'")
            else:
                available_sheets = ", ".join(excel_file.sheet_names)
                raise ValueError(
                    f"La hoja '{sheet_name}' no existe en {path}. "
                    f"Hojas disponibles: {available_sheets}"
                )
        
        # Leer la hoja específica
        df = pd.read_excel(path, sheet_name=actual_sheet)
        
        if df.empty:
            print(f"Advertencia: La hoja '{actual_sheet}' en {path} está vacía.")
        
        return df
        
    except FileNotFoundError:
        raise
    except Exception as e:
        raise ValueError(f"Error al leer {path}, hoja '{sheet_name}': {str(e)}")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normaliza los nombres de columnas: elimina espacios y estandariza a mayúsculas.
    
    Args:
        df: DataFrame a normalizar
        
    Returns:
        DataFrame con columnas normalizadas
    """
    df = df.copy()
    # Eliminar espacios al inicio y final, convertir a mayúsculas
    df.columns = df.columns.str.strip().str.upper()
    return df


def validate_structure(df_base: pd.DataFrame, df_final: pd.DataFrame, 
                      key_column: str = "SKU_HIJO") -> Tuple[bool, List[str]]:
    """
    Valida la estructura de los DataFrames.
    
    Args:
        df_base: DataFrame de BASE
        df_final: DataFrame de FINAL
        key_column: Nombre de la columna clave
        
    Returns:
        Tupla (es_valido, lista_errores)
    """
    errors = []
    
    # Verificar que los DataFrames no estén vacíos
    if df_base.empty:
        errors.append("El DataFrame BASE está vacío.")
    
    if df_final.empty:
        errors.append("El DataFrame FINAL está vacío.")
    
    # Verificar existencia de la columna clave
    if key_column not in df_base.columns:
        errors.append(f"La columna '{key_column}' no existe en BASE.")
    
    if key_column not in df_final.columns:
        errors.append(f"La columna '{key_column}' no existe en FINAL.")
    
    return len(errors) == 0, errors


def find_duplicates(df: pd.DataFrame, key_column: str) -> pd.DataFrame:
    """
    Detecta registros duplicados por la columna clave.
    Excluye valores NaN del análisis de duplicados.
    
    Args:
        df: DataFrame a analizar
        key_column: Nombre de la columna clave
        
    Returns:
        DataFrame con los registros duplicados (solo los duplicados, excluyendo NaN)
    """
    if key_column not in df.columns:
        return pd.DataFrame()
    
    # Filtrar valores no nulos para detectar duplicados reales
    df_with_values = df[df[key_column].notna()].copy()
    
    if df_with_values.empty:
        return pd.DataFrame()
    
    # Detectar duplicados solo en valores no nulos
    duplicates = df_with_values[df_with_values.duplicated(subset=[key_column], keep=False)]
    return duplicates.sort_values(by=key_column)


def normalize_value_for_comparison(value) -> str:
    """
    Normaliza un valor para comparación (strings: trim y lowercase).
    
    Args:
        value: Valor a normalizar
        
    Returns:
        String normalizado para comparación
    """
    if pd.isna(value):
        return ""
    
    if isinstance(value, (int, float)):
        # Para números, convertir a string sin decimales si es entero
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)
    
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    
    # Para strings: trim y lowercase
    return str(value).strip().lower()


def normalize_sku_hijo(value) -> str:
    """
    Normaliza SKU_HIJO para comparación consistente.
    Convierte números (int/float) a string sin decimales.
    
    Args:
        value: Valor de SKU_HIJO
        
    Returns:
        String normalizado
    """
    if pd.isna(value):
        return ""
    
    # Si es numérico, convertir a int primero para eliminar decimales
    if isinstance(value, (int, float)):
        return str(int(value))
    
    # Si es string, eliminar espacios y puntos decimales al final
    str_value = str(value).strip()
    if str_value.endswith('.0'):
        str_value = str_value[:-2]
    
    return str_value


def format_value_for_display(value) -> str:
    """
    Formatea un valor para mostrarlo claramente en el reporte.
    Preserva el formato completo de números grandes.
    
    Args:
        value: Valor a formatear
        
    Returns:
        String formateado para mostrar
    """
    if pd.isna(value):
        return "(vacío)"
    
    if isinstance(value, (int, float)):
        # Si es un float que es entero, mostrar sin decimales y sin notación científica
        if isinstance(value, float) and value.is_integer():
            # Convertir a int para evitar notación científica en números grandes
            int_value = int(value)
            return str(int_value)
        # Para floats con decimales, mostrar con formato adecuado
        return f"{value:.10f}".rstrip('0').rstrip('.')
    
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    
    # Para strings, mantener el valor original
    return str(value)


def compare_data(df_base: pd.DataFrame, df_final: pd.DataFrame, 
                key_column: str = "SKU_HIJO") -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Compara los DataFrames y detecta NUEVOS, ELIMINADOS y MODIFICADOS.
    Solo trabaja con registros que tienen SKU_HIJO válido (no NaN).
    
    Args:
        df_base: DataFrame de BASE
        df_final: DataFrame de FINAL
        key_column: Nombre de la columna clave
        
    Returns:
        Tupla (df_nuevos, df_eliminados, df_modificados)
    """
    # Filtrar solo registros con SKU_HIJO válido (no NaN)
    df_base_validos = df_base[df_base[key_column].notna()].copy()
    df_final_validos = df_final[df_final[key_column].notna()].copy()
    
    # Normalizar SKU_HIJO para comparación consistente (convertir números a string sin decimales)
    df_base_validos[key_column + "_normalized"] = df_base_validos[key_column].apply(normalize_sku_hijo)
    df_final_validos[key_column + "_normalized"] = df_final_validos[key_column].apply(normalize_sku_hijo)
    
    # Crear sets de SKU_HIJO normalizados para comparación rápida
    sku_base = set(df_base_validos[key_column + "_normalized"])
    sku_final = set(df_final_validos[key_column + "_normalized"])
    
    # NUEVOS: están en FINAL pero no en BASE
    nuevos_sku = sku_final - sku_base
    df_nuevos = df_final_validos[df_final_validos[key_column + "_normalized"].isin(nuevos_sku)].copy()
    # Eliminar columna temporal
    if key_column + "_normalized" in df_nuevos.columns:
        df_nuevos = df_nuevos.drop(columns=[key_column + "_normalized"])
    
    # ELIMINADOS: están en BASE pero no en FINAL
    eliminados_sku = sku_base - sku_final
    df_eliminados = df_base_validos[df_base_validos[key_column + "_normalized"].isin(eliminados_sku)].copy()
    # Eliminar columna temporal
    if key_column + "_normalized" in df_eliminados.columns:
        df_eliminados = df_eliminados.drop(columns=[key_column + "_normalized"])
    
    # MODIFICADOS: están en ambos, pero hay diferencias
    comunes_sku = sku_base & sku_final
    
    modificados_list = []
    
    for sku_normalized in comunes_sku:
        # Obtener las filas correspondientes
        rows_base = df_base_validos[df_base_validos[key_column + "_normalized"] == sku_normalized]
        rows_final = df_final_validos[df_final_validos[key_column + "_normalized"] == sku_normalized]
        
        # Si hay múltiples filas con el mismo SKU_HIJO, tomar la primera y advertir
        if len(rows_base) > 1:
            # Advertir pero continuar con la primera
            pass
        if len(rows_final) > 1:
            # Advertir pero continuar con la primera
            pass
        
        row_base = rows_base.iloc[0]
        row_final = rows_final.iloc[0]
        
        # Usar el SKU de FINAL para mostrar (más actualizado)
        sku_original = row_final[key_column] if key_column in row_final else row_base[key_column]
        
        # Comparar todas las columnas excepto la clave y la columna temporal
        columns_to_compare = [col for col in df_base_validos.columns 
                             if col != key_column and col != key_column + "_normalized"]
        
        for col in columns_to_compare:
            if col not in df_final_validos.columns:
                continue
            
            val_base = row_base[col]
            val_final = row_final[col]
            
            # Normalizar para comparación
            norm_base = normalize_value_for_comparison(val_base)
            norm_final = normalize_value_for_comparison(val_final)
            
            # Si hay diferencia, registrar el cambio
            if norm_base != norm_final:
                # Determinar tipo de cambio
                tipo_cambio = "texto"
                if isinstance(val_base, (int, float)) or isinstance(val_final, (int, float)):
                    tipo_cambio = "numérico"
                elif isinstance(val_base, datetime) or isinstance(val_final, datetime):
                    tipo_cambio = "fecha"
                
                # Formatear valores para mostrar claramente qué cambió
                valor_base_str = format_value_for_display(val_base)
                valor_final_str = format_value_for_display(val_final)
                
                modificados_list.append({
                    key_column: sku_original,
                    "COLUMNA": col,
                    "VALOR_BASE": valor_base_str,
                    "VALOR_FINAL": valor_final_str,
                    "ESTADO": "MODIFICADO",
                    "TIPO_CAMBIO": tipo_cambio
                })
    
    df_modificados = pd.DataFrame(modificados_list)
    
    return df_nuevos, df_eliminados, df_modificados


def generate_excel_output(df_nuevos: pd.DataFrame, df_eliminados: pd.DataFrame,
                         df_modificados: pd.DataFrame, timestamp: str, 
                         output_dir: str) -> str:
    """
    Genera el archivo Excel de salida con las tres hojas.
    
    Args:
        df_nuevos: DataFrame de registros nuevos
        df_eliminados: DataFrame de registros eliminados
        df_modificados: DataFrame de registros modificados
        timestamp: Timestamp para el nombre del archivo
        output_dir: Directorio de salida
        
    Returns:
        Ruta del archivo generado
    """
    filename = f"RESULTADOS_COMPARACION_{timestamp}.xlsx"
    filepath = os.path.join(output_dir, filename)
    
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        # Hoja NUEVOS
        if not df_nuevos.empty:
            df_nuevos.to_excel(writer, sheet_name="NUEVOS", index=False)
        else:
            pd.DataFrame(columns=df_nuevos.columns if not df_nuevos.empty else ["SKU_HIJO"]).to_excel(
                writer, sheet_name="NUEVOS", index=False
            )
        
        # Hoja ELIMINADOS
        if not df_eliminados.empty:
            df_eliminados.to_excel(writer, sheet_name="ELIMINADOS", index=False)
        else:
            pd.DataFrame(columns=df_eliminados.columns if not df_eliminados.empty else ["SKU_HIJO"]).to_excel(
                writer, sheet_name="ELIMINADOS", index=False
            )
        
        # Hoja MODIFICADOS
        if not df_modificados.empty:
            df_modificados.to_excel(writer, sheet_name="MODIFICADOS", index=False)
        else:
            pd.DataFrame(columns=["SKU_HIJO", "COLUMNA", "VALOR_BASE", "VALOR_FINAL", "ESTADO", "TIPO_CAMBIO"]).to_excel(
                writer, sheet_name="MODIFICADOS", index=False
            )
    
    return filepath


def generate_report(df_base: pd.DataFrame, df_final: pd.DataFrame,
                   df_nuevos: pd.DataFrame, df_eliminados: pd.DataFrame,
                   df_modificados: pd.DataFrame, duplicates_info: Dict[str, pd.DataFrame],
                   timestamp: str, output_dir: str, errors: List[str] = None,
                   nan_base: int = 0, nan_final: int = 0) -> str:
    """
    Genera el archivo de reporte de texto.
    
    Args:
        df_base: DataFrame de BASE
        df_final: DataFrame de FINAL
        df_nuevos: DataFrame de registros nuevos
        df_eliminados: DataFrame de registros eliminados
        df_modificados: DataFrame de registros modificados
        duplicates_info: Diccionario con información de duplicados
        timestamp: Timestamp para el nombre del archivo
        output_dir: Directorio de salida
        errors: Lista de errores encontrados
        
    Returns:
        Ruta del archivo generado
    """
    filename = f"REPORTE_COMPARACION_{timestamp}.txt"
    filepath = os.path.join(output_dir, filename)
    
    # Calcular estadísticas
    total_nuevos = len(df_nuevos)
    total_eliminados = len(df_eliminados)
    total_modificados_sku = df_modificados["SKU_HIJO"].nunique() if not df_modificados.empty else 0
    total_diferencias_columna = len(df_modificados)
    
    # Obtener timestamp legible
    timestamp_readable = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write("=" * 80 + "\n")
        f.write("REPORTE DE COMPARACIÓN DE PLANILLAS EXCEL\n")
        f.write("=" * 80 + "\n\n")
        
        f.write(f"Fecha y hora de ejecución: {timestamp_readable}\n\n")
        
        f.write("ARCHIVOS COMPARADOS:\n")
        f.write("-" * 80 + "\n")
        f.write("BASE.xlsx – Hoja BASE\n")
        f.write("FINAL.xlsx – Hoja FINAL\n\n")
        
        f.write("ESTADÍSTICAS:\n")
        f.write("-" * 80 + "\n")
        f.write(f"Total de filas en BASE: {len(df_base) + nan_base} (procesadas: {len(df_base)}, excluidas (NaN): {nan_base})\n")
        f.write(f"Total de filas en FINAL: {len(df_final) + nan_final} (procesadas: {len(df_final)}, excluidas (NaN): {nan_final})\n")
        f.write(f"Total de registros NUEVOS: {total_nuevos}\n")
        f.write(f"Total de registros ELIMINADOS: {total_eliminados}\n")
        f.write(f"Total de registros MODIFICADOS (SKU únicos): {total_modificados_sku}\n")
        f.write(f"Total de diferencias a nivel de columna: {total_diferencias_columna}\n\n")
        
        # Información sobre valores NaN en SKU_HIJO
        if nan_base > 0 or nan_final > 0:
            f.write("REGISTROS EXCLUIDOS (SKU_HIJO vacío/NaN):\n")
            f.write("-" * 80 + "\n")
            if nan_base > 0:
                f.write(f"  BASE: {nan_base} registros excluidos (SKU_HIJO vacío)\n")
            if nan_final > 0:
                f.write(f"  FINAL: {nan_final} registros excluidos (SKU_HIJO vacío)\n")
            f.write("  Nota: Solo se procesan registros con SKU_HIJO válido (no vacío).\n\n")
        
        # Sección de duplicados (solo valores reales, excluyendo NaN)
        base_dups = duplicates_info.get("base_duplicates")
        final_dups = duplicates_info.get("final_duplicates")
        
        if base_dups is not None and not base_dups.empty:
            f.write("REGISTROS DUPLICADOS EN BASE (valores reales):\n")
            f.write("-" * 80 + "\n")
            dup_sku_base = base_dups["SKU_HIJO"].unique()
            for sku in dup_sku_base:
                count = len(base_dups[base_dups["SKU_HIJO"] == sku])
                f.write(f"  SKU_HIJO: {sku} (aparece {count} veces)\n")
            f.write(f"  Total de registros duplicados: {len(base_dups)}\n")
            f.write(f"  Total de SKU_HIJO únicos duplicados: {len(dup_sku_base)}\n\n")
        else:
            f.write("REGISTROS DUPLICADOS EN BASE:\n")
            f.write("-" * 80 + "\n")
            f.write("  No se encontraron duplicados (excluyendo valores vacíos).\n\n")
        
        if final_dups is not None and not final_dups.empty:
            f.write("REGISTROS DUPLICADOS EN FINAL (valores reales):\n")
            f.write("-" * 80 + "\n")
            dup_sku_final = final_dups["SKU_HIJO"].unique()
            for sku in dup_sku_final:
                count = len(final_dups[final_dups["SKU_HIJO"] == sku])
                f.write(f"  SKU_HIJO: {sku} (aparece {count} veces)\n")
            f.write(f"  Total de registros duplicados: {len(final_dups)}\n")
            f.write(f"  Total de SKU_HIJO únicos duplicados: {len(dup_sku_final)}\n\n")
        else:
            f.write("REGISTROS DUPLICADOS EN FINAL:\n")
            f.write("-" * 80 + "\n")
            f.write("  No se encontraron duplicados (excluyendo valores vacíos).\n\n")
        
        # Sección de errores
        if errors:
            f.write("ERRORES Y ADVERTENCIAS:\n")
            f.write("-" * 80 + "\n")
            for error in errors:
                f.write(f"  - {error}\n")
            f.write("\n")
        
        f.write("=" * 80 + "\n")
        f.write("Fin del reporte\n")
        f.write("=" * 80 + "\n")
    
    return filepath


def save_to_history(base_file: str, final_file: str, excel_output: str, 
                   report_output: str, timestamp: str, base_dir: str) -> str:
    """
    Crea una carpeta con timestamp y guarda los archivos BASE, FINAL y reportes.
    
    Args:
        base_file: Ruta al archivo BASE.xlsx
        final_file: Ruta al archivo FINAL.xlsx
        excel_output: Ruta al archivo Excel de resultados
        report_output: Ruta al archivo de reporte
        timestamp: Timestamp para el nombre de la carpeta
        base_dir: Directorio base donde crear la carpeta
        
    Returns:
        Ruta de la carpeta creada
    """
    # Crear nombre de carpeta con timestamp
    history_folder = f"HISTORIAL_{timestamp}"
    history_path = os.path.join(base_dir, history_folder)
    
    # Crear la carpeta
    os.makedirs(history_path, exist_ok=True)
    
    # Copiar archivos
    try:
        # Copiar BASE.xlsx
        if os.path.exists(base_file):
            shutil.copy2(base_file, os.path.join(history_path, "BASE.xlsx"))
        
        # Copiar FINAL.xlsx
        if os.path.exists(final_file):
            shutil.copy2(final_file, os.path.join(history_path, "FINAL.xlsx"))
        
        # Copiar Excel de resultados
        if os.path.exists(excel_output):
            shutil.copy2(excel_output, os.path.join(history_path, os.path.basename(excel_output)))
        
        # Copiar reporte de texto
        if os.path.exists(report_output):
            shutil.copy2(report_output, os.path.join(history_path, os.path.basename(report_output)))
        
        print(f"\n✓ Historial guardado en: {history_path}")
        return history_path
        
    except Exception as e:
        print(f"\n⚠ Error al guardar historial: {str(e)}")
        return history_path


def parse_args():
    """
    Parsea los argumentos de línea de comandos.
    
    Returns:
        Namespace con los argumentos parseados
    """
    parser = argparse.ArgumentParser(
        description="Comparador de Planillas Excel (BASE vs FINAL)"
    )
    parser.add_argument(
        '--test',
        action='store_true',
        help='Ejecutar en modo test (usar archivos de prueba)'
    )
    parser.add_argument(
        '--save',
        action='store_true',
        help='Guardar archivos BASE, FINAL y reportes en carpeta de historial con timestamp'
    )
    return parser.parse_args()


def main():
    """
    Función principal que orquesta todo el flujo.
    """
    # Capturar argumentos de la línea de comandos, pero ignorar los de uvicorn
    # Esto es crucial cuando se ejecuta el script a través de uvicorn
    original_argv = sys.argv
    sys.argv = [original_argv[0]]  # Solo el nombre del script
    try:
        args = parse_args()
    except SystemExit:
        # Si argparse falla, usar valores por defecto
        import argparse
        args = argparse.Namespace(test=False, save=False)
    finally:
        sys.argv = original_argv  # Restaurar
    
    # Determinar directorio de trabajo
    # Primero verificar si hay archivos en el directorio actual (para uso desde API)
    cwd = os.getcwd()
    if os.path.exists(os.path.join(cwd, "BASE.xlsx")) and os.path.exists(os.path.join(cwd, "FINAL.xlsx")):
        base_dir = cwd
        print(f"Usando directorio de trabajo: {cwd}")
    elif args.test:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        print("Modo TEST activado")
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Rutas de archivos
    base_file = os.path.join(base_dir, "BASE.xlsx")
    final_file = os.path.join(base_dir, "FINAL.xlsx")
    output_dir = base_dir
    
    # Timestamp para nombres de archivos
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    print("=" * 80)
    print("COMPARADOR DE PLANILLAS EXCEL (BASE vs FINAL)")
    print("=" * 80)
    print(f"\nDirectorio de trabajo: {base_dir}")
    print(f"Timestamp: {timestamp}\n")
    
    errors = []
    df_base = None
    df_final = None
    
    try:
        # 1. Cargar archivos
        print("Cargando archivos...")
        try:
            df_base = load_excel(base_file, "BASE")
            print(f"✓ BASE.xlsx cargado: {len(df_base)} filas")
        except Exception as e:
            errors.append(f"Error al cargar BASE.xlsx: {str(e)}")
            print(f"✗ Error: {str(e)}")
            sys.exit(1)
        
        try:
            df_final = load_excel(final_file, "FINAL", fallback_sheet="Hoja1")
            print(f"✓ FINAL.xlsx cargado: {len(df_final)} filas")
        except Exception as e:
            errors.append(f"Error al cargar FINAL.xlsx: {str(e)}")
            print(f"✗ Error: {str(e)}")
            sys.exit(1)
        
        # 2. Normalizar columnas
        print("\nNormalizando columnas...")
        df_base = normalize_columns(df_base)
        df_final = normalize_columns(df_final)
        print("✓ Columnas normalizadas")
        
        # 2.5. Filtrar registros válidos (solo con SKU_HIJO no nulo)
        print("\nFiltrando registros válidos (SKU_HIJO no nulo)...")
        nan_base = df_base["SKU_HIJO"].isna().sum()
        nan_final = df_final["SKU_HIJO"].isna().sum()
        df_base_validos = df_base[df_base["SKU_HIJO"].notna()].copy()
        df_final_validos = df_final[df_final["SKU_HIJO"].notna()].copy()
        
        print(f"  BASE: {len(df_base)} filas totales, {nan_base} con SKU_HIJO vacío → {len(df_base_validos)} registros válidos")
        print(f"  FINAL: {len(df_final)} filas totales, {nan_final} con SKU_HIJO vacío → {len(df_final_validos)} registros válidos")
        print("✓ Filtrado completado (solo se procesarán registros con SKU_HIJO válido)")
        
        # 3. Validar estructura
        print("\nValidando estructura...")
        is_valid, validation_errors = validate_structure(df_base_validos, df_final_validos, "SKU_HIJO")
        if not is_valid:
            errors.extend(validation_errors)
            print("✗ Errores de validación:")
            for error in validation_errors:
                print(f"  - {error}")
            sys.exit(1)
        print("✓ Estructura validada")
        
        # 4. Detectar duplicados (solo en registros válidos)
        print("\nDetectando duplicados (solo registros válidos)...")
        duplicates_base = find_duplicates(df_base_validos, "SKU_HIJO")
        duplicates_final = find_duplicates(df_final_validos, "SKU_HIJO")
        
        if not duplicates_base.empty:
            sku_duplicados_base = duplicates_base["SKU_HIJO"].nunique()
            print(f"⚠ Advertencia: {len(duplicates_base)} registros duplicados encontrados en BASE ({sku_duplicados_base} SKU_HIJO únicos)")
        else:
            print("✓ No hay duplicados en BASE")
        
        if not duplicates_final.empty:
            sku_duplicados_final = duplicates_final["SKU_HIJO"].nunique()
            print(f"⚠ Advertencia: {len(duplicates_final)} registros duplicados encontrados en FINAL ({sku_duplicados_final} SKU_HIJO únicos)")
        else:
            print("✓ No hay duplicados en FINAL")
        
        duplicates_info = {
            "base_duplicates": duplicates_base,
            "final_duplicates": duplicates_final
        }
        
        # 5. Comparar datos (solo registros válidos)
        print("\nComparando datos (solo registros con SKU_HIJO válido)...")
        df_nuevos, df_eliminados, df_modificados = compare_data(df_base_validos, df_final_validos, "SKU_HIJO")
        print(f"✓ Comparación completada:")
        print(f"  - NUEVOS: {len(df_nuevos)}")
        print(f"  - ELIMINADOS: {len(df_eliminados)}")
        print(f"  - MODIFICADOS: {df_modificados['SKU_HIJO'].nunique() if not df_modificados.empty else 0} SKU únicos")
        print(f"  - Diferencias de columna: {len(df_modificados)}")
        
        # 6. Generar archivos de salida
        print("\nGenerando archivos de salida...")
        
        excel_path = generate_excel_output(
            df_nuevos, df_eliminados, df_modificados, timestamp, output_dir
        )
        print(f"✓ Excel generado: {excel_path}")
        
        report_path = generate_report(
            df_base_validos, df_final_validos, df_nuevos, df_eliminados, df_modificados,
            duplicates_info, timestamp, output_dir, errors, nan_base, nan_final
        )
        print(f"✓ Reporte generado: {report_path}")
        
        # 7. Guardar en historial si se solicita
        if args.save:
            save_to_history(
                base_file, final_file, excel_path, report_path, timestamp, base_dir
            )
        
        # 8. Resumen final
        print("\n" + "=" * 80)
        print("RESUMEN FINAL")
        print("=" * 80)
        print(f"Total de filas en BASE: {len(df_base)} (válidos: {len(df_base_validos)}, NaN: {nan_base})")
        print(f"Total de filas en FINAL: {len(df_final)} (válidos: {len(df_final_validos)}, NaN: {nan_final})")
        print(f"Registros NUEVOS: {len(df_nuevos)}")
        print(f"Registros ELIMINADOS: {len(df_eliminados)}")
        print(f"Registros MODIFICADOS: {df_modificados['SKU_HIJO'].nunique() if not df_modificados.empty else 0}")
        print(f"Diferencias de columna: {len(df_modificados)}")
        print("\n✓ Proceso completado exitosamente")
        print("=" * 80)
        
    except KeyboardInterrupt:
        print("\n\nProceso interrumpido por el usuario.")
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ Error inesperado: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

