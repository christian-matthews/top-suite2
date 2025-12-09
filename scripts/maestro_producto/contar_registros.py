#!/usr/bin/env python3
"""
Script para contar registros en la columna A de la hoja BASE GS1 (2)
"""

import pandas as pd

def contar_registros_columna_a():
    """
    Cuenta los registros en la columna A de la hoja BASE GS1 (2)
    """
    archivo_excel = "/Users/christianmatthews/Library/Mobile Documents/com~apple~CloudDocs/CURSOR/TOP/MAESTRO PRODUCTO/Test1.xlsx"
    
    try:
        # Leer la hoja BASE GS1 (2)
        print("📊 Leyendo hoja 'BASE GS1 (2)'...")
        df = pd.read_excel(archivo_excel, sheet_name="BASE GS1 (2)")
        
        # Obtener información de la columna A (primera columna)
        columna_a = df.iloc[:, 0]  # Primera columna (índice 0)
        nombre_columna_a = df.columns[0]
        
        print(f"\n📋 Información de la Columna A:")
        print(f"   🏷️  Nombre: '{nombre_columna_a}'")
        print(f"   📏 Total de filas en el DataFrame: {len(df)}")
        print(f"   📊 Registros totales en columna A: {len(columna_a)}")
        print(f"   ✅ Registros no vacíos: {columna_a.notna().sum()}")
        print(f"   ❌ Registros vacíos (NaN): {columna_a.isna().sum()}")
        print(f"   🔍 Registros únicos: {columna_a.nunique()}")
        
        print(f"\n📈 Estadísticas adicionales:")
        print(f"   📍 Primer valor: {columna_a.iloc[0] if len(columna_a) > 0 else 'N/A'}")
        print(f"   📍 Último valor: {columna_a.iloc[-1] if len(columna_a) > 0 else 'N/A'}")
        
        # Mostrar algunos valores de ejemplo
        print(f"\n🔍 Primeros 10 valores:")
        for i, valor in enumerate(columna_a.head(10)):
            print(f"   Fila {i+1}: {valor}")
        
        # Mostrar valores únicos si no son demasiados
        valores_unicos = columna_a.unique()
        if len(valores_unicos) <= 20:
            print(f"\n🎯 Valores únicos encontrados ({len(valores_unicos)}):")
            for i, valor in enumerate(valores_unicos, 1):
                print(f"   {i}. {valor}")
        else:
            print(f"\n🎯 Hay {len(valores_unicos)} valores únicos (demasiados para mostrar todos)")
            print("   Primeros 10 valores únicos:")
            for i, valor in enumerate(valores_unicos[:10], 1):
                print(f"   {i}. {valor}")
        
        return len(columna_a), columna_a.notna().sum()
        
    except Exception as e:
        print(f"❌ Error al leer el archivo: {e}")
        return None, None

if __name__ == "__main__":
    print("=" * 60)
    print("🔍 CONTADOR DE REGISTROS - COLUMNA A")
    print("=" * 60)
    
    total, no_vacios = contar_registros_columna_a()
    
    if total is not None:
        print(f"\n🎉 RESUMEN:")
        print(f"   📊 Total de registros en columna A: {total:,}")
        print(f"   ✅ Registros con datos: {no_vacios:,}")
        print(f"   ❌ Registros vacíos: {(total - no_vacios):,}")
    else:
        print("\n❌ No se pudo contar los registros")


