#!/bin/bash
# Script de build optimizado para Render
# Fuerza el uso de wheels precompilados cuando sea posible

set -e

echo "🔧 Instalando dependencias..."

# Instalar pip, setuptools y wheel actualizados primero
pip install --upgrade pip
pip install --upgrade setuptools wheel build

# Instalar dependencias (sin forzar only-binary para evitar errores)
pip install -r requirements.txt

echo "✅ Dependencias instaladas"

