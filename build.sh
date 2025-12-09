#!/bin/bash
# Script de build optimizado para Render
# Fuerza el uso de wheels precompilados cuando sea posible

set -e

echo "🔧 Instalando dependencias..."

# Instalar pip, setuptools y wheel actualizados primero
pip install --upgrade pip setuptools wheel

# Instalar dependencias con preferencia por wheels
pip install --only-binary :all: --prefer-binary -r requirements.txt || pip install -r requirements.txt

echo "✅ Dependencias instaladas"

