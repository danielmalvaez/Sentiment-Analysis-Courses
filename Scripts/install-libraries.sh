#!/bin/bash

# Definir las bibliotecas necesarias
libraries=("numpy" "pandas" "seaborn" "matplotlib" "deep-translator" "nltk" "unidecode" "reportlab" "datetime")

# Verificar si Python está instalado
if command -v python3 &>/dev/null; then
    for library in "${libraries[@]}"; do
        # Verificar si la biblioteca ya está instalada
        if python3 -c "import $library" 2>/dev/null; then
            echo "$library ya está instalada."
        else
            echo "Instalando $library..."
            pip install $library
        fi
    done
else
    echo "Python 3 no está instalado en este sistema."
fi