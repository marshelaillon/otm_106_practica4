import pandas as pd
import matplotlib.pyplot as plt
import os

# Crear carpeta output si no existe
if not os.path.exists('output'):
    os.makedirs('output')

# ========================================
# PARTE 1: GRÁFICA DE BARRAS CON DATOS ESTÁTICOS
# ========================================
print("Generando gráfica de ejemplo con datos estáticos...")

# Datos de ejemplo
categorias = ['Producto A', 'Producto B', 'Producto C', 'Producto D']
valores = [45, 30, 25, 50]

# Crear gráfica de barras
plt.figure(figsize=(10, 6))
plt.bar(categorias, valores, color='skyblue', edgecolor='navy')
plt.xlabel('Productos')
plt.ylabel('Valores')
plt.title('Gráfica de Barras - Datos de Ejemplo Estáticos')
plt.grid(axis='y', alpha=0.3)

# Guardar la gráfica
plt.savefig('output/grafica_ejemplo_estatica.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Gráfica estática guardada: output/grafica_ejemplo_estatica.png\n")

# ========================================
# PARTE 2: LEER ARCHIVO EXCEL Y GENERAR GRÁFICAS
# ========================================

# Ruta del archivo Excel
archivo_excel = 'datos/datos_base.xlsx'

try:
    # Obtener los nombres de todas las hojas del archivo
    xls = pd.ExcelFile(archivo_excel)
    hojas = xls.sheet_names
    
    print(f"Archivo encontrado: {archivo_excel}")
    print(f"Hojas detectadas: {hojas}\n")
    
    # Iterar sobre cada hoja
    for hoja in hojas:
        print(f"Procesando hoja: {hoja}")
        
        # Leer datos de la hoja actual
        df = pd.read_excel(archivo_excel, sheet_name=hoja)
        
        print(f"  - Datos leídos: {len(df)} filas")
        print(f"  - Columnas: {list(df.columns)}\n")
        
        # Verificar que existan las columnas necesarias
        # Ajusta estos nombres según las columnas reales de tu archivo
        columna_categoria = df.columns[0]  # Primera columna (categorías/etiquetas)
        columna_valores = 'Valores' if 'Valores' in df.columns else df.columns[1]
        columna_porcentaje = 'Porcentaje' if 'Porcentaje' in df.columns else df.columns[2]
        
        # ========================================
        # GRÁFICA DE BARRAS con columna Valores
        # ========================================
        plt.figure(figsize=(10, 6))
        plt.bar(df[columna_categoria], df[columna_valores], 
                color='steelblue', edgecolor='black', alpha=0.8)
        plt.xlabel(columna_categoria, fontsize=12)
        plt.ylabel(columna_valores, fontsize=12)
        plt.title(f'Gráfica de Barras - {hoja}', fontsize=14, fontweight='bold')
        plt.xticks(rotation=45, ha='right')
        plt.grid(axis='y', alpha=0.3)
        plt.tight_layout()
        
        # Guardar gráfica de barras
        nombre_archivo_barras = f'output/barras_{hoja.replace(" ", "_")}.png'
        plt.savefig(nombre_archivo_barras, dpi=300, bbox_inches='tight')
        plt.close()
        print(f"  ✓ Gráfica de barras guardada: {nombre_archivo_barras}")
        
        # ========================================
        # GRÁFICA DE PASTEL con columna Porcentaje
        # ========================================
        plt.figure(figsize=(10, 8))
        
        # Crear gráfica de pastel
        colores = plt.cm.Set3.colors
        explode = [0.05] * len(df)  # Separar ligeramente cada porción
        
        plt.pie(df[columna_porcentaje], 
                labels=df[columna_categoria],
                autopct='%1.1f%%',
                startangle=90,
                colors=colores,
                explode=explode,
                shadow=True)
        
        plt.title(f'Gráfica de Pastel - {hoja}', fontsize=14, fontweight='bold')
        plt.axis('equal')
        
        # Guardar gráfica de pastel
        nombre_archivo_pastel = f'output/pastel_{hoja.replace(" ", "_")}.png'
        plt.savefig(nombre_archivo_pastel, dpi=300, bbox_inches='tight')
        plt.close()
        print(f"  ✓ Gráfica de pastel guardada: {nombre_archivo_pastel}\n")
    
    print("="*60)
    print("✓ Proceso completado exitosamente")
    print(f"✓ Total de hojas procesadas: {len(hojas)}")
    print(f"✓ Gráficas generadas en la carpeta 'output/'")
    print("="*60)

except FileNotFoundError:
    print(f"ERROR: No se encontró el archivo '{archivo_excel}'")
    print("Por favor, asegúrate de:")
    print("  1. Tener el archivo 'datos_base.xlsx' en la carpeta 'datos/'")
    print("  2. El nombre del archivo sea exactamente 'datos_base.xlsx'")
    
except Exception as e:
    print(f"ERROR: Ocurrió un error al procesar el archivo")
    print(f"Detalle: {str(e)}")