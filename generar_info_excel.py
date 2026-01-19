import pandas as pd
import os

def generate_excel():
    output_file = "Guia_Informativa_Proyecto.xlsx"
    
    # 1. OBJETIVOS Y ALCANCE
    df_objetivos = pd.DataFrame({
        "Sección": ["Planteamiento", "Objetivo General", "Alcance", "Ciclo de Vida"],
        "Descripción": [
            "La enseñanza tradicional requiere validación manual exhaustiva, generando frustración y demora.",
            "Optimizar la validación de actividades contables estudiantiles mediante la automatización digital de balances financieros básicos.",
            "Abarca desde la lectura de libros hasta la generación de Estados Financieros para fines educativos.",
            "Relevamiento -> Diseño -> Implementación -> Despliegue/Retroalimentación"
        ]
    })

    # 2. DICCIONARIO DE VARIABLES
    df_variables = pd.DataFrame({
        "Variable": ["diaryRows", "ledgerRows", "trialBalance", "type", "incomeStatement", "isBalanced"],
        "Tipo": ["Array[Obj]", "Array[Obj]", "Array[Obj]", "String", "Object", "Boolean"],
        "Descripción": [
            "Filas normalizadas del Libro Diario.",
            "Filas normalizadas del Libro Mayor.",
            "Saldos procesados para el Balance de Comprobación.",
            "'REAL' (Balance Gral) o 'NOMINAL' (Resultados).",
            "Totales de ingresos, gastos y utilidad.",
            "Determina si Total Debe es igual al Total Haber."
        ]
    })

    # 3. MANUAL DE USUARIO
    df_manual = pd.DataFrame({
        "Paso": [1, 2, 3, 4, 5],
        "Acción": ["Preparación", "Validación Inicial", "Análisis", "Corrección", "Entrega Final"],
        "Instrucción": [
            "Asegúrate de que tu tarea de Libro Diario esté en un archivo Excel (.xlsx).",
            "Carga tu Libro Diario; el sistema validará la partida doble de inmediato.",
            "Revisa la sección de Clasificación para entender el destino de cada cuenta.",
            "Si ves un mensaje naranja, corrige tus sumas en Excel y vuelve a cargar.",
            "Al lograr el balance cuadrado (Verde), exporta tu trabajo profesionalmente."
        ]
    })

    # 4. CARACTERÍSTICAS TÉCNICAS
    df_tecnico = pd.DataFrame({
        "Categoría": ["Frontend", "Procesamiento", "Base de Datos", "Escritorio", "Fuentes"],
        "Tecnología": ["HTML5, CSS3, JS (ES6+)", "XLSX.js", "Firebase Firestore (Opcional)", "Electron", "Google Fonts (Inter)"]
    })

    # Crear el archivo Excel con múltiples hojas
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_objetivos.to_excel(writer, sheet_name='Objetivos y Alcance', index=False)
        df_variables.to_excel(writer, sheet_name='Diccionario Variables', index=False)
        df_manual.to_excel(writer, sheet_name='Manual de Usuario', index=False)
        df_tecnico.to_excel(writer, sheet_name='Stack Técnico', index=False)

    print(f"Archivo '{output_file}' generado exitosamente en {os.getcwd()}")

if __name__ == "__main__":
    generate_excel()
