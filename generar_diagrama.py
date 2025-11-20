"""
Generador de Diagrama de Flujo - Calculador de Balance General
Usa Graphviz para crear una imagen del diagrama de flujo

Instalación requerida:
pip install graphviz

También necesitas tener Graphviz instalado en tu sistema:
- Windows: https://graphviz.org/download/
- Después de instalar, agrega Graphviz a tu PATH
"""

from graphviz import Digraph

def create_flowchart():
    # Crear diagrama con orientación vertical
    dot = Digraph(comment='Calculador de Balance General', format='png')
    dot.attr(rankdir='TB', size='12,16')
    dot.attr('node', shape='box', style='rounded,filled', fillcolor='lightblue', fontname='Arial')
    
    # Nodos de inicio y fin
    dot.node('start', 'INICIO DE APLICACIÓN', shape='ellipse', fillcolor='lightgreen')
    dot.node('end', 'FIN', shape='ellipse', fillcolor='lightcoral')
    
    # Sistema de Licencias
    dot.node('check_license', '¿Licencia válida\nalmacenada?', shape='diamond', fillcolor='lightyellow')
    dot.node('show_license_dialog', 'Mostrar diálogo\nde activación')
    dot.node('enter_license', 'Usuario ingresa\ncódigo de licencia')
    dot.node('validate_license', '¿Licencia\nválida?', shape='diamond', fillcolor='lightyellow')
    dot.node('show_invalid', 'Mostrar error:\nLicencia inválida', fillcolor='lightcoral')
    dot.node('save_license', 'Guardar licencia\nen localStorage', fillcolor='lightgreen')
    
    # Carga de aplicación
    dot.node('load_app', 'Cargar Interfaz\nWeb/Electron')
    dot.node('load_libs', 'Cargar Librerías:\nXLSX, Firebase')
    dot.node('show_ui', 'Mostrar Interfaz\nPrincipal')
    
    # Carga de archivos
    dot.node('wait_input', '¿Usuario selecciona\narchivos Excel?', shape='diamond', fillcolor='lightyellow')
    dot.node('diary_file', 'Cargar Libro Diario')
    dot.node('ledger_file', 'Cargar Libro Mayor')
    dot.node('check_files', '¿Ambos archivos\ncargados?', shape='diamond', fillcolor='lightyellow')
    dot.node('enable_btn', 'Habilitar botón\nGenerar Balance', fillcolor='lightgreen')
    
    # Procesamiento
    dot.node('wait_compute', 'Usuario presiona\nGenerar Balance', shape='diamond', fillcolor='lightyellow')
    dot.node('read_diary', 'Leer archivo Diario\ncon XLSX.read')
    dot.node('read_ledger', 'Leer archivo Mayor\ncon XLSX.read')
    dot.node('parse_diary', 'Parsear JSON\ndel Diario')
    dot.node('parse_ledger', 'Parsear JSON\ndel Mayor')
    dot.node('normalize', 'Normalizar nombres\nde columnas')
    dot.node('extract', 'Extraer datos:\nCuenta, Debe, Haber')
    dot.node('validate_format', '¿Formatos\nválidos?', shape='diamond', fillcolor='lightyellow')
    dot.node('show_error', 'Mostrar mensaje\nde error', fillcolor='lightcoral')
    
    # Cálculos
    dot.node('merge', 'Agrupar cuentas\npor nombre')
    dot.node('sum_debits', 'Sumar Débitos\npor cuenta')
    dot.node('sum_credits', 'Sumar Créditos\npor cuenta')
    dot.node('calc_balance', 'Calcular Saldo:\nDebe - Haber')
    dot.node('calc_totals', 'Calcular totales:\nTotal Debe y Haber', fillcolor='lightgreen')
    
    # Validación de balance cuadrado
    dot.node('check_balanced', '¿Total Debe =\nTotal Haber?', shape='diamond', fillcolor='lightyellow')
    dot.node('show_warning', '⚠️ Advertencia:\nBalance descuadrado', fillcolor='orange')
    dot.node('show_balanced', '✅ Balance\ncuadrado', fillcolor='lightgreen')
    
    # Renderizado
    dot.node('sort', 'Ordenar cuentas\nalfabéticamente')
    dot.node('render', 'Renderizar tabla\nHTML')
    dot.node('show_options', '¿Opciones\nhabilitadas?', shape='diamond', fillcolor='lightyellow')
    dot.node('apply_currency', 'Aplicar formato\nde moneda')
    dot.node('show_totals', 'Mostrar fila\nde totales')
    dot.node('enable_export', 'Habilitar botones:\nExportar y Guardar', fillcolor='lightgreen')
    
    # Acciones del usuario
    dot.node('user_action', '¿Qué acción elige\nel usuario?', shape='diamond', fillcolor='lightyellow')
    
    # Exportar Excel
    dot.node('check_platform', '¿Plataforma?', shape='diamond', fillcolor='lightyellow')
    dot.node('show_save_dialog', 'Mostrar diálogo\nde guardado')
    dot.node('select_path', '¿Usuario selecciona\nubicación?', shape='diamond', fillcolor='lightyellow')
    dot.node('save_excel', 'Guardar archivo Excel')
    dot.node('download_excel', 'Descargar archivo\nbalance_general.xlsx')
    dot.node('prepare_export', 'Preparar datos:\nTraducir headers')
    dot.node('format_columns', 'Configurar anchos\nde columnas')
    dot.node('apply_number_format', 'Aplicar formato\nnumérico #,##0.00')
    dot.node('write_file', 'Escribir archivo XLSX')
    dot.node('show_success', 'Mostrar mensaje\nde éxito', fillcolor='lightgreen')
    
    # Firebase
    dot.node('check_firebase', '¿Firebase\nconfigurado?', shape='diamond', fillcolor='lightyellow')
    dot.node('show_firebase_error', 'Mostrar error:\nFirebase no configurado', fillcolor='lightcoral')
    dot.node('init_firebase', 'Inicializar Firebase')
    dot.node('create_doc', 'Crear documento:\ntimestamp + balance')
    dot.node('save_firestore', 'Guardar en Firestore')
    dot.node('show_firebase_success', 'Mostrar ID del\ndocumento', fillcolor='lightgreen')
    
    # Conexiones - Sistema de Licencias
    dot.edge('start', 'check_license')
    dot.edge('check_license', 'show_license_dialog', label='No')
    dot.edge('check_license', 'load_app', label='Sí')
    dot.edge('show_license_dialog', 'enter_license')
    dot.edge('enter_license', 'validate_license')
    dot.edge('validate_license', 'show_invalid', label='No')
    dot.edge('show_invalid', 'show_license_dialog')
    dot.edge('validate_license', 'save_license', label='Sí')
    dot.edge('save_license', 'load_app')
    
    # Conexiones - Carga
    dot.edge('load_app', 'load_libs')
    dot.edge('load_libs', 'show_ui')
    dot.edge('show_ui', 'wait_input')
    dot.edge('wait_input', 'diary_file', label='Diario')
    dot.edge('wait_input', 'ledger_file', label='Mayor')
    dot.edge('diary_file', 'check_files')
    dot.edge('ledger_file', 'check_files')
    dot.edge('check_files', 'wait_input', label='No')
    dot.edge('check_files', 'enable_btn', label='Sí')
    dot.edge('enable_btn', 'wait_compute')
    
    # Conexiones - Procesamiento
    dot.edge('wait_compute', 'read_diary')
    dot.edge('read_diary', 'read_ledger')
    dot.edge('read_ledger', 'parse_diary')
    dot.edge('parse_diary', 'parse_ledger')
    dot.edge('parse_ledger', 'normalize')
    dot.edge('normalize', 'extract')
    dot.edge('extract', 'validate_format')
    dot.edge('validate_format', 'show_error', label='No')
    dot.edge('show_error', 'wait_input')
    dot.edge('validate_format', 'merge', label='Sí')
    
    # Conexiones - Cálculos
    dot.edge('merge', 'sum_debits')
    dot.edge('sum_debits', 'sum_credits')
    dot.edge('sum_credits', 'calc_balance')
    dot.edge('calc_balance', 'calc_totals')
    dot.edge('calc_totals', 'check_balanced')
    dot.edge('check_balanced', 'show_warning', label='No')
    dot.edge('check_balanced', 'show_balanced', label='Sí')
    dot.edge('show_warning', 'sort')
    dot.edge('show_balanced', 'sort')
    
    # Conexiones - Renderizado
    dot.edge('sort', 'render')
    dot.edge('render', 'show_options')
    dot.edge('show_options', 'apply_currency', label='Moneda')
    dot.edge('show_options', 'show_totals', label='Totales')
    dot.edge('show_options', 'enable_export', label='Ninguna')
    dot.edge('apply_currency', 'enable_export')
    dot.edge('show_totals', 'enable_export')
    dot.edge('enable_export', 'user_action')
    
    # Conexiones - Exportar
    dot.edge('user_action', 'check_platform', label='Exportar')
    dot.edge('check_platform', 'show_save_dialog', label='Electron')
    dot.edge('check_platform', 'download_excel', label='Navegador')
    dot.edge('show_save_dialog', 'select_path')
    dot.edge('select_path', 'save_excel', label='Sí')
    dot.edge('select_path', 'enable_export', label='No')
    dot.edge('save_excel', 'prepare_export')
    dot.edge('download_excel', 'prepare_export')
    dot.edge('prepare_export', 'format_columns')
    dot.edge('format_columns', 'apply_number_format')
    dot.edge('apply_number_format', 'write_file')
    dot.edge('write_file', 'show_success')
    dot.edge('show_success', 'enable_export')
    
    # Conexiones - Firebase
    dot.edge('user_action', 'check_firebase', label='Firebase')
    dot.edge('check_firebase', 'show_firebase_error', label='No')
    dot.edge('show_firebase_error', 'enable_export')
    dot.edge('check_firebase', 'init_firebase', label='Sí')
    dot.edge('init_firebase', 'create_doc')
    dot.edge('create_doc', 'save_firestore')
    dot.edge('save_firestore', 'show_firebase_success')
    dot.edge('show_firebase_success', 'enable_export')
    
    # Conexiones - Otras acciones
    dot.edge('user_action', 'wait_input', label='Nuevo')
    dot.edge('user_action', 'end', label='Cerrar')
    
    return dot

if __name__ == '__main__':
    # Crear el diagrama
    flowchart = create_flowchart()
    
    # Guardar como archivo
    output_file = 'diagrama_flujo_balance_general'
    
    # Renderizar y guardar
    flowchart.render(output_file, view=True, cleanup=True)
    
    print(f"✅ Diagrama generado exitosamente: {output_file}.png")
    print(f"📁 También se guardó el código fuente: {output_file}")
