# Instrucciones para Generar el Diagrama de Flujo

## Requisitos

### 1. Instalar Graphviz en tu sistema

**Windows:**
1. Descarga Graphviz desde: https://graphviz.org/download/
2. Ejecuta el instalador
3. Durante la instalación, marca la opción "Add Graphviz to the system PATH"
4. Si no lo hiciste, agrega manualmente a PATH:
   - Busca "Variables de entorno" en Windows
   - Edita la variable PATH
   - Agrega: `C:\Program Files\Graphviz\bin`

### 2. Instalar la librería Python

```bash
pip install graphviz
```

## Uso

### Generar el diagrama

```bash
python generar_diagrama.py
```

Esto creará:
- `diagrama_flujo_balance_general.png` - La imagen del diagrama
- El diagrama se abrirá automáticamente después de generarse

## Personalización

Puedes modificar el script `generar_diagrama.py` para:

### Cambiar el formato de salida
```python
# En la línea donde se crea el Digraph
dot = Digraph(comment='...', format='pdf')  # Opciones: png, pdf, svg, jpg
```

### Cambiar colores
```python
# Modificar los fillcolor en los nodos
dot.node('nombre', 'Texto', fillcolor='lightblue')  # Cambiar color
```

### Cambiar orientación
```python
# En dot.attr
dot.attr(rankdir='LR')  # LR = Left to Right (horizontal)
dot.attr(rankdir='TB')  # TB = Top to Bottom (vertical)
```

### Cambiar tamaño
```python
# En dot.attr
dot.attr(size='16,20')  # ancho,alto en pulgadas
```

## Colores Disponibles

- `lightblue` - Azul claro (procesos normales)
- `lightgreen` - Verde claro (éxito, inicio)
- `lightyellow` - Amarillo claro (decisiones)
- `lightcoral` - Coral claro (errores, fin)
- `orange` - Naranja (advertencias)

## Solución de Problemas

### Error: "graphviz executables not found"
- Asegúrate de que Graphviz esté instalado en tu sistema
- Verifica que esté en el PATH
- Reinicia tu terminal/IDE después de instalar

### Error: "No module named 'graphviz'"
```bash
pip install graphviz
```

### La imagen no se abre automáticamente
- Busca el archivo `diagrama_flujo_balance_general.png` en la carpeta del proyecto
- Ábrelo manualmente

## Exportar a otros formatos

```python
# PDF
flowchart.format = 'pdf'
flowchart.render('diagrama_flujo_balance_general', view=True)

# SVG (vectorial, escalable)
flowchart.format = 'svg'
flowchart.render('diagrama_flujo_balance_general', view=True)

# JPG
flowchart.format = 'jpg'
flowchart.render('diagrama_flujo_balance_general', view=True)
```
