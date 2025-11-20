"""
Generador de Códigos de Licencia
Genera códigos únicos con checksum para el Calculador de Balance General
"""

import random
import string
import hashlib

def generate_license_key():
    """
    Genera un código de licencia único con formato: XXXX-XXXX-XXXX-XXXX
    Los primeros 3 bloques son aleatorios, el último es un checksum
    """
    # Generar 3 bloques aleatorios
    blocks = []
    for i in range(3):
        block = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
        blocks.append(block)
    
    # Calcular checksum para el 4to bloque
    data = ''.join(blocks)
    checksum = hashlib.md5(data.encode()).hexdigest()[:4].upper()
    blocks.append(checksum)
    
    return '-'.join(blocks)

def validate_license_key(key):
    """
    Valida que un código de licencia sea correcto
    """
    if not key or len(key) != 19:
        return False
    
    parts = key.split('-')
    if len(parts) != 4:
        return False
    
    # Verificar que cada parte tenga 4 caracteres
    for part in parts:
        if len(part) != 4:
            return False
    
    # Verificar checksum
    data = ''.join(parts[:3])
    expected_checksum = hashlib.md5(data.encode()).hexdigest()[:4].upper()
    
    return parts[3] == expected_checksum

if __name__ == '__main__':
    print("=" * 60)
    print("GENERADOR DE LICENCIAS - CALCULADOR DE BALANCE GENERAL")
    print("=" * 60)
    print()
    
    # Generar 10 licencias válidas
    print("📋 CÓDIGOS DE LICENCIA VÁLIDOS:")
    print("-" * 60)
    
    licenses = []
    for i in range(10):
        license_key = generate_license_key()
        licenses.append(license_key)
        print(f"{i+1:2d}. {license_key}")
    
    print()
    print("=" * 60)
    print("✅ VERIFICACIÓN DE CÓDIGOS")
    print("=" * 60)
    
    # Verificar que todos sean válidos
    for i, key in enumerate(licenses, 1):
        is_valid = validate_license_key(key)
        status = "✅ VÁLIDO" if is_valid else "❌ INVÁLIDO"
        print(f"{i:2d}. {key} - {status}")
    
    print()
    print("=" * 60)
    print("📝 INSTRUCCIONES")
    print("=" * 60)
    print()
    print("1. Copia cualquiera de los códigos de arriba")
    print("2. Abre la aplicación (web o Electron)")
    print("3. Aparecerá un modal pidiendo la licencia")
    print("4. Pega el código y presiona 'Activar'")
    print("5. La licencia se guardará automáticamente")
    print()
    print("💾 La licencia se guarda en localStorage y persiste entre sesiones")
    print()
    
    # Guardar licencias en archivo
    with open('LICENCIAS_VALIDAS.txt', 'w', encoding='utf-8') as f:
        f.write("CÓDIGOS DE LICENCIA VÁLIDOS\n")
        f.write("Calculador de Balance General\n")
        f.write("=" * 60 + "\n\n")
        for i, key in enumerate(licenses, 1):
            f.write(f"{i:2d}. {key}\n")
        f.write("\n" + "=" * 60 + "\n")
        f.write("Usa cualquiera de estos códigos para activar la aplicación.\n")
        f.write("La licencia se guardará automáticamente en localStorage.\n")
    
    print("✅ Licencias guardadas en: LICENCIAS_VALIDAS.txt")
    print()
