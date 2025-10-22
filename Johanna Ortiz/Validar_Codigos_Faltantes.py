def comparar_archivos_txt(archivo1, archivo2, salida):
    # Leer los códigos del primer archivo
    with open(archivo1, 'r', encoding='utf-8') as f1:
        codigos1 = {line.strip() for line in f1 if line.strip()}
    
    # Leer los códigos del segundo archivo
    with open(archivo2, 'r', encoding='utf-8') as f2:
        codigos2 = {line.strip() for line in f2 if line.strip()}
    
    # Encontrar los que están en ambos archivos
    codigos_comunes = sorted(codigos1.intersection(codigos2))
    
    # Guardar el resultado en un nuevo archivo
    with open(salida, 'w', encoding='utf-8') as f_out:
        for codigo in codigos_comunes:
            f_out.write(codigo + '\n')
    
    print(f"✅ Se encontraron {len(codigos_comunes)} coincidencias.")
    print(f"Archivo generado: {salida}")

# Ejemplo de uso
comparar_archivos_txt('lista1.txt', 'lista2.txt', 'coincidencias.txt')
