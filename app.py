def pedir_nombre():
    """Solicita al usuario que introduzca su nombre y lo devuelve."""
    nombre = input("Introduce tu nombre: ")
    return nombre

if __name__ == "__main__":
    nombre_usuario = pedir_nombre()
    print(f"¡Hola, {nombre_usuario}!")