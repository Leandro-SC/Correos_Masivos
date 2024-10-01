import tkinter as tk
from tkinter import filedialog, messagebox



def crearInterfazEnvio(fun_enviar_correo, fun_descargar_estado):
    # Crear la interfaz gráfica
    root = tk.Tk()
    root.title("Aplicación para Enviar Correos")
    root.geometry("500x500")

    # Campo para ingresar el correo remitente
    tk.Label(root, text="Correo Remitente:").pack(pady=5)
    entrada_remitente = tk.Entry(root, width=50)
    entrada_remitente.pack(pady=5)

    # Campo para ingresar la contraseña
    tk.Label(root, text="Contraseña:").pack(pady=5)
    entrada_contraseña = tk.Entry(root, show="*", width=50)
    entrada_contraseña.pack(pady=5)

    # Campo para redactar el cuerpo del correo
    tk.Label(root, text="Cuerpo del correo (usa [nombre_estudiante] para personalizar):").pack(pady=10)
    texto_cuerpo = tk.Text(root, height=10, width=50)
    texto_cuerpo.pack(pady=10)

    # Botón para enviar correos
    btn_enviar = tk.Button(root, text="Cargar Excel y Enviar Correos", command=fun_enviar_correo)
    btn_enviar.pack(pady=20)

    # Botón para descargar el estado de los envíos (inicialmente deshabilitado)
    btn_descargar = tk.Button(root, text="Descargar Estado de Envíos", state=tk.DISABLED, command=fun_descargar_estado)
    btn_descargar.pack(pady=10)

    # Ejecutar la aplicación
    root.mainloop()

if __name__ == "__main__":
    crearInterfazEnvio("enviando...", "descargando...")
    




























