import os
import smtplib
import pandas as pd
import tkinter as tk
import time
from tkinter import filedialog, messagebox
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage  # Importa MIMEImage
from email import encoders
import imgkit  # Importa imgkit para convertir HTML en imagen
import base64
import random


# Variable global para almacenar la ruta del archivo de estado
ruta_salida_global = None

ruta_wkhtmltoimage = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltoimage.exe" # Reemplaza con la ruta correcta en tu sistema
config = imgkit.config(wkhtmltoimage=ruta_wkhtmltoimage)


# Función para convertir HTML a imagen
def html_a_imagen(cuerpo_html, nombre_imagen):
    opciones = {
        'format': 'png',
        'encoding': 'UTF-8'
    }
    imgkit.from_string(cuerpo_html, nombre_imagen, options=opciones, config=config)

# Función para enviar correos personalizados con archivos adjuntos
def enviar_correo(destinatario, nombre_estudiante, archivo_adjuntar, correo_remitente, contraseña):
    try:
        # Crear el mensaje
        mensaje = MIMEMultipart('related')
        mensaje['From'] = correo_remitente
        mensaje['To'] = destinatario
        mensaje['Subject'] = f'Constancia "Chambea IESRP" 2024 - {nombre_estudiante}'

        # Cuerpo del mensaje con HTML y CSS en línea
        cuerpo_correo = f"""
        <html>
            <body style="font-family: Verdana, sans-serif; background-color: #f4f4f4; margin: 0; padding: 0;">
                <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f4f4f4; padding: 20px;">
                    <tr>
                        <td align="center">
                            <table width="600" cellpadding="0" cellspacing="0" style="background-color: #ffffff; padding: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);">
                                <tr>
                                    <td align="center" style="background-color: #001f60; color: white; padding: 10px; border-radius: 10px 10px 0 0;">
                                        <h2 style="text-transform: uppercase; color: white; margin: 0;">Constancia "Chambea IESRP" 2024</h2>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 20px; line-height: 1.6; color: #333333; text-align: justify;">
                                        <p>Estimado/a <strong>{nombre_estudiante}</strong>,</p>
                                        <p>
                                            Queremos agradecerte por tu valiosa participación en la Feria Laboral "Chambea IESRP" 2024. 
                                            Confiamos en que esta experiencia haya sido enriquecedora y haya contribuido a tu desarrollo profesional.

                                        </p>
                                        <p>
                                            Adjunto encontrarás tu constancia de participación. 
                                            Lamentamos la demora en la entrega y te aseguramos que estamos trabajando para mejorar continuamente nuestros procesos.
                                        </p>
                                        <p>
                                          Te animamos a seguir aprovechando las diversas oportunidades que IESRP pone a tu disposición para continuar fortaleciendo tu perfil profesional.

                                        </p>
                                        <p>Atentamente,</p>
                                        <p text-align: center;>
                                        <br/>
                                        <br/>
                                        <strong>Luz Ramos</strong><br/>
                                        <strong>Analista de Empleabilidad y Relaciones Empresariales</strong><br/>
                                        <strong>Instituto de Educación Superior Ricardo Palma</strong>
                                        </p>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="center" style="padding: 20px;">
                                        <img src="cid:logo_instituto" alt="logo instituto" style="max-width: 100%; height: auto;"/>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </body>
        </html>
        """

        # Adjuntar el cuerpo HTML al mensaje
        parte_html = MIMEText(cuerpo_correo, 'html')
        mensaje.attach(parte_html)

        # Adjuntar la imagen al correo y establecer un Content-ID
        with open('./img/logo_i.png', 'rb') as img_file:
            mime_image = MIMEImage(img_file.read())
            mime_image.add_header('Content-ID', '<logo_instituto>')  # El Content-ID debe estar entre ángulos
            mensaje.attach(mime_image)

        # Adjuntar archivo (certificado o documento personalizado)
        with open(archivo_adjuntar, 'rb') as adjunto:
            parte = MIMEBase('application', 'pdf')
            parte.set_payload(adjunto.read())
            encoders.encode_base64(parte)
            parte.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(archivo_adjuntar)}"')
            mensaje.attach(parte)

        # Configurar servidor SMTP de Zoho y enviar el correo
        # print("Configurando el servidor SMTP...")
        # servidor = smtplib.SMTP('smtp.zoho.com', 587)
        # servidor.starttls()
        # print("Iniciando sesión en el servidor SMTP...")
        # servidor.login(correo_remitente, contraseña)
        # texto = mensaje.as_string()
        # print(f"Enviando correo a {destinatario}...")
        # servidor.sendmail(correo_remitente, destinatario, texto)
        # servidor.quit()
        # print("Correo enviado con éxito.")
        # Configurar servidor SMTP de Gmail y enviar el correo
        print("Configurando el servidor SMTP...")
        servidor = smtplib.SMTP('smtp.gmail.com', 587)
        servidor.starttls()
        print("Iniciando sesión en el servidor SMTP...")
        servidor.login(correo_remitente, contraseña)  # Usa una contraseña de aplicación para mayor seguridad
        texto = mensaje.as_string()
        print(f"Enviando correo a {destinatario}...")
        servidor.sendmail(correo_remitente, destinatario, texto)
        servidor.quit()
        print("Correo enviado con éxito.")


        return True

    except smtplib.SMTPAuthenticationError:
        print(f'Error de autenticación. Verifica el correo y la contraseña de {correo_remitente}.')
        return False

    except Exception as e:
        print(f'Error al enviar correo a {destinatario}: {str(e)}')
        return False


# Función para cargar el archivo Excel
def cargar_excel():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo:
        try:
            df = pd.read_excel(archivo)
            return df, archivo
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar el archivo: {e}")
            return None, None
    return None, None

# Función para enviar correos desde la interfaz y generar el reporte
def enviar_correos():
    global ruta_salida_global
    df, archivo = cargar_excel()
    if df is not None:
        if 'nombre_estudiante' in df.columns and 'correo_estudiante' in df.columns and 'ruta_certificado_base' in df.columns and 'nombre_archivo' in df.columns:
            # Obteniendo el remitente y la contraseña desde la interfaz
            correo_remitente = entrada_remitente.get()
            contraseña = entrada_contraseña.get()

            if not correo_remitente or not contraseña:
                messagebox.showerror("Error", "Debe ingresar el correo remitente y la contraseña")
                return

            estados_envio = []

            for index, row in df.iterrows():
                nombre_estudiante = row['nombre_estudiante']
                correo_estudiante = row['correo_estudiante']
                ruta_certificado_base = row['ruta_certificado_base']
                nombre_archivo = row['nombre_archivo']
                
                # Generar el nombre del archivo personalizado
                archivo_personalizado = os.path.join(ruta_certificado_base, f"{nombre_archivo}.pdf")
                tiempo_random = random.randint(80, 100)
                time.sleep(tiempo_random)
                
                if not os.path.exists(archivo_personalizado):
                    messagebox.showerror("Error", f"No se encontró el archivo: {archivo_personalizado}")
                    estados_envio.append("No enviado")
                    continue

                # Enviar el correo y guardar el estado
                enviado = enviar_correo(correo_estudiante, nombre_estudiante, archivo_personalizado, correo_remitente, contraseña)
                estados_envio.append("Enviado" if enviado else "No enviado")
                
            # Agregar columna de estado de envío y guardar el archivo
            df['estado_envio'] = estados_envio
            ruta_salida_global = archivo.replace(".xlsx", "_resultado.xlsx")
            df.to_excel(ruta_salida_global, index=False)
            time.sleep(2)

            messagebox.showinfo("Éxito", f"Correos enviados correctamente. Resultado guardado en: {ruta_salida_global}")
            btn_descargar.config(state=tk.NORMAL)  # Habilitar el botón de descarga

        else:
            messagebox.showerror("Error", "El archivo debe contener las columnas: nombre_estudiante, correo_estudiante, ruta_certificado_base, nombre_archivo")
    else:
        messagebox.showerror("Error", "No se pudo cargar el archivo")

# Función para descargar el archivo de estado de los envíos
def descargar_estado():
    if ruta_salida_global:
        # Abrir el explorador para descargar el archivo generado
        filedialog.asksaveasfilename(initialfile=ruta_salida_global, defaultextension=".xlsx")
        messagebox.showinfo("Descarga", f"El archivo ha sido guardado en: {ruta_salida_global}")
    else:
        messagebox.showerror("Error", "No hay archivo disponible para descargar")

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

# Botón para enviar correos
btn_enviar = tk.Button(root, text="Cargar Excel y Enviar Correos", command=enviar_correos)
btn_enviar.pack(pady=20)

# Botón para descargar el estado de los envíos (inicialmente deshabilitado)
btn_descargar = tk.Button(root, text="Descargar Estado de Envíos", state=tk.DISABLED, command=descargar_estado)
btn_descargar.pack(pady=10)

# Ejecutar la aplicación
root.mainloop()








