import sys
import tempfile
import os
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import PyPDF2
import openpyxl

# Autenticación y acceso a Google Drive con almacenamiento de sesión
def authenticate_drive():
    gauth = GoogleAuth()

    # Intentar cargar las credenciales guardadas para evitar autenticaciones repetitivas
    gauth.LoadCredentialsFile("mycreds.txt")
    
    if gauth.credentials is None:
        # Si no existen credenciales guardadas, realizar la autenticación
        gauth.LocalWebserverAuth()
        # Guardar las credenciales para futuros usos
        gauth.SaveCredentialsFile("mycreds.txt")
    elif gauth.access_token_expired:
        # Si el token ha expirado, refrescar la sesión
        gauth.Refresh()
    else:
        # Si las credenciales son válidas, cargarlas
        gauth.Authorize()
    
    drive = GoogleDrive(gauth)
    return drive

# Descargar el archivo PDF y extraer el nombre del estudiante
def extract_student_name(pdf_file):
    # Crear un archivo temporal local para el PDF
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        pdf_file.GetContentFile(temp_pdf.name)  # Descargar el archivo PDF localmente

        # Abrir y leer el archivo PDF con PyPDF2
        with open(temp_pdf.name, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            text = ""

            # Extraer el texto de todas las páginas del PDF
            for page in reader.pages:
                text += page.extract_text()

        # Filtrar el texto para obtener el nombre del estudiante
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        
        # Devolver la primera línea que no esté vacía como el nombre del estudiante
        student_name = lines[0] if lines else "Nombre desconocido"

    # Eliminar el archivo temporal
    os.remove(temp_pdf.name)

    return student_name

# Renombrar y mover el archivo a la carpeta de destino en Google Drive
def rename_and_move_files(pdf_files, target_folder_id, drive):
    student_file_pairs = []

    for pdf_file in pdf_files:
        # Extraer el nombre del estudiante directamente del archivo PDF
        student_name = extract_student_name(pdf_file)
        new_file_name = f"{student_name}.pdf"  # Nombre del archivo basado en el nombre del estudiante

        # Renombrar el archivo en Google Drive
        pdf_file['title'] = new_file_name
        pdf_file.Upload()  # Subir el archivo renombrado

        # Obtener el ID del padre actual (carpeta original)
        current_parents = pdf_file['parents'][0]['id']
        
        # Mover el archivo a la nueva carpeta (quitar el padre actual y agregar el nuevo)
        drive.auth.service.files().update(
            fileId=pdf_file['id'],
            addParents=target_folder_id,
            removeParents=current_parents,
            fields='id, parents'
        ).execute()

        # Guardar la información del nombre del estudiante y el archivo
        student_file_pairs.append((student_name, new_file_name))

    return student_file_pairs

# Crear archivo Excel con el nombre del estudiante y nombre del archivo
def create_excel(student_file_pairs, excel_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(['Nombre del Estudiante', 'Archivo PDF'])
    
    for student, file in student_file_pairs:
        sheet.append([student, file])
    
    workbook.save(excel_file)
    print(f'Archivo Excel guardado como {excel_file}')

def main():
    drive = authenticate_drive()

    # IDs de las carpetas en Google Drive
    source_folder_id = '1a0MZOUtS9ZH7ud1GxsEdtS2Xstkrkd67'  # Carpeta origen
    target_folder_id = '1Ha4lyJqM3NSP1cToj2rYk_S9dGF2HaCC'  # Carpeta destino

    # Descargar archivos PDF desde la carpeta origen
    pdf_files = drive.ListFile({'q': f"'{source_folder_id}' in parents and mimeType='application/pdf'"}).GetList()

    if not pdf_files:
        print("No se encontraron archivos PDF en la carpeta de origen.")
        sys.exit()

    # Renombrar y mover archivos
    student_file_pairs = rename_and_move_files(pdf_files, target_folder_id, drive)

    # Crear archivo Excel con los nombres de los estudiantes
    excel_file = 'reporte_estudiantes.xlsx'
    create_excel(student_file_pairs, excel_file)

    # Finalizar el proceso
    sys.exit()

if __name__ == '__main__':
    main()
