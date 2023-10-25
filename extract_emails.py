import win32com.client
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Color
import re
from datetime import datetime, timedelta
from dateutil.parser import parse

def format_date_time(datetime_str):
    try:
        datetime_obj = parse(datetime_str)
        formatted_date_time = datetime_obj.strftime('%d/%m/%Y - %I:%M %p')
    except ValueError:
        formatted_date_time = "Invalid Date/Time Format"
    
    return formatted_date_time

def extract_and_save_parts(output_filename, sender_email):
    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Carpeta de la bandeja de entrada

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "EXCESO DE VELOCIDAD"

    # Encabezados de la columna
    sheet["A1"] = "N°"
    sheet["B1"] = "INT"
    sheet["C1"] = "DESCRIPCIÓN"
    sheet["D1"] = "SUCESO"
    sheet["E1"] = "FECHA / HORA"
    sheet["F1"] = "UBICACIÓN"

    # Definir el formato de hipervínculo
    hyperlink_format = Font(color=Color(rgb="0000FF"), underline="single")

    row = 2  # Iniciar en la segunda fila
    count = 1  # Contador de sucesos

    for email in inbox.Items:
        email_body = email.Body
        # Obtiene la dirección de correo del remitente
        email_sender = email.SenderEmailAddress

        # Verificar si el correo es del remitente deseado
        if email_sender == sender_email:
            # Verificar si el correo contiene la información deseada
            if "Vehículo:" in email_body and "Suceso:" in email_body and "Descripción:" in email_body and "Fecha y hora:" in email_body:
                # Utilizar expresiones regulares para extraer la información
                vehicle_match = re.search(r"Vehículo: (\d+)", email_body)
                description_match = re.search(
                    r"Descripción: (.+?) -", email_body)
                event_match = re.search(r"Suceso: (.+?) -", email_body)
                datetime_match = re.search(r"Fecha y hora: ([A-Za-z]{3} \d{1,2} \d{4}(?:\s*)\d{1,2}:\d{2}[APap][Mm])", email_body)

                location_match = re.search(
                    r"Ubicación:\s*(https://[^\n]+)", email_body)

                if vehicle_match and description_match and event_match and datetime_match and location_match:
                    vehicle = vehicle_match.group(1)
                    description = description_match.group(1)
                    event = event_match.group(1)
                    datetime_str = datetime_match.group(1)
                    location = location_match.group(1)

                    # Buscar el índice del primer espacio después del número del vehículo
                    space_index = email_body.find(vehicle) + len(vehicle) + 1

                    # Buscar el índice del primer guion después del espacio encontrado
                    dash_index = email_body.find("-", space_index)

                    # Extraer el modelo del vehículo desde el espacio hasta el guion
                    vehicle_model = email_body[space_index:dash_index].strip()

                    # Llenar la hoja de cálculo con la información
                    sheet[f"A{row}"] = count
                    sheet[f"B{row}"] = vehicle
                    sheet[f"C{row}"] = vehicle_model
                    # Cambio a la descripción completa
                    sheet[f"D{row}"] = description
                    sheet[f"E{row}"] = format_date_time(datetime_str)

                    # Agregar el enlace de ubicación y aplicar formato de hipervínculo
                    sheet[f"F{row}"] = location
                    sheet[f"F{row}"].hyperlink = location
                    sheet[f"F{row}"].font = hyperlink_format

                    row += 1
                    count += 1

    workbook.save(output_filename)


if __name__ == "__main__":
    output_file = "EXCESO DE VELOCIDAD.xlsx"  # Nombre del archivo de salida Excel
    # Dirección de correo del remitente deseado
    sender_email = "test@example.com"

    extract_and_save_parts(output_file, sender_email)
