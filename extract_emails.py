import win32com.client
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Color
import re
from datetime import datetime, timedelta

def format_date_time(datetime_str):
    # Split the datetime_str into lines
    lines = datetime_str.split('\n')

    # Initialize a list to store the formatted date-time strings
    formatted_date_times = []

    # Process pairs of lines to reformat them
    for i in range(0, len(lines), 2):
        date_str = lines[i]
        time_str = lines[i + 1]

        # Process and reformat date and time
        date_time_match = re.search(r"(\w{3} \d{1,2} \d{4})\s*(\d{1,2}:\d{2}[APap][Mm])", f"{date_str} {time_str}")
        if date_time_match:
            date, time = date_time_match.groups()

            # Convert the abbreviated month to a numeric month
            months = {
                'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
                'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
            }
            month_abbr, day, year = date.split()
            formatted_date = f"{day}/{months[month_abbr]}/{year}"

            formatted_date_time = f"{formatted_date} - {time}"
            formatted_date_times.append(formatted_date_time)

    # Join the formatted date-time strings with line breaks
    formatted_datetime_str = '\n'.join(formatted_date_times)

    return(formatted_datetime_str)

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

                    print(format_date_time(datetime_str))

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
    sender_email = "daemon@lojack.com.ar"

    extract_and_save_parts(output_file, sender_email)
