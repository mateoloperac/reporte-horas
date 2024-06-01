import win32com.client
import openpyxl
from datetime import datetime, timedelta
import os

# Obtener la fecha actual
today = datetime.now()

# Calcular el lunes de la semana actual (offset-naive)
start_date = datetime(today.year, today.month, today.day) - timedelta(days=today.weekday())
# Calcular el domingo de la semana actual (offset-naive)
end_date = start_date + timedelta(days=6, hours=23, minutes=59, seconds=59)

# Formatear las fechas para el filtro en formato de Outlook
start_date_str = start_date.strftime("%d/%m/%Y %H:%M %p")
end_date_str = end_date.strftime("%d/%m/%Y %H:%M %p")
# Formatear la fecha del viernes en formato AAAA-MM-DD
friday_date_str = end_date.strftime("%Y-%m-%d")
print("start date: " + str(start_date_str))
print("end date: " + str(end_date_str))

# Conectar a Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Obtener el calendario del usuario
calendar = namespace.GetDefaultFolder(9)  # 9 es el índice para la carpeta del calendario

# Crear un filtro para las fechas
restriction = f"[Start] >= '{start_date_str}' AND [Start] <= '{end_date_str}'"
print("restriction: " + restriction)
calendar_items = calendar.Items.Restrict(restriction)
calendar_items.IncludeRecurrences = True
calendar_items.Sort("[Start]")

# Crear un archivo Excel
output_file = "reporte-semanal.xlsx"

# Eliminar el archivo existente si existe
if os.path.exists(output_file):
    os.remove(output_file)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "reporte"
ws.append(["timeSpent", "date", "consecutivo", "concept", "comment", "tareaId"])

# Diccionario para agrupar eventos por categorías
grouped_events = {}

# Agregar eventos al archivo Excel después de filtrarlos nuevamente
for appointment in calendar_items:
    appointment_start = appointment.Start
    appointment_end = appointment.End

    # Convertir appointment_start y appointment_end a offset-naive si son offset-aware
    if appointment_start.tzinfo is not None:
        appointment_start = appointment_start.replace(tzinfo=None)
    if appointment_end.tzinfo is not None:
        appointment_end = appointment_end.replace(tzinfo=None)
    
    if start_date <= appointment_start <= end_date:
        name = appointment.Subject
        categories = appointment.Categories
        duration_hours = (appointment_end - appointment_start).total_seconds() / 3600  # Duración en horas
        
        # Dividir las categorías en una lista y seleccionar solo las dos primeras categorías
        categories_list = categories.split(";") if categories else []
        concept = categories_list[0].strip() if len(categories_list) > 0 else ""
        consecutivo = ""
        if len(categories_list) > 1:
            category2_parts = categories_list[1].strip().split(" ")
            if len(category2_parts) > 0 and category2_parts[0].isdigit():
                consecutivo = category2_parts[0].strip()

        # Verificar si la segunda categoría está vacía antes de agregar al diccionario
        if consecutivo:
            # Generar una clave única para agrupar eventos por categorías
            key = f"{concept}-{consecutivo}"
            
            # Agregar el evento al diccionario o actualizar la duración si ya existe
            if key in grouped_events:
                grouped_events[key]["comment"] += f". {name}"
                grouped_events[key]["timeSpent"] += duration_hours
            else:
                grouped_events[key] = {
                    "comment": name,
                    "concept": concept,
                    "consecutivo": consecutivo,
                    "date": friday_date_str,
                    "timeSpent": duration_hours,
                    "tareaId": ""
                }

# Agregar eventos agrupados al archivo Excel
for event_key, event_data in grouped_events.items():
    ws.append([round(event_data["timeSpent"], 2), event_data["date"], event_data["consecutivo"], event_data["concept"], event_data["comment"], event_data["tareaId"]])

# Guardar el archivo Excel
wb.save(output_file)
print(f"Los eventos del calendario se han exportado a {output_file}")
