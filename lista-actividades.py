import win32com.client
import openpyxl
from datetime import datetime, timedelta
import os

# Obtener la fecha actual
today = datetime.now()

# Calcular el lunes de la semana actual (offset-naive)
start_date = datetime(today.year, today.month, today.day) - timedelta(days=today.weekday())
# Calcular el viernes de la semana actual (offset-naive)
end_date = start_date + timedelta(days=4, hours=23, minutes=59, seconds=59)

# Formatear las fechas para el filtro
start_date_str = start_date.strftime("%Y-%m-%d %H:%M")
end_date_str = end_date.strftime("%Y-%m-%d %H:%M")

# Conectar a Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Obtener el calendario del usuario
calendar = namespace.GetDefaultFolder(9)  # 9 es el índice para la carpeta del calendario

# Crear un filtro para las fechas
restriction = f"[Start] >= '{start_date_str}' AND [Start] <= '{end_date_str}'"
calendar_items = calendar.Items.Restrict(restriction)
calendar_items.IncludeRecurrences = True
calendar_items.Sort("[Start]")

# Crear un archivo Excel
output_file = "reporte.xlsx"

# Eliminar el archivo existente si existe
if os.path.exists(output_file):
    os.remove(output_file)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Eventos del Calendario"
ws.append(["timeSpent", "date", "consecutivo", "concept", "comment", "tareaId"])

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
        categories = appointment.Categories.split(';') if appointment.Categories else []
        if len(categories) == 2:  # Solo incluir eventos con exactamente dos categorías
            category2_number = categories[1].split()[0]  # Tomar solo la parte numérica de la segunda categoría
            event_date = appointment_start.strftime("%Y-%m-%d")  # Formatear solo la fecha sin la hora
            duration_hours = (appointment_end - appointment_start).total_seconds() / 3600  # Duración en horas
            ws.append([round(duration_hours, 2), event_date, category2_number, categories[0], name, ""])

# Guardar el archivo Excel
wb.save(output_file)

print(f"Los eventos del calendario se han exportado a {output_file}")
