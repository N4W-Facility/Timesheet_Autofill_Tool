"""
  Nature For Water Facility - The Nature Conservancy
  -------------------------------------------------------------------------
  Python 3.11
  -------------------------------------------------------------------------
                            BASIC INFORMATION
 --------------------------------------------------------------------------
  Author        : Jonathan Nogales Pimentel
                  Carlos A. Rogéliz Prada
  Email         : jonathan.nogales@tnc.org
  Date          : November, 2024
"""

# ----------------------------------------------------------------------------------------------------------------------
# Load libraries
# ----------------------------------------------------------------------------------------------------------------------
import locale
import pytz
import calendar
import datetime as dt
import os
import numpy as np
import pandas as pd
import win32com.client
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
import win32com.client
import threading
import time
from datetime import datetime, timedelta
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import traceback
from pytz import timezone, utc
from tzlocal import get_localzone
from datetime import datetime

# Función para crear directorio
def create_folder(dir):
    if not os.path.exists(dir):
        os.makedirs(dir)

# ----------------------------------------------------------------------------------------------------------------------
# Progress bar
# ----------------------------------------------------------------------------------------------------------------------
# Función para mostrar la ventana de progreso
def show_progress_window(max_value):
    global progress_window, progress_bar
    progress_window = tk.Toplevel(App)
    progress_window.title("Progress")
    progress_window.geometry("300x100")
    progress_window.resizable(False, False)

    # Etiqueta en la ventana
    tk.Label(progress_window, text="Updating categories, please wait...").pack(pady=10)

    # Barra de progreso
    progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=250, mode="determinate")
    progress_bar.pack(pady=10)
    progress_bar['value'] = 0
    progress_bar['maximum'] = max_value
    progress_window.update_idletasks()

# Función para ocultar la ventana de progreso
def hide_progress_window():
    progress_window.destroy()

def readDataBase(filepath):
    df1 = pd.read_excel(filepath, sheet_name='TNC-Employee')
    df2 = pd.read_excel(filepath, sheet_name='N4W-Projects')
    df3 = pd.read_excel(filepath, sheet_name='TNC-Projects')
    return pd.concat([df1, df2, df3], ignore_index=True)

# ----------------------------------------------------------------------------------------------------------------------
# Update categories
# ----------------------------------------------------------------------------------------------------------------------
# Función para actualizar categorías en Outlook
def update_categories(filepath):
    try:
        # Leer el archivo Excel
        df = readDataBase(filepath)
        df = df.dropna(subset=['Pegasys ID']).fillna(0)

        # Verificar si existen las columnas necesarias
        required_columns = ['Category', 'Include']
        for column in required_columns:
            if column not in df.columns:
                raise ValueError(f"El archivo Excel debe contener una columna llamada '{column}'.")

        # Mostrar la ventana de progreso
        show_progress_window(len(df))

        # Conectar con Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        categories = outlook.Session.Categories

        # Lista de categorías existentes en Outlook
        existing_categories = [cat.Name for cat in categories]

        # Procesar categorías
        for i, row in df.iterrows():
            category_name = row['Category']
            include = row['Include']
            color_index = row.get('ColorIndex', i % 25 + 1)

            if include == 1:
                if category_name not in existing_categories:
                    categories.Add(category_name, color_index)
            elif include == 0:
                if category_name in existing_categories:
                    categories.Remove(category_name)
            time.sleep(0.25)

            # Actualizar el progreso
            progress_bar['value'] = i + 1
            progress_window.update_idletasks()

        # Ocultar la ventana de progreso
        hide_progress_window()
        messagebox.showinfo("Completed", "Category update completed.")
    except Exception as e:
        hide_progress_window()
        messagebox.showerror("Error", f"Error updating categories: {e}")

# Función para ejecutar el proceso en un hilo separado
def run_update_categories(filepath):
    threading.Thread(target=update_categories, args=(filepath,), daemon=True).start()

# ----------------------------------------------------------------------------------------------------------------------
# Get meetings
# ----------------------------------------------------------------------------------------------------------------------
# Detectar y establecer automáticamente la configuración regional del sistema operativo
#locale.setlocale(locale.LC_TIME, locale.getdefaultlocale()[0])  # Configuración regional del sistema operativo

def get_calendar(start_date, end_date,i1=25,i2=25):
    # Conexión a Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = outlook.GetDefaultFolder(9)  # Carpeta predeterminada de Calendario

    # Detectar zona horaria del sistema
    local_tz = get_localzone()

    # Convertir fechas a formato datetime con zona horaria
    start_date = datetime.strptime(start_date, '%Y-%m-%d').replace(tzinfo=local_tz)
    end_date = datetime.strptime(end_date, '%Y-%m-%d').replace(tzinfo=local_tz)

    # Obtener elementos del calendario
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    start_date1 = start_date - dt.timedelta(days=i1)
    end_date1 = end_date + dt.timedelta(days=i2)

    # Filtrar reuniones por rango de fechas
    start_str = start_date1.strftime('%m/%d/%Y %H:%M')
    end_str = end_date1.strftime('%m/%d/%Y %H:%M')
    restriction = f"[Start] >= '{start_str}' AND [End] <= '{end_str}'"
    restricted_items = items.Restrict(restriction)

    # Crear diccionario para almacenar reuniones por día y Category
    meetings = []

    for item in restricted_items:
        try:
            # Verificar si el elemento está dentro del rango explícitamente
            meeting_start = item.Start
            meeting_end = item.End
            if remove_timezone(start_date) <= remove_timezone(meeting_start) <= remove_timezone(end_date):
                category = item.Categories if item.Categories else "Sin Category"
                meeting_date = meeting_start.date()
                duration = (meeting_end - meeting_start).total_seconds() / 3600  # Duración en horas

                meetings.append({'Date': meeting_date, 'Category': category, 'Hours': duration})
        except AttributeError:
            # Saltar elementos que no tengan las propiedades necesarias
            continue

    # Crear DataFrame base
    df = pd.DataFrame(meetings)

    return df

def remove_timezone(date):
    """
    Convierte un objeto datetime timezone-aware a timezone-naive.
    """
    return date.replace(tzinfo=None)

# Función para calcular días laborables del mes
def calculate_workdays(year, month):
    _, total_days = calendar.monthrange(year, month)
    workdays = sum(1 for day in range(1, total_days + 1)
                   if dt.datetime(year, month, day).weekday() < 5)  # Excluye sábados y domingos
    return workdays

# Función para procesar la columna `Category`
def process_category(category):
    keywords = [
        "REGULAR", "LWOP", "MATERNITY", "ADMIN LEAVE", "PARENTAL LEAVE",
        "Compensation", "FURLOUGH", "PUBLIC HOLIDAY", "Medical Leave",
        "Personal Leave Day", "SICK", "VACATION"
    ]

    # Buscar si alguna palabra clave está en el texto
    found_keyword = next((keyword for keyword in keywords if re.search(keyword, category, flags=re.IGNORECASE)),
                         "REGULAR")

    # Si se encuentra la palabra clave, eliminarla del texto
    if found_keyword != "REGULAR":
        category = re.sub(found_keyword, "", category, flags=re.IGNORECASE)

    # Eliminar comas y espacios adicionales
    category = category.replace(",", "").strip()

    # Eliminar comas y espacios adicionales
    category = category.replace(";", "").strip()

    return found_keyword, category


# Función principal para generar el reporte
def generate_report():
    try:
        # Detectar zona horaria del sistema
        local_tz = get_localzone()

        # Obtener las fechas seleccionadas
        start_date = datetime.strptime(start_date_entry.get(), '%Y-%m-%d')
        end_date = datetime.strptime(end_date_entry.get(), '%Y-%m-%d')
        NameDataBase = ProjectsDataBase.get()

        # Validar que la fecha de inicio sea menor o igual a la de finalización
        if start_date > end_date:
            messagebox.showerror("Error", "Start date cannot be after end date.")
            return

        end_date = end_date + dt.timedelta(days=1)

        Aditivos = [13, 2, 3, 5, 7, 11, 17, 19, 23, 29, 31]
        for i1 in Aditivos:
            for i2 in Aditivos:
                # Cargar datos del calendario
                results = get_calendar(start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'),i1,i2)
                #results = get_appointments(raw_data)

                if len(results.columns) is not 0:
                    break
            if len(results.columns) is not 0:
                break

        results[['Earning', 'Category']] = results['Category'].apply(lambda x: pd.Series(process_category(x)))

        # Agregar resultados
        tmp = results.groupby(by=['Date', 'Category', 'Earning'], as_index=False)['Hours'].sum()
        tmp = tmp.pivot(index=['Category', 'Earning'], columns='Date', values='Hours').fillna(0)
        tmp = tmp.reset_index(level='Earning')

        # Crear reporte con fechas completas del rango
        report = pd.DataFrame(columns=pd.date_range(start_date, end_date, freq='D'))

        # Asegúrate de que las columnas de `tmp` que representan fechas sean datetime
        tmp.columns = pd.to_datetime(tmp.columns, errors='coerce')

        # Asegúrate de que las columnas de `report` sean datetime
        report.columns = pd.to_datetime(report.columns, errors='coerce')

        # Concatenar y llenar con ceros
        report = pd.concat([report, tmp], axis=0).fillna(0)

        # Formatear las columnas de fechas en el DataFrame final
        report.columns = report.columns.map(
            lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if isinstance(x, pd.Timestamp) else x)

        report.index = [texto.split('|')[0].strip() for texto in report.index.values]

        # Leer códigos del N4W Facility
        N4WCodes = readDataBase(NameDataBase)
        N4WCodes = N4WCodes.dropna(subset=['Pegasys ID']).fillna(0).replace('XXXXXX', 0)
        N4WCodes['Activity ID'] = N4WCodes['Activity ID'].astype(int)
        N4WCodes['Project ID'] = N4WCodes['Project ID'].astype(str)
        N4WCodes['Award ID'] = N4WCodes['Award ID'].astype(str)
        N4WCodes = N4WCodes.set_index(['Pegasys ID'])

        # Combinar con el reporte generado
        ruta_directorio = os.path.dirname(NameDataBase)
        Value = pd.merge(N4WCodes, report, left_index=True, right_index=True)
        Value = Value.drop(columns=['Description', 'Category', 'Include'])
        Value.columns = [str(col) if pd.notnull(col) else 'Earning' for col in Value.columns]

        # Mover columna 'Earning' a una posición específica
        column_to_move = 'Earning'
        new_position = 3
        cols = Value.columns.tolist()
        cols.insert(new_position, cols.pop(cols.index(column_to_move)))
        Value = Value[cols]

        # Crear un DataFrame con los textos a reemplazar
        data = {'Earning': [
            'REGULAR',
            'LWOP',
            'MATERNITY',
            'ADMIN LEAVE',
            'PARENTAL LEAVE',
            'Compensation',
            'FURLOUGH',
            'PUBLIC HOLIDAY',
            'Medical Leave',
            'Personal Leave Day',
            'SICK',
            'VACATION'
        ]}

        df = pd.DataFrame(data)

        # Crear el diccionario de mapeo (texto -> ID)
        mapping = {
            'REGULAR': '1',
            'LWOP': '17',
            'MATERNITY': '301',
            'ADMIN LEAVE': '6',
            'PARENTAL LEAVE': '69',
            'Compensation': 'C',
            'FURLOUGH': 'FRL',
            'PUBLIC HOLIDAY': 'H',
            'Medical Leave': 'ML',
            'Personal Leave Day': 'PLD',
            'SICK': 'S',
            'VACATION': 'V'
        }

        # Reemplazar los textos en la columna 'Category' por los IDs
        Value['Earning'] = Value['Earning'].map(mapping)

        # eliminar el dia adicional que se incluye
        Value = Value.drop(columns=Value.columns[-1])

        # Guardar resultados
        output_folder = os.path.join(ruta_directorio)
        os.makedirs(output_folder, exist_ok=True)

        results.to_excel(os.path.join(output_folder, '01-Report.xlsx'))
        Value.to_csv(os.path.join(output_folder, '02-Deltek.csv'), index_label='Pegasys ID')

        tk.messagebox.showinfo(message="Process Completed", title="Status")
    except Exception as e:
        messagebox.showerror("General Error", f"An unexpected error occurred: {e}")
        traceback.print_exc()


# Función para seleccionar archivo
def select_file(entry_field):
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if filepath:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, filepath)

# Función para Fill_Deltek
def fill_deltek():
    try:
        PoPo        = int(Posi_entry_deltek.get())
        LoginID     = email_entry_deltek.get()
        Password    = password_entry_deltek.get()
        Domain      = 'TNC.ORG'
        Path        = r'chromedriver.exe'
        NameDataBase= ProjectsDataBase.get()
        # ----------------------------------------------------------------------------------------------------------------------
        # Leer Datos
        # ----------------------------------------------------------------------------------------------------------------------
        ruta_directorio = os.path.dirname(NameDataBase)
        # Leer códigos del N4W Facility
        output_folder = os.path.join(ruta_directorio)
        Value = pd.read_csv(os.path.join(output_folder, '02-Deltek.csv'), index_col=0)
        Value = Value.groupby(['Project ID', 'Activity ID', 'Award ID','Earning'], as_index=False).sum()
        Deltek_Data = Value[['Project ID', 'Activity ID', 'Award ID','Earning']]
        Value = Value.drop(columns=['Project ID', 'Activity ID', 'Award ID','Earning'])
        Value[np.isnan(Value)] = 0
        Value.columns = pd.to_datetime(Value.columns)

        # ----------------------------------------------------------------------------------------------------------------------
        # Abrir Google Chrome operado con Selenium cambiando la carpeta de descargas
        # ----------------------------------------------------------------------------------------------------------------------
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--start-maximized')
        chrome_options.add_argument('--desable-extensions')
        try:
            chrome_options.add_experimental_option("detach", True)
            service = Service(executable_path=Path)
            driver = webdriver.Chrome(service=service, options=chrome_options)
        except:
            driver = webdriver.Chrome(Path, chrome_options=chrome_options)

        # ----------------------------------------------------------------------------------------------------------------------
        # Abrir la página de Deltek
        # ----------------------------------------------------------------------------------------------------------------------
        #driver.get('https://tnc.hostedaccess.com/DeltekTC/welcome.msv')
        driver.get("https://tnc.hostedaccess.com/DeltekTC/TimeCollection.msv")

        Ntime = 10
        # ----------------------------------------------------------------------------------------------------------------------
        # Introducir dominio
        # ----------------------------------------------------------------------------------------------------------------------
        WebDriverWait(driver, Ntime).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#uid'))).send_keys(LoginID)

        # ----------------------------------------------------------------------------------------------------------------------
        # Introducir contraseña
        # ----------------------------------------------------------------------------------------------------------------------
        WebDriverWait(driver, Ntime).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#passField'))).send_keys(
            Password)

        # ----------------------------------------------------------------------------------------------------------------------
        # Introducir dominio
        # ----------------------------------------------------------------------------------------------------------------------
        WebDriverWait(driver, Ntime).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#dom'))).send_keys(Domain)

        # ----------------------------------------------------------------------------------------------------------------------
        # Entrar a deltek
        # ----------------------------------------------------------------------------------------------------------------------
        WebDriverWait(driver, Ntime).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#loginButton'))).click()

        # ----------------------------------------------------------------------------------------------------------------------
        # Borrar los Projects que esten por defecto
        # ----------------------------------------------------------------------------------------------------------------------
        driver.switch_to.frame(1)
        WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "allRowSelector"))).click()
        WebDriverWait(driver, Ntime).until(EC.element_to_be_clickable((By.ID, "deleteLine"))).click()
        time.sleep(0.5)

        # JavaScript para mover el scroll, usando el porcentaje de entrada
        js_script = """
                let scroller = document.getElementById('udtScroller'); // Elemento que tiene el scroll
                let scrollContent = document.getElementById('udtScrollerContent'); // Contenido dentro del scroll

                if (scroller && scrollContent) {
                    // Calcular el desplazamiento total posible
                    let maxScrollLeft = scrollContent.offsetWidth - scroller.offsetWidth;
                    let scrollAmount = maxScrollLeft * arguments[0]; // Mover el porcentaje ingresado
                    scroller.scrollLeft = scrollAmount; // Ajustar el desplazamiento horizontal
                }
                """

        js_scrip2 = """
                let scroller = document.getElementById('udtScroller'); // Elemento que tiene el scroll
                if (scroller) {
                    scroller.scrollLeft = 0; // Volver al inicio horizontalmente
                    scroller.scrollTop = 0;  // Volver al inicio verticalmente (opcional)
                }
                """

        # ----------------------------------------------------------------------------------------------------------------------
        # Diligenciar los Projects - ID
        # ----------------------------------------------------------------------------------------------------------------------
        # 4. Navegar a la segunda página
        # driver.get("https://tnc.hostedaccess.com/DeltekTC/TimeCollection.msv")
        """
        Book = Book_deltek.get()
        """
        for i in range(0, Deltek_Data["Project ID"].size):
            # Project ID
            WebDriverWait(driver, Ntime).until(
                EC.element_to_be_clickable((By.ID, "udt" + str(i + PoPo) + "_1"))).click()
            WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).send_keys(
                Deltek_Data["Project ID"][i])

            # Ejecutar el script en la página con el porcentaje como argumento
            driver.execute_script(js_script, 0.2)

            # GeoOrigen
            WebDriverWait(driver, Ntime).until(
                EC.element_to_be_clickable((By.ID, "udt" + str(i + PoPo) + "_3"))).click()

            # Ejecutar el script en la página con el porcentaje como argumento
            driver.execute_script(js_script, 0.3)

            # Award ID
            WebDriverWait(driver, Ntime).until(
                EC.element_to_be_clickable((By.ID, "udt" + str(i + PoPo) + "_4"))).click()
            WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).clear()
            WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).send_keys(
                str(Deltek_Data["Award ID"][i]))

            # Ejecutar el script en la página con el porcentaje como argumento
            driver.execute_script(js_script, 0.4)

            # Activity
            WebDriverWait(driver, Ntime).until(
                EC.element_to_be_clickable((By.ID, "udt" + str(i + PoPo) + "_5"))).click()
            WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).clear()
            WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).send_keys(
                str(Deltek_Data["Activity ID"][i]))
            time.sleep(0.1)

            # Ejecutar el script en la página con el porcentaje como argumento
            driver.execute_script(js_script, 0.5)

            # Indicador
            WebDriverWait(driver, Ntime).until(
                EC.element_to_be_clickable((By.ID, "udt" + str(i + PoPo) + "_6"))).click()
            WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).clear()
            WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).send_keys(
                str(Deltek_Data["Earning"][i]))
            time.sleep(0.1)

            # Ejecutar el script en la página con el porcentaje como argumento
            driver.execute_script(js_scrip2)

            """
            # Book
            WebDriverWait(driver, Ntime).until(
                EC.element_to_be_clickable((By.ID, "udt" + str(i + PoPo) + "_7"))).click()
            WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).clear()
            WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).send_keys(
                str(Book))
            time.sleep(0.1)
            """

        for j in range(Value.columns.size):
            for i in range(0, Deltek_Data["Project ID"].size):
                WebDriverWait(driver, Ntime).until(
                    EC.element_to_be_clickable((By.ID, "hrs" + str(i + PoPo) + "_" + str(j)))).click()
                WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).clear()
                WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).send_keys(
                    str(Value.iloc[i, j]))
                time.sleep(0.05)

        # ----------------------------------------------------------------------------------------------------------------------
        # Final processing
        # ----------------------------------------------------------------------------------------------------------------------
        print("Final - OK")
        # Mensaje de error
        tk.messagebox.showinfo(message="Process Completed", title="Status")

    except Exception as e:
        messagebox.showerror("General Error", f"An unexpected error occurred: {e}")
        traceback.print_exc()

# Función para Fill_Pegasys
def fill_pegasys():
    try:
        LoginID     = email_entry_pegasys.get()
        Password    = password_entry_pegasys.get()
        Path        = r'chromedriver.exe'
        NameDataBase = ProjectsDataBase.get()
        ruta_directorio = os.path.dirname(NameDataBase)

        # Leer códigos del N4W Facility
        output_folder = os.path.join(ruta_directorio)
        Value = pd.read_csv(os.path.join(output_folder, '02-Deltek.csv'),index_col=0)
        Value = Value.drop(columns=['Project ID', 'Activity ID', 'Award ID','Earning'])
        Value[np.isnan(Value)] = 0
        Value.columns = pd.to_datetime(Value.columns)
        Value = Value.groupby(Value.index).sum()

        ListDate = Value.columns
        ListPro  = Value.index.values

        # ----------------------------------------------------------------------------------------------------------------------
        # Abrir Google Chrome operado con Selenium cambiando la carpeta de descargas
        # ----------------------------------------------------------------------------------------------------------------------
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--start-maximized')
        chrome_options.add_argument('--desable-extensions')
        chrome_options.add_experimental_option("detach", True)
        try:
            service = Service(executable_path=Path)
            driver = webdriver.Chrome(service=service, options=chrome_options)
        except:
            driver = webdriver.Chrome(Path,chrome_options=chrome_options)

        # ----------------------------------------------------------------------------------------------------------------------
        # Abrir la página de Deltek
        # ----------------------------------------------------------------------------------------------------------------------
        driver.get('https://time.pegasys.co.za/trs/index.php')

        Ntime = 10
        # ----------------------------------------------------------------------------------------------------------------------
        # Introducir correo
        # ----------------------------------------------------------------------------------------------------------------------
        WebDriverWait(driver, Ntime).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#email.form-control'))).send_keys(LoginID)

        # ----------------------------------------------------------------------------------------------------------------------
        # Introducir contraseña
        # ----------------------------------------------------------------------------------------------------------------------
        WebDriverWait(driver, Ntime).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#pass.form-control'))).send_keys(Password)

        # ----------------------------------------------------------------------------------------------------------------------
        # Entrar al timesheet de pegasys
        # ----------------------------------------------------------------------------------------------------------------------
        WebDriverWait(driver, Ntime).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input.button.glossy.orange'))).click()

        # ----------------------------------------------------------------------------------------------------------------------
        # Selecciona una fecha
        # ----------------------------------------------------------------------------------------------------------------------
        for Datei in ListDate:
            if np.sum(Value[Datei]) == 0.0:
                continue

            # Convertir la fecha a un objeto datetime
            fecha = datetime.strptime(Datei.strftime("%Y-%m-%d"), "%Y-%m-%d")

            # Calcular la fecha del lunes de esa semana
            # weekday() devuelve 0 para lunes, 1 para martes, ..., 6 para domingo
            NumDay = fecha.weekday()
            lunes = fecha - timedelta(days=NumDay)
            Datejj = f"{lunes.strftime('%Y-%m-%d')}"

            # Seleccionar la semana en la cual se deben cargar los tiempos
            Dateii = WebDriverWait(driver, Ntime).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'select'))).get_attribute("value")

            if Datejj != Dateii:
                WebDriverWait(driver, Ntime).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'input.button.glossy.orange'))).click()
                WebDriverWait(driver, Ntime).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'select'))).send_keys(
                    Datejj)

            # Esperar hasta que la tabla esté presente
            tabla = WebDriverWait(driver, Ntime).until(
                EC.presence_of_element_located((By.XPATH, "//form[@action='/trs/updatetimesheet.php']/table")))

            # Extraer todas las filas de la tabla (excluyendo la fila de encabezado)
            filas = tabla.find_elements(By.TAG_NAME, "tr")[1:]

            # Iterar sobre cada fila para buscar la columna "Task" con el valor NameProj
            for ProjectName in ListPro:
                if Value[Datei][ProjectName] == 0:
                    continue

                for fila in filas:
                    celdas = fila.find_elements(By.TAG_NAME, "td")

                    # Verificar si la fila tiene suficientes columnas para evitar errores de índice
                    if len(celdas) < 2:
                        continue

                    # Obtener el texto de la columna "Task" (índice 1)
                    task_texto = celdas[1].text

                    # Verificar si el texto coincide con NameProj
                    if task_texto == ProjectName:
                        # Llenar los campos de entrada de esa fila con el valor '1'
                        celda = celdas[4 + NumDay]
                        input_tag = celda.find_element(By.TAG_NAME, "input")
                        input_tag.clear()
                        input_tag.send_keys('%f' % Value[Datei][ProjectName])
                        # Salir del bucle después de llenar la fila
                        break

        # Guardar al finalizar
        WebDriverWait(driver, Ntime).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input.button.glossy.orange'))).click()

        # Mensaje de error
        tk.messagebox.showinfo(message="Process Completed", title="Status")

    except Exception as e:
        messagebox.showerror("General Error", f"An unexpected error occurred: {e}")
        traceback.print_exc()

# Crear la ventana principal
App = tk.Tk()
App.title("App Timesheet Autofill Tool")

# ----------------------------------------------------------------------------------------------------------------------
# Módulo 1 - Update Categories in Outlook
# ----------------------------------------------------------------------------------------------------------------------
module1 = tk.LabelFrame(App, text="Step 1 - Update Categories in Outlook")
module1.pack(fill="x", padx=10, pady=5)

ProjectsDataBase = tk.Entry(module1, width=50)
ProjectsDataBase.pack(side="left", padx=5, pady=5)

Button_LoadDataBase = tk.Button(module1, text="Load Project DataBase", command=lambda: select_file(ProjectsDataBase))
Button_LoadDataBase.pack(side="left", padx=5, pady=5)

Button_UpdateCategories = tk.Button(module1, text="Update Categories", command=lambda: run_update_categories(ProjectsDataBase.get()))
Button_UpdateCategories.pack(side="left", padx=5, pady=5)

# ----------------------------------------------------------------------------------------------------------------------
# Módulo 2 - Read outlook meetings
# ----------------------------------------------------------------------------------------------------------------------
module2 = tk.LabelFrame(App, text="Step 2 - Read outlook meetings")
module2.pack(fill="x", padx=10, pady=5)

# Selector de fecha de inicio
start_date_label = tk.Label(module2, text="Start Date:")
start_date_label.pack(side="left", padx=5, pady=5)

start_date_entry = DateEntry(module2, width=12, background='darkblue',
                             foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
start_date_entry.pack(side="left", padx=5, pady=5)

# Selector de fecha de finalización
end_date_label = tk.Label(module2, text="End Date:")
end_date_label.pack(side="left", padx=5, pady=5)

end_date_entry = DateEntry(module2, width=12, background='darkblue',
                           foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
end_date_entry.pack(side="left", padx=5, pady=5)

# Botón para leer reuniones
read_button = tk.Button(module2, text="Read meetings", command=lambda: generate_report())
read_button.pack(side="left", padx=5, pady=5)

# ----------------------------------------------------------------------------------------------------------------------
# Module 3 - Automatic filling of time sheet in Deltek
# ----------------------------------------------------------------------------------------------------------------------
module3_deltek = tk.LabelFrame(App, text="Step 3 - Automatic filling of time sheet in Deltek")
module3_deltek.pack(fill="x", padx=10, pady=5)

email_entry_deltek = tk.Entry(module3_deltek, width=30)
email_entry_deltek.pack(side="left", padx=5, pady=5)
email_entry_deltek.insert(0, "Login ID")

password_entry_deltek = tk.Entry(module3_deltek, width=30, show="*")
password_entry_deltek.pack(side="left", padx=5, pady=5)
password_entry_deltek.insert(0, "Password")

fill_deltek_button = tk.Button(module3_deltek, text="Fill Deltek", command=lambda: fill_deltek())
fill_deltek_button.pack(side="left", padx=5, pady=5)

"""
Book_deltek = tk.Entry(module3_deltek, width=8)
Book_deltek.pack(side="left", padx=5, pady=5)
Book_deltek.insert(0, "100")
"""

Posi_entry_deltek = tk.Entry(module3_deltek, width=3)
Posi_entry_deltek.pack(side="left", padx=5, pady=5)
Posi_entry_deltek.insert(0, "0")

# ----------------------------------------------------------------------------------------------------------------------
# Module 4 - Automatic filling of time sheet in Pegasys
# ----------------------------------------------------------------------------------------------------------------------
module3_pegasys = tk.LabelFrame(App, text="Step 4 - Automatic filling of time sheet in Pegasys")
module3_pegasys.pack(fill="x", padx=10, pady=5)

email_entry_pegasys = tk.Entry(module3_pegasys, width=30)
email_entry_pegasys.pack(side="left", padx=5, pady=5)
email_entry_pegasys.insert(0, "Email")

password_entry_pegasys = tk.Entry(module3_pegasys, width=30, show="*")
password_entry_pegasys.pack(side="left", padx=5, pady=5)
password_entry_pegasys.insert(0, "Password")

fill_pegasys_button = tk.Button(module3_pegasys, text="Fill Pegasys",
                                command=lambda: fill_pegasys())
fill_pegasys_button.pack(side="left", padx=5, pady=5)

App.mainloop()
