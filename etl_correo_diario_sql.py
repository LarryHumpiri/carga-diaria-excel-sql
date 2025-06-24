"""

 ██╗      ██╗  ██╗ ██╗  ██╗███████╗
 ██║      ██║ ██╔╝ ██║  ██║██╔══██╗
 ██║      █████═╝  ███████║██║  ██║
 ██║      ██╔═██╗  ██╔══██║██║  ██║
 ███████╗ ██║  ██╗ ██║  ██║███████║
 ╚══════╝ ╚═╝  ╚═╝ ╚═╝  ╚═╝╚══════╝
  👨‍💻 Lion King HO - Desarrollador Python
  📂 Proyecto: Carga automática diaria a SQL Server desde un correo electrónico. 
  💼 https://github.com/LarryHumpiri
  © 2025 | LK
"""

import os
import time
import logging
import pandas as pd
from datetime import datetime, timedelta
import pyodbc
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext

# --------------------------------------------------------------------------------------------
#                                   Configuración de Logging
# --------------------------------------------------------------------------------------------
log_dir = r"D:\Logs\Reporte" #Carpeta donde se almacenara el Log.
if not os.path.exists(log_dir):
    os.makedirs(log_dir)
log_file = os.path.join(log_dir, "etl_log.txt")

logger = logging.getLogger()
logger.setLevel(logging.INFO)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
console_handler.setFormatter(console_formatter)
logger.addHandler(console_handler)

file_handler = logging.FileHandler(log_file)
file_handler.setLevel(logging.ERROR)
file_formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
file_handler.setFormatter(file_formatter)
logger.addHandler(file_handler)
# --------------------------------------------------------------------------------------------


# --------------------------------------------------------------------------------------------
#                                   Funciones del Proceso ETL
# --------------------------------------------------------------------------------------------
def authenticate_sharepoint(site_url, username, password):
    """
    🔄 Autentica al usuario en SharePoint y retorna el contexto autenticado.
    
    Esta función permite conectarse a un sitio de SharePoint Online
    utilizando credenciales de usuario y contraseña, obteniendo un contexto
    válido para realizar operaciones como descargar archivos o listar carpetas.
    
    Parámetros:
        site_url (str): URL del sitio de SharePoint.
        username (str): Correo electrónico del usuario con acceso.
        password (str): Contraseña del usuario.
    
    Retorna:
        ClientContext: Objeto autenticado de SharePoint para operaciones posteriores.
    
    Excepciones:
        Lanza una excepción si no se puede autenticar.
    """
    try:
        ctx_auth = AuthenticationContext(site_url)
        if ctx_auth.acquire_token_for_user(username, password):
            ctx = ClientContext(site_url, ctx_auth)
            ctx.load(ctx.web)
            ctx.execute_query()
            logger.info(f"Autenticación exitosa en: {ctx.web.properties.get('Url', 'Desconocida')}")
            return ctx
        else:
            raise Exception("Tokken no adquirido")
    except Exception as e:
        logger.error(f"Error de autenticación: {e}")
        raise
# --------------------------------------------------------------------------------------------
def list_files_in_folder(ctx, folder_url):
    """
    📁 Lista los archivos de una carpeta específica de SharePoint.
    
    Esta funcion es clave para interactuar con SharePoint. Recibe un contexto autenticado (ctx)
    y la ruta de una carpeta (folder_url) para localizarla. Su proposito es extraer los nombres de 
    todos los archivos que contiene esa carpeta y, además, devolver el objeto de la carpeta en sí. 
    
    Esto último es útil si necesitas realizar más acciones con la carpeta.
    Si algo sale mal durante este proceso, la función avisará lanzando una excepción.
    """
    try:
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        files = folder.files.get().execute_query()
        file_names = [f.name for f in files]
        logger.info(f"Archivos en '{folder_url}': {file_names}")
        return file_names, folder
    except Exception as e:
        logger.error(f"Error al listar archivos: {e}")
        raise

# --------------------------------------------------------------------------------------------
def download_excel_sharepoint(ctx, server_path, local_path):
    """
    Esta función te permite descargar un archivo específico desde SharePoint y guardarlo en tu 
    computadora. Necesitas proporcionarle tres datos clave: un contexto autenticado (ctx) para 
    acceder a SharePoint, la ubicación exacta del archivo en el servidor de SharePoint 
    (server_path), y la ruta completa en tu máquina donde quieres guardarlo (local_path). 
    La función te indicará si la descarga se realizó con éxito (True) o si falló (False).

        """
    try:
        file = ctx.web.get_file_by_server_relative_url(server_path)
        with open(local_path, "wb") as local_file:
            file.download(local_file)
            ctx.execute_query()
        logger.info("Archivo descargado desde SharePoint.")
        return True
    except Exception as e:
        logger.error(f"Error al descargar archivo: {e}")
        return False

# --------------------------------------------------------------------------------------------
def extract_data(excel_path):
    """
    Esta función se encarga de leer datos de un archivo Excel (cuya ubicación se especifica con 
    excel_path). Su tarea principal es asegurarse de que las columnas necesarias estén 
    presentes y correctas.
        
    Al finalizar, te entregará dos cosas: los datos ya procesados y validados en un formato    
    DataFrame, y el tiempo exacto que tardó en realizar toda la operación. Si por alguna razón el 
    archivo Excel no tiene todas las columnas que se esperan, la función generará un error 
    (ValueError) para avisarte.
    """
    start_time = time.time()
    try:
        df = pd.read_excel(excel_path, engine="openpyxl")
        expected_columns = [
            "Cod. Doc", "Oficina", "Producto", "NombreProducto",
            "Fecha", "Saldo", "Monto", "Estado"
        ]
        missing_cols = [col for col in expected_columns if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Faltan columnas esperadas: {missing_cols}")
        end_time = time.time()
        logger.info(f"Datos extraídos: {len(df)} registros en {end_time - start_time:.2f}s")
        return df, end_time - start_time
    except Exception as e:
        logger.error(f"Error al leer archivo: {e}")
        raise
    
# --------------------------------------------------------------------------------------------
def transform_data(df, fecha_vencimiento_param):
    """
    Esta función toma un DataFrame con datos sin procesar (df) y los prepara meticulosamente 
    para ser cargados en una base de datos SQL.
    Al final, la función entrega un DataFrame completamente transformado y optimizado, listo
    para su inserción en SQL.
    """
    start_time = time.time()
    try:
        # Agregar columna de fecha deproceso
        df["FechaProceso"] = datetime.now().date()

        # Convertir tipos de datos
        df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce")
        df["Saldo"] = pd.to_numeric(df["Saldo"].astype(str).str.replace(",", "."), errors="coerce")
        df["Monto"] = pd.to_numeric(df["Monto"].astype(str).str.replace(",", "."), errors="coerce")

        # Renombrar columnas
        df.rename(columns={
            "Cod. Doc": "CodDocumento",
            "Oficina": "Oficina",
            "Producto": "Producto",
            "NombreProducto": "NombreProduct",
            "Fecha": "Fecha",
            "Saldo": "Saldo",
            "Monto": "MontoSaldo",
            "Estado": "EstadoDocumento"
        }, inplace=True)

        # Truncar texto
        df["CodDocumento"] = df["CodDocumento"].astype(str).str.slice(0, 20)
        df["Oficina"] = df["Oficina"].astype(str).str.slice(0, 10)
        df["Producto"] = df["Producto"].astype(str).str.slice(0, 50)

        # Filtrar datos
        filtro_fecha = (df["Fecha"].dt.date == fecha_vencimiento_param.date())
        filtro_saldo = df["Saldo"] != 0
        filtro_estado = df["EstadoDocumento"].str.contains("Activo|Pendiente", case=False, na=False)
        df_filtrado = df.loc[filtro_fecha & filtro_saldo & filtro_estado].copy()

        # Validación: Fechas no deben ser mayores a hoy + 90 días
        max_fecha = datetime.now() + timedelta(days=90)
        df_filtrado = df_filtrado[df_filtrado["Fecha"] <= max_fecha]

        # Eliminar campos no requeridos
        if "Saldo" in df_filtrado.columns:
            df_filtrado.drop(columns=["Saldo"], inplace=True)

        end_time = time.time()
        logger.info(f"Transformación completada: {len(df_filtrado)} registros en {end_time - start_time:.2f}s")
        return df_filtrado, end_time - start_time
    except Exception as e:
        logger.error(f"Error al transformar datos: {e}")
        raise

# --------------------------------------------------------------------------------------------
def load_data_pyodbc(df, server, database, trusted_connection):
    """
    Esta función se encarga de insertar los datos ya preparados (en el DataFrame proporcionado) 
    directamente en una base de datos SQL Server utilizando la biblioteca PyODBC.
    Para funcionar, necesita que le indiques tres cosas clave sobre tu base de datos: el servidor 
    (server), el nombre de la base de datos (database) y el método de autenticación 
    (trusted_connection).

    Al finalizar, la función te devolverá el tiempo exacto que tardó en completar la carga de todos los datos.
    """
    start_time = time.time()
    try:
        df["Fecha"] = df["Fecha"].apply(lambda x: x.strftime("%Y-%m-%d") if pd.notnull(x) else None)
        df["FechaProceso"] = df["FechaProceso"].apply(lambda x: x.strftime("%Y-%m-%d") if pd.notnull(x) else None)

        conn = pyodbc.connect(
            f"DRIVER={{SQL Server}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"Trusted_Connection={trusted_connection};"
        )
        cursor = conn.cursor()
        cursor.fast_executemany = True

        insert_query = """
        INSERT INTO [dbo].[ReporteVencidosDiarios]
        (Producto, CodDocumento, NombreProduct, Oficina, Fecha, FechaProceso, MontoSaldo, EstadoDocumento)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """

        data_tuples = [tuple(row) for row in df[
            ["Producto", "CodDocumento", "NombreProduct", "Oficina", "Fecha", "FechaProceso", "MontoSaldo", "EstadoDocumento"]
        ].values]

        cursor.executemany(insert_query, data_tuples)
        conn.commit()
        conn.close()

        end_time = time.time()
        logger.info(f"{len(df)} registros cargados exitosamente.")
        return end_time - start_time
    except Exception as e:
        logger.error(f"Error al cargar datos: {e}")
        raise

# --------------------------------------------------------------------------------------------
#                                      Proceso Principal (ETL)
# --------------------------------------------------------------------------------------------
def main():
    site_url = "https://tuempresa.sharepoint.com/"  #Url del Equipo Sharepoint
    username = "usuario@tudominio.com" #Correo
    password = "Passwoard" #Contraseña
    folder_path = "Documentos compartidos/Reportes" # Carpeta donde se almacena el archivo, esto se logra con el flujo en Power Automate.
    excel_server_relative_url = folder_path + "/reporte.xlsx" 
    local_excel = "reporte.xlsx" #Nombre del archivo a procesar.
    fecha_actual = datetime.now() #Fecha Actual

    fecha_vencimiento_param = fecha_actual.strftime("%d/%m/%Y")

    server = "localhost"
    database = "BD_Ejemplo"
    trusted_connection = "YES"

    try:
        ctx = authenticate_sharepoint(site_url, username, password)
        file_names, folder = list_files_in_folder(ctx, folder_path)

        if "runETL.txt" not in file_names:
            logger.error("No se encontró el runETL.txt. Proceso abortado.")
            return

        logger.info("runETL encontrado. Iniciando ETL.")

        if not download_excel_sharepoint(ctx, excel_server_relative_url, local_excel):
            logger.error("Fallo al descargar el archivo.")
            return

        df, _ = extract_data(local_excel)
        df_transformado, _ = transform_data(df, fecha_vencimiento_param)
        load_time = load_data_pyodbc(df_transformado, server, database, trusted_connection)

        runETL = folder.files.get_by_url("runETL.txt")
        runETL.delete_object()
        ctx.execute_query()
        logger.info("ETL completado. Archivo runETL eliminado.")

    except Exception as e:
        logger.error(f"Error en el proceso ETL: {e}")

if __name__ == "__main__":
    main()