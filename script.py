import pandas as pd
import smtplib
from dotenv import load_dotenv 
import os 
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- CONFIGURACIÓN Y CARGA DE VARIABLES DE ENTORNO ---
load_dotenv() 

# Cargamos las variables, convirtiendo el puerto a entero
GMAIL_USER = os.getenv("GMAIL_USER")
GMAIL_PASSWORD = os.getenv("GMAIL_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT"))
ARCHIVO_DATOS = os.getenv("ARCHIVO_DATOS")
PLANTILLA_EMAIL = os.getenv("PLANTILLA_EMAIL")


def cargar_plantilla(ruta_plantilla):
    """
    Carga y retorna el contenido de la plantilla de email desde un archivo de texto.
    """
    try:
        # Usamos 'with open' para abrir el archivo en modo lectura ('r') y asegurar que se cierre.
        with open(ruta_plantilla, 'r', encoding='utf-8') as archivo:
            contenido = archivo.read()
        return contenido
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo de plantilla en la ruta: {ruta_plantilla}")
        return None
    except Exception as e:
        print(f"Error al leer la plantilla: {e}")
        return None


def enviar_email(destinatario, asunto, cuerpo):
    """
    Establece la conexión SMTP, crea el mensaje MIME y lo envía.
    Retorna True si el envío fue exitoso, False en caso contrario.
    """
    
    # 1. Crear el objeto MIMEMultipart
    mensaje = MIMEMultipart()
    mensaje['From'] = GMAIL_USER 
    mensaje['To'] = destinatario
    mensaje['Subject'] = asunto

    # 2. Adjuntar el cuerpo del mensaje como texto plano
    mensaje.attach(MIMEText(cuerpo, 'plain')) 

    try:
        # 3. Conectar al servidor SMTP de Gmail (Puerto 587)
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            
            # Iniciar la seguridad TLS (Transport Layer Security)
            server.starttls()
            
            # 4. Iniciar sesión con las credenciales (App Password)
            server.login(GMAIL_USER, GMAIL_PASSWORD)
            
            # 5. Enviar el correo
            server.sendmail(GMAIL_USER, destinatario, mensaje.as_string())
            
            print(f"Éxito: Email enviado a {destinatario}")
            return True
            
    except Exception as e:
        # 6. Manejo de errores de envío o autenticación
        print(f"Fallo al enviar el email a {destinatario}. Error: {e}")
        return False


def procesar_datos_y_enviar():
    """
    Lógica principal: lee Excel, itera, personaliza, envía y actualiza el seguimiento en el Excel.
    """
    
    # --- Preparación ---
    asunto = "Propuesta de Colaboración" # Asunto fijo
    cuerpo_plantilla = cargar_plantilla(PLANTILLA_EMAIL)
    
    if cuerpo_plantilla is None:
        return # Sale si la plantilla no se cargó

    try:
        # Carga el archivo de Excel usando la ruta definida en .env
        df = pd.read_excel(ARCHIVO_DATOS, sheet_name=0, engine='openpyxl')
        
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo de datos en la ruta: {ARCHIVO_DATOS}")
        return
    except Exception as e:
        print(f"Error al cargar el Excel: {e}")
        return

    # --- Bucle de Envío ---
    for indice, fila in df.iterrows():
        
        # 1. Control de Seguimiento: Ignorar si ya se envió
        if fila.get('Enviado') == 'SI':
            print(f"Saltando contacto {fila['Contacto']} ({fila['Empresa']}): Ya fue enviado.")
            continue
            
        # 2. Personalización
        try:
            # Reemplaza los placeholders en la plantilla (Asumimos columnas: Contacto, Empresa, Email_Destino)
            cuerpo_personalizado = cuerpo_plantilla.format(
                Contacto=fila['Contacto'], 
                Empresa=fila['Empresa']
            )
        except KeyError as e:
            print(f"Error de formato: Falta la columna {e} en el Excel o en la plantilla. Saltando fila.")
            continue

        # 3. Envío
        destinatario = fila['Email_Destino']
        envio_exitoso = enviar_email(destinatario, asunto, cuerpo_personalizado)
        
        # 4. Actualización del Seguimiento
        if envio_exitoso:
            # Actualiza el DataFrame en memoria para marcar como enviado
            df.loc[indice, 'Enviado'] = 'SI' 
            
    # --- Finalización y Guardado ---
    try:
        # Sobreescribe el archivo original con los cambios de seguimiento
        df.to_excel(ARCHIVO_DATOS, index=False, engine='openpyxl')
        print("\nPROCESO COMPLETADO: El archivo de seguimiento ha sido actualizado y guardado.")
    except Exception as e:
        print(f"ADVERTENCIA: No se pudo guardar el archivo de Excel. Asegúrese de que no esté abierto. Error: {e}")


if __name__ == "__main__":
    procesar_datos_y_enviar()