Proyecto de Automatización de Email y Seguimiento

🚀 Resumen del Proyecto

Este proyecto es una solución de automatización diseñada para simplificar y optimizar el proceso de envío de correos electrónicos masivos y personalizados (Mail Merge) a una lista de contactos definida en un archivo de Excel (.xlsx). La herramienta garantiza la personalización del mensaje y, fundamentalmente, implementa un sistema de seguimiento para evitar envíos duplicados.

Tecnología: Python 3.x
Librerías Clave: pandas, python-dotenv, smtplib.

⚙️ Características Principales

Personalización y Seguimiento: Personaliza mensajes con campos de Excel (Empresa, Contacto). Utiliza la columna Enviado para gestionar el seguimiento y prevenir duplicados, actualizando el estado de NO a SI tras el envío.

Seguridad y Conexión: Las credenciales se manejan de forma segura a través del archivo .env (requiere App Password de Google). La autenticación SMTP utiliza smtplib con TLS para una conexión cifrada y segura con Gmail.

📦 Estructura del Proyecto

El proyecto sigue una estructura modular para facilitar la gestión de archivos y plantillas:

/SSAEED (Directorio Raíz)
├── .env                          # Variables de entorno y credenciales (IGNORAR en Git)
├── script.py                     # Lógica principal y funciones de envío/carga
├── Template/                     # Directorio para plantillas de email
│   └── email_template.txt        # Plantilla del cuerpo del correo
└── ExcelData/                    # Directorio para bases de datos
    └── data.xlsx                 # Archivo de contactos con seguimiento



🛠 Instalación y Configuración

Siga estos pasos para configurar y ejecutar el script:

1. Requisitos de Python

Instale las librerías necesarias dentro de su entorno virtual:

pip install pandas openpyxl python-dotenv


2. Configuración de Credenciales (.env)

Cree y complete el archivo .env en la raíz. Importante: GMAIL_PASSWORD debe ser una Contraseña de Aplicación de 16 caracteres de Google.

GMAIL_USER="tu_correo_de_envio@gmail.com"
GMAIL_PASSWORD="tu_clave_de_aplicacion_16_caracteres"
SMTP_SERVER="smtp.gmail.com"
SMTP_PORT=587
ARCHIVO_DATOS="ExcelData/data.xlsx"
PLANTILLA_EMAIL="Template/email_template.txt"


3. Preparación del Archivo de Datos (data.xlsx)

El Excel debe contener las siguientes columnas exactas para la correcta personalización y seguimiento:

Columna

Propósito

Ejemplo de Dato

Empresa

Personalización

Tech Innovators S.A.

Contacto

Personalización

Juan Pérez

Email_Destino

Dirección de envío

juan.perez@techin.com

Enviado

Seguimiento (NO / SI)

NO / SI

▶️ Uso del Script

Asegúrese de que el entorno virtual esté activo.

Verifique que el archivo data.xlsx no esté abierto para evitar errores al intentar escribir en él.

Ejecute el script desde la terminal en el directorio raíz:

python script.py



Salida Esperada

El script generará un log en la consola indicando el estado de cada contacto:

Saltando contacto Marcos Gómez (FutureCorp): Ya fue enviado.
Éxito: Email enviado a carlos.ruiz@softlabs.com
Éxito: Email enviado a elena.diaz@innovatech.com

PROCESO COMPLETADO: El archivo de seguimiento ha sido actualizado y guardado.



- Solución de Problemas Comunes

Error: KeyError: 'Email_Destino'
Causa Probable: Encabezados de Excel inconsistentes (mayúsculas/espacios).
Solución: Asegúrese de que las columnas en data.xlsx coincidan exactamente con las claves en el script.


Error: 535 5.7.8 Username and Password not accepted
Causa Probable: La GMAIL_PASSWORD es incorrecta o no es App Password.
Solución: Genere una nueva Contraseña de Aplicación (16 caracteres) y actualice el .env.


Error: FileNotFoundError
Causa Probable: Error en las rutas relativas.
Solución: Verifique que ARCHIVO_DATOS y PLANTILLA_EMAIL en el .env apunten a las carpetas correctas.# SSAEED
# SSAEED
