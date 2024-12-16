# Fugas de Tiempo

## Descripción

Este proyecto permite separar la información obtenida mediante los formularios de Google Forms en archivos individuales por estudiante y agrupar las respuestas de cada sección en su respectiva carpeta para facilitar el posterior envío de correos a los estudiantes. Además, coloriza las celdas de las respuestas de acuerdo a la relevancia de cada fuga de tiempo para facilitar la identificación de problemas para los estudiantes.

## Instrucciones de instalación y uso

1. Abrir la carpeta del proyecto en un editor de código como Visual Studio Code o PyCharm.
2. Crear un entorno virtual con `python -m venv .venv`.
3. Activar el entorno virtual según su sistema operativo.
   - Windows: `.\.venv\Scripts\activate`
   - Linux: `source ./.venv/bin/activate`
4. Instalar las dependencias con `pip install -r requirements.txt`.
5. Descargar el archivo de respuestas de Google Forms en formato `.xlsx`.
6. Renombrar el archivo a `respuestas.xlsx` y moverlo a la carpeta del proyecto.
7. Ejecutar el script con `python main.py`.
8. Revisar que se hayan generado los archivos para cada estudiantes en la carpeta `estudiantes`.

## ¿Qué hago ahora?

Envía un correo a cada estudiante con sus respuestas antes de la clase para que puedan revisarlas y preparar sus dudas.
