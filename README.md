## Descripción de la aplicación

Esta aplicación permite convertir documentos de Office (como Word y Excel) a archivos en formato Markdown (.md). Es útil para desarrolladores, escritores técnicos y cualquier persona que necesite transformar documentos de Office en un formato más adecuado para la web o sistemas de control de versiones.

## Pasos para la instalación y ejecución

1. Crear un entorno virtual de Python:
   ```bash
   python -m venv venv
   ```

2. Activar el entorno virtual:
   - En Windows:
     ```bash
     ./venv/Scripts/Activate.ps1
     ```
   - En macOS/Linux:
     ```bash
     source venv/bin/activate
     ```

3. Instalar las dependencias necesarias:
   ```bash
   pip install -r requirements.txt
   ```

4. Ejecutar el conversor:
   ```bash
   python conversor.py
   ```

5. Desactivar el entorno virtual al finalizar:
   ```bash
   deactivate
   ```

## Uso de la conversión de documentos

Esta herramienta utiliza la biblioteca [MarkItDown](https://github.com/microsoft/markitdown) para realizar la conversión de documentos de Office a Markdown. Asegúrate de que los documentos de entrada estén en un formato compatible antes de iniciar la conversión.