# Imprimir Etiquetas V1

## Descripción del Proyecto
Esta aplicación permite monitorear una carpeta específica para detectar archivos ZIP con un nombre específico ("Etiquetas - "). Al detectar un archivo ZIP, la aplicación extrae su contenido (archivos PDF) y los envía a imprimir a la impresora seleccionada por el usuario. El programa se puede configurar para usar **Adobe Acrobat Reader** como el software de impresión.

## Instrucciones de Uso

### Requisitos Previos
- **Python 3.x** debe estar instalado en el sistema.
- **Adobe Acrobat Reader** debe estar instalado y accesible en el sistema.

### Archivos incluidos:
- `app_imprenta.py`: Código fuente principal de la aplicación.
- `config.ini`: Archivo de configuración que guarda la ruta de **Adobe Acrobat Reader**.
- `logo_atmosfera_sport.jpg`: Imagen del logo que se muestra en la interfaz gráfica.
- `README.txt`: Archivo de documentación con instrucciones de uso.
- `impresion_etiquetas.log`: Archivo de log que guarda los registros de errores.

### Configuración y Uso:
1. **Abrir la aplicación**:
   - Ejecuta `app_imprenta.py` directamente usando Python o abre `app_imprenta.exe` si has convertido el script en una aplicación de escritorio.

2. **Primera ventana (Configuración inicial)**:
   - **Seleccionar carpeta**: Haz clic en el botón "Seleccionar carpeta" y elige la carpeta que deseas monitorear.
   - **Escribir el nombre de la impresora**: Escribe el nombre exacto de la impresora que deseas usar. Asegúrate de que el nombre coincida con el nombre registrado en tu sistema operativo.
   - **Ruta de Adobe Acrobat Reader**: Si es la primera vez que usas la aplicación, se te pedirá que selecciones la ubicación del archivo `AcroRd32.exe` de **Adobe Acrobat Reader**. Esta ruta se guardará en el archivo `config.ini` para usos futuros.

3. **Segunda ventana (Estado de la impresión)**:
   - Esta ventana muestra los registros de las acciones realizadas, como detección de archivos ZIP, extracción de archivos PDF y estado de impresión.

### Cómo convertir `app_imprenta.py` en una aplicación de escritorio:
Para convertir el script Python en un ejecutable, sigue estos pasos:

1. **Instalar PyInstaller**:
   Asegúrate de tener **PyInstaller** instalado. Si no lo tienes, instálalo usando:
   ```bash
   pip install pyinstaller
   ```

2. **Crear el ejecutable**:
   Ejecuta el siguiente comando en la terminal desde el directorio donde se encuentra `app_imprenta.py`:
   ```bash
   pyinstaller --onefile --windowed --add-data "logo_atmosfera_sport.jpg;." app_imprenta.py
   ```

   - `--onefile`: Genera un único archivo ejecutable.
   - `--windowed`: Evita que se abra la consola al ejecutar la aplicación.
   - `--add-data`: Asegura que el archivo `logo_atmosfera_sport.jpg` se incluya en el ejecutable.

3. **Archivos generados**:
   Los archivos se generarán en la carpeta `dist`. El archivo ejecutable `app_imprenta.exe` estará dentro de esa carpeta.

4. **Distribuir la aplicación**:
   Copia el archivo `app_imprenta.exe` y el archivo `config.ini` (si es necesario) a la ubicación de tu elección. No es necesario copiar los archivos de la carpeta `build` o el archivo `.spec`.

## Mantenimiento y Actualizaciones
- Si necesitas cambiar la ruta de **Adobe Acrobat Reader**, elimina el archivo `config.ini` y vuelve a ejecutar la aplicación para que te pida seleccionar la ruta nuevamente.

## Registro de Errores
Los errores se registran en el archivo `impresion_etiquetas.log`, que se encuentra en el mismo directorio que `app_imprenta.py`. Revisa este archivo si encuentras algún problema.

## Contacto
Para preguntas contacta con https://www.linkedin.com/in/carlos-luis-fern%C3%A1ndez-santana/
