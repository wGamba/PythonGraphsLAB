# Proyecto de Análisis de Datos y Generación de Gráficas

Este proyecto tiene como objetivo la generación de gráficas a partir de datos experimentales, utilizando Python, la biblioteca `matplotlib` para las gráficas, y `python-docx` para insertar las gráficas en un archivo Word. El flujo del proyecto incluye la recopilación de información desde una planilla, la creación de gráficas y la sustitución de etiquetas en un documento de Word.

## Flujo del Proyecto

### 1. Recopilación de Datos
Los datos experimentales se recopilan desde una planilla que contiene valores en formato de tabla. La información se utiliza para generar las gráficas de acuerdo con los diferentes experimentos y transformaciones realizadas sobre los datos.

### 2. Generación de Gráficas
Se generan las siguientes gráficas a partir de los datos:

- **Gráfica 1**: Relación entre m [g] y H [cm].
- **Gráfica 2**: Relación entre D [cm] y m [g].
- **Gráfica 3**: Relación entre D² [cm²] y m [g].
- **Gráfica 4**: Relación logarítmica entre log(D [cm]) y log(m [g]).
- **Gráfica 5**: Relación entre D [cm] y m [g].
- **Gráfica 6**: Relación entre D² [cm²] y m [g].
- **Gráfica 7**: Relación entre D³ [cm³] y m [g].
- **Gráfica 8**: Relación logarítmica entre log(D [cm]) y log(m [g]).

Cada una de estas gráficas es generada utilizando la librería `matplotlib` y se guarda en una carpeta `imagenes/`.

### 3. Reemplazo de Etiquetas en el Documento
El siguiente paso es abrir un documento de Word predefinido (plantilla), en el que las gráficas se insertarán en las posiciones correspondientes. Las etiquetas `<<GRAFICA1>>`, `<<GRAFICA2>>`, etc., son reemplazadas por las imágenes generadas en el paso anterior.

El flujo es el siguiente:

- La plantilla de Word contiene marcadores de posición, como `<<GRAFICA1>>`, `<<GRAFICA2>>`, etc.
- El script de Python busca estos marcadores en el documento y reemplaza cada uno con la imagen correspondiente generada en el paso anterior.
- Al final, se guarda el documento con el nombre `informe_generado.docx`.

### 4. Herramientas y Bibliotecas
Este proyecto utiliza las siguientes bibliotecas de Python:

- `matplotlib`: Para la creación de gráficas.
- `python-docx`: Para la manipulación de documentos de Word.
- `numpy`: Para realizar operaciones matemáticas, como el logaritmo o el cuadrado de los valores.

## Uso

1. **Recopilar los datos**: Asegúrate de que los datos estén correctamente organizados en la planilla.
2. **Ejecutar el script**: El script generará las gráficas y las insertará en el archivo de Word.
3. **Verificar el resultado**: El archivo `informe_generado.docx` se guardará con las gráficas insertadas.

## Instalación

1. Clona el repositorio o descarga los archivos.
2. Instala las dependencias necesarias:
   ```bash
   pip install matplotlib python-docx numpy
