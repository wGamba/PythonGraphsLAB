import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches

# -------- Gráfica 1 --------
x1 = [1, 2, 3, 4, 5, 6]
y1 = [8.65, 17.30, 25.95, 34.63, 43.31, 51.95]

plt.plot(x1, y1, marker='o')
plt.xticks(range(1, 8))
plt.yticks(range(10, 70, 10))
plt.xlabel('m [g]')
plt.ylabel('H [cm]')
plt.grid(True)
plt.savefig('imagenes/grafica1.png', dpi=300)
plt.close()

# -------- Gráfica 2 --------
x2 = [1, 2, 3, 4, 5, 6]
y2 = [1.22, 4.90, 10.40, 19.52, 30.71, 43.75]

plt.plot(x2, y2, marker='o')
plt.xticks(range(1, 8))
plt.yticks(range(10, 70, 10))
plt.xlabel('D [cm]')
plt.ylabel('m [g]')
plt.grid(True)
plt.savefig('imagenes/grafica2.png', dpi=300)
plt.close()

# -------- Insertar en DOCX --------
doc = Document('plantilla.docx')

for para in doc.paragraphs:
    if '<<GRAFICA1>>' in para.text:
        para.clear()
        para.add_run().add_picture('imagenes/grafica1.png', width=Inches(4))
    if '<<GRAFICA2>>' in para.text:
        para.clear()
        para.add_run().add_picture('imagenes/grafica2.png', width=Inches(4))

doc.save('informePythonFrancoPrieto.docx')
