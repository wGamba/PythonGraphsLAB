from docx import Document
from docx.shared import Inches
import matplotlib.pyplot as plt

# -------- Gráfica 1 --------
x1 = [1, 2, 3, 4, 5, 6]  # m [g] (Eje X)
y1 = [8.65, 17.30, 25.95, 34.63, 43.31, 51.95]  # H [cm] (Eje Y)

plt.plot(x1, y1, marker='o')
plt.xticks(range(1, 8))  # 1 a 7
plt.yticks(range(10, 70, 10))  # 10 a 60
plt.xlabel('m [g]')
plt.ylabel('H [cm]')
plt.grid(True)
plt.savefig('imagenes/grafica1.png', dpi=300)
plt.close()

# -------- Insertar en DOCX --------
doc = Document('plantilla.docx')

for para in doc.paragraphs:
    if '<<GRAFICA1>>' in para.text:
        para.clear()
        para.add_run().add_picture('imagenes/grafica1.png', width=Inches(4))

doc.save('informePythonFrancoPrieto.docx')
