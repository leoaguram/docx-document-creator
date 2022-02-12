import docx
from docx.shared import Pt as size


file = open("DB.csv", "r")  # Abriendo archivo DB.csv
stri = file.readlines()  # Leyendo Archivo DB.csv
file.close()

nuevaLista = stri[1:]  # Copiando lista de csv a nueva lista desde elemento 1 al final

sublista = []  # Creando lista Auxiliar para eliminar salto de linea


for i in range(1, len(stri)):
    sublista.append(stri[i].rstrip("\n"))

for i in range(0, len(sublista)):
    doc = docx.Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = size(12)
    auxlist = sublista[i].split(",")
    srts = ""
    strs = "Hola, buenas tardes " + auxlist[0]
    strs += " con la edad de " + auxlist[1] + " años"
    strs += " se le notifica que usted ha sido seleccionado para la entrevista de trabajo, "
    strs += "favor de presentarse en " + auxlist[2]
    strs += " el dia " + auxlist[4]
    strs += " cualquier duda, comunicarse al " + auxlist[3]
    paragraph = doc.add_paragraph(strs)
    paragraph.alignment = 3
    doc.save(auxlist[0]+".docx")
