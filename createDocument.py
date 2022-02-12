#
# * Author: Aguirre Ramírez Leonardo
# * Date: 11-06-2020
# * Description: This program reads a csv file and create a docx documents with information
# * contained on the csv file.
# * This program can be used when you need create documents with similar information but
# * the name and personal information of the people to whom is adressed is different,
# * so you can create these documents of masively way with this script.
# * The library docx is implemented to read/write a docx file
#

import docx
from docx.shared import Pt as size


file = open("DB.csv", "r")  # Opening file DB.csv
stri = file.readlines()  # Reading file DB.csv
file.close()                #Closing file DB.csv

nuevaLista = stri[1:]  # Copying csv list to a new list from the first to the last element

sublista = []  # Creating aux list to delete line breaks


for i in range(1, len(stri)):
    sublista.append(stri[i].rstrip("\n"))  #The line breaks are deleted with rstrip

for i in range(0, len(sublista)):  #The number documents correspond to the number on information lines on the csv file
    doc = docx.Document()
    style = doc.styles['Normal'] #Setting normal style to docx document
    font = style.font
    font.name = 'Arial'  #Setting font name to document
    font.size = size(12) #Setting font size to document
    auxlist = sublista[i].split(",") #Dividing each element on a line into a list
    srts = "" #Creating a string
    strs = "Hola, buenas tardes " + auxlist[0]
    strs += " con la edad de " + auxlist[1] + " años"
    strs += " se le notifica que usted ha sido seleccionado para la entrevista de trabajo, "
    strs += "favor de presentarse en " + auxlist[2]
    strs += " el dia " + auxlist[4]
    strs += " cualquier duda, comunicarse al " + auxlist[3]
    paragraph = doc.add_paragraph(strs) #Adding paragraphs to document
    paragraph.alignment = 3
    doc.save(auxlist[0]+".docx") #Saving the document
