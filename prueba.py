'''
Descarga el xlsx de la pagina de la Poli
'''

#Libreria para hacer peticiones http/https
import requests

#se asigna el archivo .html en la variable
URL_Poli = requests.get('https://www.pol.una.py/?q=horario_clases')

#ubicar la posicion del href en donde se tiene el link de descarga del excel
indice_extension = URL_Poli.text.find('.xlsx')

#ubicar los indices de las comillas que encierran el link de descarga del excel
primeraComilla = URL_Poli.text.rfind('"', 0, indice_extension) + 1
segundaComilla = URL_Poli.text.find('"', indice_extension)

#se saca el link mediante los indices ubicados
link_descarga_excel = URL_Poli.text[ primeraComilla : segundaComilla ]

#se asigna el archivo a la variable
URL_Excel = requests.get(link_descarga_excel)

#se escribe en el disco
open('Horario.xlsx', 'wb').write(URL_Excel.content)

