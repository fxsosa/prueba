''' Descarga el xlsx de la pagina de la Poli '''

#Libreria para hacer peticiones http/https (instalacion $python -m pip install requests)
import requests

def obtener_horario():
    '''Funcion para extraer el link de descarga y descargar el horario'''
    link_descarga_excel = obtener_link()

    #se asigna el archivo a la variable
    sheets_poli = requests.get(link_descarga_excel)

    extension = obtener_extension(link_descarga_excel)

    #se escribe en el disco
    open(f'horario.{extension}', 'wb').write(sheets_poli.content)

def obtener_link():
    '''obtiene el link de descarga del sheets_horario'''
    #se asigna el archivo .html en la variable
    html_poli = requests.get('https://www.pol.una.py/?q=horario_clases')

    #ubicar la posicion del href en donde se tiene el link de descarga del excel
    indice_extension = html_poli.text.find('.xl')

    #ubicar los indices de las comillas que encierran el link de descarga del excel
    primera_comilla_indice = html_poli.text.rfind('"', 0, indice_extension) + 1
    segunda_comilla_indice = html_poli.text.find('"', indice_extension)

    #se saca el link mediante los indices ubicados
    return html_poli.text[ primera_comilla_indice : segunda_comilla_indice ]

def obtener_extension(link_archivo):
    '''Simple codigo para obtener la extension del sheets'''
    indice_extension = link_archivo.find('.')

    return link_archivo[ indice_extension + 1 ]
