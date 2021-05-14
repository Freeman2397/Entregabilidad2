from datetime import date
import os
import re

import DiccionarioPalabras
import editpyxl

informe_revision = True
numero_reporte = DiccionarioPalabras.numero_informe
DiccionarioPalabras.numero_informe += 1
lista_parseo_global = []
lista_reportes = []
lista_puntuaciones = []
lista_informacion = []
nombre_cliente = "Pablo Romero "
email_cliente = "pabloromero@brekiadata.com"
nombre_informe = "Informe " + nombre_cliente + str(date.today()) + ".xlsx"

# Estos tres primeros parámetros rigen la puntuación del analisis.
# Se dibidite el total de puntuació, por defecto 100, entre el total de funciones que se ejecutan.
# Cada funcioón que de un error se resta su parte a la nota final.
NumeroDeFunciones = 14  # Este parámetro hay que ajustarlo en función de las funciones que se ejecuten. Es sencillo contarlas en el método main.
Puntuacion = 100
NotaParaRestar = Puntuacion / NumeroDeFunciones
lista_posiciones = []
lista_ranking = []
lista_fuentes = []
lista_titulos = []
lista_errores = []


# INSTRUCCIONES
# Para usar el código hay que modificar tres ficheros .txt que se encuentran en la misma carpeta que este
# script. Cada fichero tiene un nombre indicativo y hay que poner lo que se debe en cada cual. Así mismo, en el fichero
# cargaAsunto.txt se coloca el asunto del email. En el fichero cargaCuerpo.txt se coloca el cuerpo de email. Por último
# en el fichero cargaHTML.txt habrá que poner el código html del email al completo. El resultado del analisis se escribe
# automaticamente en el fichero Informe + nombre usuario + fecha.


#       ***************************VERSION ACROBAT***********************************

# Documentos en dos pasos para el cliente, resultado inicial y diferencia con el final.
# Chequear todas las mayusculas y minusculas de las palabras. Tambien las tildes.
# Añadir al final de excel una explicación en profundidad de los valores del test.
# Escibir información adicional en la cabecera.
# Incorporar datos del cliente. Nombre + email.
# Incorporar fecha automática hecho
# Incorporar sitema de revisiones y veriones de ficheros. empezado
# Crear una hoja secundaria con datos para revisar. empezado

def leeTXT(fichero):
    archivo = open(fichero, "r")
    datos = archivo.read()
    archivo.close()
    return datos


def puntua(puntos):
    if puntos == 0:
        lista_puntuaciones.append("RatingIndA")
    elif puntos == 1:
        lista_puntuaciones.append("RatingIndB")
    elif puntos == 2:
        lista_puntuaciones.append("RatingIndC")
    elif puntos == 3:
        lista_puntuaciones.append("RatingIndD")
    elif puntos == 4:
        lista_puntuaciones.append("RatingIndE")
    elif puntos == 5:
        lista_puntuaciones.append("RatingIndF")
    elif puntos == 6:
        lista_puntuaciones.append("RatingIndG")
    else:
        lista_puntuaciones.append("RatingIndH")


def buscaExclamacionesIncorrectas(texto, lugar, filaRating, filaDato, titulo):
    totalIncorectas = 0
    global Puntuacion
    global lista_informacion

    for s in ("!!", "!!!", "!!!!", "!!!!!", "¡¡", "¡¡¡", "¡¡¡¡", "¡¡¡¡¡"):
        if s in texto:
            texto1 = "Se han encontrado señales de exclamación indebidas en el " + lugar
            lista_reportes.append(texto1)
            totalIncorectas += 1

    if totalIncorectas > 2:
        Puntuacion = Puntuacion - NotaParaRestar
        puntua(3)
    else:
        puntua(totalIncorectas)

    lista_informacion.append(totalIncorectas)

    cargaListas(filaRating, filaDato, titulo)
    cargaDatos(totalIncorectas)


def chequeaPalabrasGancho(cuerpo, filaRating, filaDato, titulo):
    global lista_informacion
    cuentaGancho = 0

    for i in DiccionarioPalabras.dic_palabras:
        if i.lower() in cuerpo.lower():
            cuentaGancho += 1
            texto = "Se ha encontrado una palabra indebida: " + i
            lista_reportes.append(texto)

    global Puntuacion

    if 3 > cuentaGancho < 6:
        puntua(4)
        Puntuacion = Puntuacion - NotaParaRestar / 2

    elif cuentaGancho > 6:

        Puntuacion = Puntuacion - NotaParaRestar
        puntua(6)
    else:
        puntua(cuentaGancho)

    lista_informacion.append(cuentaGancho)

    cargaListas(filaRating, filaDato, titulo)
    cargaDatos(cuentaGancho)


def chequeaTamañoCuerpo(cuerpo, filaRating, filaDato, titulo):
    global Puntuacion
    global lista_informacion
    if len(cuerpo) < 500:
        texto = "El tamaño del cuerpo es de " + str(len(cuerpo)) + " caracteres. Es un tamaño comedido y correcto."
        lista_reportes.append(texto)
        puntua(0)
    elif 500 < len(cuerpo) < 1500:
        texto = "El tamaño del cuerpo es de " + str(
            len(cuerpo)) + " caracteres. Es un tamaño medio pero no tiene porqué empeorar la entrega."
        lista_reportes.append(texto)
        puntua(1)
        Puntuacion = Puntuacion - NotaParaRestar / 3
    elif len(cuerpo) > 1500:
        texto = "El tamaño del cuerpo es de " + str(
            len(cuerpo)) + " caracteres. Es un tamaño muy grande y puede empeorar la entrega."
        lista_reportes.append(texto)
        puntua(3)
        Puntuacion = Puntuacion - NotaParaRestar

        lista_informacion.append(len(cuerpo))
    cargaListas(filaRating, filaDato, titulo)
    cargaDatos(len(cuerpo))


def cuentaMayusculasYMinusculas(texto, parte, filaRating, filaDato, titulo):
    indice = 0
    mayusculas = 0
    minusculas = 0
    global Puntuacion
    global lista_informacion
    while indice < len(texto):
        letra = texto[indice]
        if letra.isupper():
            mayusculas += 1
        else:
            minusculas += 1
        indice += 1

    if mayusculas > (minusculas / 2):
        text1 = "El " + parte + " tiene más de la mitad de letras en mayusculas, se debe corregir."
        lista_reportes.append(text1)
        Puntuacion = Puntuacion - NotaParaRestar / 4
        puntua(2)

    elif mayusculas > (minusculas / 4):
        text1 = "El " + parte + " tiene más de un cuarto de letras en mayusculas, es una cantidad un poco alta."
        lista_reportes.append(text1)
        Puntuacion = Puntuacion - NotaParaRestar / 2
        puntua(4)
    else:
        puntua(0)

    lista_informacion.append(minusculas / mayusculas)
    cargaListas(filaRating, filaDato, titulo)
    dato = str((mayusculas / minusculas) * 100) + "%"
    cargaDatos(dato)


def chequeaTamanoAsunto(asunto, filaRating, filaDato, titulo):
    global Puntuacion
    global lista_informacion
    if 4 < len(asunto) <= 15:
        lista_reportes.append(
            "Ratio de apertura por tamaño de cabecera de 15.2%, ratio de click de 3.1%. MEJOR APERTURA.")
        puntua(0)
    elif 16 < len(asunto) <= 27:
        lista_reportes.append("Ratio de apertura por tamaño de cabecera de 11.6%, ratio de click de 3.8%. CASO MEDIO.")
        puntua(2)
    elif 28 < len(asunto) <= 39:
        lista_reportes.append(
            "Ratio de apertura por tamaño de cabecera de 12.2%, ratio de click de 4.0%. MEJOR RATIO DE CLICKS.")
        puntua(0)
    elif 40 < len(asunto) <= 50:
        lista_reportes.append("Ratio de apertura por tamaño de cabecera de 11.9%, ratio de click de 2.8%. CASO MEDIO.")
        puntua(2)
    elif len(asunto) > 51:
        lista_reportes.append("Ratio de apertura por tamaño de cabecera de 10.4%, ratio de click de 1.8%. PEOR CASO.")
        puntua(4)
        Puntuacion = Puntuacion - NotaParaRestar / 2
    cargaListas(filaRating, filaDato, titulo)
    cargaDatos(len(asunto))


def chequeaTbody(html, filaRating, filaDato, titulo):
    global lista_informacion
    etiquetas = re.findall('tbody', html)
    print("Tbody: ", len(etiquetas))

    if len(etiquetas) > 0:
        global Puntuacion
        Puntuacion = Puntuacion - NotaParaRestar / 2
        puntua(2)

        lista_reportes.append("Se ha encontrado una etiqueta <tbody>, no deberia estar, "
                              "ya que hace referencia a tablas que pueden disminuir la calidad de la entrega.")
    else:
        puntua(0)

    cargaListas(filaRating, filaDato, titulo)
    cargaDatos(len(etiquetas))


def chequeaForm(html, filaRating, filaDato, titulo):
    global lista_informacion
    etiquetas = re.findall('<form', html)
    print("Form: ", len(etiquetas))

    if len(etiquetas) > 0:
        global Puntuacion
        Puntuacion = Puntuacion - NotaParaRestar / 2
        puntua(2)
        lista_reportes.append("Se ha encontrado un formulario, no deberia estar.")
    else:
        puntua(0)

    cargaListas(filaRating, filaDato, titulo)
    cargaDatos(len(etiquetas))


def chequeaMargenImagen(html, filaRating, filaDato, titulo):
    global lista_informacion
    if "img border" in html:
        if "img border=0" not in html:
            global Puntuacion
            Puntuacion = Puntuacion - NotaParaRestar
            puntua(4)
            lista_reportes.append("Se ha encontrado un borde de imagen incorrecto.")
    else:
        puntua(0)

    cargaListas(filaRating, filaDato, titulo)
    cargaDatos(" ")


def chequeaUNSUBSCRIBE(html, filaRating, filaDato, titulo):
    global lista_informacion
    html.lower()

    if any(s in html for s in ("desuscribir", "desuscribirse", "darse de baja", "unsubscribe")):

        puntua(0)
        cargaDatos(1)
    else:
        global Puntuacion
        Puntuacion = Puntuacion - NotaParaRestar
        puntua(4)
        cargaDatos(0)
        lista_reportes.append("Necesita añadir una forma de desuscribirse.")

    cargaListas(filaRating, filaDato, titulo)


def chequeaColores(html, filaRating, filaDato, titulo):
    global lista_informacion
    if "#FF0000" in html:
        global Puntuacion
        Puntuacion = Puntuacion - NotaParaRestar
        puntua(4)
        cargaDatos("1")
        lista_reportes.append(
            "Se está haciendo uso de un color rojo en alguna parte del email. Esto empeora la entrega.")
    else:
        puntua(0)

    cargaListas(filaRating, filaDato, titulo)
    cargaDatos(" ")


def cuentaImagenes(html, filaRating, filaDato, titulo):
    global lista_informacion
    cantidadImages = str(html.count("<img"))
    texto = "Aparecen " + cantidadImages + " imagenes"
    lista_reportes.append(texto)
    if int(cantidadImages) > 4:
        puntua(2)
    else:
        puntua(0)

    cargaListas(filaRating, filaDato, titulo)
    cargaDatos(cantidadImages)


def parseaStringNoOptimo(cadena):
    ListaParsea = []
    global lista_informacion
    while cadena.find("font") > 0:

        cadena = cadena.partition("font")[2]
        indice2 = cadena.partition(">")[0]
        indice3 = cadena.partition("}")[0]

        if len(indice2) < len(indice3):
            ListaParsea.append(indice2)
        else:
            ListaParsea.append(indice3)

    return ListaParsea


def parseaStringPorInicioYFinal(cadena, inicio, final):
    ListaParsea = []
    global lista_informacion

    while cadena.find(inicio) > 0:
        cadena = cadena.partition(inicio)[2]
        cadenaFinal = cadena.partition(final)[0]
        ListaParsea.append(cadenaFinal)

    return ListaParsea


def chequeaFuente(html, filaRating, filaDato, titulo):
    TotalFuentes = 0
    TotalFuentesIncorrectas = 0
    global lista_informacion

    html = html.lower()

    divisor = 0
    for cadena in lista_parseo_global:

        if divisor == 0:
            for fuenteCorrecta in DiccionarioPalabras.dic_fuentes:
                if fuenteCorrecta in cadena:
                    TotalFuentes += 1
                    divisor = 1

        if divisor == 0:
            for fuenteIncorrecta in lista_fuentes:

                if fuenteIncorrecta in cadena:
                    TotalFuentesIncorrectas += 1

        divisor = 0

    lista_informacion.append(TotalFuentesIncorrectas)
    print("Incorrectas", TotalFuentesIncorrectas)
    print("Correctas", TotalFuentes)

    if TotalFuentes == 0:
        global Puntuacion
        Puntuacion = Puntuacion - NotaParaRestar
        puntua(4)
        lista_reportes.append("No hemos encontrado ninguna fuente segura, por favor cambie la fuente.")
    else:
        puntua(0)

    cargaListas(filaRating, filaDato, titulo)
    cargaDatos(TotalFuentesIncorrectas)


def chequeaJavaScript(html, filaRating, filaDato, titulo):
    global Puntuacion
    global lista_informacion
    if "javascript" in html:
        Puntuacion = Puntuacion - NotaParaRestar / 2
        puntua(2)
        lista_reportes.append("Hemos encontrados referencias a javascript en el html. No deberian estar.")
    else:
        puntua(0)
    cargaListas(filaRating, filaDato, titulo)
    cargaDatos(html.count("javascript"))


def chequeaFlash(html, filaRating, filaDato, titulo):
    global Puntuacion
    global lista_informacion
    if "flash" in html:
        Puntuacion = Puntuacion - NotaParaRestar / 2
        puntua(2)
        lista_reportes.append("Hemos encontrados referencias a flash en el html. No deberian estar.")
    else:
        puntua(0)

    cargaListas(filaRating, filaDato, titulo)
    cargaDatos(html.count("flash"))


def escribeLog():
    file = open("resultadoHTML_Email_Check.txt", "w")
    for reporte in lista_reportes:
        file.write(reporte + os.linesep)

    file.close()


def cargaXLSX():
    # Escribir resultados ampliados en la columna I, u otra fuera de impresión.

    """wb = load_workbook('Plantilla Entregabilidad Newsletter.xlsx')
    ws = wb['Informe']
    #df = pd.read_excel('file.xlsx', sheet_name='sheet1')
    for index in range(0,len(lista_puntuaciones)):

        cell1 = 'E%d' % (lista_ranking[index])
        ws[cell1] = lista_puntuaciones[index]


    wb.save("Informe Entregabilidad.xlsx")
"""
    wb = editpyxl.Workbook()
    source_filename = r'Plantilla Entregabilidad Newsletter.xlsx'

    wb.open(source_filename)
    ws = wb.active
    for index, puntuacion in enumerate(lista_puntuaciones):
        cell1 = 'E%d' % (lista_ranking[index])
        ws.cell(cell1).value = str(puntuacion)
        cell2 = 'D%d' % (lista_posiciones[index])
        ws.cell(cell2).value = str(lista_errores[index])

    ws.cell('A15').value = int(Puntuacion) / 100
    ws.cell('A15').number_format = '0.00%'
    ws.cell('A9').value = nombre_cliente
    ws.cell('A12').value = email_cliente
    ws.cell('G9').value = date.today()
    ws2 = wb.get_sheet_by_name("Errores")
    indiceHoja1 = 0

    for i in lista_titulos:
        cellT = 'A%d' % (indiceHoja1 + 1)
        ws2[cellT] = i
        indiceHoja1 += 1

        """for j in lista_errores:
            for y in j:
                cellT = 'A%d' % (indiceHoja1)
                indiceHoja1 += 1
                ws2[cellT] = y
        cellT = 'A%d' % (indiceHoja1)
        indiceHoja1 += 1
        ws2[cellT] = """""

    wb.save(nombre_informe)
    wb.close()


def cargaFuentes():
    global lista_fuentes
    f = open('Fuentes texto.txt', 'r')
    mensaje = f.read()
    lista_fuentes = []
    f.close()
    lista_fuentes.append(mensaje)
    parseo = lista_fuentes[0].split("\n")
    for i in parseo:
        if not i.isnumeric():
            lista_fuentes.append(i.lower())


def parseoGlobal(html):
    global lista_parseo_global
    ListaAux = []
    lista_parseo_global = parseaStringPorInicioYFinal(html, "<", ">")
    for i in lista_parseo_global:
        if len(i) < 7:
            pass
        else:
            ListaAux.append(i.lower())
    lista_parseo_global = ListaAux


def cargaDatos(dato):
    global lista_errores
    lista_errores.append(dato)


def cargaListas(filaRating, filaDato, titulo):
    global lista_ranking
    global lista_posiciones
    global lista_titulos

    lista_ranking.append(filaRating)
    lista_posiciones.append(filaDato)
    lista_titulos.append(titulo)


def mideTamañoFichero(fichero):
    sizefile = os.stat(fichero).st_size
    print("El tamaño de ", fichero, " es de ", sizefile)
    return sizefile


if __name__ == '__main__':
    html = leeTXT("cargaHTML.txt")
    cuerpo = leeTXT("cargaCuerpo.txt")
    asunto = leeTXT("cargaAsunto.txt")

    parseoGlobal(html)

    cargaFuentes()
    chequeaForm(html, 22, 23, "A01 Uso de formularios")
    chequeaMargenImagen(html, 24, 25, "A02 Margenes de las imágenes")
    chequeaTbody(html, 27, 29, "A03-Etiqueta tbody")
    chequeaColores(html, 30, 31, "A04 Uso de colores")
    chequeaFuente(html, 32, 33, "A05 Uso de fuentes de texto alternativas")
    cuentaImagenes(html, 34, 36, "A06 Cantidad de imágenes")
    chequeaJavaScript(html, 37, 38, "A07 Uso de javascript")
    chequeaFlash(html, 39, 40, "A08 Uso de flash")
    cuentaMayusculasYMinusculas(cuerpo, "cuerpo", 41, 42, "B01 Mayúsculas en el cuerpo")
    buscaExclamacionesIncorrectas(cuerpo, "cuerpo", 43, 44, "B02 Exclamaciones en el cuerpo")
    chequeaUNSUBSCRIBE(cuerpo, 45, 46, "B03 Boton de desuscripción")
    chequeaPalabrasGancho(cuerpo, 47, 48, "B04 Uso de palabras gancho")
    chequeaTamañoCuerpo(cuerpo, 49, 50, "B05 Tamaño del cuerpo")
    cuentaMayusculasYMinusculas(asunto, "asunto", 51, 52, "C01*Mayúsculas en el asunto")
    buscaExclamacionesIncorrectas(asunto, "asunto", 53, 54, "C02 Exclamaciones en el asunto")
    chequeaTamanoAsunto(asunto, 55, 56, "C03 Tamaño del asunto")

    print(lista_puntuaciones)

    texto = "La puntuación del analisis es de " + str(Puntuacion) + " sobre 100"
    lista_reportes.append(texto)
    print(len(lista_posiciones), len(lista_informacion), len(lista_ranking), len(lista_titulos),
          len(lista_puntuaciones), len(lista_errores))
    cargaXLSX()
    escribeLog()
    mideTamañoFichero(nombre_informe)
