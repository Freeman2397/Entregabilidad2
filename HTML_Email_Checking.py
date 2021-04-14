import os
import DiccionarioPalabras
ListaDeReportes = []

#INSTRUCCIONES
# Para usar el código simplemente hay que modificar tres ficheros .txt que se ecuentran en la misma carpeta que este
# script. Cada fichero tienen un nombre indicativo y hay que poner lo que se debe en cada cual. Así mismo, en el fichero
# cargaAsunto.txt se coloca el asunto del email. En el fichero cargaCuerpo.txt se coloca el cuerpo de email. Por último
# en el fichero carga HTML habrá que poner el código html del email al completo. El resultado del analisis se escribe
# automaticamente en el fichero resultadoHTML_Email_Check.txt

def leeTXT(fichero):
    
    archivo = open(fichero, "r")
    datos = archivo.read()
    archivo.close()
    return datos

def chequeaPalabrasGancho(cuerpo):
    for i in DiccionarioPalabras.dic_palabras:
        if i in cuerpo:
            texto = "Se ha encontrado una palabra indebida: " + i
            ListaDeReportes.append(texto)

def chequeaUNSUBSCRIBE(html):
    
    if "desuscribir" or "desuscribirse" or "darse de baja" or "unsubscribe" not in html:
        ListaDeReportes.append("Necesita añadir una forma de desuscribirse.")

def cuentaMayusculasYMinusculas(texto,parte):
    
    indice = 0
    mayusculas = 0
    minusculas = 0

    while indice < len(texto):
        letra = texto[indice]
        if letra.isupper() == True:
            mayusculas += 1
        else:
            minusculas += 1
        indice += 1

    if mayusculas>(minusculas/2):
        text1 = "El " + parte + " tiene más de la mitad de letras en mayusculas, se debe corregir."
        ListaDeReportes.append(text1)

    elif mayusculas>(minusculas/4):
        text1 = "El " + parte + " tiene más de un cuarto de letras en mayusculas, es una cantidad un poco alta."
        ListaDeReportes.append(text1)

def chequeaTamañoAsunto(asunto):
    if 4< len(asunto)<=15:
        ListaDeReportes.append("Ratio de apertura por tamaño de cabecera de 15.2%, ratio de click de 3.1%. MEJOR APERTURA.")
    elif 16< len(asunto)<=27:
        ListaDeReportes.append("Ratio de apertura por tamaño de cabecera de 11.6%, ratio de click de 3.8%. CASO MEDIO.")
    elif 28< len(asunto)<=39:
        ListaDeReportes.append("Ratio de apertura por tamaño de cabecera de 12.2%, ratio de click de 4.0%. MEJOR RATIO DE CLICKS.")
    elif 40< len(asunto)<=50:
        ListaDeReportes.append("Ratio de apertura por tamaño de cabecera de 11.9%, ratio de click de 2.8%. CASO MEDIO.")
    elif len(asunto)>51:
        ListaDeReportes.append("Ratio de apertura por tamaño de cabecera de 10.4%, ratio de click de 1.8%. PEOR CASO.")

def chequeaTbody(html):
    
    if "tbody" in html:
        ListaDeReportes.append("Se ha encontrado una etiqueta <tbody>, no deberia estar, ya que hace referencia a tablas que pueden disminuir la calidad de la entrega.")

def chequeaForm(html):
    
    if "<form" in html:
        ListaDeReportes.append("Se ha encontrado un formulario, no deberia estar.")

def chequeaMargenImagen(html):
    
    if "img border" in html:
        if "img border=0" not in html:
            ListaDeReportes.append("Se ha encontrado un borde de imagen incorrecto.")

def chequeaColores(html):

    if "#FF0000" in html:
        ListaDeReportes.append("Se está haciendo uso de un color rojo en alguna parte del email. Esto empeora la entrega.")

def cuentaImagenes(html):
    texto = "Aparecen " + str(html.count("<img")) + " imagenes"
    ListaDeReportes.append(texto)

def chequeaFuente(html):
    TotalFuentes = 0
    for fuente in DiccionarioPalabras.dic_fuentes:

        if fuente in html:
            TotalFuentes =+ 1
            texto = "Hemos encontrado la fuente " + fuente + ", es una fuente segura"
            ListaDeReportes.append(texto)

    if TotalFuentes == 0:
        ListaDeReportes.append("No hemos encontrado ninguna fuente segura, por favor cambie la fuente.")
def escribeLog():

    file = open("resultadoHTML_Email_Check.txt", "w")
    for reporte in ListaDeReportes:
        file.write(reporte + os.linesep)


    file.close()

if __name__ == '__main__':
    html = leeTXT("cargaHTML.txt")
    cuerpo = leeTXT("cargaCuerpo.txt")
    asunto = leeTXT("cargaAsunto.txt")

    chequeaForm(html)
    chequeaMargenImagen(html)
    chequeaTbody(html)
    chequeaColores(html)
    chequeaFuente(html)
    cuentaImagenes(html)
    cuentaMayusculasYMinusculas(cuerpo,"cuerpo")
    cuentaMayusculasYMinusculas(asunto,"asunto")
    chequeaUNSUBSCRIBE(cuerpo)
    chequeaPalabrasGancho(cuerpo)
    chequeaTamañoAsunto(asunto)
    escribeLog()
