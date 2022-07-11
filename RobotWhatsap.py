from cgi import print_arguments
import webbrowser as web
import time
import pandas as pd
import pyautogui as pg
import pymysql

data = pd.read_excel('D:\Marco\CodigosPython\RobotWhatsapp\contactos.xlsx')
data_dict = data.to_dict('list')
celulares = data_dict['celular']
placaGenerales = data_dict['placaGeneral']
fechas = data_dict['fecha']
combo = zip(celulares, placaGenerales, fechas)
combo2 = zip(celulares, placaGenerales, fechas)
first = True
contador = 0
total = len(list(combo2))

for celular, placaGeneral, fecha in combo:

    contador = contador + 1
    print("total :" + str(total) + " contador :" + str(contador) + " placa :" + str(placaGeneral) + " celular :" + str(celular))

    miConexion = pymysql.connect( host='localhost', user= 'root', passwd='', db='confianza_general' )
    cur = miConexion.cursor()
    consulta = "SELECT placa,id_subagente_venta FROM t_datos_soat WHERE placa='" + str(placaGeneral) + "'"
    cur.execute(consulta)
    result = cur.fetchone()
    if result is not None:
        consulta2 = "SELECT id_subagente FROM t_subagentes_ventas WHERE id='" + str(result[1]) + "'"
        cur.execute(consulta2)
        result2 = cur.fetchone()
        consulta3 = "SELECT * FROM t_subagentes WHERE id='" + str(result2[0]) + "'"
        cur.execute(consulta3)
        result3 = cur.fetchone()
        mensaje_final = "¡Hola!👋🏻 Somos *CONFIANZA Y VIDA*👩🏻‍💼Asesores y Corredores de Seguros, te recordamos que el día *" + str(fecha)+ "* vence tu SOAT de la placa *" + str(placaGeneral)+"* puedes renovarlo en el Punto de Venta donde lo adquiriste%0A*Nombre de PV :* "+str(result3[3].upper())+"%0A*Dirección :* "+str(result3[6].upper())+"%0A*Contacto Celular :* "+str(result3[5].upper())+"%0Ao te puedes comunicar con nosotros y gustosamente te atenderemos.☺️Evita multas, sanciones o retención de tu vehículo 👮🏻‍♂️🚘🏍️"
        web.open("whatsapp://send?phone=51"+str(celular)+"&text="+str(mensaje_final))
        time.sleep(12)
        width, height = pg.size()
        pg.click(x=1883, y=1007)
        time.sleep(1)
        miConexion.close()
    else :        
        mensaje_final = "Recibe un cordial saludo a nombre del equipo de trabajo de *Confianza y Vida Asesores y Corredores de Seguros*%0A¡El SOAT de tu auto con placa *" + str(placaGeneral) + "* va a vencer!!! el dia *" + str(fecha) + "* ya que el año pasado lo adquiriste en algunas de nuestras oficinas o puntos de venta. Evita multas, sanciones y la retención de tu vehículo 🚗%0AAprovecha, renueva o adquiere tu SOAT en sólo 5 minutos, estés donde estés😉%0AEnvíanos un mensaje a nuestro WhatsApp y consulta todas las promociones que tenemos en Seguros👍"
        web.open("whatsapp://send?phone=51"+str(celular)+"&text="+str(mensaje_final))
        time.sleep(12)
        width, height = pg.size()
        pg.click(x=1883, y=1007)
        time.sleep(1)
        miConexion.close()


   
