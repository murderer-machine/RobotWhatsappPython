import webbrowser as web
import time
import pandas as pd
import pyautogui as pg
data = pd.read_excel('D:\Marco\CodigosPython\RobotWhatsapp\contactos.xlsx')
data_dict = data.to_dict('list')
celulares = data_dict['celular']
placas = data_dict['placa']
fechas = data_dict['fecha']
combo = zip(celulares, placas, fechas)
combo2 = zip(celulares, placas, fechas)
first = True
contador = 0
total = len(list(combo2))
for celular, placa, fecha in combo:
    contador = contador + 1
    print("total :" + str(total) + " contador :" + str(contador) + " placa :" + str(placa) + " celular :" + str(celular))
    mensaje_final = "Recibe un cordial saludo a nombre del equipo de trabajo de *Confianza y Vida Asesores y Corredores de Seguros*%0A¡El SOAT de tu auto con placa *" + str(placa) + "* va a vencer!!! el dia *" + str(fecha) + "* ya que el año pasado lo adquiriste en algunas de nuestras oficinas o puntos de venta. Evita multas, sanciones y la retención de tu vehículo 🚗%0AAprovecha, renueva o adquiere tu SOAT en sólo 5 minutos, estés donde estés😉%0AEnvíanos un mensaje a nuestro WhatsApp y consulta todas las promociones que tenemos en Seguros👍"
    web.open("whatsapp://send?phone=51"+str(celular)+"&text="+str(mensaje_final))
    time.sleep(12)
    width, height = pg.size()
    pg.click(x=1883, y=1007)
    time.sleep(1)
   
