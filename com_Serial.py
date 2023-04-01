# -*- coding: utf-8 -*-
"""
Created on Thu Jul  4 21:09:51 2019

@author: PC-13
"""
import serial
import time
import pandas as pd
from pandas import ExcelWriter
from tkinter import *
from tkinter import filedialog
import matplotlib.pyplot as plt
#
try :
    if not arduino.isOpen():
        arduino = serial.Serial('COM5', 9600)#Iniciamos comunicación serial con arduino por el puerto 5
except(NameError):
    serial.Serial("COM5",9600).close()
    arduino = serial.Serial('COM5', 9600)
#arduino = serial.Serial('COM4', 9600)#Iniciamos comunicación serial con arduino por el puerto 5
time.sleep(2)
Fecha=""
Hora=""
#
#arduino.write(b"r")
def FuncionPrincipal():
    ColTemp=[]
    ColHum=[]
    cont=0
    while True:
#    **** Extrayendo datos enviados por Arduino
#        Hora=time.strftime("%H:%M:%S")
        Fecha=time.strftime("%d_%m_%y")
        
        Temp = arduino.readline()
#        print(Temp)
        Humedad= arduino.readline()
        Enc_Apag=arduino.readline()
#Decodificando la información recibida en bits
        Temp=Temp.decode()
        Humedad=Humedad.decode()
        Enc_Apag=Enc_Apag.decode()
        
        E_A.set(Enc_Apag)
   
#    Manipulación de cadenas  ******
        TitTemp=Temp[0:13]
        a=int(Temp.find("°C"))
        ValTemp=Temp[13:a]
        TitHum=Humedad[0:9]
        b=int(Humedad.find("%"))
        ValHum=Humedad[9:b]
            
#        print(Temp)
#        print(TitTemp)
#        print(ValTemp)
#        print(Humedad)
#        print(TitHum)
#        print(ValHum)
#  *****************Usamos el método set para reemplazar la variable ValTemp por TempPant      
        TempPant.set(ValTemp)
        HumPant.set(ValHum)
        ColTemp.append(ValTemp)
        ColHum.append(ValHum)
        Tabla={TitTemp:ColTemp,TitHum:ColHum}
#        print(Tabla)
#-------------Enviamos Datos a una tabla de excel     
        Tablaexcel=pd.DataFrame(Tabla)
        Tablaexcel.to_excel("Tabla_Humedad_Temperatura.xlsx")
#******************Crear un archivo excel y guardar los datos  ("DECOMENTAR" Para utilizar , 
#Poner la dirección donde desea guardar los datos)       
        
#        writer=ExcelWriter("C:/Users/PC-13/Desktop/Avance/Datos "+Fecha+".xlsx")
#        Tablaexcel.to_excel(writer,"Hoja ", index=False)
#        writer.save()
###        
        
        TablaTemp={"Temperatura":ColTemp}
        GrafTemp=pd.DataFrame(TablaTemp)
#        print(GrafTemp)
        
        GrafTemp["Temperatura"]=GrafTemp.Temperatura.astype(float)
        
#        print(GrafTemp)
    
        TablaHum={"Humedad":ColHum}
        GrafHum=pd.DataFrame(TablaHum)
#        print(GrafHum)
    
        GrafHum["Humedad"]=GrafHum.Humedad.astype(float)
#       

    
        graficaTemp=GrafHum.plot.line()
        plt.title("Gráfica de Humedad Relativa")
        plt.xlabel("Tiempo(cada 5 segundos)")
        plt.ylabel("Temperatura")
        plt.savefig("Gráf. Hum"+Fecha+".png",dpi=100,bbox_inches="tight"  )
#        plt.show()
        
        graficaTemp=GrafTemp.plot.line()
        plt.title("Gráfica de Temperatura ")
        plt.xlabel("Tiempo(cada 5 segundos)")
        plt.ylabel("Humedad")
        plt.savefig("Gráf. Temp"+Fecha+".png",dpi=100,bbox_inches="tight"  )
#        plt.show()
    
        Interfaz.update()
        time.sleep(1)

def cerrar():
    arduino.close()
    Interfaz.destroy()
    
def AbrirArch():
##    if abrir.get()==1:
#        
#        print("funciona")
    fichero=filedialog.askopenfilename(title="Abrir",filetypes=(("Tabla de Datos","*.xlsx"),("Gráficas","*.png")))
#    return fichero
   

#Abrir y cerrar puerta
def Abrirpuerta():
    arduino.write("a".encode())
    
def Cerrarpuerta():
    arduino.write("c".encode())

#Abrir y cerrar cochera     
def Abrircochera():
    arduino.write("d".encode())
    
def Cerrarcochera():
    arduino.write("e".encode())     


#Encender y apagar luz cochera
def EncLed1():
    arduino.write("1".encode())
def ApagLed1():
    arduino.write("2".encode())


#Encender y apagar luz Sala    
def EncLed2():
    arduino.write("3".encode())
def ApagLed2():
    arduino.write("4".encode())

#Encender y apagar luz cuarto
def EncLed3():
    arduino.write("5".encode())
def ApagLed3():
    arduino.write("6".encode())
    
#Encender y apagar Todas las luces
def EncLedall():
    arduino.write("r".encode())
def ApagLedall():
    arduino.write("n".encode())


            
Interfaz=Tk()

TempPant=StringVar()
HumPant=StringVar()
E_A=StringVar()
Textpuer=StringVar()

Interfaz.title("Interfaz gráfica")
Interfaz.iconbitmap("UNTlogo.ico")
Interfaz.geometry("1400x1000")
Interfaz.config(bg="black")


Fondo=PhotoImage(file="Casa.png")
Fondo=Fondo.subsample(1,1)
label=Label(image=Fondo)
label.place(x=0,y=0,relwidth=1.0,relheight=1.0)


TextTemp=Label(Interfaz,text="Temperatura : ",bg="blue",fg="red",font=("Times New Roman",20))
TextTemp.grid(row=1,column=1)
TextValTemp=Label(textvariable=TempPant,bg="white",fg="black",font=("Times New Roman",20)).grid(row=2,column=1)


TextHum=Label(Interfaz,text="Humedad : ",bg="blue",fg="red",font=("Times New Roman",20))
TextHum.grid(row=1,column=2,padx=50,pady=5)
TextValHum=Label(textvariable=HumPant,bg="white",fg="black",font=("Times New Roman",20))
TextValHum.grid(row=2,column=2,padx=50,pady=5)

TextVent=Label(Interfaz,text="Ventilador ",bg="blue",fg="red",font=("Times New Roman",20))
TextVent.grid(row=3,column=2,padx=50,pady=5)
TextEnc_Apag=Label(textvariable=E_A,bg="white",fg="black",font=("Times New Roman",20))
TextEnc_Apag.grid(row=4,column=2)


boton1=Button(Interfaz, text="cerrar", command=lambda:cerrar())
boton1.grid(row=20,column=1)
#
#Enc_led


Verdat=Button(Interfaz,text="Ver Datos",command=AbrirArch)
Verdat.grid(row=6,column=1)

TextPuerta=Label(Interfaz,text="Puerta",bg="blue", fg="black",font=("Times New Roman",20))
TextPuerta.grid(row=35,column=20)
#
A_puerta=Button(Interfaz,text="Abrir puerta",command=Abrirpuerta)
A_puerta.grid(row=36,column=20)
C_puerta=Button(Interfaz,text="Cerrar puerta",command=Cerrarpuerta)
C_puerta.grid(row=37,column=20)

Textcochera=Label(Interfaz,text="Puerta Cochera",bg="blue", fg="black",font=("Times New Roman",20))
Textcochera.grid(row=18,column=46)
#
A_cochera=Button(Interfaz,text="Abrir puerta",command=Abrircochera)
A_cochera.grid(row=19,column=46)
C_cochera=Button(Interfaz,text="Cerrar puerta",command=Cerrarcochera)
C_cochera.grid(row=20,column=46)


TextLed1=Label(Interfaz,text="Luz Cohera",bg="blue",fg="black",font=("Times New Roman",20))
TextLed1.grid(row=14,column=44)
LedsENC1=Button(Interfaz,text="Encender",command=EncLed1)
LedsENC1.grid(row=15,column=44)
LedsAPAG1=Button(Interfaz,text="Apagar",command=ApagLed1)
LedsAPAG1.grid(row=16,column=44)

TextLed2=Label(Interfaz,text="Luz Sala",bg="blue",fg="black",font=("Times New Roman",20))
TextLed2.grid(row=25,column=5)
LedsENC2=Button(Interfaz,text="Encender",command=EncLed2)
LedsENC2.grid(row=26,column=5)
LedsAPAG2=Button(Interfaz,text="Apagar",command=ApagLed2)
LedsAPAG2.grid(row=27,column=5)

TextLed3=Label(Interfaz,text="Luz Cuarto",bg="blue",fg="black",font=("Times New Roman",20))
TextLed3.grid(row=1,column=45)
LedsENC3=Button(Interfaz,text="Encender",command=EncLed3)
LedsENC3.grid(row=2,column=45)
LedsAPAG3=Button(Interfaz,text="Apagar",command=ApagLed3)
LedsAPAG3.grid(row=3,column=45)

TextLedall=Label(Interfaz,text="Encender Luces",bg="blue",fg="black",font=("Times New Roman",20))
TextLedall.grid(row=5,column=2)
LedsENCall=Button(Interfaz,text="Encender Todas las luces",command=EncLedall)
LedsENCall.grid(row=6,column=2)
LedsAPAGall=Button(Interfaz,text="Apagar todas las luces",command=ApagLedall)
LedsAPAGall.grid(row=7,column=2)

Interfaz.after(10,FuncionPrincipal)

Interfaz.mainloop()
   
arduino.close()