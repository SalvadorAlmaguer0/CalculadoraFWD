# Calculadora de FWD
# Puntos a considerar:
# Antes de la ejecucion de todo el codigo (en el caso de que se ejecute en la terminal de VSCode), es importante 
# primero ejecutar unicamente las librerias, despues ya el codigo no deberia tener ningun error 

# Antes de ejecutar el codigo, es importante mantener una cierta configuracoin debido al formato de la fecha.
# Para ete codigo lo unico a tomar en cuenta en Visual Studio Code, es mantener el programa en español
# esto debido al formato de la fecha que se maneja.

# Ademas de tener instaladas las librerias utilizadas.
# Es posible que al ejecutar el codigo y no se tenga el Excel modificado (INPUTS.xlsx) en la misma carpeta, no logre
# hacer los calculos correctamente, por lo que se sugiere utilizar la base de datos modificada, enla cual solo se 
# agrego algunas fechas en las tablas donde hay dias fesivos y se interpolaron los valores para no afectar en la grafica.

# El codigo y las librerias siguen estando mal optimizadas. (Falta mejorar las practicas)
from tkinter import *
from tkinter import ttk
from tkcalendar import Calendar
import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta
from datetime import datetime
from math import exp
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

rgb='#B0B0B0' # Cambiar color de fondo y letra
letra='black'
redo = 2 # redondeo

# Interfas

ventana = Tk() # Para que se cree la ventana y lo demas son declaracion de variables, widget y posisionamiento de los widget
ventana.title('Productos Financieros Derivados')

ventana.resizable(0,0) # No ajustar tamaño manualmente
ventana.geometry('920x555') # Tamaño de la ventana
ventana.config(bg=rgb) # Color fondo

# Declaracion variables
s_0 = DoubleVar(value='')
tiempo = IntVar(value='')
nodo0 = IntVar(value='')
valor0 = DoubleVar(value='')
dia0 = IntVar(value='')
dividendo0 = DoubleVar(value='')
fecha1 = StringVar()
fecha2 = StringVar()
fechaI = StringVar()
fechaF = StringVar()
empdiv = StringVar()
pnaV = StringVar()
cantidad = IntVar(value=1)
fwd =StringVar()
casillaest = IntVar()
casillae = IntVar()
# Calendario

feIE = Calendar(ventana, year=2022, month=1)

def print_date(date):
    global fecha
    fecha = date
    if fecha[1] == '/':

        fecha = '0'+fecha
    if fecha[4] == '/':
        fecha = fecha[0:3] + '0' + fecha[3:5]+'20'+fecha[5:]

    if fecha[6:8] != '20':
        fecha = fecha[0:6]+'20'+fecha[6:]
    
    print(f'{fecha}')


def selfechI():
    global fecha, fechaI, fecha1d, fecha1m, fecha1a, fechaIr
    
    fecha1.set(fecha)
    
    fecha1d = fecha[0:2]
    fecha1m = fecha[3:5]
    fecha1a = fecha[6:10]
    fechILE.config(state='disabled',textvariable=fecha1)
    fechaIr = fecha[6:10] + '-' + fecha[3:5] +'-' + fecha[0:2]
    fechaI = datetime.strptime(fechaIr, '%Y-%m-%d')
    fechaIr = str(fecha[6:10]) + str(fecha[3:5]) + str(fecha[0:2])

def selfechF():
    global fecha, fechaF, fecha2a, fecha2d, fecha2m, fechaFr, fechaFrs
    
    fecha2.set(fecha)
    fecha2d = fecha[0:2]
    fecha2m = fecha[3:5]
    fecha2a = fecha[6:10]
    tiempoE.config(state='disabled',textvariable=fecha2)
    fechaFrs = fecha[6:10] + '-' + fecha[3:5] +'-' + fecha[0:2]
    fechaF = datetime.strptime(fechaFrs, '%Y-%m-%d')
    fechaFr = str(fecha[6:10]) + str(fecha[3:5]) + str(fecha[0:2])


def fempdiv(emp):
    global fecha1d, fecha1m, fecha1a

    fech = fecha1a+'-'+fecha1m+'-'+fecha1d
    s_0.set(pd.read_excel('INPUTS.xlsx',sheet_name='PRECIOS').set_index('Dates').at[fech,empE.get()])
    s_0E.config(state='disabled', textvariable = s_0)

    if empE.get() == 'BIMBO' or empE.get() == 'CEMEX' or empE.get() == 'JOSECUERVO' or empE.get() == 'FEMSA' or empE.get() == 'TELEVISA' or empE.get() == 'VOLARIS':
        estado.place(x=492, y=2000)
        valorB.config(state='normal')
    else:
        estado.place(x=520, y=130)
        valorB.config(state='disabled')


# Precio no arbitraje

def precioNA():
    global fechaI, fechaF, fecha1d, fecha1m, fecha1a, fechaIr, fechaFr, nodo, ttvalor, s_0v, tsumdiv, fechf, s_0d

    # relativedelta(months=1)   Sumar un mes 


    if (fechaF-fechaI).days <1:
        print('Fechas invalidas')
        return
    
    if fechaI.day != 1:
        
        fechd = fechaI + relativedelta(months=1)
        fech = str(fechd.year) +'-'+ str(fechd.month) +'-01'  # fecha para buscar en el excel, (yyyymmdd)
    else:
        fechd = fechaI
        fech = str(fechd.year) +'-'+ str(fechd.month) +'-01'

    fechdf =  fechaF
    if int(fechd.month) <10 : 
        fech = fech[0:5]+'0'+fech[5:]

    if fechaF.day != 1:
        fechdf = fechaF
        fechf = str(fechdf.year) +'-'+ str(fechdf.month) +'-01'  # fecha para buscar en el excel, (yyyymmdd)
    else:
        fechdf = fechaF
        fechf = str(fechdf.year) +'-'+ str(fechdf.month) +'-01'

    if int(fechdf.month) <10 : 
        fechf = fechf[0:5]+'0'+fechf[5:]

    
    # Calcular la diferencia entre las dos fechas en dias
    nodo = (fechaF - fechaI).days

    
    if empE.get() == 'BIMBO' or empE.get() == 'CEMEX' or empE.get() == 'JOSECUERVO' or empE.get() == 'FEMSA' or empE.get() == 'TELEVISA' or empE.get() == 'VOLARIS' or empE.get() == 'IPC':

        print(f'\ncargando datos')
        ttvalor = pd.read_excel('INPUTS.xlsx',sheet_name='CURVAMXN').set_index('Node')
        tvalor = ttvalor[int(fechaIr)]
        valor = round(ttvalor.at[nodo, int(fechaIr)],6)
        print('datos cargados')

        # solo para los que pagan dividendos estandar


###################################################################################
        print(f'{fech}\n{fechf}\n')



        if empE.get() == 'BIMBO' or empE.get() == 'CEMEX' or empE.get() == 'JOSECUERVO' or empE.get() == 'FEMSA' or empE.get() == 'TELEVISA' or empE.get() == 'VOLARIS':
            
            tsumdiv = pd.read_excel('INPUTS.xlsx',sheet_name='DIVIDENDOS')  # 1
            indicesI = tsumdiv.iloc[(tsumdiv['Dates'] == fech).values].index.values[0]
            indicesF = tsumdiv.iloc[(tsumdiv['Dates'] == fechf).values].index.values[0]
            sumdiv = tsumdiv.loc[indicesI:indicesF,empE.get()] # Creo una lista con los dividendos que entran en el fwd

            sumdiv.reset_index(drop=True, inplace=True) # reinicie el indice para moverme por index relacinandolo con el nodo nodo

            sumaa = 0
            for i in range(sumdiv.shape[0]):
        
                difdpm = (((fechd - relativedelta(day=1)) + relativedelta(months=i)) - fechaI).days # Diferencia de dias entre el inicio del fdw y el primer mes
                rd = tvalor[difdpm] # tvalor es la tabla unicamente del dia t=0

                sumaa=sumaa + (sumdiv[int(i)]*exp(-rd*difdpm/360))
                

            print(f'\nPresio:\t{s_0.get()}\nT: \t{nodo}\nr(0,T):\t{valor}\nI:\t{sumaa}\nCant:\t{cantidad.get()}\n')
            s_0d = ((s_0.get()-sumaa)*(exp(valor*nodo/360)))
            s_0v = cantidad.get()*((s_0.get()-sumaa)*(exp(valor*nodo/360)))
            fwd.set(str(round(s_0v,redo))+' MXN')
            return


        else: # Se selecciono IPC y tiene tasa de dividendos 
            
            
            # s_0 x;   valor x;   calcular delta ?;  nodo x;

            tdelta = pd.read_excel('INPUTS.xlsx',sheet_name='TASA DE DIVIDENDOS')[['Dates', 'IPC']].set_index('Dates')
            fdelta = str(fechaI.year) + '-'
            if fechaI.month<10:
                fdelta = fdelta + '0' + str(fechaI.month) +'-'
            else:
                fdelta = fdelta + str(fechaI.month)+'-'

            if fechaI.day<10:
                fdelta = fdelta + '0' + str(fechaI.day)

            delta = tdelta.loc[fdelta,'IPC']
            print(f'\nPresio:\t{s_0.get()}\nT: \t{nodo}\nr(0,T):\t{valor}\ndelta:\t{delta}\nCant:\t{cantidad.get()}')
            fwd.set(str(round(
                cantidad.get()*(s_0.get()*(exp((valor - delta)*(nodo/360))))
                ,redo))+' MXN')
            return

    
    if empE.get() == 'AMAZON (USD)' or empE.get() == 'S&P500' or empE.get() == 'ORO (USD)':
    
        print(f'\ncargando datos')
        valor = round(pd.read_excel('INPUTS.xlsx',sheet_name='CURVAUSD').set_index('Node').at[nodo, int(fechaIr)],6)
        print('datos cargados')
        

        fcambio = fecha1a+'-'+fecha1m+'-'+fecha1d
        cambio = pd.read_excel('INPUTS.xlsx',sheet_name='PRECIOS').set_index('Dates').at[fcambio,'MXN/USD']

        if empE.get() == 'AMAZON (USD)' or empE.get() == 'S&P500':
            tdelta = pd.read_excel('INPUTS.xlsx',sheet_name='TASA DE DIVIDENDOS')[['Dates', str(empE.get())]].set_index('Dates')
            fdelta = str(fechaI.year) + '-'
            if fechaI.month<10:
                fdelta = fdelta + '0' + str(fechaI.month) +'-'
            else:
                fdelta = fdelta + str(fechaI.month)+'-'

            if fechaI.day<10:
                fdelta = fdelta + '0' + str(fechaI.day)
            else:
                fdelta = fdelta + str(fechaI.day)

            delta = tdelta.loc[fdelta,str(empE.get())]
            
            print(f'\nPresio:\t{s_0.get()}\nT: \t{nodo}\nr(0,T):\t{valor}\ndelta:\t{delta}\nCant:\t{cantidad.get()}\nCambio:\t{cambio}')
           
            fwd.set((str(round(
                cambio*cantidad.get()*(s_0.get()*(exp((valor - delta)*(nodo/360)))),redo))+' MXN\n'+
                str(round(
                cantidad.get()*(s_0.get()*(exp((valor - delta)*(nodo/360)))),redo))+' USD'))
            resultado.place(x=700, y= 50)
            return

        else: # se escojio oro
            tsumdiv = pd.read_excel('INPUTS.xlsx',sheet_name='COSTOS ALMACEN')
            tvalor = pd.read_excel('INPUTS.xlsx',sheet_name='CURVAUSD').set_index('Node')[int(fechaIr)]

            indicesI = tsumdiv.iloc[(tsumdiv['Dates'] == fech).values].index.values[0]
            indicesF = tsumdiv.iloc[(tsumdiv['Dates'] == fechf).values].index.values[0]
            sumdiv = tsumdiv.loc[indicesI:indicesF,empE.get()] 
            sumdiv.reset_index(drop=True, inplace=True)
            
            sumaa = 0
            for i in range(sumdiv.shape[0]):
        
                difdpm = (((fechd - relativedelta(day=1)) + relativedelta(months=i)) - fechaI).days # Diferencia de dias entre el inicio del fdw y el primer mes
                rd = tvalor[difdpm]

                sumaa=sumaa + (sumdiv[int(i)]*exp(-rd*difdpm/360))

            print(f'\nPresio:\t{s_0.get()}\nT: \t{nodo}\nr(0,T):\t{valor}\nI:\t{sumaa}\nCant:\t{cantidad.get()}')
            fwd.set((str(round(
                cambio*cantidad.get()*((s_0.get()+sumaa)*(exp(valor*nodo/360))),redo))+' MXN\n'+
                str(round(
                cantidad.get()*((s_0.get()+sumaa)*(exp(valor*nodo/360))),redo))+' USD'))
            resultado.place(x=700, y= 50)
            return

    # DIVISAS :c
    if empE.get() == 'MXN/USD' or empE.get() == 'MXN/EUR' or empE.get() == 'USD/EUR':       
 
        if empE.get() == 'MXN/USD':
 
            r_loc = round(pd.read_excel('INPUTS.xlsx',sheet_name='CURVAMXN').set_index('Node').at[nodo, int(fechaIr)],6)

            r_ext = round(pd.read_excel('INPUTS.xlsx',sheet_name='CURVAUSD').set_index('Node').at[nodo, int(fechaIr)],6)

            fwd.set(str(round(
                cantidad.get()*(s_0.get()*exp((r_loc-r_ext)*nodo))
            ,redo))+' MNX')
            return

        elif empE.get() == 'MXN/EUR':
            
            r_loc = round(pd.read_excel('INPUTS.xlsx',sheet_name='CURVAMXN').set_index('Node').at[nodo, int(fechaIr)],6)

            r_ext = round(pd.read_excel('INPUTS.xlsx',sheet_name='CURVAEUR').set_index('Node').at[nodo, int(fechaIr)],6)

            fwd.set(str(round(
                cantidad.get()*(s_0.get()*exp((r_loc-r_ext)*nodo))
            ,redo))+' MXN')
            return

        else: # USD/EUR
            
            r_loc = round(pd.read_excel('INPUTS.xlsx',sheet_name='CURVAUSD').set_index('Node').at[nodo, int(fechaIr)],6)

            r_ext = round(pd.read_excel('INPUTS.xlsx',sheet_name='CURVAEUR').set_index('Node').at[nodo, int(fechaIr)],6)

            fwd.set(str(round(
                cantidad.get()*(s_0.get()*exp((r_loc-r_ext)*nodo))
                ,redo))+' USD')
            return

def casilla():
    global nodo
    if casillaest.get() == 1:

        casillae.set(nodo)
        casillaE.config(state='disabled')
    else:
        casillaE.config(state='normal')

def valorfwd():
    global nodo, fechaIr, ttvalor, fechaI, s_0d, tsumdiv, fechaF, fechf, fechaFrs
    salir = 0
    valor_fwd = [0]

    if casillae.get() > nodo:
        print('El tiempo seleccionado esta feura del plazo del contraro')
        return
    elif casillae.get() < 1:
        print('Seleccione una fecha mayor a 1')
        return
    
    indice = ttvalor.columns.get_loc(int(fechaIr))
    ttvalor.reset_index(drop=True, inplace=True)
    valorr=[]
    fechasv=[]

    # dias festivos: *Se realizo una interpolacoin en la tabla con ellos
    # 7 febrero
    # 21 marzo
    # 14 y 15 Abril
    # 16 septiembre
    # 2 y 21 Noviembre
    # 12 Diciembre
    # 6 Febrero 2023
    # 6 y 7 Abril
    # ---
    
    tprecio = pd.read_excel('INPUTS.xlsx',sheet_name='PRECIOS').set_index('Dates')
    
    if fechaI.strftime("%A") == 'Monday':

        for i in range(4):

            valorr.append(round(ttvalor.iloc[(nodo-2-i), (indice+i+1)],6))
            
            nindice = indice+i
            fechv = str(ttvalor.columns[indice+i+1])
            tvalor = ttvalor[int(fechv)]
            fechvs = fechv[0:4]+'-'+fechv[4:6]+'-'+fechv[6:8]
            fechdI = datetime(int(fechv[0:4]),int(fechv[4:6]),int(fechv[6:8])) # fecha de ese dia, t

            fechasv.append(pd.to_datetime(fechvs))

            s_i = tprecio.at[fechvs, empE.get()]

            fecha_indiced = datetime(int(fechv[0:4]),int(fechv[4:6]),1) + relativedelta(months=1)
            fecha_indice = fecha_indiced.strftime("%Y")+'-'+fecha_indiced.strftime("%m")+'-'+fecha_indiced.strftime("%d")
            indicesI = tsumdiv.iloc[(tsumdiv['Dates'] == fecha_indice).values].index.values[0]
            indicesF = tsumdiv.iloc[(tsumdiv['Dates'] == fechf).values].index.values[0]

            sumdiv = tsumdiv.loc[indicesI:indicesF, empE.get()]
            fechasdiv = tsumdiv.loc[indicesI:indicesF, 'Dates']

            fechasdiv.reset_index(drop=True,inplace=True)

            sumdiv.reset_index(drop=True,inplace=True)
            
            sumaa = 0
            for j in range(sumdiv.shape[0]):

                difdpm = (fechasdiv[j]-fechdI).days - j
                rd = tvalor[difdpm-1]
                sumaa = sumaa + (sumdiv[int(j)]*exp(-rd*difdpm/360))
                

            if nodo-2-i >= 0:

                rd =tvalor[nodo-2-i]
                f_it = (s_i-sumaa)*exp(rd*(nodo-i-1)/360)
                
                valor_fwd.append((s_0d-f_it)*exp(-rd*(nodo-1-i)/360))
                print(f'\n')
            else:
                valor_fwd.append(s_0d-s_i)
                salir = 1
            
        nnodo = nodo-7
        dias_transcurridos = 4 # Dias que ya conseguimos su valor

    elif fechaI.strftime("%A") == 'Tuesday':

        for i in range(3):

            valorr.append(round(ttvalor.iloc[(nodo-2-i), (indice+i+1)],6))
            
            nindice = indice+i
            fechv = str(ttvalor.columns[indice+i+1])
            tvalor = ttvalor[int(fechv)]
            fechvs = fechv[0:4]+'-'+fechv[4:6]+'-'+fechv[6:8]
            fechdI = datetime(int(fechv[0:4]),int(fechv[4:6]),int(fechv[6:8])) # fecha de ese dia, t

            fechasv.append(pd.to_datetime(fechvs))

            s_i = tprecio.at[fechvs, empE.get()]

            fecha_indiced = datetime(int(fechv[0:4]),int(fechv[4:6]),1) + relativedelta(months=1)
            fecha_indice = fecha_indiced.strftime("%Y")+'-'+fecha_indiced.strftime("%m")+'-'+fecha_indiced.strftime("%d")
            indicesI = tsumdiv.iloc[(tsumdiv['Dates'] == fecha_indice).values].index.values[0]
            indicesF = tsumdiv.iloc[(tsumdiv['Dates'] == fechf).values].index.values[0]
            sumdiv = tsumdiv.loc[indicesI:indicesF, empE.get()]
            fechasdiv = tsumdiv.loc[indicesI:indicesF, 'Dates']
            fechasdiv.reset_index(drop=True,inplace=True)
            sumdiv.reset_index(drop=True,inplace=True)
            sumaa = 0
            for j in range(sumdiv.shape[0]):
                difdpm = (fechasdiv[j]-fechdI).days - j
                rd = tvalor[difdpm-1]
                sumaa = sumaa + (sumdiv[int(j)]*exp(-rd*difdpm/360))

            if nodo-2-i >= 0:

                rd =tvalor[nodo-2-i]
                f_it = (s_i-sumaa)*exp(rd*(nodo-i-1)/360)
                valor_fwd.append((s_0d-f_it)*exp(-rd*(nodo-1-i)/360))

            else:
                valor_fwd.append(s_0d-s_i)
                salir = 1
        
        nnodo = nodo-6
        dias_transcurridos = 3

    elif fechaI.strftime("%A") == 'Wednesday':
        
        for i in range(2):

            valorr.append(round(ttvalor.iloc[(nodo-2-i), (indice+i+1)],6))
            
            nindice = indice+i
            fechv = str(ttvalor.columns[indice+i+1])
            tvalor = ttvalor[int(fechv)]
            fechvs = fechv[0:4]+'-'+fechv[4:6]+'-'+fechv[6:8]
            fechdI = datetime(int(fechv[0:4]),int(fechv[4:6]),int(fechv[6:8])) # fecha de ese dia, t

            fechasv.append(pd.to_datetime(fechvs))

            s_i = tprecio.at[fechvs, empE.get()]

            fecha_indiced = datetime(int(fechv[0:4]),int(fechv[4:6]),1) + relativedelta(months=1)
            fecha_indice = fecha_indiced.strftime("%Y")+'-'+fecha_indiced.strftime("%m")+'-'+fecha_indiced.strftime("%d")
            indicesI = tsumdiv.iloc[(tsumdiv['Dates'] == fecha_indice).values].index.values[0]
            indicesF = tsumdiv.iloc[(tsumdiv['Dates'] == fechf).values].index.values[0]

            sumdiv = tsumdiv.loc[indicesI:indicesF, empE.get()]
            fechasdiv = tsumdiv.loc[indicesI:indicesF, 'Dates']
            fechasdiv.reset_index(drop=True,inplace=True)
            sumdiv.reset_index(drop=True,inplace=True)
            
            sumaa = 0
            for j in range(sumdiv.shape[0]):

                #difdpm = nodo-1-j-i
                difdpm = (fechasdiv[j]-fechdI).days - j
                rd = tvalor[difdpm-1]
                sumaa = sumaa + (sumdiv[int(j)]*exp(-rd*difdpm/360))

            if nodo-2-i >= 0:

                rd =tvalor[nodo-2-i]
                f_it = (s_i-sumaa)*exp(rd*(nodo-i-1)/360)

                valor_fwd.append((s_0d-f_it)*exp(-rd*(nodo-1-i)/360))

            else:
                valor_fwd.append(s_0d-s_i)
                salir = 1
            
        nnodo = nodo-5
        dias_transcurridos = 2



    elif fechaI.strftime("%A") == 'Thursday':

        for i in range(1):

            valorr.append(round(ttvalor.iloc[(nodo-2-i), (indice+i+1)],6))
            
            nindice = indice+i
            fechv = str(ttvalor.columns[indice+i+1])
            tvalor = ttvalor[int(fechv)]
            fechvs = fechv[0:4]+'-'+fechv[4:6]+'-'+fechv[6:8]
            fechdI = datetime(int(fechv[0:4]),int(fechv[4:6]),int(fechv[6:8])) # fecha de ese dia, t
            fechasv.append(pd.to_datetime(fechvs))
            s_i = tprecio.at[fechvs, empE.get()]
            fecha_indiced = datetime(int(fechv[0:4]),int(fechv[4:6]),1) + relativedelta(months=1)
            fecha_indice = fecha_indiced.strftime("%Y")+'-'+fecha_indiced.strftime("%m")+'-'+fecha_indiced.strftime("%d")
            indicesI = tsumdiv.iloc[(tsumdiv['Dates'] == fecha_indice).values].index.values[0]
            indicesF = tsumdiv.iloc[(tsumdiv['Dates'] == fechf).values].index.values[0]
            sumdiv = tsumdiv.loc[indicesI:indicesF, empE.get()]
            fechasdiv = tsumdiv.loc[indicesI:indicesF, 'Dates']
            fechasdiv.reset_index(drop=True,inplace=True)
            sumdiv.reset_index(drop=True,inplace=True)
            
            sumaa = 0
            for j in range(sumdiv.shape[0]):
                difdpm = (fechasdiv[j]-fechdI).days - j
                rd = tvalor[difdpm-1]
                sumaa = sumaa + (sumdiv[int(j)]*exp(-rd*difdpm/360))

            if nodo-2-i >= 0:

                rd =tvalor[nodo-2-i]
                f_it = (s_i-sumaa)*exp(rd*(nodo-i-1)/360)
                valor_fwd.append((s_0d-f_it)*exp(-rd*(nodo-1-i)/360))
            else:
                valor_fwd.append(s_0d-s_i)
                salir = 1

        nnodo = nodo-4
        dias_transcurridos = 1


    elif fechaI.strftime("%A") == 'Friday':
        nnodo = nodo-3
        nindice= indice-1
        dias_transcurridos = 0

    numDiasNoAvil = (int(nnodo/7)*2)+2
    tprecio = pd.read_excel('INPUTS.xlsx',sheet_name='PRECIOS').set_index('Dates') #.at[fech,empE.get()]
    for i in range((casillae.get()-numDiasNoAvil-dias_transcurridos-1)):

        valorr.append(round(ttvalor.iloc[nnodo-i-1-(int(i/5)*2), nindice+i+2],6))
        fechv = str(ttvalor.columns[nindice+i+2])
        fechvs = fechv[0:4]+'-'+fechv[4:6]+'-'+fechv[6:8]
        fechasv.append(pd.to_datetime(fechvs))

        s_i = tprecio.at[fechvs, empE.get()]

        fecha_indice = datetime(int(fechv[0:4]),int(fechv[4:6]),1) + relativedelta(months=1)
        fecha_indice = fecha_indice.strftime("%Y")+'-'+fecha_indice.strftime("%m")+'-'+fecha_indice.strftime("%d")
        indicesI = tsumdiv.iloc[(tsumdiv['Dates'] == fecha_indice).values].index.values[0]
        indicesF = tsumdiv.iloc[(tsumdiv['Dates'] == fechf).values].index.values[0]
        
        sumdiv = tsumdiv.loc[indicesI:indicesF, empE.get()]
        sumdiv.reset_index(drop=True,inplace=True)
        tvalor = ttvalor[int(fechv)]
        
        sumaa = 0
        for j in range(sumdiv.shape[0]):

            difdpm = nnodo-i-(int(i/5)*2)
            rd = tvalor[difdpm-1]
            sumaa = sumaa + (sumdiv[int(j)]*exp(-rd*difdpm/360))
        
        f_it = (s_i-sumaa)*exp(valorr[i]*(nnodo-i-(int(i/5)*2))/360)
        valor_fwd.append((s_0d-f_it)*exp((-valorr[i])*(nnodo-i-(int(i/5)*2))/360))
    if salir == 0:
        s_i = tprecio.at[fechaFrs, empE.get()]
        valor_fwd.append(s_0d-s_i)

    if largo_cortoE.get() == 'Largo':
        largCort = 1
    else:
        largCort = -1
    #print(valor_fwd,'\n',cantidad.get()*largCort, '\n\n')
    valor_fwd = list(map(lambda x: x * (cantidad.get()*largCort) , valor_fwd))
    print(valor_fwd,'\n',nodo)

    if empE.get() == 'BIMBO':
        col = 'r'
    elif empE.get() == 'CEMEX':
        col = 'b'
    elif empE.get() == 'JOSECUERVO':
        col = 'orange'
    elif empE.get() == 'FEMSA':
        col = 'g'
    elif empE.get() == 'TELEVISA':
        col = 'y'
    elif empE.get() == 'VOLARIS':
        col = 'm'

    lista = list(range(0, (len(valor_fwd)-1) + 1, 1))
    line, = ax.plot(lista, valor_fwd, color=col)
    canvas.draw()

# Estructur Calendario

feIL = Label(ventana, text='Seleción de fechas ',bg=rgb,fg=letra,font=50)
feIL.place(x=50, y=60)

feIE.bind('<<CalendarSelected>>', lambda e: print_date(feIE.get_date()))
feIE.place(x=40, y=85)

botf = Button(ventana, text='Fecha Inicio',command=selfechI,width=10)
botf.place(x=125, y=275)

botdf = Button(ventana, text='Fecha Final',command=selfechF, width=10)
botdf.place(x=212, y=275)

# Estructur (Ya no es necesario)
encabezado = Label(ventana, text=' Calculadora de FORDWAR', bg=rgb, fg=letra)
encabezado.place(x=40, y=15)
encabezado.config(font=("Calibri", 25, "bold"), justify='center')

####################################################################################################################

fechIL= Label(ventana, text='Fecha inicial',bg=rgb,fg=letra,font=50)
fechIL.place(x=325,y=60)
fechILE = Entry(ventana, textvariable=fecha1)
fechILE.place(x=325,y=85)

tiempoL = Label(ventana,text='Fecha final',bg=rgb,fg=letra,font=50)
tiempoL.place(x=325,y=120)
tiempoE = Entry(ventana,textvariable=tiempo)
tiempoE.place(x=325,y=145)

empL = Label(ventana,text='Empresa/Divisa',bg=rgb,fg=letra,font=50)
empL.place(x=325,y=180)
empE = ttk.Combobox(ventana,state='readonly',values=pd.read_excel('INPUTS.xlsx',sheet_name='PRECIOS').columns.tolist()[1:], textvariable=empdiv)
empE.place(x=325,y=205) 
empE.bind("<<ComboboxSelected>>", fempdiv)

largo_cortoL = Label(ventana,text='Tipo de contrato',bg=rgb,fg=letra,font=50)
largo_cortoL.place(x=325,y=335)
largo_cortoE = ttk.Combobox(ventana,state='readonly',values=['Largo', 'Corto'])
largo_cortoE.place(x=325,y=360)
largo_cortoE.current(0)

s_0L = Label(ventana,text='Precio',bg=rgb,fg=letra,font=50,)
s_0L.place(x=325,y=240)
s_0E = Entry(ventana,textvariable=s_0)
s_0E.place(x=325,y=265)

cantidadL = Label(ventana,text='Cantidad',bg=rgb,fg=letra,font=50,)
cantidadL.place(x=325,y=300)
cantidadE = Entry(ventana,textvariable=cantidad, width=7)
cantidadE.place(x=403,y=300)

pnaB = Button(ventana, text='Calcular', command=precioNA,width=10)
pnaB.place(x=325, y=400)

resultadoL = Label(ventana, text='Precio de\nno arbitraje:',bg=rgb,fg=letra)
resultadoL.config(font=("Calibri", 20, "bold"))
resultadoL.place_forget()
resultadoL.place(x=500, y= 50)

resultado = Label(ventana, textvariable= fwd,bg=rgb,fg=letra)
resultado.place(x=700, y= 70)
resultado.config(font=("Calibri", 20, "bold"), justify='center')

casillaL = Label(ventana, text='Calcular valor de FWD',bg=rgb,fg=letra)
casillaL.place(x=500, y=160)
casillaL.config(font=("Calibri", 15, "bold"), justify='center')

casillaE = Entry(ventana, textvariable=casillae)
casillaE.place(x=695, y=167)

casillaC = Checkbutton(ventana, bg=rgb,fg=letra, command=casilla, variable=casillaest)
casillaC.place(x=690, y=190)

casillaCL = Label(text='(Todos los dias del contrato)',bg=rgb,fg=letra)
casillaCL.place(x=492, y=190)
casillaCL.config(font=("Calibri", 12, "bold"), justify='center')

valorB= Button(ventana,text='Calcular', command=valorfwd)
valorB.place(x=830, y=164)

estado = Label(text='Por ahora solo con empresas Mexicanas',bg=rgb,fg='red')
estado.config(font=("Calibri", 15, "bold"), justify='center')

###### GRAFICA 
fig, ax = plt.subplots(dpi=90, figsize=(5,3.5),facecolor=rgb)
frame = Frame(ventana)
frame.place(x=480, y=220)
ax.set_facecolor(rgb)
ax.axhline(linewidth=1, c=letra)
ax.axvline(linewidth=1, c=letra)
canvas = FigureCanvasTkAgg(fig, master = frame)  # Crea el area de dibujo en Tkinter
canvas.get_tk_widget().grid()

#limp_grafica = Button # con el metodo clf o clear (clear quiza tamiben borra los ejes)
ventana.mainloop()
