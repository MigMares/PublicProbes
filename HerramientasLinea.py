print('% >>>>>>>>>>>>>>>>>> Herramientas Linea <<<<<<<<<<<<<<<<<<< %')
print('Opcion 1: Generar poligonal alrededor de una linea(s) (box)')
print('Opcion 2: Generar linea por azimut y lineas paralelas')
print('Opcion 3: Calcular estaciones entre dos puntos')
print('Opcion 4: Extender estaciones desde extremos de una linea')
print('% >>>>>>>>>>>>>>>>>>>>>>>>>>> - <<<<<<<<<<<<<<<<<<<<<<<<<<< %')
# print('% --- por Miguel Angel Mares Aguero --- %')
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
import xlsxwriter
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import math as mth
import tkinter as tk
from tkinter import filedialog
val = False
while(val == False):
      try:
          op = int(input('¿Que opcion deseas utilizar?: '))
          val = True
      except:
          print('¡Opcion no valida, introduce un numero!')
          val = False 
      if (op < 0 or op > 4):
          print('¡Opcion no valida, introduce del 1 al 4!')
          val = False 
proceso = False
while(proceso == False):
      if(op == 1):
          print('%---------------------------------------%')
          print('% ------ Generar box para lineas ------ %')
          # print('% --- por Miguel Angel Mares Aguero --- %')
          print('%---------------------------------------%')
          print('%- El programa calcula los vertices de un rectangulo alrededor de una linea -%')
          print('%- para recortar el perimetro de un voxel en Geosoft con PLY (Coord UTM). ---%') 
          print('%---------------------------------------%')
          print('-Si eliges importar tabla el formato es-')
          print('%------ Pix(L-1) \t Piy(L-1) ------%')
          print('%------ Pfx(L-1) \t Pfy(L-1) ------%')
          print('%------ Pix(L-2) \t Piy(L-2) ------%')
          print('%------ Pfx(L-2) \t Pfy(L-2) ------%')
          print('%---------- ....   ...    .... ---------%')
          print('%---------- ....   ...    .... ---------%')
          print('%------ Pix(L-N) \t Piy(L-N) ------%')
          print('%------ Pfx(L-N) \t Pfy(L-N) ------%')
          print('%---------------------------------------%')
          ing = int(input('Ingresar puntos por (1 = linea o 2 = tabla): '))
          check = 1
          while(check == 1):
               if(ing == 1):
                   l = int(input('Numero de lineas a calcular "box": '))
                   P = np.zeros((l,4,2))
                   V = np.zeros((l,4,2))
                   print('>> Ingresar informacion por linea <<')
                   n = 0
                   while(n <= l-1):
                        P[n][0][:] = list(map(float,input('Coordenadas punto inicial x y: ').split()))
                        P[n][1][:] = list(map(float,input('Coordenadas punto final x y: ').split()))
                        if(n < l-1):
                          print('------ Meter siguiente linea -------')
                          n += 1
                        else:
                          print('% >>>>>>>>>>>> Calculando <<<<<<<<<<<<< %')
                          n += 1
                   check = 2
               elif(ing == 2):
                   root = tk.Tk()
                   root.withdraw()
                   tabla = np.asmatrix(pd.read_fwf(filedialog.askopenfilename(),header=None),dtype=float)
                   l = int(len(tabla[:,0]) / 2)
                   P = np.zeros((l,4,2))
                   V = np.zeros((l,4,2))
                   n = 0
                   m = 0
                   while(n <= l-1):
                        P[n][0][0] = tabla[m,0]
                        P[n][0][1] = tabla[m,1]
                        P[n][1][0] = tabla[m+1,0]
                        P[n][1][1] = tabla[m+1,1]
                        n += 1
                        m += 2
                   check = 2
               else: 
                   print('--------- Opcion no valida ---------')
                   ing = int(input('Ingresar puntos por 1 = linea o 2 = tabla: '))
                   check = 1
          r = float(input('Metros a la redonda que se quiere cubrir: '))
          re_ev = 2
          while(re_ev == 2):
               for n in range(1,l+1):
                   a = P[n-1,1,1] - P[n-1,0,1]
                   b = P[n-1,1,0] - P[n-1,0,0]
                   if(b == 0):
                       m = mth.inf
                   else:
                       m = a / b
                   h = mth.sqrt(a**2 + b**2)
                   ang = mth.asin(a/h)
                   h2 = h + r
                   P[n-1,2,1] = h2 * mth.sin(ang) + P[n-1,0,1]
                   if(m == 0):
                       if(P[n-1,0,0] < P[n-1,1,0]):
                           P[n-1,2,0] = P[n-1,0,0] - r
                       else:
                           P[n-1,2,0] = P[n-1,0,0] + r
                   else:               
                       P[n-1,2,0] = ((P[n-1,2,1] - P[n-1,0,1]) / m) + P[n-1,0,0]  
                   P[n-1,3,1] = h2 * mth.sin(-ang) + P[n-1,1,1]
                   if(m == 0):
                       if(P[n-1,0,0] < P[n-1,1,0]):
                           P[n-1,3,0] = P[n-1,1,0] + r
                       else:
                           P[n-1,3,0] = P[n-1,1,0] - r
                   else:  
                       P[n-1,3,0] = ((P[n-1,3,1] - P[n-1,1,1]) / m) + P[n-1,1,0]
                   if(m == 0):
                       mp = mth.inf
                   else:
                       mp = -1 / m
                   V[n-1,0,1] = r * mth.cos(ang) + P[n-1,2,1]
                   if(abs(mp) == 0):
                       if(P[n-1,0,0] < P[n-1,1,0]):
                           V[n-1,0,0] = P[n-1,2,0] - r
                       else:
                           V[n-1,0,0] = P[n-1,2,0] + r
                   else:
                       V[n-1,0,0] = ((V[n-1,0,1] - P[n-1,2,1]) / mp) + P[n-1,2,0]
                   V[n-1,1,1] = r * mth.cos(mth.pi - ang) + P[n-1,2,1]          
                   if(abs(mp) == 0):
                       if(P[n-1,0,0] < P[n-1,1,0]):
                           V[n-1,1,0] = P[n-1,2,0] + r
                       else:
                           V[n-1,1,0] = P[n-1,2,0] - r
                   else:
                       V[n-1,1,0] = ((V[n-1,1,1] - P[n-1,2,1]) / mp) + P[n-1,2,0]
                   V[n-1,3,1] = r * mth.cos(ang) + P[n-1,3,1]
                   if(abs(mp) == 0):
                       if(P[n-1,0,0] < P[n-1,1,0]):
                           V[n-1,3,0] = P[n-1,3,0] - r
                       else:
                           V[n-1,3,0] = P[n-1,3,0] + r
                   else:  
                       V[n-1,3,0] = ((V[n-1,3,1] - P[n-1,3,1]) / mp) + P[n-1,3,0]
                   V[n-1,2,1] = r * mth.cos(mth.pi - ang) + P[n-1,3,1]
                   if(abs(mp) == 0):
                       if(P[n-1,0,0] < P[n-1,1,0]):
                           V[n-1,2,0] = P[n-1,3,0] + r
                       else:
                           V[n-1,2,0] = P[n-1,3,0] - r
                   else:
                       V[n-1,2,0] = ((V[n-1,2,1] - P[n-1,3,1]) / mp) + P[n-1,3,0]
               ing = input('Quieres ver el grafico? (y/n): ')
               if(ing == 'y'):
                   plt.close()
                   xinf = np.min(V[:,:,0])
                   xsup = np.max(V[:,:,0])
                   yinf = np.min(V[:,:,1])
                   ysup = np.max(V[:,:,1])
                   plt.style.use('seaborn-paper')
                   plt.axis([xinf, xsup, yinf, ysup])
                   for n in range(1,l+1):
                       plt.plot([V[n-1,0,0], V[n-1,3,0]],[V[n-1,0,1], V[n-1,3,1]],'-r')
                       plt.plot([V[n-1,0,0], V[n-1,1,0]],[V[n-1,0,1], V[n-1,1,1]],'-r')
                       plt.plot([V[n-1,2,0], V[n-1,3,0]],[V[n-1,2,1], V[n-1,3,1]],'-r')
                       plt.plot([V[n-1,2,0], V[n-1,1,0]],[V[n-1,2,1], V[n-1,1,1]],'-r')
                       plt.plot([P[n-1,0,0], P[n-1,1,0]],[P[n-1,0,1], P[n-1,1,1]],'b')
                   plt.minorticks_on()
                   plt.grid(b=True,which='major',axis='both',linestyle='-')
                   plt.grid(b=True,which='minor',axis='both',linestyle='--')
                   plt.xlabel('X')
                   plt.ylabel('Y')
                   plt.title('Box')
                   plt.draw()
                   plt.pause(.001)
                   plt.show(block=False)
                   ing = int(input('Replantear box? (1 = si, 2 = no): '))
                   if(ing == 1):
                       r = float(input('Metros a la redonda que se quiere cubrir: '))
                       re_ev = 2
                   else: 
                       ing = int(input('Guardar vertices? (1 = si, 2 = no): '))
                       re_ev = 1
               else:
                   ing = int(input('Guardar vertices? (1 = si, 2 = no): '))
                   re_ev = 1
          if(ing == 1):
              df = pd.DataFrame({'X':V[0,:,0],'Y':V[0,:,1]})
              np.savetxt("Box.txt",df.values,fmt='%.2f')
              vacio = np.array([0,0]).reshape((1,2))
              for n in range(1,l):
                  with open('Box.txt','a') as box:
                      df = pd.DataFrame({'X':V[n,:,0],'Y':V[n,:,1]})
                      np.savetxt(box,vacio,fmt='%.2f')
                      np.savetxt(box,df.values,fmt='%.2f')
              print('%---------------------------------------%')
              print('%                 Box.txt               %')
              print('%   Copiar vertices a cada archivo PLY  %')
              print('%---------------------------------------%')
          cont = False
          while(cont == False):
                pregu = input('Quieres usar otra opcion del programa? (y/n):')
                if(pregu == 'y'):
                    op = int(input('¿Que opcion deseas utilizar? (1 al 4): '))
                    proceso = False
                    cont = True 
                elif(pregu == 'n'):
                    proceso = True
                    cont = True
                else: 
                    print('¡Opcion no valida, introduce y o n!')
                    cont = False               
      elif(op == 2):
          print('%---------------------------------------%')
          print('% --- Trazo de linea(s) por azimut ---- %')
          # print('% --- por Miguel Angel Mares Aguero --- %')
          print('%---------------------------------------%')
          print('%- El programa calcula las estaciones en una linea tomando en cuenta -%')
          print('%- longitud y azimut determinados, posteriormente puede trazar lineas-%')
          print('%- paralelas en sentido horario (rotacion derextral) o viceversa. ----%') 
          print('%---------------------------------------------------------------------%')
          P1 = np.zeros((1,2))
          P1[:] = list(map(float,input('> Punto inicial/referencia x y: ').split()))
          azi = float(input('> Azimut = '))
          h = float(input('> Longitud de linea = '))
          spcSTN = float(input('> Espacio entre estaciones = '))
          azir = (azi * mth.pi) / 180
          dy = h * mth.cos(azir)
          P2 = np.zeros((1,2))
          P2[0,1] = P1[0,1] + dy
          if(azir >= 0 and azir <= mth.pi): 
              P2[0,0] = mth.sqrt(h**2 - dy**2) + P1[0,0]
          elif(azir > mth.pi and azir < 2*mth.pi):
              P2[0,0] = -mth.sqrt(h**2 - dy**2) + P1[0,0]
          b = P2[0,0] - P1[0,0] 
          pts = h / spcSTN
          S = np.zeros((1,2))
          S[0] =  b / pts
          S[0,1] =  dy / pts
          if(round(pts) < pts):
                  if(round(pts)%2 == 0):
                      if((pts-round(pts)) < .01):
                          aux = int(round(pts)+1)
                      else:
                          aux = int(round(pts)+2)
                  else:            
                      aux = int(round(pts)+2)
          else:
                  if(round(pts)%2 == 0):
                      if((round(pts)-pts) < .01):
                          aux = int(round(pts)+1)
                      else:
                          aux = int(round(pts)+2)
                  else:            
                      aux = int(round(pts)+2)
          STN = np.zeros((aux,2))
          STN[0,:] = P1[:]
          STN[aux-1,:] = P2[:]
          for n in range (1,aux-1):
              STN[n,0] = STN[n-1,0] + S[0,0]
              STN[n,1] = STN[n-1,1] + S[0,1]
          STN_aux = STN
          ing = input('Quieres ver el grafico? (y/n): ')
          if(ing == 'y'):   
              plt.style.use('seaborn-paper')
              plt.plot(STN[:,0],STN[:,1],'*k')
              plt.minorticks_on()
              plt.grid(b=True,which='major',axis='both',linestyle='-')
              plt.grid(b=True,which='minor',axis='both',linestyle='--')
              plt.xlabel('X')
              plt.ylabel('Y')
              plt.title('Estaciones')
              plt.draw()
              plt.pause(.001)
              plt.show(block=False)
          ml = False
          while(ml == False):
               ing = input('Trazar lineas paralelas? (y/n): ')
               if(ing == 'y'):
                   check = True
                   ml = True
               elif(ing == 'n'): 
                   check = False
                   ml = True
               else:
                   print('¡Respuesta invalida seleciona y = SI o n = NO!')
                   ml = False
          if(ml == True and check == True):  
              while(check == True):
                   if(ml == True):
                       val = False
                       while(val == False):
                            dir = input('Proyecciones de linea...(a = anti-horario o h = horario): ')
                            if(dir == 'a'):
                                val = True
                            elif(dir == 'h'):
                                val = True
                            else:
                                print('¡Opcion no valida a = anti-horario o h = horario!')
                                val = False
                       val = False
                       while(val == False):
                            try:
                                dis = float(input('Angulo en que se trazaran las lineas paralelas [ej.(1-179) 90 ⊥ direccion linea ]: '))
                                val = True
                            except:
                                print('¡Error! Introduce un valor numerico')
                                val = False
                       val = False
                       disr = (dis * mth.pi) / 180
                       while(val == False):
                            try:
                                nlext = int(input('Lineas a trazar = '))
                                val = True
                            except:
                                print('¡Error! Introduce un valor numerico')
                                val = False   
                       val = False
                       while(val == False):
                            try:
                                slext = float(input('Separacion entre lineas = '))
                                val = True
                            except:
                                print('¡Error! Introduce un valor numerico')
                                val = False 
                       if(b == 0):
                           m = mth.inf
                       else:
                           m = dy / b
                       STN = np.zeros((nlext+1,aux,2))
                       STN[0,:,:] = STN_aux
                       if(m < mth.inf):
                           if(m != 0):
                               if(dir == 'a'):
                                   ang = azir - disr 
                               else:
                                   ang = azir + disr
                               h2 = nlext * slext
                               dy2 = h2 * mth.cos(ang)
                               if(azi < 90 or azi > 270):
                                   if((ang >= 0 and ang <= mth.pi) or (ang >= 2*mth.pi and ang <= 3*mth.pi)):
                                       STN[nlext,0,0] = mth.sqrt(h2**2 - dy2**2) + STN[0,0,0,]
                                       STN[nlext,aux-1,0] = mth.sqrt(h2**2 - dy2**2) + STN[0,aux-1,0]
                                   else:
                                       STN[nlext,0,0] = -mth.sqrt(h2**2 - dy2**2) + STN[0,0,0,]
                                       STN[nlext,aux-1,0] = -mth.sqrt(h2**2 - dy2**2) + STN[0,aux-1,0]
                               else:
                                   if((ang >= 0 and ang <= mth.pi) or (ang >= 2*mth.pi and ang <= 3*mth.pi)):
                                       STN[nlext,0,0] = mth.sqrt(h2**2 - dy2**2) + STN[0,0,0,]
                                       STN[nlext,aux-1,0] = mth.sqrt(h2**2 - dy2**2) + STN[0,aux-1,0]
                                   else:
                                       STN[nlext,0,0] = -mth.sqrt(h2**2 - dy2**2) + STN[0,0,0,]
                                       STN[nlext,aux-1,0] = -mth.sqrt(h2**2 - dy2**2) + STN[0,aux-1,0]                     
                               b2 = STN[nlext,0,0] - STN[0,0,0] 
                               S2 = np.zeros((1,2))
                               S2[0] =  b2 / nlext
                               S2[0,1] =  dy2 / nlext
                               for n in range (1,nlext+1):
                                   STN[n,0,0] = STN[n-1,0,0] + S2[0,0]
                                   STN[n,0,1] = STN[n-1,0,1] + S2[0,1]
                                   STN[n,aux-1,0] = STN[n-1,aux-1,0] + S2[0,0]
                                   STN[n,aux-1,1] = STN[n-1,aux-1,1] + S2[0,1]
                               for n in range (1,nlext+1):
                                   for j in range (1,aux-1):
                                       STN[n,j,0] = STN[n,j-1,0] + S[0,0]
                                       STN[n,j,1] = STN[n,j-1,1] + S[0,1]
                           elif(m != 0):
                               for n in range (1,nlext+1):
                                   STN[n,:,0] = STN[n-1,:,0]
                                   if(dir == 'a'):
                                       STN[n,:,1] = STN[n-1,:,1] + slext
                                   else:
                                       STN[n,:,1] = STN[n-1,:,1] - slext
                       else:
                           for n in range (1,nlext+1):
                               STN[n,:,1] = STN[n-1,:,1]
                               if(dir == 'a'):
                                   STN[n,:,0] = STN[n-1,:,0] - slext
                               else:
                                   STN[n,:,0] = STN[n-1,:,0] + slext
                       ing = input('Quieres ver el grafico? (y/n): ')
                       if(ing == 'y'):
                           plt.close()
                           xinf = np.min(STN[:,:,0]) - slext
                           xsup = np.max(STN[:,:,0]) + slext
                           yinf = np.min(STN[:,:,1]) - slext
                           ysup = np.max(STN[:,:,1]) + slext
                           plt.style.use('seaborn-paper')
                           plt.axis([xinf, xsup, yinf, ysup])
                           plt.plot([STN[0,0,0], STN[0,aux-1,0]],[STN[0,0,1], STN[0,aux-1,1]],'-r')
                           for n in range (1,nlext+1):
                               plt.plot([STN[n,0,0], STN[n,aux-1,0]],[STN[n,0,1], STN[n,aux-1,1]],'-c')
                           for n in range (0,nlext+1):
                               plt.plot(STN[n,:,0],STN[n,:,1],'*k')
                           plt.minorticks_on()
                           plt.grid(b=True,which='major',axis='both',linestyle='-')
                           plt.grid(b=True,which='minor',axis='both',linestyle='--')
                           plt.xlabel('X')
                           plt.ylabel('Y')
                           plt.title('Plano de lineas')
                           plt.draw()
                           plt.pause(.001)
                           plt.show(block=False)
                       ing = input('Replantear lineas? (y/n): ')
                       if(ing == 'y'):
                           ml = True
                       else:
                           df = pd.DataFrame({'X':STN[0,:,0],'Y':STN[0,:,1]})
                           np.savetxt("Survey.txt",df.values,fmt='%.2f')
                           vacio = np.array([0,0]).reshape((1,2))
                           for n in range(1,nlext+1):
                               with open('Survey.txt','a') as survey:
                                   df = pd.DataFrame({'X':STN[n,:,0],'Y':STN[n,:,1]})
                                   np.savetxt(survey,vacio,fmt='%.2f')
                                   np.savetxt(survey,df.values,fmt='%.2f')
                           ml = False
                           check = False          
          elif(ml == True and check == False):
              df = pd.DataFrame({'X':STN[:,0],'Y':STN[:,1]})
              np.savetxt("Survey.txt",df.values,fmt='%.2f')
              check = False
          print('%---------------------------------------%')
          print('%               Survey.txt              %')
          print('%     Estaciones por linea separadas    %')
          print('%    con ceros en el archivo de texto   %')
          print('%---------------------------------------%')        
          cont = False
          while(cont == False):
                pregu = input('Quieres usar otra opcion del programa? (y/n):')
                if(pregu == 'y'):
                    op = int(input('¿Que opcion deseas utilizar? (1 al 4): '))
                    proceso = False
                    cont = True 
                elif(pregu == 'n'):
                    proceso = True
                    cont = True
                else: 
                    print('¡Opcion no valida, introduce y o n!')
                    cont = False
      elif(op == 3):
          print('%---------------------------------------%')
          print('% ------- Calculo de estaciones ------- %')
          # print('% --- por Miguel Angel Mares Aguero --- %')
          print('%---------------------------------------%')
          print('%- El programa calcula las estaciones equidistantes entre 2 puntos. ---------------------%')
          print('%- Procurar que el intervalo sea un divisible proporcional a la distancia entre puntos. -%') 
          print('%---------------------------------------%')
          print('-Si eliges importar tabla el formato es-')
          print('%------ Pix(L-1) \t Piy(L-1) ------%')
          print('%------ Pfx(L-1) \t Pfy(L-1) ------%')
          print('%------ Pix(L-2) \t Piy(L-2) ------%')
          print('%------ Pfx(L-2) \t Pfy(L-2) ------%')
          print('%---------- ....   ...    .... ---------%')
          print('%---------- ....   ...    .... ---------%')
          print('%------ Pix(L-N) \t Piy(L-N) ------%')
          print('%------ Pfx(L-N) \t Pfy(L-N) ------%')
          print('%---------------------------------------%')
          ing = int(input('Ingresar puntos por (1 = linea o 2 = tabla): '))
          check = 1
          while(check == 1):
               if(ing == 1):
                   l = int(input('Numero de lineas a segmentar: '))
                   C = np.zeros((l,2,2))
                   print('>>>>>>> Meter datos por linea <<<<<<')
                   n = 0
                   while(n <= l-1):
                        C[n][0][:] = list(map(float,input('Coordenadas punto inicial x y: ').split()))
                        C[n][1][:] = list(map(float,input('Coordenadas punto final x y: ').split()))
                        if(n < l-1):
                          print('------ Meter siguiente linea -------')
                          n += 1
                        else:
                          n += 1
                   check = 2
               elif(ing == 2):
                   root = tk.Tk()
                   root.withdraw()
                   tabla = np.asmatrix(pd.read_fwf(filedialog.askopenfilename(),header=None),dtype=float)
                   l = int(len(tabla[:,0]) / 2)
                   C = np.zeros((l,2,2))
                   n = 0
                   m = 0
                   while(n <= l-1):
                        C[n][0][0] = tabla[m,0]
                        C[n][0][1] = tabla[m,1]
                        C[n][1][0] = tabla[m+1,0]
                        C[n][1][1] = tabla[m+1,1]
                        n += 1
                        m += 2
                   check = 2
               else: 
                   print('--------- Opcion no valida ---------')
                   ing = int(input('Ingresar puntos por 1 = linea o 2 = tabla: '))
                   check = 1
          di = float(input('Distancia entre puntos: '))
          S = np.zeros((l,2))
          h = np.zeros((l))
          pts = np.zeros((l))
          IndL = np.zeros((l))
          for n in range(1,l+1):
              a = C[n-1,1,0] - C[n-1,0,0]
              b = C[n-1,1,1] - C[n-1,0,1]
              h[n-1] = mth.sqrt(a**2 + b**2)
              pts[n-1] = h[n-1] / di
              S[n-1,0] = a / pts[n-1]
              S[n-1,1] = b / pts[n-1]
              if(round(pts[n-1]) < pts[n-1]):
                  if((round(pts[n-1])%2) == 0):
                      if((pts[n-1]-round(pts[n-1])) < .02):
                          aux = int(round(pts[n-1])+1)
                      else:
                          aux = int(round(pts[n-1])+2)
                  else:            
                      aux = int(round(pts[n-1])+2)
              else:
                  if((round(pts[n-1])%2) == 0):
                      if((round(pts[n-1])-pts[n-1]) < .02):
                          aux = int(round(pts[n-1])+1)
                      else:
                          aux = int(round(pts[n-1])+2)
                  else:            
                      aux = int(round(pts[n-1])+2)
              STN = np.zeros((aux,2))
              STN[0,:] = C[n-1,0,:]
              STN[aux-1,:] = C[n-1,1,:]
              for m in range(1,aux-1):
                  STN[m,0] = STN[m-1,0] + S[n-1,0]
                  STN[m,1] = STN[m-1,1] + S[n-1,1]
              if(n == 1):
                  lineas = STN
              else:
                  lineas = np.append(lineas,STN,axis=0)
              linea = 'L' + str(n)
              df = pd.DataFrame({'X':STN[:, 0],'Y':STN[:, 1]})
              if(n == 1):
                 excel = pd.ExcelWriter('STN.xlsx',engine='xlsxwriter')
                 df.to_excel(excel, sheet_name=linea)
                 excel.save()
              else:
                 excel = openpyxl.load_workbook(filename='STN.xlsx')
                 excel.create_sheet(linea)
                 pestana = excel[linea]
                 for r in dataframe_to_rows(df, index=True, header=False):
                  pestana.append(r)
                 for cell in pestana['A'] + pestana[1]:
                     cell.style = 'Pandas'     
                 excel.save(filename='STN.xlsx')
          print('%---------------------------------------%')
          print('%                STN.xlsx               %')
          print('%Se encuentra linea por hoja de calculo %')
          print('%---------------------------------------%')
          ing = input('Quieres ver el grafico? (y/n): ')
          if(ing == 'y'):     
              plt.style.use('seaborn-paper')
              plt.plot(lineas[:,0],lineas[:,1],'b*')
              plt.minorticks_on()
              plt.grid(b=True,which='major',axis='both',linestyle='-')
              plt.grid(b=True,which='minor',axis='both',linestyle='--')
              plt.xlabel('X')
              plt.ylabel('Y')
              plt.title('Estaciones')
              plt.draw()
              plt.pause(.001)
              plt.show(block=False)
          cont = False
          while(cont == False):
                pregu = input('Quieres usar otra opcion del programa? (y/n):')
                if(pregu == 'y'):
                    op = int(input('¿Que opcion deseas utilizar? (1 al 4): '))
                    proceso = False
                    cont = True 
                elif(pregu == 'n'):
                    proceso = True
                    cont = True
                else: 
                    print('¡Opcion no valida, introduce y o n!')
                    cont = False
      elif(op == 4):
          print('%---------------------------------------%')
          print('%---- Extender estaciones de linea -----%')
          # print('% --- por Miguel Angel Mares Aguero --- %')
          print('%---------------------------------------%')
          print('%- El programa calcula estaciones externas a los vertices de lineas previas -%')
          print('%- Introducir las coordenadas de los extremos de la linea. ------------------%') 
          print('%---------------------------------------%')
          print('-Si eliges importar tabla el formato es-')
          print('%------ Pix(L-1) \t Piy(L-1) ------%')
          print('%------ Pfx(L-1) \t Pfy(L-1) ------%')
          print('%------ Pix(L-2) \t Piy(L-2) ------%')
          print('%------ Pfx(L-2) \t Pfy(L-2) ------%')
          print('%---------- ....   ...    .... ---------%')
          print('%---------- ....   ...    .... ---------%')
          print('%------ Pix(L-N) \t Piy(L-N) ------%')
          print('%------ Pfx(L-N) \t Pfy(L-N) ------%')
          print('%---------------------------------------%')
          ing = int(input('Ingresar puntos por (1 = linea o 2 = tabla): '))
          check = 1
          while(check == 1):
               if(ing == 1):
                   l = int(input('Numero de lineas a segmentar: '))
                   C = np.zeros((l,4,2))
                   print('>>>>>>> Meter datos por linea <<<<<<')
                   n = 0
                   while(n <= l-1):
                        C[n][0][:] = list(map(float,input('Coordenadas punto inicial x y: ').split()))
                        C[n][1][:] = list(map(float,input('Coordenadas punto final x y: ').split()))
                        if(n < l-1):
                          print('------ Meter siguiente linea -------')
                          n += 1
                        else:
                          n += 1
                   check = 2
               elif(ing == 2):
                   root = tk.Tk()
                   root.withdraw()
                   tabla = np.asmatrix(pd.read_fwf(filedialog.askopenfilename(),header=None),dtype=float)
                   l = int(len(tabla[:,0]) / 2)
                   C = np.zeros((l,4,2))
                   n = 0
                   m = 0
                   while(n <= l-1):
                        C[n][0][0] = tabla[m,0]
                        C[n][0][1] = tabla[m,1]
                        C[n][1][0] = tabla[m+1,0]
                        C[n][1][1] = tabla[m+1,1]
                        n += 1
                        m += 2
                   check = 2
               else: 
                   print('--------- Opcion no valida ---------')
                   ing = int(input('Ingresar puntos por 1 = linea o 2 = tabla: '))
                   check = 1
          ext = np.zeros((l,2))
          m = np.zeros((1,l))
          h = np.zeros((l,3))
          S = np.zeros((l,4))
          n = 0
          while(n <= l-1):
               ext[n,0] = int(input('Estaciones que se desean extender respecto a la ultima stn de la L'+str(n+1)+' = '))
               ext[n,1] = int(input('Estaciones que se desean extender respecto a la primera stn de la L'+str(n+1)+' = '))
               inter = float(input('Separacion entre estaciones = '))
               a = C[n,1,1] - C[n,0,1]
               b = C[n,1,0] - C[n,0,0]
               if(b == 0):
                     m[0,n] = mth.inf
               else:
                     m[0,n] = a / b
               h[n,0] = mth.sqrt(a**2 + b**2)
               ang = mth.asin(a / h[n,0])
               h[n,1] = h[n,0] + (ext[n,0] * inter)           
               h[n,2] = h[n,0] + (ext[n,1] * inter)    
               C[n,2,1] = h[n,1] * mth.sin(ang) + C[n,0,1]
               if(m[0,n] == 0):
                   if(C[n,0,0] < C[n,1,0]):
                       C[n,2,0] = C[n,0,0] - (h[n,1] - h[n,0])
                   else: 
                       C[n,2,0] = C[n,0,0] + (h[n,1] - h[n,0])
               else: 
                   C[n,2,0] = ((C[n,2,1] - C[n,0,1]) / m[0,n]) + C[n,0,0]
               C[n,3,1] = h[n,2] * mth.sin(-ang) + C[n,1,1]
               if(m[0,n] == 0):
                   if(C[n,0,0] < C[n,1,0]):
                       C[n,3,0] = C[n,1,0] + (h[n,2] - h[n,0])
                   else:
                       C[n,3,0] = C[n,1,0] - (h[n,2] - h[n,0])
               else:
                   C[n,3,0] = ((C[n,3,1] - C[n,1,1]) / m[0,n]) + C[n,1,0]
               aux1 = ext[n,0].astype('int32')
               aux2 = ext[n,1].astype('int32')
               a = C[n,2,1] - C[n,1,1]
               if(a == 0):
                   if(C[n,0,0] < C[n,1,0]):
                       b =  h[n,0] - h[n,1]
                   else:
                       b =  h[n,1] - h[n,0]
               else:
                   b = C[n,2,0] - C[n,1,0]
               if(aux1 != 0):
                   S[n,0] =  b / ext[n,0]
                   S[n,1] =  a / ext[n,0]
                   STN1 = np.zeros((aux1,2))
                   STN1[aux1-1,:] = C[n,2,:]      
                   if(a == 0):
                       STN1[0,0] = C[n,0,0] + S[n,0]
                   else:
                       STN1[0,0] = C[n,1,0] + S[n,0]
                   STN1[0,1] = C[n,1,1] + S[n,1]      
                   for i in range(1,aux1-1):
                       STN1[i,0] = STN1[i-1,0] + S[n,0]
                       STN1[i,1] = STN1[i-1,1] + S[n,1]   
               a = C[n,3,1] - C[n,0,1]
               if(a == 0):
                   if(C[n,0,0] < C[n,1,0]):
                       b = h[n,2] - h[n,0]
                   else:
                       b = h[n,0] - h[n,2]
               else:                
                   b = C[n,3,0] - C[n,0,0]
               if(aux2 != 0):
                   S[n,2] =  b / ext[n,1]
                   S[n,3] =  a / ext[n,1]
                   STN2 = np.zeros((aux2,2))
                   STN2[aux2-1,:] = C[n,3,:]       
                   if(a == 0):
                       STN2[0,0] = C[n,1,0] + S[n,2]            
                   else:
                       STN2[0,0] = C[n,0,0] + S[n,2]
                   STN2[0,1] = C[n,0,1] + S[n,3]              
                   for i in range(1,aux2-1):
                       STN2[i,0] = STN2[i-1,0] + S[n,2]
                       STN2[i,1] = STN2[i-1,1] + S[n,3]
               if((aux1 != 0) and (aux2 != 0)): 
                   STN = np.concatenate((STN1,STN2), axis=0)
               elif(aux1 != 0):
                   STN = STN1
               elif(aux2 != 0):
                   STN = STN2
               vacio = np.array([0,0]).reshape((1,2))
               if(n == 0):
                   df = pd.DataFrame({'X':STN[:,0],'Y':STN[:,1]})
                   np.savetxt("Extension.txt",df.values,fmt='%.2f')
                   STN_aux = STN
               else:    
                   with open('Extension.txt','a') as extension:
                       df = pd.DataFrame({'X':STN[:,0],'Y':STN[:,1]})
                       np.savetxt(extension,vacio,fmt='%.2f')
                       np.savetxt(extension,df.values,fmt='%.2f')
                   STN_aux = np.concatenate((STN_aux,STN), axis=0)
               plt.plot([C[n,0,0], C[n,1,0]],[C[n,0,1], C[n,1,1]],'-r')
               plt.plot([STN[:,0]],[STN[:,1]],'*b')
               n += 1
          print('%---------------------------------------%')
          print('%              Extension.txt            %')
          print('%    Coord separadas por ceros en txt   %')
          print('%---------------------------------------%')
          xinf = np.min(STN_aux[:,0]) - inter
          xsup = np.max(STN_aux[:,0]) + inter
          yinf = np.min(STN_aux[:,1]) - inter
          ysup = np.max(STN_aux[:,1]) + inter
          plt.style.use('seaborn-paper')
          plt.axis([xinf, xsup, yinf, ysup])
          plt.minorticks_on()
          plt.grid(b=True,which='major',axis='both',linestyle='-')
          plt.grid(b=True,which='minor',axis='both',linestyle='--')
          plt.xlabel('X')
          plt.ylabel('Y')
          plt.title('Extensiones')
          plt.draw()
          plt.pause(.001)
          plt.show(block=False)          
          cont = False
          while(cont == False):
                pregu = input('Quieres usar otra opcion del programa? (y/n):')
                if(pregu == 'y'):
                    op = int(input('¿Que opcion deseas utilizar? (1 al 4): '))
                    proceso = False
                    cont = True 
                elif(pregu == 'n'):
                    proceso = True
                    cont = True
                else: 
                    print('¡Opcion no valida, introduce y o n!')
                    cont = False
input('Presiona cualquier tecla para cerrar programa')