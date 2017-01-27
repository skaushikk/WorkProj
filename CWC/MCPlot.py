# -*- coding: utf-8 -*-
"""
Created on Mon Jan 16 12:20:31 2017

@author: ksirvole
"""

# -*- coding: utf-8 -*-
"""
Created on Sun Jan 15 15:06:32 2017

@author: ksirvole
"""

import math
import xlwings as xl
import numpy as np
import matplotlib.pyplot as plt
from stresstrain import *
from reducedstrain import *
import time



OD = 508  #nominal diameter
Nl = 5  #Number of Layers

#ResidualAxial force
Fresi = 0
Tl=np.zeros(Nl)
for i in range(Nl):
    Tl[i] = 0
Tcon = 40
Tacl=3.2
Tst =22.2
Tcra = 0
Tcore = OD/2 - Tst - Tcra
Tl=[Tcon, Tacl, Tst, Tcra, Tcore]

#Diameters
ODl = np.zeros(Nl)

ODl[-1]=Tl[-1]*2
arr = np.arange(Nl)[::-1][:-1]  #arr = [4,3,2,1,0]

for i in arr:
    ODl[i-1]=ODl[i]+2*Tl[i-1]

Rl = [x/2 for x in ODl]

#Layer Thicknesses
nl=np.zeros(Nl)
nl[0] = 100
nl[1] = 50
nl[2] = 100
nl[3] = 0
nl[4] = 600
nl.astype(int)
t0= time.time()

ts = Tl/nl

#Ns = [Ncon, Nacl, Nst, Ncore]
Nt = nl*0
Ntot = np.sum(nl).astype(int)

#==============================================================================
#                               MATERIAL PROPERTIES
## CONCRETE
emax_con = 0.2e-2 # Strain corresponding to the max stress
Scom_con = 27.6e3 # kpa Compressive strength of concrete kpa
ConcProp = [Scom_con, emax_con]

## Anti Corrosion Layer
Ssh_acl = 30 # kpa 5 PSI Bond Strength - Shear strength of the ACL material (kpa)
Sy_acl = Ssh_acl * np.sqrt(3) # Yield stress of the ACL material
E_acl = 889700 # kpa Elastic modulus of the ACL is assumed to be 1 percent of the Steel elastic modulus

# Cladding
E_cra = 206800e3  #kPa

#Line Pipe - Steel
Esteel = 207000e3 #kPa
Systeel = 448e3   #kPa
UTSsteel = 531e3  #kPa
SteelProp = [Esteel, Systeel, UTSsteel]
#==============================================================================

temp=0
for i in range(0,Nl):
    temp = temp+nl[i]
    Nt[i]=temp

Nt=Nt.astype(int)

LtoFJ = 2500 #Distance of the section that moment is calculated from the field joint
Fshc_max = (np.pi)*OD*LtoFJ*Ssh_acl/(10**(6)) #The shear force that stops concrete fom sliding

#
i=0
#
t=np.zeros(Ntot)
y=np.zeros(Ntot)
theta = np.zeros((Ntot,Nl))
area = np.zeros((Ntot,Nl))
AreaS = np.zeros((Ntot,Nl))
MomentArea=np.zeros((Ntot*2,Nl))

i=0
j=0
for j in range(Nl):    
    for n in range(i,Ntot):
       if (n < Nt[j]):
           t[i]= ts[j]
           y[i]= Rl[0]-(np.sum(t)-t[i]/2)
           i+=1
    j+=1

j=0   
for R in Rl:
     for i in range(0,Ntot):
         # projection angle calculation
         if y[i]>R:
             theta[i,j]=0.0  #assign theta =0 for lever arm greater than radius
         else:
             theta[i,j] = 2*math.acos((y[i]-t[i]/2)/R)  #calculate theta for the strip
         # area calculation
         if (i == Nt[j-1] or i == 0):
             area[i,j] = (R**2*(theta[i,j]-math.sin(theta[i,j]))/2)
         else:
             area[i,j] = (R**2*(theta[i,j]-math.sin(theta[i,j]))/2) - (R**2*(theta[i-1,j]-math.sin(theta[i-1,j]))/2)
     j=j+1
#     print(j)     

# Area assignment
for j in range(0,Nl):
    try:
        AreaS[:,j] = area[:,j]-area[:,j+1] #assign areas to sections
    except:
        AreaS[:,j] = area[:,j]    #assign area to inner core
AreaTot = np.concatenate((AreaS,AreaS[::-1]),axis=0)
yTot = np.concatenate((-y,y[::-1]),axis=0)

# Moment of area calcualtion
for j in range(0,Nl):
    MomentArea[:,j] = AreaTot[:,j]*yTot[:]**2

## Moment Calculation for Global Strain
#==============================================================================
#==============================================================================
Egmax = 0.05e-1
estep = 0.05e-2
Eg = np.arange(0,Egmax,estep)
point = 0
yNA = Eg*0
Mplot = Eg*0
Kplot = Eg*0
Fsteelplot =Eg*0
Faclplot = Eg*0
Fconplot= Eg*0

yi = yTot


#for k in range(0,5):
for k in range(len(Eg)):
    
    error = []
    dl=10 #neutral axis movement steps
    er = 100
    tol=1e-2
    count=0
    yTot = yi
    kappag = 2*Eg[k]/OD #curvature based on global strain
    ei= np.round(yi*kappag,6)

    while abs(er) > tol: #check for neutral axis position, check = 1 means neutral axis position reached
             
         e = np.round(yTot*kappag,6)
    #     if k==1:
    #         check = 0
         Sst = Scon = Sacl = e*0
         Sax = AreaTot*0
    #==============================================================================
    # ## Stress Strain Relation for different materials
    #           Steel:
    #                  Ramberg-Osgood Parameters
    #                   K = 0.00276			
    #                   n = 25.8467
    #           ACL coating: strain % = x
    #                   stress = 71650x3 - 13159x2 + 783.26x + 0.7016
    #           Concrete: baed on Eurocode 2
    #                   
    #                   
    #==============================================================================
    ##         ## Stress Calculation for steel, acl and concrete
         for i in range(0,2*Ntot):
    
             # Steel Stress
             Sax[i][2] = np.sign(e[i])*findstressSteel2(abs(e[i]),SteelProp)
                 
             # ACL stress
             if abs(e[i])<Sy_acl/E_acl:
                 Sax[i][1]= E_acl*(e[i])
             else:
                 Sax[i][1]= Sy_acl*np.sign(e[i])
                 
             #CRA stress
             if abs(e[i])<0.2e-2:
                 Sax[i][3]= E_cra*(e[i])
             else:
                 Sax[i][3]= E_cra*0.2e-2
                 
                 
             # Concrete Stress
             Sax[i][0]=findstressConc(e[i],ConcProp)  #crushing limit
    
    #==============================================================================
    ##  calculate total axial force in each layer
         Fax =AreaTot*0            
         Fax = AreaTot*Sax/10**6
               
    
    #     Check if concrete slips
         if Tacl > 0:
             econ = e*0
             Fax_con_tot = np.sum(Fax[:,0])
             if abs(Fax_con_tot) > Fshc_max:
                 slide = True
                 econ=findreducedstrain(e,AreaTot[:,0],Fshc_max) #recalculated strains based on ACL slippage
                 reducedConcload = findforceConc(econ,AreaTot[:,0])
                 Sax[:,0] = reducedConcload[1]
                 Fax[:,0] = reducedConcload[2]
                 print(slide)
        #==============================================================================
        # Recalculate Stresses and Forces
#             for i in range(0,2*Ntot):
#        
#                 # Steel Stress
#                 Sax[i][2] = np.sign(e[i])*findstressSteel2(abs(e[i]),SteelProp)
#                     
#                 # ACL stress
#                 if abs(e[i])<Sy_acl/E_acl:
#                     Sax[i][1]= E_acl*(e[i])
#                 else:
#                     Sax[i][1]= Sy_acl*np.sign(e[i])
#                     
#                 #CRA stress
#                 if abs(e[i])<0.2e-2:
#                     Sax[i][3]= E_cra*(e[i])
#                 else:
#                     Sax[i][3]= E_cra*0.2e-2  
    #normalize area
         striparray = np.divide(AreaTot,AreaTot)
         striparray[np.isnan(striparray)]=0
    
         Sax=striparray*Sax              
                 
    # Recalculate the forces
         Fax = AreaTot*Sax/10**6
                
        
    #==============================================================================
    
     ## Calculate Moments 
    #     for i in range(0,2*Ntot):
         Ftot = np.sum(Fax,axis=1)
         Mtot = np.multiply(Ftot,yTot)
         Mtotsum = np.sum(Mtot)
         Mst = np.multiply(Fax[:,2],yTot)
         Mstsum = np.sum(Mst)
         Ftotsum = np.sum(Ftot)
         Fsum = np.sum(Fax,axis=0)

         
         er = Ftotsum-Fresi
         error.append(er)
              
         if abs(er) > tol:
             if (np.sign(error[count])/np.sign(error[count-1]))==-1:
                 dl=dl/2
             yTot = yTot - np.sign(er)*dl
             count+=1
         print(point,count,er,dl)
#         print(Sax[1,0])
  
    Mplot[point] = Mtotsum
    Kplot[point] = kappag
    yNA[point] = yTot[0]-yi[0]
    Fsteelplot[point]=Fsum[2]
    Faclplot[point]=Fsum[1]
    Fconplot[point]=Fsum[0]
    point+=1
#    print(point)
         
    t1 = time.time()
    print(t1-t0)    