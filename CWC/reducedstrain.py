# -*- coding: utf-8 -*-
"""
Created on Tue Jan 17 17:45:43 2017

@author: ksirvole
"""
import numpy as np
import time
from matplotlib import pyplot as plt
from scipy.optimize import fsolve


def findforceConc (strain, area):
    N = np.size(strain)
    e = strain
    A = area
    F = e*0
    S = e*0
    Scom_con = 27.6e3
    emax_con = 0.2e-2
    
    for i in range(0,N):
        
        if e[i]>=0:
            S[i] = 0
        
        elif (e[i]<0 and e[i]>-0.002):
            S[i]= -2*0.9*Scom_con*(abs(e[i])/emax_con)/(1+(abs(e[i])/emax_con)**2) #paraboic stress strain curve Eurocode 2
             
        elif e[i]<=-0.002:
            S[i]=-0.9*Scom_con #crushing limit
    
        F[i]=A[i]*S[i]/10**6
    
    Fa = np.sum(F)

    return (Fa,S,F)
    
    
def findreducedstrain(strain,area,targetforce):
    faci =0.1
    e = strain
    A = area
    TF = targetforce
    
    func = lambda Fac : abs(findforceConc(e*abs(Fac),A)[0]) - TF
    
    Factor = fsolve(func,faci)
    
    ec = e*abs(Factor)
#    print(Factor)
    return ec

    ##

def findstressConc(strain,ConcProp):
    e=strain
    Scom_con = ConcProp[0]
    emax_con = ConcProp[1]
    if e>=0:
     S = 0

    elif (e<0 and e>-0.002):
        S = -2*0.9*Scom_con*(abs(e)/emax_con)/(1+(abs(e)/emax_con)**2)
    
    elif e<=-emax_con:
        S=-0.9*Scom_con
    return S
        
    
if __name__ == "__main__":
    t0 = time.time()
#    reducer = findreducedstrain(e,AreaTot[:,0],Fshc_max)
#    print(reducer)
    e=np.arange(0,0.01,0.01e-2) 
    S=e*0
    for i in range(len(e)):
        S[i]=findstressConc(-e[i],[27.6e3,0.2e-2])
#    print(S)
    plt.plot(e,-S,'r-')
    plt.xlabel('strain (%)')
    plt.ylabel('stress (kPa)')
      
    t1 = time.time()
    t = t1-t0
#    print(Stress)
    print(t)