# -*- coding: utf-8 -*-
"""
Created on Mon Jan 16 15:41:11 2017

@author: ksirvole
"""
#
import time
import numpy as np
from scipy.optimize import fsolve
from scipy.optimize import root
import matplotlib.pyplot as plt

def findstrainSteel (stress):
    E = 207000e3 # KPa
    Sy = 448e3
    uts = 531e3
    elong = 0.22 - 0.00175*Sy/6.89475908677537/1000
    K = 0.005 - Sy/E
    n = (np.log(elong - uts / E) - np.log(0.005 - Sy / E)) / np.log(uts / Sy)
    
    return (stress/E)+K*(stress/Sy)**n

def findstressCoating (strain):
    A = 71650
    B = -13159
    C = 783.26
    D = 0.7016
    e = strain
    stress = A*e**3+B*e**2+C*e+D
    return stress
    

def findstressSteel2 (strain, SteelProp):
    
    E,Sy,uts = SteelProp # KPa
    
    elong = 0.22 - 0.00175*Sy/6.89475908677537/1000
    K = 0.005 - Sy/E
    n = (np.log(elong - uts / E) - np.log(0.005 - Sy / E)) / np.log(uts / Sy)

    func = lambda stress : (abs(stress)/E)+K*(abs(stress)/Sy)**n - strain
    stress_initial_guess=600
    stress_solu = fsolve(func,stress_initial_guess,xtol=0.5e-6)
    
    return abs(stress_solu)


def findstressSteel3 (strain, SteelProp):
    
    E,Sy,uts = SteelProp # KPa
    
    elong = 0.22 - 0.00175*Sy/6.89475908677537/1000
    K = 0.005 - Sy/E
    n = (np.log(elong - uts / E) - np.log(0.005 - Sy / E)) / np.log(uts / Sy)

    func = lambda stress : ((stress)/E)+K*((stress)/Sy)**n - strain
    stress_initial_guess=600
    stress_solu =  root(func,stress_initial_guess)                        
    return (stress_solu.x)

    
if __name__ == "__main__":
    t0 = time.time()
    SteelProp = [207000e3, 448e3, 531e3]
    e=np.arange(0,0.01,0.01e-2)
    S=e*0
    ec=e*0
    for i in range(len(e)):
        S[i] = findstressSteel3(e[i],SteelProp)
        ec[i]=findstrainSteel(S[i])
#    print(S)
#    St = findstrainSteel(S)
#    print(St)
    plt.plot(e,S,'r-')
    plt.xlabel('strain (%)')
    plt.ylabel('stress (kPa)')
    
    t1 = time.time()
    t = t1-t0
#    print(stress)
    print(t)
    
