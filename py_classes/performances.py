# -*- coding: utf-8 -*-
"""
Created on Wed Dec  7 21:33:12 2022

@author: marko
"""

from py_classes.market import Market
from py_classes.contract import Contract
from py_classes.arbre import Arbre
import time
from py_classes.functions import pricer,BS
import xlwings as xw
import matplotlib.pyplot as plt
    
def run_tim_to_precision(S, K, r, sigma,type_op,ex_op,today,tomorow,div_date=None,div=0):
  
    wb=xw.Book("Pricer_Python.xlsm")
    ws1 = wb.sheets("Pricer")    
    tps=[]
    m1=Market(S,r,sigma,div_date,div)
    c1=Contract(today,tomorow,K,type_op,ex_op)
    N=[10,20,30,40,50,60,70,80,90,100,200,300,400,500,600,700,800,900,1000]
    prix_bs=BS(S, K, c1.maturity, r, sigma,type_op)
    error=[]
    for k in N:
        start = time.time()
        ar1=Arbre(m1,c1,k)      
        prix=pricer(ar1)        
        end = time.time()
        elapsed = end - start
        tps.append(elapsed)
        err=abs(prix-prix_bs)
        A=[err,elapsed]
        error.append(A)
    B=sorted(error, key=lambda x: x[1]) 
    elapsed_list=[k[1] for k in B ]
    err_list=[k[0] for k in B ]      
    XY=[err_list,elapsed_list]
    for k in range(len(XY[0])):
        ws1.range("C55").offset(k,0).value=XY[0][k]
        ws1.range("C55").offset(k,1).value=XY[1][k]
    
    return XY 

def run_tim_to_steps(S, K, r, sigma,type_op,ex_op,today,tomorow,div_date=None,div=0):
  
    wb=xw.Book("Pricer_Python.xlsm")
    ws1 = wb.sheets("Pricer")    
    tps=[]
    m1=Market(S,r,sigma,div_date,div)
    c1=Contract(today,tomorow,K,type_op,ex_op)
    N=[10,20,30,40,50,60,70,80,90,100,200,300,400,500,600,700,800,900,1000]
    
    for k in N:        
        start = time.time()
        ar1=Arbre(m1,c1,k)      
        prix=pricer(ar1)        
        end = time.time()
        elapsed = end - start
        tps.append(elapsed)
    XY=[N,tps]
    for k in range(len(XY[0])):
        ws1.range("C31").offset(k,0).value=XY[0][k]
        ws1.range("C31").offset(k,1).value=XY[1][k]
    return XY 

    