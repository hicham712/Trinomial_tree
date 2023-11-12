# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 20:29:01 2022

@author: milos
"""
import time
from py_classes.arbre import Arbre
from py_classes.market import Market
from py_classes.contract import Contract
from py_classes.functions import pricer,BS
import xlwings as xw
from py_classes.affichage import afficher_proba,afficher_prix,afficher_variance,afficher_forward,afficher_arbre

@xw.func
def main():
    
    wb=xw.Book("Pricer_Python.xlsm")
    ws1 = wb.sheets("Pricer")

    today=ws1.range("L3").value
    S=ws1.range("L4").value
    r=ws1.range("L5").value
    sigma=ws1.range("L6").value
    dividende=ws1.range("L7").value
    div_date=ws1.range("L8").value
    tomorow=ws1.range("L9").value
    type_op=ws1.range("L10").value
    ex_op=ws1.range("L11").value
    K=ws1.range("L12").value
    N=int(ws1.range("L13").value)
    affichage=ws1.range("L14").value
    m1=Market(S,r,sigma,div_date,dividende)
    c1=Contract(today,tomorow,K,type_op,ex_op)
    start = time.time()
    ar1=Arbre(m1,c1,N)
    end = time.time()
    elapsed = end - start
    ws1.range("L18").value=elapsed
    start = time.time()
    prix=pricer(ar1)
    end = time.time()
    elapsed = end - start
    ws1.range("L19").value=elapsed
    ws1.range("L16").value=prix
    if ex_op=="EU":    
        ws1.range("L15").value=BS(S, K, c1.maturity, r, sigma,type_op)
    else:
        ws1.range("L15").value="None"
    if ws1.range("L21").value !="Oui":
        ws1.range("L21").value="Non"
    if affichage=="Arbre":
        afficher_arbre(ar1)
    if affichage=="Prix":
        afficher_prix(ar1)
    if affichage=="Forward":
        afficher_forward(ar1)
    if affichage=="Proba":
        afficher_proba(ar1)
    if affichage=="Variance":
        afficher_variance(ar1) 
    if affichage=="All":
        afficher_arbre(ar1)
        afficher_prix(ar1)
        afficher_forward(ar1)
        afficher_proba(ar1)
        afficher_variance(ar1)
        
    
    
    
if __name__=='__main__':
    
    main()
    
    
    
    
    
    
    
    