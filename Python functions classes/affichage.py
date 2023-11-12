# -*- coding: utf-8 -*-
"""
Created on Mon Dec 12 16:37:42 2022

@author: marko
"""
import xlwings as xw
import math as ma

def afficher_arbre(arbre):
    wb=xw.Book("Pricer_Python.xlsm")
    ws1 = wb.sheets("Arbre")
    ws1.clear()
    node=arbre.racine 
    node_mid=node
    N=arbre.n_steps+1
    for k in range(N):
        
        ws1.range("A1").offset(N,k).value=node_mid.v
        node=node.voisin_up
        i=1
        while node!=None:
            ws1.range("A1").offset(N-i,k).value=node.v
            i=i+1
            node=node.voisin_up
        i=1
        node=node_mid
        node=node.voisin_down 
        while node!=None:
            ws1.range("A1").offset(N+i,k).value=node.v
            i=i+1
            node=node.voisin_down
        node_mid=node_mid.next_mid
        node=node_mid
    ws1.range("A1").offset(N,N+2).value="Underlying price"
        

def afficher_proba(arbre):
    wb=xw.Book("Pricer_Python.xlsm")
    ws1 = wb.sheets("Proba")
    ws1.clear()
    node=arbre.racine 
    node_mid=node
    N=arbre.n_steps+1
    for k in range(N):
        
        ws1.range("A1").offset(N,k).value=node_mid.proba_next_up
        node=node.voisin_up
        i=1
        while node!=None:
            ws1.range("A1").offset(N-i,k).value=node.proba_next_up
            i=i+1
            node=node.voisin_up
        i=1
        node=node_mid
        node=node.voisin_down 
        while node!=None:
            ws1.range("A1").offset(N+i,k).value=node.proba_next_up
            i=i+1
            node=node.voisin_down
        node_mid=node_mid.next_mid
        node=node_mid
    ws1.range("A1").offset(N,N+2).value="Proba next up"
    node=arbre.racine 
    node_mid=node  
    for k in range(N):
        
        ws1.range("A1").offset(3*N,k).value=node_mid.proba_next_down
        node=node.voisin_up
        i=1
        while node!=None:
            ws1.range("A1").offset(3*N-i,k).value=node.proba_next_down
            i=i+1
            node=node.voisin_up
        i=1
        node=node_mid
        node=node.voisin_down 
        while node!=None:
            ws1.range("A1").offset(3*N+i,k).value=node.proba_next_down
            i=i+1
            node=node.voisin_down
        node_mid=node_mid.next_mid
        node=node_mid
    ws1.range("A1").offset(3*N,N+2).value="Proba next down"
    node=arbre.racine 
    node_mid=node 
    for k in range(N):
        
        ws1.range("A1").offset(5*N,k).value=node_mid.proba_next_mid
        node=node.voisin_up
        i=1
        while node!=None:
            ws1.range("A1").offset(5*N-i,k).value=node.proba_next_mid
            i=i+1
            node=node.voisin_up
        i=1
        node=node_mid
        node=node.voisin_down 
        while node!=None:
            ws1.range("A1").offset(5*N+i,k).value=node.proba_next_mid
            i=i+1
            node=node.voisin_down
        node_mid=node_mid.next_mid
        node=node_mid
        ws1.range("A1").offset(5*N,N+2).value="Proba next mid"
        
        
def afficher_prix(arbre):
    wb=xw.Book("Pricer_Python.xlsm")
    ws1 = wb.sheets("Prix")
    ws1.clear()
    node=arbre.racine 
    node_mid=node
    N=arbre.n_steps+1
    for k in range(N):
        
        ws1.range("A1").offset(N,k).value=node_mid.v2
        node=node.voisin_up
        i=1
        while node!=None:
            ws1.range("A1").offset(N-i,k).value=node.v2
            i=i+1
            node=node.voisin_up
        i=1
        node=node_mid
        node=node.voisin_down 
        while node!=None:
            ws1.range("A1").offset(N+i,k).value=node.v2
            i=i+1
            node=node.voisin_down
        node_mid=node_mid.next_mid
        node=node_mid
    ws1.range("A1").offset(N,N+2).value="option price"       
        
  
def afficher_variance(arbre):
    
    wb=xw.Book("Pricer_Python.xlsm")
    ws1 = wb.sheets("Variance")
    ws1.clear()
    node=arbre.racine 
    node_mid=node
    N=arbre.n_steps+1
    r=arbre.market.int_rate
    dt=arbre.dt
    sigma=arbre.market.vol
    for k in range(N):
        var= ((node_mid.v)**2) * ma.exp(2*r*dt) * (ma.exp(sigma**2*dt) -1)
        ws1.range("A1").offset(N,k).value=var
        node=node.voisin_up
        i=1
        while node!=None:
            var= ((node.v)**2) * ma.exp(2*r*dt) * (ma.exp(sigma**2*dt) -1)
            ws1.range("A1").offset(N-i,k).value=var
            i=i+1
            node=node.voisin_up
        i=1
        node=node_mid
        node=node.voisin_down 
        while node!=None:
            var= ((node.v)**2) * ma.exp(2*r*dt) * (ma.exp(sigma**2*dt) -1)
            ws1.range("A1").offset(N+i,k).value=var
            i=i+1
            node=node.voisin_down
        node_mid=node_mid.next_mid
        node=node_mid
    ws1.range("A1").offset(N,N+2).value="Variance"       
        
def afficher_forward(arbre):
    
    wb=xw.Book("Pricer_Python.xlsm")
    ws1 = wb.sheets("Forward")
    ws1.clear()
    node=arbre.racine 
    node_mid=node
    N=arbre.n_steps+1
    r=arbre.market.int_rate
    dt=arbre.dt
    current_date=0
    dividende=arbre.market.dividende
    D=0
    if arbre.market.div_date!=None:
        dividende_date=arbre.market.div_date
        T_div_date= ((dividende_date-arbre.contract.pricing_date).days)/365
    else:
        T_div_date=arbre.contract.maturity+10000000
    
    for k in range(N):
        if T_div_date>current_date and T_div_date<=(current_date+dt) :  # dividende
            D=dividende
            
        forward=node_mid.v*ma.exp(r*dt)-D 
        ws1.range("A1").offset(N,k).value=forward
        node=node.voisin_up
        i=1
        while node!=None:
            forward=node.v*ma.exp(r*dt)-D 
            ws1.range("A1").offset(N-i,k).value=forward
            i=i+1
            node=node.voisin_up
        i=1
        node=node_mid
        node=node.voisin_down 
        while node!=None:
            forward=node.v*ma.exp(r*dt)-D 
            ws1.range("A1").offset(N+i,k).value=forward
            i=i+1
            node=node.voisin_down
        node_mid=node_mid.next_mid
        node=node_mid
        current_date=current_date+dt
        D=0
    ws1.range("A1").offset(N,N+2).value="forward"       
         
        
        
        
        