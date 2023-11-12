# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 20:25:55 2022

@author: milos
"""
import math as ma
from noeud import Noeud
import xlwings as xw


class Arbre: 
        
    def __init__(self,market,contract,n_steps,racine=None):
        
        self.racine=racine
        self.n_steps=n_steps
        self.contract=contract
        self.market=market
        N=n_steps
        self.dt=self.contract.maturity/N
        self.alpha=ma.exp(market.int_rate*self.dt+market.vol*ma.sqrt(3*self.dt))
        self.generer_arbre(N)

              
    def AfficherArbre(self):
        wb=xw.Book("Pricer_Python.xlsm")
        ws1 = wb.sheets("Arbre")
        ws1.clear()
        node=self.racine 
        node_mid=node
        N=self.n_steps+1
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

    def generer_arbre(self,N):   
              
        if self.market.div_date!=None:
            dividende_date=self.market.div_date
            T_div_date= ((dividende_date-self.contract.pricing_date).days)/365
        else:
            T_div_date=self.contract.maturity+10000000
              
        S=self.market.stock_price
        r=self.market.int_rate
        dividende=self.market.dividende
        D=0
        dt=self.dt
        current_date=0
        noeud=Noeud(S,arbre=self)       
        self.racine=noeud
        noeud_tronc=noeud
        
        for k in range(1,N+1,1):
            if T_div_date>current_date and T_div_date<=(current_date+dt) :  # dividende
                D=dividende
            last_next_mid=Noeud(noeud_tronc.v*ma.exp(r*dt)-D,arbre=self)
            
            while noeud!=None:
                last_next_mid=noeud.good_next_mid(noeud.v*ma.exp(r*dt)-D,last_next_mid,D)
                noeud=noeud.voisin_up                
            noeud=noeud_tronc
            last_next_mid=noeud_tronc.next_mid
            while noeud!=None:
                last_next_mid= noeud.good_next_mid(noeud.v*ma.exp(r*dt)-D,last_next_mid,D)
                noeud=noeud.voisin_down
            noeud=noeud_tronc
            noeud=noeud.next_mid
            noeud_tronc=noeud
            current_date=current_date+dt
            D=0
                    



    











