# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 20:25:00 2022

@author: milos
"""
import math as ma
from functions import find_next_mid
import xlwings as xw

class Noeud: 
    
    def __init__(self, v=None, next_up = None, next_down = None , next_mid= None,v2=None,proba_next_up=None,proba_next_down=None,proba_next_mid=None,voisin_behind=None,voisin_up=None,voisin_down=None,arbre=None): 
        self.next_up = next_up 
        self.next_down = next_down
        self.next_mid = next_mid 
        self.v = v
        self.v2 = v2
        self.proba_next_up = proba_next_up
        self.proba_next_down = proba_next_down
        self.proba_next_mid = proba_next_mid
        self.voisin_behind=voisin_behind
        self.voisin_up=voisin_up
        self.voisin_down=voisin_down
        self.arbre=arbre

    
    def move_up(self,alpha,div=False):
        
        if self.voisin_up != None:
            return self.voisin_up
        else:
            self.voisin_up=Noeud(self.v*alpha,arbre=self.arbre)
            self.voisin_up.voisin_down=self
            return self.voisin_up

    
    def move_down(self,alpha,div=False):
       
        if self.voisin_down != None:
            return self.voisin_down
        else:
            self.voisin_down=Noeud(self.v/alpha,arbre=self.arbre)
            self.voisin_down.voisin_up=self
            return self.voisin_down


    def add_probability(self,D):
        r=self.arbre.market.int_rate
        dt=self.arbre.dt 
        alpha=self.arbre.alpha
        sigma=self.arbre.market.vol
        Espe=self.v*ma.exp(r*dt)-D        
        var= ((self.v)**2) * ma.exp(2*r*dt) * (ma.exp(sigma**2*dt) -1) 
        p_down=(self.next_mid.v**(-2)*(var+Espe**2)-1-(alpha+1)*(self.next_mid.v**(-1)*Espe-1))/((1-alpha)*(alpha**(-2)-1))
        p_up=(self.next_mid.v**(-1)*Espe-1-(alpha**(-1)-1)*p_down)/(alpha-1)
        p_mid=1-p_up-p_down
        if p_mid<0 or p_up<0 or p_down<0:
            wb=xw.Book("Pricer_Python.xlsm")
            ws1 = wb.sheets("Pricer")
            ws1.range("L21").value="Oui"
        self.proba_next_down=p_down
        self.proba_next_up=p_up
        self.proba_next_mid=p_mid

    
    def good_next_mid(self,forward,last_next_mid,D):
        alpha=self.arbre.alpha
        self.next_mid=find_next_mid(forward,alpha,last_next_mid)
        self.next_mid.voisin_behind=self
        self.next_mid.voisin_up=self.next_mid.move_up(alpha)
        self.next_mid.voisin_down=self.next_mid.move_down(alpha)
        self.add_probability(D)
        return self.next_mid










