# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 20:26:45 2022

@author: milos
"""
import math as ma
import numpy as np
from scipy.stats import norm



def find_next_mid(forward,alpha,node): # node = next_mid du tronc - Dt  =St
                                       # forward = St_-1*exp() -Dt 
    while forward>node.v*(1+alpha)/2:
        node=node.move_up(alpha)
    while forward <=node.v*(1+1/alpha)/2:
        if forward<0:
            break
        node=node.move_down(alpha)
    return node

def pricer(arbre,type_option=None,style_option=None):
    
        r=arbre.market.int_rate
        N=arbre.n_steps 
        T=arbre.contract.maturity    
        K=arbre.contract.strike
        dt=T/N
        d_f=ma.exp(-r*dt)
        if type_option==None:
            type_op=arbre.contract.op_type
        else:
            type_op=type_option
        if style_option==None:   
            op_exercice=arbre.contract.op_exercice
        else:
            op_exercice=style_option
        op_multiplicator=1
        if type_op=="Put":
            op_multiplicator=-1        
        last_node=arbre.racine        
        while last_node.next_mid!=None:
            last_node=last_node.next_mid            
        comput_payoff(op_multiplicator,last_node,K)
        pricing(last_node,d_f,op_exercice,K,op_multiplicator)                                        
        return arbre.racine.v2
    
    
def comput_payoff(op_multiplicator,last_node,K):
    
    current_node=last_node
    while current_node!=None:
        current_node.v2=max((current_node.v-K)*op_multiplicator,0)
        current_node=current_node.voisin_up
    current_node=last_node
    while current_node!=None:
        current_node.v2=max((current_node.v-K)*op_multiplicator,0)
        current_node=current_node.voisin_down
           
def pricing(last_node,d_f,op_exercice,K,op_multiplicator):
    
    while last_node!=None:
        last_node=last_node.voisin_behind
        current_node=last_node        
        while current_node!=None: 
            u=current_node.proba_next_up
            pm=current_node.proba_next_mid
            d=current_node.proba_next_down  
            val=(u*current_node.next_mid.voisin_up.v2+pm*current_node.next_mid.v2+d*current_node.next_mid.voisin_down.v2)*d_f            
            if op_exercice=="US":
                current_node.v2=max(val,max((current_node.v-K)*op_multiplicator,0))
            else:
                current_node.v2=val                
            current_node=current_node.voisin_up
        current_node=last_node
        while current_node!=None: 
            u=current_node.proba_next_up
            pm=current_node.proba_next_mid
            d=current_node.proba_next_down  
            val=(u*current_node.next_mid.voisin_up.v2+pm*current_node.next_mid.v2+d*current_node.next_mid.voisin_down.v2)*d_f           
            if op_exercice=="US":
                current_node.v2=max(val,max((current_node.v-K)*op_multiplicator,0))
            else:
                current_node.v2=val
                
            current_node=current_node.voisin_down

def BS(S, K, T, r, sigma,type_op):
    N = norm.cdf
    d1 = (np.log(S/K) + (r + sigma**2/2)*T) / (sigma*np.sqrt(T))
    d2 = d1 - sigma * np.sqrt(T)
    if type_op=="Call":
        return S * N(d1) - K * np.exp(-r*T)* N(d2)
    else :
        return K*np.exp(-r*T)*N(-d2) - S*N(-d1) 
 

    
        
        
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    