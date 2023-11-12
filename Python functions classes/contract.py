# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 20:23:32 2022

@author: milos
"""

class Contract:
    
    def __init__(self,pricing_date,maturity_date,strike,op_type=None,op_exercice=None):
          
        self.maturity_date=maturity_date
        self.pricing_date=pricing_date
        
        self.maturity=((self.maturity_date-self.pricing_date).days)/365
        self.strike=strike
        self.op_type=op_type
        self.op_exercice=op_exercice