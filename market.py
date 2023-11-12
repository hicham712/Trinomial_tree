# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 20:24:41 2022

@author: milos
"""

class Market:
    
    def __init__(self,stock_price,int_rate,vol,div_date=None,dividende=0):
        
        self.stock_price=stock_price
        self.int_rate=int_rate
        self.vol=vol
        self.div_date=div_date
        self.dividende=dividende