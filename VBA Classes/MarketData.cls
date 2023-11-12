VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MarketData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Classe market qui contiendra les donnees de marche
Public price_data As Double
Public volatility As Double
Public rfRate As Double
Public DivRate As Double
Public divDate As Double

Public Sub InitializeMarketData(ByVal Price As Double, ByVal vol As Double, _
    ByVal initRate As Double, ByVal divY As Double, ByVal divD As Date)
    'Initialisation de notre classe marketdata
    Let Me.price_data = Price
    Let Me.volatility = vol
    Let Me.rfRate = initRate
    Let Me.DivRate = divY
    Let Me.divDate = CDbl(divD)
End Sub

