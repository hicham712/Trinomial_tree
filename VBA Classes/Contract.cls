VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Contract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Strike As Double      'strike de notre option
Public maturity As Date      'date de maturit_$
Public option_type As String        'type d'option
Private Const CallOption As String = "Call"
Private Const PutOption As String = "Put"
Public EU_US As String
Public pricing_date As Date
Public relative_prdate As Double

Public Sub InitializeOption(ByVal callPut As String, ByVal k As Double, _
    ByVal matD As Date, ByVal Opt_EU As String, ByVal prD As Date)
    'initialisaiton de notre classe Option
    Let Me.Strike = k
    Let Me.maturity = matD
    Let Me.option_type = callPut
    Let Me.EU_US = Opt_EU
    Let Me.pricing_date = prD
    Let Me.relative_prdate = CDbl((Me.maturity - Me.pricing_date)) / 365

End Sub

Public Function isCall() As Boolean
    'retourne en boolean si l'option est un call ou non
    Let isCall = UCase(Me.option_type) = "CALL"
End Function

Public Function is_US() As Boolean
    'retourne en boolean si l'option est européenne ou non
    Let is_US = UCase(Me.EU_US) = "US"
End Function

