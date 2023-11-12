Attribute VB_Name = "Price_function"
Option Explicit
Public Function Price_For_Greeks(ByVal St As Double) As Double
    Dim Tree1 As New Tree
    Dim node_init As New Node
    Dim mkt As New MarketData
    Dim Opt As New Contract
    
    ' Methode permettant le calcul du prix de l'option pour les grecques (directement sur excel)
    Application.Calculation = xlCalculationManual

    Set mkt = MakeMarketData_Tri(St, Range("Vol"), Range("IntRate"), Range("Divamount"), Range("divDate"))
    Set Opt = MakeOption(Range("OptType"), Range("Strike"), Range("Mat"), Range("EU_US"), Range("Pr_date"))
    Set Tree1 = MakeTree_Tri(Range("Pr_Date"), Range("Steps"), mkt, Opt)
    Set node_init = MakeNode_Tri(St, Tree1)
    
    Call Tree1.InitializeRoot(node_init)
    
    Call Tree1.MakeTree
    
    Application.Calculation = xlCalculationAutomatic
    
    Price_For_Greeks = Tree1.Opt_price
    
End Function
