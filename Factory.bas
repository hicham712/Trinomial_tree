Attribute VB_Name = "Factory"
Option Explicit
Option Base 1
Public Function MakeMarketData_Tri(ByVal Price As Double, ByVal vol As Double, _
    ByVal initRate As Double, ByVal divY As Double, ByVal divD As Date) As MarketData
    'Initialization of the market data object
    Dim mkt As New MarketData
    Call mkt.InitializeMarketData(Price, vol, initRate, divY, divD)
    Set MakeMarketData_Tri = mkt

End Function
Public Function MakeOption(ByVal callPut As String, ByVal k As Double, _
    ByVal matD As Date, ByVal EU_US As String, ByVal prD) As Contract
    'Initialization of the option object
    Dim Opt As New Contract
    Call Opt.InitializeOption(callPut, k, matD, EU_US, prD)
    Set MakeOption = Opt
End Function
Public Function MakeTree_Tri(ByVal Pr_Date As Date, ByVal Nb_Steps As Double, _
    ByRef Mkt_Data As MarketData, ByRef Opt_Contract As Contract) As Tree
    
    'Initialization of the tree object
    Dim Tree_Tri As New Tree
    Call Tree_Tri.InitializeTree(Pr_Date, Nb_Steps, Mkt_Data, Opt_Contract)
    Set MakeTree_Tri = Tree_Tri

End Function
Public Function MakeNode_Tri(ByVal Price As Double, _
    ByRef Init_Tree As Tree) As Node
    
    'Initialization of the node object
    Dim NodeU As New Node
    Call NodeU.InitializeNode(Price, Init_Tree)
    Set MakeNode_Tri = NodeU

End Function
Public Sub Price_opt()
    Dim Tree1 As New Tree
    Dim node_init As New Node
    Dim mkt As New MarketData
    Dim Opt As New Contract
    Dim before As Single, elapsed As Single
    
    ' Methode permettant le calcul du prix de l'option en passant par toutes les étapes du code
    Application.Calculation = xlCalculationManual
    
    before = Timer()
    ' Initialisation de toutes les classes
    
    Set mkt = MakeMarketData_Tri(Range("St"), Range("Vol"), Range("IntRate"), Range("Divamount"), Range("divDate"))
    Set Opt = MakeOption(Range("OptType"), Range("Strike"), Range("Mat"), Range("EU_US"), Range("Pr_date"))
    Set Tree1 = MakeTree_Tri(Range("Pr_Date"), Range("Steps"), mkt, Opt)
    Set node_init = MakeNode_Tri(Range("St"), Tree1)
    
    Call Tree1.InitializeRoot(node_init)
    
    Call Tree1.MakeTree
    elapsed = Timer() - before
    
    Range("OptPrice") = Tree1.Opt_price
    Range("Time_VBA") = elapsed
    Application.Calculation = xlCalculationAutomatic
    
    
    
End Sub
Public Sub recalculate_greeks()

' Recalcul des greeks

Application.CalculateFull


End Sub


