VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public PrDate As Date
Public NbSteps As Double
Public Alpha As Double
Private Const Multi As Double = 3 ^ (1 / 2)
Private Const Days As Integer = 365
Public Root As Node
Public MktData As MarketData
Public Opt As Contract
Public Delta_T As Double
Public Opt_price As Double
Public Sub InitializeTree(ByVal Pr_Date As Date, ByVal Nb_Steps As Double, _
    ByRef Mkt_Data As MarketData, ByRef Opt_Contract As Contract)
    ' Initialisation de la classe arbre
    Let Me.PrDate = Pr_Date
    Let Me.NbSteps = Nb_Steps
    Set Me.MktData = Mkt_Data
    Set Me.Opt = Opt_Contract
    Let Me.Delta_T = Opt.relative_prdate / Me.NbSteps
    Let Me.Alpha = Exp(MktData.rfRate * Me.Delta_T + MktData.volatility * Multi * Sqr(Me.Delta_T))
End Sub
Public Sub InitializeRoot(ByRef init_node As Node)

    Set Me.Root = init_node

End Sub

Public Sub AddProba(ByRef Node1 As Node, ByVal Divammount As Double)
    Dim Esp As Double
    Dim Var As Double
    ' Ajout des proba
    ' Calcul de l'espérance et de la variance des noeuds en fonction du forward
    Esp = Node1.SPrice * Exp(MktData.rfRate * Me.Delta_T) - Divammount
    Var = Node1.SPrice ^ 2 * Exp(2 * MktData.rfRate * Me.Delta_T) * (Exp(MktData.volatility ^ 2 * Me.Delta_T) - 1)
    
    'Attribution des probas
    Node1.ProbaDown = (Node1.NextMid.SPrice ^ (-2) * (Var + Esp ^ 2) - 1 - (Me.Alpha + 1) _
        * (Node1.NextMid.SPrice ^ (-1) * Esp - 1)) _
        / ((1 - Me.Alpha) * (Me.Alpha ^ (-2) - 1))
    
    Node1.ProbaUp = (Node1.NextMid.SPrice ^ (-1) * Esp - 1 - (Me.Alpha ^ (-1) - 1) * Node1.ProbaDown) _
        / (Me.Alpha - 1)
    
    Node1.ProbaMid = 1 - Node1.ProbaUp - Node1.ProbaDown

End Sub
Public Sub MakeTree()
    ' Sub qui permet de creer un arbre à l'aide des autres classes
    Dim NextMidNode As Node, Nodel As Node, Nodel_Tronc As Node, Node_ephemere As Node, _
    OptPrice As Double, Div_Date As Double, Div_amount As Double, Current_Date As Double, i As Integer
    
    ' Gestion d'un manque de date de div
    If MktData.divDate <> 0 Then
        Div_Date = ((MktData.divDate - Opt.pricing_date) / Days)
    Else
        Div_Date = Opt.maturity + 10 ^ 3
    End If
    
    ' Initialization du tronc
    Set Nodel = Me.Root
    Set Nodel_Tronc = Nodel
    
    ' Boucle pour créer les noeuds de l'arbre
    For i = 1 To Me.NbSteps
        
        ' Gestion de la date de dividende pour avoir le bon forward
        If (Div_Date > Current_Date) And (Div_Date <= Current_Date + Me.Delta_T) Then
            Div_amount = MktData.DivRate
        End If
        
        ' Noeud ephemere, auquel on check si son NextMid est bon
        Set Node_ephemere = MakeNode_Tri(Nodel_Tronc.SPrice * Exp(MktData.rfRate * Me.Delta_T) - Div_amount, Me)
                            
        ' Creation des noeuds en haut
        Do Until Nodel Is Nothing
            Call Next_Mid(Nodel.SPrice * Exp(MktData.rfRate * Me.Delta_T) - Div_amount, Nodel, _
            Node_ephemere, Div_amount)
                          
            Set Nodel = Nodel.LeftNode
        Loop
        
        ' retour au tronc
        Set Nodel = Nodel_Tronc
        
        ' Creation des noeuds en bas
        Do Until Nodel Is Nothing
            Call Next_Mid(Nodel.SPrice * Exp(MktData.rfRate * Me.Delta_T) - Div_amount, Nodel, _
            Node_ephemere, Div_amount)
                          
            Set Nodel = Nodel.RightNode
        Loop
        
        ' Retour au tronc et incrémentation des dates
        Set Nodel = Nodel_Tronc
        Set Nodel = Nodel.NextMid
        Set Nodel_Tronc = Nodel
        Current_Date = Current_Date + Me.Delta_T
        Div_amount = 0
        
    Next i
    
    ' Appel du pricing
    
    Me.Opt_price = Pricer()

End Sub

Public Function Good_Next_Mid(ByVal forward As Double, ByRef Nodel As Node) As Node
' Verification de la bonne attribution du NextMid des noeuds
' Et creation de nouveaux noeuds le cas échéant
    Do While forward > Nodel.SPrice * (1 + Me.Alpha) / 2
        Set Nodel = Nodel.MoveUp
    Loop
    Do While forward <= Nodel.SPrice * (1 + 1 / Me.Alpha) / 2
        Set Nodel = Nodel.MoveDown
    Loop
    
    Set Good_Next_Mid = Nodel
End Function
Public Sub Next_Mid(ByVal forward As Double, ByRef Nodel As Node, ByRef Nodel_ephemere As Node, _
    ByVal Div_ammount As Double)
    ' Creation du NextMid des noeuds avec vérification d'adéquation de ce dit nextMid
    Set Nodel.NextMid = Good_Next_Mid(forward, Nodel_ephemere)
    Set Nodel.NextMid.DownNode = Nodel
    Set Nodel.NextMid.LeftNode = Nodel.NextMid.MoveUp
    Set Nodel.NextMid.RightNode = Nodel.NextMid.MoveDown
    
    ' Ajout des probabilités
    Call AddProba(Nodel, Div_ammount)
End Sub
Public Function Pricer() As Double
    Dim Opt_Mult As Integer, CurrentNode As Node, FinalNode As Node, DF As Double
    
    DF = Exp(-MktData.rfRate * Me.Delta_T)
    
    ' Check si l'option est un call ou non
    If Me.Opt.isCall Then Opt_Mult = 1 Else Opt_Mult = -1
    
    Set FinalNode = Me.Root
    
    ' On remonte le tronc du début à la fin
    Do Until FinalNode.NextMid Is Nothing
        Set FinalNode = FinalNode.NextMid
    Loop
    
    ' Calcul des payoffs
    Call compute_payoff(Opt_Mult, FinalNode)
       
    ' Boucle pour calculer le prix de l'option
    Do Until FinalNode Is Nothing
        ' Recul dans l'arbre
        Set FinalNode = FinalNode.DownNode
        Set CurrentNode = FinalNode
        
        Do Until CurrentNode Is Nothing
            ' Calcul du prix de l'option
            Call Get_OptPrice(CurrentNode, DF, Opt_Mult)
            Set CurrentNode = CurrentNode.LeftNode
        Loop
        
        ' Retour au tronc
        Set CurrentNode = FinalNode
        
        Do Until CurrentNode Is Nothing
        ' Calcul du prix de l'option à chaque noeuds du bas
            Call Get_OptPrice(CurrentNode, DF, Opt_Mult)
            Set CurrentNode = CurrentNode.RightNode
        Loop
    Loop
    
    ' Attribution du prix à la racine = le prix aujourd'hui
    Pricer = Me.Root.Payoff
    
End Function

Public Sub Get_OptPrice(ByRef CurrentNode As Node, ByVal DF As Double, ByVal Opt_Mult As Double)
    Dim OptPrice As Double
' Calcul du prix de l'option à chaque noeuds
    With CurrentNode
        OptPrice = DF * (.ProbaUp * .NextMid.LeftNode.Payoff + .ProbaMid * _
            .NextMid.Payoff + .ProbaDown * .NextMid.RightNode.Payoff)
    End With
    
    ' Vérification si l'option est européenne ou américaine et calcul de son prix le cas échéant
    If Opt.is_US Then
        With WorksheetFunction
            CurrentNode.Payoff = .Max(OptPrice, .Max((CurrentNode.SPrice - Opt.Strike) * Opt_Mult, 0))
        End With
    Else
        CurrentNode.Payoff = OptPrice
    
    End If
            
End Sub
Public Sub compute_payoff(ByVal Opt_Mult As Double, ByRef Nodel As Node)
    
    ' Calcul des payoffs de l'option à chaque noeud en partant de la fin du tronc
    
    Dim CurrentNode As Node

    Set CurrentNode = Nodel
    
    Do Until CurrentNode Is Nothing
        CurrentNode.Payoff = WorksheetFunction.Max((CurrentNode.SPrice - Opt.Strike) * Opt_Mult, 0)
        Set CurrentNode = CurrentNode.LeftNode
    Loop
    
    Set CurrentNode = Nodel
    
    Do Until CurrentNode Is Nothing
        CurrentNode.Payoff = WorksheetFunction.Max((CurrentNode.SPrice - Opt.Strike) * Opt_Mult, 0)
        Set CurrentNode = CurrentNode.RightNode
    Loop
    
End Sub
