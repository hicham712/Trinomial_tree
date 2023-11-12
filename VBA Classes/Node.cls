VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Classe Node qui definira chaque noeud
Public NextUp As Node
Public NextDown As Node
Public NextMid As Node
Public LeftNode As Node
Public RightNode As Node
Public DownNode As Node
Public ProbaUp As Double
Public ProbaDown As Double
Public ProbaMid As Double
Public SPrice As Double
Public Payoff As Double

Public VBTree As Tree
Public Sub InitializeNode(ByVal Price As Double, ByRef Init_Tree As Tree)
'Initialisation d'un noeud
    Let Me.SPrice = Price
    Set Me.VBTree = Init_Tree
End Sub

Public Function MoveUp() As Node
'Creation ou deplacement d'un noeud a celui au dessus de lui-meme
    If Me.LeftNode Is Nothing Then
        Set Me.LeftNode = MakeNode_Tri(Me.SPrice * VBTree.Alpha, VBTree)
        Set Me.LeftNode.RightNode = Me
    End If
    
    Set MoveUp = Me.LeftNode
    
End Function
Public Function MoveDown() As Node
'Creation ou deplacement d'un noeud a celui en dessous de lui-meme
    If Me.RightNode Is Nothing Then
        
        Set Me.RightNode = MakeNode_Tri(Me.SPrice / VBTree.Alpha, VBTree)
        Set Me.RightNode.LeftNode = Me
    End If
    
    Set MoveDown = Me.RightNode
End Function
