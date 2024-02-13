VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private myName As String
Private myIncome As Integer

Public Property Get Name() As String
    Name = myName
End Property
Public Property Let Name(vName As String)
    myName = vName
End Property

Public Property Get Income() As Integer
    Income = myIncome
End Property
Public Property Let Income(vIncome As Integer)
    myIncome = vIncome
End Property

Public Function Grade() As String
   Dim strGrade As String
   
   If myIncome >= 1000 Then
      strGrade = "鼻盟類"
   ElseIf myIncome >= 600 Then
      strGrade = "醞骯類"
   Else: strGrade = "憮團類"
   End If
   Grade = strGrade
End Function
