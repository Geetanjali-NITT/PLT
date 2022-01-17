VERSION 5.00
Begin VB.Form frmMultipleof7 
   Caption         =   "Multiple of 7"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   8040
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Calculate1st 2nd and 4th  Multiple of 7  which gives 1 as remainder when divided by 2 3 4 5 and 6"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   2880
      Width           =   4095
   End
End
Attribute VB_Name = "frmMultipleof7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
Dim n As Integer
Dim i As Integer

n = 7
For i = 1 To 100
If (n Mod 2 = 1 And n Mod 3 = 1 And n Mod 4 = 1 And n Mod 5 = 1 And n Mod 6 = 1) Then
Print "n"
n = n * i
End If
Next




End Sub
