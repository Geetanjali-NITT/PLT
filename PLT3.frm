VERSION 5.00
Begin VB.Form frmEvenOdd 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the Number"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmEvenOdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub txtNumber_Change()

Dim num As Integer

num = Val(txtNumber.Text)
If (num Mod 2 = 0) Then
MsgBox num & " " & "is an Even number"
Else
MsgBox num & " " & "is an Odd number"
End If

End Sub
