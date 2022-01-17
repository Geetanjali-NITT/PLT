VERSION 5.00
Begin VB.Form frmNumtoWords 
   Caption         =   "Display Number in Words"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Convert"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   3600
      Width           =   6135
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a  Number"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmNumtoWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim num As Integer
Dim digit As String
Dim r As Integer

num = Val(txtNumber.Text)

While num > 0
r = num Mod 10
Select Case r
Case 1
digit = "One"
Case 2
digit = "Two"
Case 3
digit = "Three"
Case 4
digit = "Four"
Case 5
digit = "Five"
Case 6
digit = "Six"
Case 7
digit = "Seven"
Case 8
digit = "Eight"
Case 9
digit = "Nine"
Case 0
digit = "Zero"

End Select
txtResult.Text = digit & " " & txtResult.Text
num = num \ 10

Wend

End Sub

