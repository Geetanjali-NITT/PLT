VERSION 5.00
Begin VB.Form frmReverse 
   Caption         =   "Reverse a Number"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Reverse"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the Number"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmReverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim n As Integer
Dim re As Integer
Dim r As Integer
n = Val(txtNumber.Text)

re = 0
r = 0

While n > 0
re = n Mod 10
r = (r * 10) + re
n = n \ 10

Wend

MsgBox r

End Sub

