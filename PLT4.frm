VERSION 5.00
Begin VB.Form frmDecimalSeparation 
   Caption         =   "Decimal Separation"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDecimal 
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtInteger 
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a decimal Number"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmDecimalSeparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim num As Double
num = Val(txtNumber.Text)
txtInteger.Text = Int(num)
txtDecimal.Text = num - Int(num)



End Sub

