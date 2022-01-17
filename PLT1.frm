VERSION 5.00
Begin VB.Form frmSimpleInterest 
   Caption         =   "PLT1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Text            =   "Result"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton CalculateSI 
      Caption         =   "Calculate SI"
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtRate 
      Height          =   495
      Left            =   9960
      TabIndex        =   2
      Text            =   "Rate"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtTime 
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Text            =   "Time"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtPrinciple 
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Text            =   "Principle"
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmSimpleInterest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CalculateSI_Click()
Dim P As Integer
Dim T As Integer
Dim r As Integer
Dim SI As Integer
P = Val(txtPrinciple.Text)
T = Val(txtTime.Text)
r = Val(txtRate.Text)
SI = (P * r * T) / 100
txtResult.Text = SI




End Sub

