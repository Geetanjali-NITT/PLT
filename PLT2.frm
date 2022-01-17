VERSION 5.00
Begin VB.Form frmSwapNumbers 
   Caption         =   "PLT2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult2 
      Height          =   495
      Left            =   12360
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtResult1 
      Height          =   495
      Left            =   12360
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdSwap 
      Caption         =   "Swap and View"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtNumber2 
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtNumber1 
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Result 2"
      Height          =   495
      Left            =   9960
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Result 1"
      Height          =   495
      Left            =   9960
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Enter the Second Number"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the First Number"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "frmSwapNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdSwap_Click()
Dim num1 As Integer
Dim num2 As Integer


num1 = Val(txtNumber1.Text)
num2 = Val(txtNumber2.Text)


txtResult1.Text = num2
txtResult2.Text = num1

End Sub

Private Sub cmdView_Click()
Dim num1 As Integer
Dim num2 As Integer

num1 = Val(txtNumber1.Text)
num2 = Val(txtNumber2.Text)

txtResult1.Text = num1
txtResult2.Text = num2


End Sub

