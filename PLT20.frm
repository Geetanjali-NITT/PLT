VERSION 5.00
Begin VB.Form frmStringReverse 
   Caption         =   "Reverse a String"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReverse 
      Caption         =   "Reverse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtOutput 
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtInput 
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Output:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a Text:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frmStringReverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReverse_Click()
Dim lLen As Integer, i As Integer
Dim sChar As String
Dim sAns As String
 
lLen = Len(txtInput.Text)
For i = lLen To 1 Step -1
    sChar = mid(txtInput.Text, i, 1)
    sAns = sAns & sChar
Next
 
txtOutput.Text = sAns
End Sub

