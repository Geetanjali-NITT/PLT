VERSION 5.00
Begin VB.Form frmPower 
   Caption         =   "Power"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPower 
      Caption         =   "Power"
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
      Left            =   4440
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtn 
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtx 
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "x raised to n:"
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
      Left            =   1800
      TabIndex        =   5
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Enter the value of n:"
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
      Left            =   1800
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the value of x:"
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
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
End
Attribute VB_Name = "frmPower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPower_Click()
Dim x As Integer
Dim n As Integer
Dim i As Integer
Dim result As Long

result = 1
x = Val(txtx.Text)
n = Val(txtn.Text)

For i = 1 To n Step 1
result = result * x
Next

txtResult = Str(result)

End Sub
