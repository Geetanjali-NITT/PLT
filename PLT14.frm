VERSION 5.00
Begin VB.Form frmDecimaltoBinary 
   Caption         =   "Decimal To Binary"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBinary 
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
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
      Left            =   5880
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Result:"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the Number:"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
End
Attribute VB_Name = "frmDecimaltoBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConvert_Click()
Dim n As Integer
Dim result As Integer
Dim r As Integer

n = Val(txtNumber.Text)
While (n > 0)
r = n Mod 2
txtBinary.Text = Str(r) + txtBinary.Text
n = n / 2
Wend


End Sub
