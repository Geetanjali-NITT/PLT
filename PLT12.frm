VERSION 5.00
Begin VB.Form frmSumofPrimes 
   Caption         =   "Sum of Primes"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSum 
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox txtDisplay 
      Height          =   1215
      Left            =   4440
      TabIndex        =   5
      Top             =   4800
      Width           =   5535
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Display Primes and Sum"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox txtn 
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtm 
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Final Sum:"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Display Primes between m and n :"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   4920
      Width           =   2535
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
      Left            =   1080
      TabIndex        =   3
      Top             =   2520
      Width           =   2355
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the value of m:"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   2355
   End
End
Attribute VB_Name = "frmSumofPrimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
Dim m As Integer
Dim n As Integer
Dim i As Integer
Dim j As Integer
Dim sum As Integer
Dim flag As Integer



sum = 0
m = Val(txtm.Text)
n = Val(txtn.Text)

For i = m To n Step 1

   For j = 2 To (i / 2) Step 1
   flag = 1
   If (i Mod j = 0) Then
   Exit For
   Else
   flag = 0
'   txtDisplay.Text = txtDisplay.Text + " " +Str(i)
   
   End If
   Next
   If (flag = 0) Then
   txtDisplay.Text = txtDisplay.Text + " " + Str(i)
   sum = sum + i
   End If
Next
   
   txtSum.Text = sum
   


End Sub

