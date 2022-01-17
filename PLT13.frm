VERSION 5.00
Begin VB.Form frmFactorial 
   Caption         =   "Factorial"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFactorial 
      Caption         =   "Factorial"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the number:"
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
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
End
Attribute VB_Name = "frmFactorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFactorial_Click()
Dim n As Integer
Dim fact As Long
Dim i As Integer

n = Val(txtNumber.Text)
fact = 1

If (n < 0) Then
MsgBox "Factorial of a negative number is not possible"

ElseIf (n = 0) Then
fact = 1
MsgBox fact

Else
For i = 1 To n Step 1
fact = fact * i
Next

MsgBox fact
End If

End Sub
