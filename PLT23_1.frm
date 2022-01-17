VERSION 5.00
Begin VB.Form frmFibonacciTriangle 
   Caption         =   "Fibonacci Triangle"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtDisplay 
      Height          =   3015
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "PLT23_1.frx":0000
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmFibonacciTriangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
Dim num1 As Integer
Dim num2 As Integer
Dim num3 As Integer


Dim i As Integer
Dim j As Integer
Dim range As Integer
Dim s As String

range = Val(txtNumber.Text)

num1 = 0
num2 = 1

txtDisplay.Text = Str(num2) & vbCrLf

For i = 2 To range
For j = 1 To i
   num3 = num1 + num2
   If (num3 <= range) Then
   s = s & Str(num3)
   num1 = num2
   num2 = num3
   End If
Next

txtDisplay.Text = txtDisplay.Text & vbCrLf & s & vbCrLf
s = ""
Next



End Sub
