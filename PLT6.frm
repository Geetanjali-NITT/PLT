VERSION 5.00
Begin VB.Form frmLargest 
   Caption         =   "Largest of 3 Numbers"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "GetLargest"
      Height          =   495
      Left            =   6720
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtNumber3 
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtNumber2 
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtNumber1 
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Enter Third Number"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Second Number"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter First Number"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmLargest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim num1 As Integer
Dim num2 As Integer
Dim num3 As Integer

num1 = Val(txtNumber1.Text)
num2 = Val(txtNumber2.Text)
num3 = Val(txtNumber3.Text)

If (num1 > num2 And num1 > num3) Then
MsgBox num1 & " " & "is the largest Number "
 If (num2 > num3) Then
 MsgBox num2 & " " & "is the Second largest Number "
 Else
 MsgBox num3 & " " & "is the Second largest Number "
 End If

ElseIf (num2 > num1 And num2 > num3) Then
MsgBox num2 & " " & "is the largest Number "
 If (num1 > num3) Then
 MsgBox num1 & " " & "is the Second largest Number "
 Else
 MsgBox num3 & " " & "is the Second largest Number "
End If

ElseIf (num3 > num1 And num3 > num2) Then
MsgBox num3 & " " & "is the largest Number "
 If (num1 > num2) Then
 MsgBox num1 & " " & "is the Second largest Number "
 Else
 MsgBox num2 & " " & "is the Second largest Number "
 End If

End If





End Sub
