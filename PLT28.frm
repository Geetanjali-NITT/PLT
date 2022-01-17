VERSION 5.00
Begin VB.Form frmIdentityMatrix 
   Caption         =   "Identity Matrix"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Input"
      Height          =   495
      Left            =   6000
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "frmIdentityMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a() As Integer
Dim i, j, row, col As Integer


Private Sub Command1_Click()

Dim i As Integer
Dim j As Integer

row = Val(InputBox("Enter the size of rows"))
col = Val(InputBox("Enter the size of columns"))
ReDim a(row, col)

For i = 1 To row
   For j = 1 To col
   
   a(i, j) = Val(InputBox("Enter the Elements of Matrix:"))
   Next
Next

Print "Elements of the Matrix:"

For i = 1 To row
   For j = 1 To col
       Print a(i, j); "";
   Next
   Print
Next

Dim flag As Integer
flag = 1
For i = 1 To row
For j = 1 To col
If (i = j And a(i, j) <> 1) Then
flag = 0
ElseIf (i <> j And a(i, j) <> 0) Then
flag = 0
End If
Next
Next

If (flag = 1) Then
MsgBox "Identity Matrix"
Else
MsgBox "Not an Identity Matrix"
End If


End Sub

Private Sub isIdentity(a() As Integer)

End Sub

