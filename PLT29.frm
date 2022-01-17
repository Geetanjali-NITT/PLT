VERSION 5.00
Begin VB.Form frmSymmetric 
   Caption         =   "Symmetric Matrix"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frmSymmetric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(), b() As Integer
Dim i, j, row, col As Integer


Private Sub Command1_Click()

Dim i As Integer
Dim j As Integer

row = Val(InputBox("Enter the size of rows"))
col = Val(InputBox("Enter the size of columns"))
ReDim a(row, col)
ReDim b(row, col)

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

Print "Transpose Matrix is:"
For i = 1 To row
For j = 1 To col
b(i, j) = a(j, i)
Print b(i, j); "";
Next
Print
Next

Dim flag As Integer
flag = 1
For i = 1 To row
For j = 1 To col
If a(i, j) <> b(i, j) Then
flag = 0
End If

Next
Next

If flag = 1 Then
MsgBox "Given Matrix is Symmetric"
Else
MsgBox "Given Matrix is not Symmetric"
End If

End Sub

