VERSION 5.00
Begin VB.Form frmLinearSearch 
   Caption         =   "Linear Search"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtn 
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
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
      Left            =   5400
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Size of the Array:"
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
      TabIndex        =   5
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
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
      TabIndex        =   3
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label lblResult 
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
      Top             =   1680
      Width           =   3495
   End
End
Attribute VB_Name = "frmLinearSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr(50) As Integer

Private Sub Command1_Click()
Dim n As Integer
Dim length As Integer
Dim index As Integer

n = Val(txtn.Text)
length = n

If Command1.Caption = "Add" Then

For index = 1 To n Step 1
arr(index) = Val(InputBox("Enter Elements"))
Next

txtn.Visible = False
Label1.Caption = "Search Element is"
Command1.Caption = "Search"
Text2.Visible = True

ElseIf Command1.Caption = "Search" Then
Dim items As Integer
items = Val(Text2.Text)

lblResult.Caption = ""
For index = 1 To n
lblResult.Caption = lblResult.Caption & Str(arr(index))

Next
Dim flag As Boolean
flag = False

For index = 1 To n Step 1
If (arr(index) = items) Then
flag = True
End If
Next

If (flag = True) Then
MsgBox ("Element Found")
Else
MsgBox ("Element Not Found")
End If
End If

End Sub

