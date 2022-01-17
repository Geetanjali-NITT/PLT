VERSION 5.00
Begin VB.Form frmBinarySearch 
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   3720
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
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Enter the Element:"
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
      Left            =   2520
      TabIndex        =   7
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
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
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "frmBinarySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr(50) As Integer

Private Sub Command1_Click()
Dim num As Integer
Dim length As Integer
Dim i As Integer
Dim temp As Integer
Dim j As Integer

num = Val(Text1.Text)
length = num

If Command1.Caption = "Add" Then

For i = 1 To num Step 1
arr(i) = Val(InputBox("Enter Elements"))
Next
'--------------------------------------------------------------------------------------
Label2.Caption = ""
For i = 1 To num
Label2.Caption = Label2.Caption & Str(arr(i))
Next

Command1.Caption = "Sort"
'-----------------------------------------------------------------------------------------

ElseIf Command1.Caption = "Sort" Then

For i = 0 To num Step 1
            For j = num To i + 1 Step -1
                If (arr(j) < arr(j - 1)) Then
                    temp = arr(j)
                    arr(j) = arr(j - 1)
                    arr(j - 1) = temp
                End If
            Next
        Next
        
Label4.Caption = ""
For i = 1 To num
Label4.Caption = Label4.Caption & Str(arr(i))
Next

Label3.Caption = "Array Elements after Sorting"
Text1.Visible = False
Label5.Caption = "Search Element is"
Command1.Caption = "Search"
Text2.Visible = True


'-----------------------------------------------------------------------
ElseIf Command1.Caption = "Search" Then
Dim items As Integer
items = Val(Text2.Text)
Dim flag As Boolean
flag = False
Dim mid As Integer

mid = num / 2
If (arr(mid) = items) Then
flag = True

ElseIf (arr(mid) > items) Then
For i = 1 To mid Step 1
If (arr(i) = items) Then
flag = True
End If
Next

ElseIf (arr(mid) < items) Then
For i = mid To num Step 1
If (arr(i) = items) Then
flag = True
End If
Next

End If

If (flag = True) Then
MsgBox ("Element Found")
Else
MsgBox ("Element Not Found")
End If

End If

End Sub

