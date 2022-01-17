VERSION 5.00
Begin VB.Form frmStringPalindrome 
   Caption         =   "String Palindrome"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStringPalindrome 
      Caption         =   "Check Palindrome"
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
      Left            =   4320
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtInput 
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a Text:"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "frmStringPalindrome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStringPalindrome_Click()
Dim lLen As Integer, lCtr As Integer
Dim sChar As String
Dim sAns As String

lLen = Len(txtInput.Text)
For lCtr = lLen To 1 Step -1
    sChar = mid(txtInput.Text, lCtr, 1)
    sAns = sAns & sChar
Next
 
If (txtInput.Text = sAns) Then
MsgBox "Palindrome"
Else
MsgBox "Not Palindrome"
End If

End Sub
