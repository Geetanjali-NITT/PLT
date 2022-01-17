VERSION 5.00
Begin VB.Form frmStudentDatabase 
   Caption         =   "Student Database"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1800
      TabIndex        =   17
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   ">>"
      Height          =   495
      Left            =   10680
      TabIndex        =   16
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<<"
      Height          =   495
      Left            =   7200
      TabIndex        =   15
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   9000
      TabIndex        =   14
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtTotal 
      Height          =   495
      Left            =   9360
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtAverage 
      Height          =   495
      Left            =   9360
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtSubject3 
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtSubject2 
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtSubject1 
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Student Database"
      Height          =   5295
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   3255
      Begin VB.Label Label5 
         Caption         =   "Student Name"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Subject3"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Subject2"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Subject1"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Label lblResult 
      Caption         =   "Result"
      Height          =   495
      Left            =   7680
      TabIndex        =   18
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Total"
      Height          =   495
      Left            =   7080
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Average"
      Height          =   495
      Left            =   6960
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "frmStudentDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type student
name As String
sub1 As Integer
sub2 As Integer
sub3 As Integer
total As Integer
avg As Double
result As String
End Type

Dim S(20) As student
Dim index As Integer
Dim ci As Integer




Private Sub txtS_Change()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub cmdClear_Click()
txtName.Text = ""
txtSubject1.Text = ""
txtSubject2.Text = ""
txtSubject3.Text = ""
txtTotal.Text = ""
txtAverage.Text = ""
lblResult.Caption = ""
End Sub

Private Sub cmdLeft_Click()
If ci > 0 Then
ci = ci - 1
getrecord (ci)

End Sub

Private Sub cmdSave_Click()
index = index + 1
ci = ci + 1
update (index)

End Sub

Private Sub update(index As Integer)
S(index.name) = txtName.Text

With S(index)
.sub1 = txtSubject1.Text
.sub2 = txtSubject2.Text
.sub3 = txtSubject3.Text
.total = .sub1 + .sub2 + .sub3
.avg = .total / 3
txtAverage.Text = .avg
txtTotal.Text = .total

If (.avg) > 60 Then
lblResult.Caption = "First Class"
ElseIf (.avg) > 50 Then
lblResult.Caption = "Second Class"
ElseIf (.avg) > 35 Then
lblResult.Caption = "Pass"
Else
lblResult.Caption = "Fail"
End If
End With

End Sub


Private Sub getrecord(index As Integer)
With S(index)
txtName.Text = .name
txtSubject1.Text = .sub1
txtSubject2.Text = .sub2
txtSubject3.Text = .sub3
txtAverage.Text = .avg
txtTotal = .total


End Sub

