VERSION 5.00
Begin VB.Form frmEmployee 
   Caption         =   "Employee Database"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNetAnnual 
      Height          =   495
      Left            =   10680
      TabIndex        =   25
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtAnnual 
      Height          =   495
      Left            =   10680
      TabIndex        =   24
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtGross 
      Height          =   495
      Left            =   10680
      TabIndex        =   23
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtTaxInvestment 
      Height          =   495
      Left            =   5640
      TabIndex        =   22
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtBonus 
      Height          =   495
      Left            =   5640
      TabIndex        =   21
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtAllowance 
      Height          =   495
      Left            =   5640
      TabIndex        =   20
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtBasic 
      Height          =   495
      Left            =   5640
      TabIndex        =   19
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtEmpId 
      Height          =   495
      Left            =   5640
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   5640
      TabIndex        =   17
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   ">>"
      Height          =   495
      Left            =   11640
      TabIndex        =   16
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   10200
      TabIndex        =   15
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<<"
      Height          =   495
      Left            =   8760
      TabIndex        =   14
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Controls"
      Height          =   3255
      Left            =   8040
      TabIndex        =   13
      Top             =   4800
      Width           =   5415
      Begin VB.Label lblIndexDisplay 
         Height          =   495
         Left            =   720
         TabIndex        =   26
         Top             =   2280
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Salary"
      Height          =   3375
      Left            =   8040
      TabIndex        =   9
      Top             =   960
      Width           =   5295
      Begin VB.Label Label9 
         Caption         =   "Anual Salary"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Anual Net Salary"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Gross Salary"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employee"
      Height          =   7215
      Left            =   3480
      TabIndex        =   0
      Top             =   960
      Width           =   3735
      Begin VB.Label Label7 
         Caption         =   "Monthly Tax Investment"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "% of Bonus"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Special Allowances"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Basic Salary"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Employee ID"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Employee
name As String
empid As Integer
basic As Long
allowance As Long
bonus As Long
bpercent As Long
taxinvestment As Long
gross As Double
annual As Double
netannual As Double
End Type

Dim index As Integer
Dim ci As Integer
Dim E(50) As Employee
Dim tax As Integer




Private Sub cmdClear_Click()
txtName.Text = ""
txtEmpId.Text = ""
txtBasic.Text = ""
txtAllowance.Text = ""
txtBonus.Text = ""
txtTaxInvestment.Text = ""
txtGross.Text = ""
txtAnnual.Text = ""
txtNetAnnual.Text = ""


End Sub

Private Sub cmdSave_Click()
index = index + 1
ci = ci - 1
update (index)
End Sub

Private Sub update(index As Integer)
E(index).name = txtName.Text

With E(index)
.empid = txtEmpId.Text
.basic = txtBasic.Text
.allowance = txtAllowance.Text
.bonus = txtBonus.Text
.taxinvestment = txtTaxInvestment.Text
.gross = .basic + .allowance


.annual = .gross + .bonus

If (((.annual) - (.taxinvestment)) <= 100000) Then
tax = 0

ElseIf (100000 < ((.annual) - (.taxinvestment)) < 150000) Then
tax = (20 * (.annual)) / 100

Else
tax = (30 * (.annual)) / 100
End If



.netannual = (.annual) - tax


txtGross.Text = .gross
txtAnnual.Text = .annual
txtNetAnnual.Text = .netannual







End With

End Sub
Private Sub checkfieldinput()
If (Len(txtName.Text) And IsNull(txtEmpId)) > 0 Then
cmdSave.Enabled = True
cmdClear.Enabled = True
End If

End Sub
Private Sub Form_Load()
cmdSave.Enabled = False
cmdLeft.Enabled = False
cmdRight.Enabled = False
cmdClear.Enabled = False

End Sub

Private Sub txtAllowance_Change()
checkfieldinput
End Sub

Private Sub txtBasic_Change()
checkfieldinput
End Sub

Private Sub txtBonus_Change()
checkfieldinput
End Sub

Private Sub txtEmpId_Change()
checkfieldinput
End Sub

Private Sub txtName_Change()
checkfieldinput
End Sub

Private Sub txtTaxInvestment_Change()
checkfieldinput
End Sub
