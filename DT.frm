VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form DT 
   Caption         =   "Date + Time"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2415
   LinkTopic       =   "Form2"
   ScaleHeight     =   4335
   ScaleWidth      =   2415
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19202049
      CurrentDate     =   37501
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1920
      Top             =   480
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Select Date From Calendar"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Millitary Time"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Second Passed Today"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Date"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Date + Time"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Current Time"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "DT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Disable timer on hideing form
Timer1.Enabled = False
'Hide the form
DT.Hide
'Show main form
Form1.Show
End Sub

Private Sub Form_Load()
'On form load enable timer
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
'Tell which text boxes to display what
Text1.Text = DateTime.Time
Text2.Text = DateTime.Date
Text3.Text = DateTime.Now
Text4.Text = DateTime.Timer
Text5.Text = DateTime.Time$
End Sub
