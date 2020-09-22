VERSION 5.00
Begin VB.Form Timers 
   Caption         =   "Timers"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2655
   LinkTopic       =   "Form2"
   ScaleHeight     =   1575
   ScaleWidth      =   2655
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   720
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reset"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Seconds/Milli"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Timers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Turn off timer on exit
Timer1.Enabled = False
'Hide timers form
Timers.Hide
'Show main form
Form1.Show
End Sub

Private Sub Command2_Click()
'Start Timer
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
'Stop Timer
Timer1.Enabled = False
End Sub

Private Sub Command4_Click()
'Stop the timer first
Timer1.Enabled = False
'Next reset the text field to 0
Text1.Text = "0"
End Sub

Private Sub Form_Load()
'Stop the timer
Timer1.Enabled = False
'Make the text field 0 and allow decimals to the hundreth place
Text1.Text = 0#
End Sub

Private Sub Timer1_Timer()
'This makes 1 second pass in text1 and counts with hundreths as well
Text1.Text = Text1.Text + 0.01
End Sub
