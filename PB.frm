VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PB 
   Caption         =   "Progress Bar"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form2"
   ScaleHeight     =   3150
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Text            =   "0"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   1320
      Top             =   1680
   End
   Begin ComctlLib.ProgressBar PB2 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1680
      Width           =   375
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   873
      _Version        =   327682
      LargeChange     =   1
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   1
      Max             =   10
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Line Line6 
      X1              =   4680
      X2              =   840
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line5 
      X1              =   840
      X2              =   840
      Y1              =   2640
      Y2              =   3000
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   840
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line3 
      X1              =   4680
      X2              =   1200
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      X1              =   1200
      X2              =   1200
      Y1              =   1560
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
End
Attribute VB_Name = "PB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Stop timer
Timer1.Enabled = False
'Reset text2 to 0
Text2.Text = 0
'Hide PB form
PB.Hide
'Show main form
Form1.Show
End Sub

Private Sub Command2_Click()
'Start Timer
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
'Stop timer
Timer1.Enabled = False
'check to see if timer is still active so u can stop it
If Timer1.Enabled = True Then Command3.Enabled = True Else Command3.Enabled = False
End Sub

Private Sub Command4_Click()
'reset the text box to 0
Text2.Text = 0
End Sub

Private Sub Form_Load()
'Set text1 to display the value of the slider which in turn is the value of the progress bar, either value would work here to display the same info
Text1.Text = Slider1.Value
End Sub

Private Sub Slider1_Change()
'The progress bar equals the value of the slider
PB1.Value = Slider1.Value
'Text1 will show the value of both objects (But takes value from slider1 in this case)
Text1.Text = Slider1.Value
End Sub

Private Sub Text2_Change()
'Validating if the stop button can be active
If Timer1.Enabled = True Then Command3.Enabled = True Else Command3.Enabled = False
'Don't want the program to crash, the bar can only hold a value of 100, reset to 0 after that
If Text2.Text = 100 Then Text2.Text = 0
'The second progress bar's value is equal to the number in text2
PB2.Value = Text2.Text
End Sub

Private Sub Timer1_Timer()
'Add another number in text2 so progress bar 2's value increases
Text2.Text = Text2.Text + 1
End Sub
