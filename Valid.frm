VERSION 5.00
Begin VB.Form Valid 
   Caption         =   "Validations"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   LinkTopic       =   "Form2"
   ScaleHeight     =   4335
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command17 
      Caption         =   "0"
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command16 
      Caption         =   "9"
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton Command15 
      Caption         =   "8"
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton Command14 
      Caption         =   "7"
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton Command13 
      Caption         =   "6"
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      Caption         =   "5"
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Caption         =   "4"
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "3"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "2"
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "1"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Reset"
      Height          =   255
      Left            =   2760
      TabIndex        =   22
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   20
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   19
      Text            =   "1234"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "then me"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "me now"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "now me"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "then me"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "me next"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Click me"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Keypad"
      Height          =   2175
      Left            =   360
      TabIndex        =   25
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Input:"
      Height          =   255
      Left            =   2520
      TabIndex        =   24
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Code:"
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   1680
      Width           =   495
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4080
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4080
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "Valid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
'After the button sequence, if check1 has a value of 1 say text
If Check1.Value = 1 Then Text1.Text = "You Have Been Validated"
'if not checked leave blank
If Check1.Value = 0 Then Text1.Text = ""
End Sub

Private Sub Command1_Click()
'On close hide valid form
Valid.Hide
'Then show main form
Form1.Show
End Sub

Private Sub Command10_Click()
'Put a number 3 next in text3
Text3.Text = Text3.Text + "3"
End Sub

Private Sub Command11_Click()
'Put a number 4 next in text3
Text3.Text = Text3.Text + "4"
End Sub

Private Sub Command12_Click()
'Put a number 5 next in text3
Text3.Text = Text3.Text + "5"
End Sub

Private Sub Command13_Click()
'Put a number 6 next in text3
Text3.Text = Text3.Text + "6"
End Sub

Private Sub Command14_Click()
'Put a number 7 next in text3
Text3.Text = Text3.Text + "7"
End Sub

Private Sub Command15_Click()
'Put a number 8 next in text3
Text3.Text = Text3.Text + "8"
End Sub

Private Sub Command16_Click()
'Put a number 9 next in text3
Text3.Text = Text3.Text + "9"
End Sub

Private Sub Command17_Click()
'Put a number 0 next in text3
Text3.Text = Text3.Text + "0"
End Sub

Private Sub Command18_Click()
'gotta reset everything for the keypad
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
Command11.Enabled = True
Command12.Enabled = True
Command13.Enabled = True
Command14.Enabled = True
Command15.Enabled = True
Command16.Enabled = True
Command17.Enabled = True
End Sub

Private Sub Command2_Click()
'Validate the next button to be able to be pressed
Command3.Enabled = True
End Sub

Private Sub Command3_Click()
'Validate the next button to be able to be pressed
Command4.Enabled = True
End Sub

Private Sub Command4_Click()
'Validate the next button to be able to be pressed
Command5.Enabled = True
End Sub

Private Sub Command5_Click()
'Validate the next button to be able to be pressed
Command6.Enabled = True
End Sub

Private Sub Command6_Click()
'Validate the next button to be able to be pressed
Command7.Enabled = True
End Sub

Private Sub Command7_Click()
'Validate the next button to be able to be pressed
Check1.Enabled = True
End Sub

Private Sub Command8_Click()
'Put a number 1 next in text3
Text3.Text = Text3.Text + "1"
End Sub

Private Sub Command9_Click()
'Put a number 2 next in text3
Text3.Text = Text3.Text + "2"
End Sub

Private Sub Form_Load()
'gotta reset the buttons up top so the trick works again when the form loads
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Check1.Enabled = False
End Sub

Private Sub Text2_Change()
'So long as the input doesn't equal the code put this text in text4
If Text3.CausesValidation Then Text4.Text = "XXXXXXXXXX"
'If the input and code match then say this in text4
If Text3.Text = Text2.Text Then Text4.Text = "Validated!"
End Sub

Private Sub Text3_Change()
'Tells text4 to say what if the codes are validated or not on text3 changing
If Text3.CausesValidation Then Text4.Text = "XXXXXXXXXX"
If Text3.Text = Text2.Text Then Text4.Text = "Validated!"
'You can make a little message box appear and say validated or w/e in it when the codes match
'You put this after the code validations so it'll say validated! in text4 before the msgbox comes up
'------------------------------------------------------------------------------------------------------
'*!VB does things in order of the code, so watch where you put each line of code!*
'------------------------------------------------------------------------------------------------------
'If text4's text says validated! then make the keypad unusable until reset
'You can also check if the codes matched to disable the buttons, either way
'Example: If text2.text = text3.text then command8.enabled = false
'And so on for each button, you get the idea.
If Text4.Text = "Validated!" Then Command8.Enabled = False
If Text4.Text = "Validated!" Then Command9.Enabled = False
If Text4.Text = "Validated!" Then Command10.Enabled = False
If Text4.Text = "Validated!" Then Command11.Enabled = False
If Text4.Text = "Validated!" Then Command12.Enabled = False
If Text4.Text = "Validated!" Then Command13.Enabled = False
If Text4.Text = "Validated!" Then Command14.Enabled = False
If Text4.Text = "Validated!" Then Command15.Enabled = False
If Text4.Text = "Validated!" Then Command16.Enabled = False
If Text4.Text = "Validated!" Then Command17.Enabled = False
End Sub

