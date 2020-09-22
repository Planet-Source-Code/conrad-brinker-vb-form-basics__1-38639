VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB For Beginners"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Validations"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Timers"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Combo Boxes"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Buttons"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Dialog Boxes"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Credits"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Text Boxes"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Date + Time"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Files"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Progress Bar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   1800
      Picture         =   "Form1.frx":0000
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Show and hide forms
PB.Show
Form1.Hide
'Reset objects when you go back into form
If PB.Timer1.Enabled = True Then PB.Command3.Enabled = True Else PB.Command3.Enabled = False
End Sub

Private Sub Command10_Click()
'Close Program
End
End Sub

Private Sub Command11_Click()
'Create a messagebox and say what's in quotes
MsgBox "This is a message box"
End Sub

Private Sub Command12_Click()
'Show buttons form
Buttons.Show
'Reset button caption
Buttons.Command1.Caption = "Test Button"
'Reset mouse point icon
Buttons.Command1.MousePointer = 0
'Hide main form
Form1.Hide
End Sub

Private Sub Command2_Click()
'Show files form
Files.Show
'Hide main form
Form1.Hide
End Sub

Private Sub Command3_Click()
'show date and time form
DT.Show
'hide main form
Form1.Hide
End Sub

Private Sub Command4_Click()
'show combo form
combo.Show
'hide main form
Form1.Hide
End Sub

Private Sub Command5_Click()
'Show timers form
Timers.Show
'Hide main form
Form1.Hide
End Sub

Private Sub Command6_Click()
'Show text form
Text.Show
'Hide main form
Form1.Hide
End Sub

Private Sub Command7_Click()
'show valid form
Valid.Show
'hide main form
Form1.Hide
'Reset all the object in the form
Valid.Command3.Enabled = False
Valid.Command4.Enabled = False
Valid.Command5.Enabled = False
Valid.Command6.Enabled = False
Valid.Command7.Enabled = False
Valid.Command8.Enabled = True
Valid.Command9.Enabled = True
Valid.Command10.Enabled = True
Valid.Command11.Enabled = True
Valid.Command12.Enabled = True
Valid.Command13.Enabled = True
Valid.Command14.Enabled = True
Valid.Command15.Enabled = True
Valid.Command16.Enabled = True
Valid.Command17.Enabled = True
Valid.Check1.Enabled = False
Valid.Text1.Text = ""
Valid.Text3.Text = ""
Valid.Text1.Text = ""
End Sub

Private Sub Command9_Click()
'Show credits form
Credits.Show
End Sub
