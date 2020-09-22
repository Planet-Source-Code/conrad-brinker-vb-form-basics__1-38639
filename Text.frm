VERSION 5.00
Begin VB.Form Text 
   Caption         =   "Text Boxes"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form2"
   ScaleHeight     =   2550
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Clear"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Click Me!"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Testing"
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4560
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "Text"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Hide text form on close
Text.Hide
'Show main form
Form1.Show
End Sub

Private Sub Command2_Click()
'Just copy text from text2 to text1
Text2.Text = Text1.Text
End Sub

Private Sub Command3_Click()
'Clear text2
Text2.Text = ""
End Sub

Private Sub Command4_Click()
'When the button is pressed write this
Text3.Text = "This is a test !Testing!"
End Sub

Private Sub Command5_Click()
'Clear text3
Text3.Text = ""
End Sub

Private Sub Form_Load()
'On form load, make text3's font size 14 so it fits the box
Text3.FontSize = 14
End Sub
