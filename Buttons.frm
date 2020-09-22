VERSION 5.00
Begin VB.Form Buttons 
   Caption         =   "Buttons"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form2"
   ScaleHeight     =   3630
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Change Pointer"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      MaxLength       =   42
      TabIndex        =   6
      Text            =   "Click Me!"
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Caption"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Test Area"
      Height          =   1695
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
      Begin VB.CommandButton Command1 
         Caption         =   "Test Button"
         Height          =   615
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show Button"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   5295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide Button"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line4 
      X1              =   5520
      X2              =   5520
      Y1              =   1200
      Y2              =   3120
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   240
      Y1              =   1200
      Y2              =   3120
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   5520
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5520
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "Buttons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
'We're gunna hide the test button
Command1.Visible = False
End Sub

Private Sub Command3_Click()
'Pressing the close button hides the buttons form and shows the main form
Buttons.Hide
Form1.Show
End Sub

Private Sub Command4_Click()
'This will show the test button
Command1.Visible = True
End Sub

Private Sub Command5_Click()
'This changes the text on the button to the text in text1
Command1.Caption = Text1.Text
End Sub

Private Sub Command6_Click()
'We change the mouseover icon to an hourglass, a different # changes the icon
Command1.MousePointer = 11
End Sub
