VERSION 5.00
Begin VB.Form combo 
   Caption         =   "Combo Boxes"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "Form2"
   ScaleHeight     =   2415
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "DELETE"
      Height          =   1215
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CLEAR LIST"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.ListBox list1 
      Height          =   1230
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "Selected Item #:"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Total # of Items:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4800
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "combo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Hide combo form, show main form
combo.Hide
Form1.Show
End Sub

Private Sub Command2_Click()
'Is there anything in the textbox to add to the list?
If Text1.Text = "" Then MsgBox "Type Something" Else list1.AddItem (Text1.Text)
'Erase after adding to list
Text1.Text = ""
'make the box refresh to say how many items on the list
Text2.Text = list1.ListCount
End Sub

Private Sub Command3_Click()
'Clear the list
list1.Clear
'Refresh the count box
Text2.Text = list1.ListCount
'Refresh the selected # box
Text3.Text = list1.SelCount
'Nothing is selected and it'll say =1, so change it to 0 to look good =)
If Text3.Text = "-1" Then Text3.Text = "0"
End Sub

Private Sub Command4_Click()
'Removes item by seeing which number is selected from textbox
list1.RemoveItem (Text3.Text)
'refreshing
Text2.Text = list1.ListCount
'refreshing
Text3.Text = list1.ListIndex
'Make it look nice again
If Text3.Text = "-1" Then Text3.Text = "0"
End Sub

Private Sub Form_Load()
'clear list on startup
list1.Clear
End Sub

Private Sub list1_Click()
'refresh
Text2.Text = list1.ListCount
'refresh
Text3.Text = list1.ListIndex
'Prettiness again =)
If Text3.Text = "-1" Then Text3.Text = "0"
End Sub
