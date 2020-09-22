VERSION 5.00
Begin VB.Form Files 
   Caption         =   "Files"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form2"
   ScaleHeight     =   4455
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   3975
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   4815
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   4080
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "File Name:"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'On close hide files form and show main form
Files.Hide
Form1.Show
End Sub

Private Sub Dir1_Change()
'The files window shows the files selected in Dir1's window
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
'If there isn't a disk in the drive or it can't connect to it, this will stop it from crashing
If Drive1.CausesValidation Then Drive1.Refresh
'this lets dir1 show folders from that drive selected
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
'Show the filename in text1
Text1.Text = File1.FileName
End Sub
