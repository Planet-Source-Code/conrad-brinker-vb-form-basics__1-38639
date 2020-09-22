VERSION 5.00
Begin VB.Form Credits 
   Caption         =   "Credits"
   ClientHeight    =   4350
   ClientLeft      =   4545
   ClientTop       =   4575
   ClientWidth     =   3375
   LinkTopic       =   "Form2"
   ScaleHeight     =   4350
   ScaleWidth      =   3375
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "http://www.chbrules.freeservers.com"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "AIM: Conrad1986"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "E-mail: Conrad_Hans_Brinker@Hotmail.com"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   240
      Picture         =   "Credits.frx":0000
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "Credits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Hide those credits
Credits.Hide
End Sub
