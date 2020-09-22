VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ahmed Essam - Encryption Work"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   4455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "About"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Poly"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Mono"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ceaser"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NTS - Encryption"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
NTSENC.Show 1, Me
End Sub

Private Sub Command2_Click()
Ceaser.Show 1, Me
End Sub

Private Sub Command3_Click()
mono.Show 1, Me
End Sub

Private Sub Command4_Click()
poly.Show 1, Me
End Sub

Private Sub Command5_Click()
About.Show 1, Me
End Sub

Private Sub Command6_Click()
End
End Sub
