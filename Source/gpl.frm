VERSION 5.00
Begin VB.Form GPL 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GNU GPL"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5610
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Height          =   4815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "gpl.frx":0000
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "GPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
