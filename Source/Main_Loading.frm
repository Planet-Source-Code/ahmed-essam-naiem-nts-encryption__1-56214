VERSION 5.00
Begin VB.Form LLoad 
   BorderStyle     =   0  'None
   Caption         =   "LOADIND . . ."
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   Picture         =   "Main_Loading.frx":0000
   ScaleHeight     =   7125
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1710
      ScaleHeight     =   240
      ScaleWidth      =   6630
      TabIndex        =   1
      Top             =   6435
      Width           =   6630
      Begin VB.Shape Shape1 
         DrawMode        =   6  'Mask Pen Not
         FillStyle       =   0  'Solid
         Height          =   330
         Left            =   0
         Top             =   -45
         Width           =   60
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   600
      Top             =   3480
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   960
      Top             =   3000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   5760
      Width           =   1560
   End
End
Attribute VB_Name = "LLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, B
Private Sub Timer1_Timer()
'Label1.Caption = Right(Label1.Caption, 1) & Left(Label1.Caption, Len(Label1.Caption) - 1)
Shape1.Width = Shape1.Width + 500
End Sub

Private Sub Timer2_Timer()
a = a + 0.5
If a >= 5 Then a = 1: B = B + 1
R = String(a, ".")
Label2.Caption = "Loading " & R

If B >= 2 Then
MainFrm.Show
Unload Me
End If
End Sub
