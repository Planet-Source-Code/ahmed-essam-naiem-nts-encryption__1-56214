VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7710
   ClientLeft      =   5385
   ClientTop       =   1410
   ClientWidth     =   3840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   514
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2640
      Top             =   7800
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1185
      ScaleWidth      =   3825
      TabIndex        =   6
      Top             =   6000
      Width           =   3855
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   10575
         Left            =   0
         ScaleHeight     =   10575
         ScaleWidth      =   3855
         TabIndex        =   7
         Top             =   0
         Width           =   3855
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1185
            Left            =   480
            Picture         =   "Aobut.frx":0000
            ScaleHeight     =   1185
            ScaleWidth      =   2865
            TabIndex        =   8
            Top             =   0
            Width           =   2865
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   $"Aobut.frx":B202
            Height          =   4695
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   3615
         End
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10575
      Left            =   1320
      ScaleHeight     =   10575
      ScaleWidth      =   3855
      TabIndex        =   2
      Top             =   10680
      Width           =   3855
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1185
         Left            =   480
         Picture         =   "Aobut.frx":B307
         ScaleHeight     =   1185
         ScaleWidth      =   2865
         TabIndex        =   3
         Top             =   0
         Width           =   2865
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   $"Aobut.frx":16509
         Height          =   4695
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"Aobut.frx":167FD
         Height          =   4695
         Left            =   0
         TabIndex        =   4
         Top             =   5880
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÚæÏÉ"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   7320
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5925
      Left            =   0
      Picture         =   "Aobut.frx":16AF0
      ScaleHeight     =   5925
      ScaleWidth      =   3840
      TabIndex        =   0
      Top             =   0
      Width           =   3840
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "åÐÇ ÇáÈÑäÇãÌ íÎÖÚ ááÑÎÕÉ Ìí Èí Ãá"
      Height          =   255
      Left            =   1080
      MouseIcon       =   "Aobut.frx":60C32
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   7320
      Width           =   2655
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ox, oy, S, V

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Select Case Button
'Case Is <> 0
'   Top = Top + (Y - oy)
'   Left = Left + (X - ox)
'Case Else
'   ox = X
'   oy = Y
'End Select

End Sub

Private Sub Label5_Click()
GPL.Show 1, Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
Case Is <> 0
    Top = Top + (Y - oy)
    Left = Left + (X - ox)
Case Else
    ox = X
    oy = Y
End Select

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
Case Is <> 0
    Top = Top + (Y - oy)
    Left = Left + (X - ox)
Case Else
    ox = X
    oy = Y
End Select
End Sub

Private Sub Timer1_Timer()
Select Case S
Case 0
    V = V - 40
    If V < -4500 Then S = 1
Case 1
    V = V + 40
    If V > 0 Then S = 0
End Select
Picture5.Top = V
End Sub
