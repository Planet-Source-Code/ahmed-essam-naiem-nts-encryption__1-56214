VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Ceaser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ceaser"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2520
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "100"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ceaser Xor"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ceaser -"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ceaser +"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1080
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Key"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   135
      Width           =   825
   End
End
Attribute VB_Name = "Ceaser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Val(Text3.Text) > 255 Or Val(Text3.Text) < 1 Then MsgBox "You have to make it less than or eqal 255 and greater than 0": Exit Sub

Dim o As Byte
CD.FileName = ""
CD.ShowOpen
If Trim(CD.FileName) = "" Then MsgBox "You have to enter the file name", vbCritical: Exit Sub
Open CD.FileName For Binary As #1
    Open CD.FileName & ".NTS-ENC" For Binary As #2
        For i = 1 To LOF(1)
            Get #1, , o
            If o + Val(Text3.Text) > 255 Then o = 255 - o Else o = o + Val(Text3.Text)
            Put #2, , o
        Next
1    Close #2
Close #1

End Sub

Private Sub Command2_Click()
If Val(Text3.Text) > 255 Or Val(Text3.Text) < 1 Then MsgBox "You have to make it less than or eqal 255 and greater than 0": Exit Sub

Dim o As Byte
CD.FileName = ""
CD.ShowOpen
If Trim(CD.FileName) = "" Then MsgBox "You have to enter the file name", vbCritical: Exit Sub
Open CD.FileName For Binary As #1
    Open CD.FileName & ".NTS-ENC" For Binary As #2
        For i = 1 To LOF(1)
            Get #1, , o
            If o - Val(Text3.Text) < 1 Then o = 255 - o Else o = o - Val(Text3.Text)
            Put #2, , o
        Next
1    Close #2
Close #1

End Sub

Private Sub Command3_Click()
If Val(Text3.Text) > 255 Or Val(Text3.Text) < 1 Then MsgBox "You have to make it less than or eqal 255 and greater than 0": Exit Sub
Dim o As Byte
CD.FileName = ""
CD.ShowOpen
If Trim(CD.FileName) = "" Then MsgBox "You have to enter the file name", vbCritical: Exit Sub
Open CD.FileName For Binary As #1
    Open CD.FileName & ".NTS-ENC" For Binary As #2
        For i = 1 To LOF(1)
            Get #1, , o
            o = o Xor Val(Text3.Text)
            Put #2, , o
        Next
1    Close #2
Close #1

End Sub
