VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form poly 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Poly"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2955
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   1200
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Poly"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command1 
         Caption         =   "Open key"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Enc"
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Dec"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "poly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Map(255, 255) As Byte
Dim KeyMap() As Byte
Dim KeyLen As Long

Private Sub Command1_Click()
CD.FileName = ""
CD.ShowOpen
If Trim(CD.FileName) = "" Then MsgBox "You have to enter the file name", vbCritical: Exit Sub

Open CD.FileName For Binary As #1
    KeyLen = LOF(1)
    ReDim KeyMap(LOF(1)) As Byte
    For i = 0 To LOF(1)
        Get #1, , KeyMap(i)
    Next
Close #1

End Sub

Private Sub Command6_Click()
Dim o As Byte
CD.FileName = ""
CD.ShowOpen
If Trim(CD.FileName) = "" Then MsgBox "You have to enter the file name", vbCritical: Exit Sub
Open CD.FileName For Binary As #1
    Open CD.FileName & ".NTS-ENC" For Binary As #2
        For i = 1 To LOF(1) Step KeyLen
            For j = 1 To KeyLen
                R = R + 1
                If R > LOF(1) Then GoTo 1
                Get #1, , o
                'MsgBox Chr(o)
                o = Map(KeyMap(j), o)
                'MsgBox Chr(o)
                Put #2, , o
            Next
        Next
1    Close #2
Close #1
End Sub

Private Sub Command7_Click()
Dim o As Byte
CD.FileName = ""
CD.ShowOpen
If Trim(CD.FileName) = "" Then MsgBox "You have to enter the file name", vbCritical: Exit Sub
Open CD.FileName For Binary As #1
    Open CD.FileName & ".NTS-ENC" For Binary As #2
        For i = 1 To LOF(1) Step KeyLen
            For j = 1 To KeyLen
                R = R + 1
                If R > LOF(1) Then GoTo 1
                Get #1, , o
                For z = 0 To 255
                    If Map(KeyMap(j), z) = o Then o = z: Exit For
                Next
                Put #2, , o
            Next
        Next
1    Close #2
Close #1
End Sub

Private Sub Form_Load()
For i = 0 To 255
    For j = 0 To 255
        If j + i > 255 Then R = ((j + i) Mod 255) Else R = j + i
        Map(i, j) = R
    Next
Next
End Sub
