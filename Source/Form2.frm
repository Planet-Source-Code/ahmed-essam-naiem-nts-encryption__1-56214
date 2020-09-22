VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form mono 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "mono"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5475
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   2520
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Key"
      Height          =   1935
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   5415
      Begin VB.Label Label1 
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5175
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mono"
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton Command4 
         Caption         =   "Encrypt / Decrypt"
         Height          =   555
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   3495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Generat Key"
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Open Key"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Save Key"
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "mono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Map(255) As Byte

Private Sub Command10_Click()
CD.FileName = ""
CD.ShowSave
If Trim(CD.FileName) = "" Then MsgBox "You have to enter the file name", vbCritical: Exit Sub
If Trim(Label1.Caption) = "" Then MsgBox "You have to generate a Key": Exit Sub
Open CD.FileName For Binary As #1
    For i = 0 To 255
        Put #1, , Map(i)
    Next
Close #1
End Sub

Private Sub Command4_Click()
On Error GoTo 1
Dim o As Byte
CD.FileName = ""
CD.ShowOpen
If Trim(CD.FileName) = "" Then MsgBox "You have to enter the file name", vbCritical: Exit Sub

Open CD.FileName For Binary As #1
    Open CD.FileName & ".NTS-ENC" For Binary As #2
        For i = 1 To LOF(1) Step 256
            For j = 0 To 255
                Get #1, , o
                o = o Xor Map(j)
                Put #2, , o
            Next
        Next
1    Close #2
Close #1
End Sub

Private Sub Command8_Click()
Label1.Caption = ""
For i = 0 To 255
    Map(i) = i
Next
 
For i = 0 To 255
    t = Int(Rnd() * 255)
    tmp = Map(i)
    Map(i) = Map(t)
    Map(t) = tmp
Next

For i = 0 To 255
    Label1.Caption = Label1.Caption & Chr(Map(i))
Next

End Sub

Private Sub Command9_Click()
CD.FileName = ""
CD.ShowOpen
If Trim(CD.FileName) = "" Then MsgBox "You have to enter the file name", vbCritical: Exit Sub

Open CD.FileName For Binary As #1
    For i = 0 To 255
        Get #1, , Map(i)
    Next
Close #1
Label1.Caption = ""
For i = 0 To 255
    Label1.Caption = Label1.Caption & Chr(Map(i))
Next
End Sub

Private Sub Form_Load()
For i = 0 To 255
    Map(i) = i
Next
End Sub
