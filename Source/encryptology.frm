VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form NTSENC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NTS - Encryptor"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "File Encryption"
      Height          =   2295
      Left            =   240
      TabIndex        =   24
      Top             =   2160
      Width           =   5895
      Begin VB.CommandButton Command14 
         Appearance      =   0  'Flat
         Caption         =   "Reserved Dec"
         Height          =   495
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command13 
         Appearance      =   0  'Flat
         Caption         =   "Dec"
         Height          =   495
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Caption         =   "Enc 1 - 8 text"
         Height          =   975
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         Caption         =   "Enc"
         Height          =   495
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Caption         =   "Reserved Enc"
         Height          =   495
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Caption         =   "&Save To"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   720
         Width           =   4335
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Caption         =   "&Open"
         Height          =   375
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   360
         Width           =   4335
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3360
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Method 2"
      Height          =   2175
      Left            =   4080
      TabIndex        =   20
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton Command12 
         Caption         =   "Dec"
         Height          =   495
         Left            =   1080
         TabIndex        =   34
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Reserved Dec"
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Enc"
         Height          =   495
         Left            =   1080
         TabIndex        =   26
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Reserved Enc"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Enter number between 00000000 - 88888888"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Method 1"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton Command9 
         Caption         =   "="
         Height          =   495
         Left            =   2760
         TabIndex        =   39
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         Caption         =   "=AK"
         Height          =   495
         Left            =   2160
         TabIndex        =   38
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Enc 1 - 8 text"
         Height          =   495
         Left            =   2160
         TabIndex        =   37
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   1
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   600
         MaxLength       =   1
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   3480
         MaxLength       =   1
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Text            =   "Allah Akbar "
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   600
         TabIndex        =   1
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   2
         Left            =   1200
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   4
         Left            =   2160
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   5
         Left            =   2640
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   6
         Left            =   3120
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   7
         Left            =   3600
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
   End
End
Attribute VB_Name = "NTSENC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text3.Text = ""
For j = 7 To 0 Step -1
    Label1(j).Caption = "1"
Next
For j = 0 To 7
    If Trim(Text1(i).Text) = "" Then MsgBox "you have to enter all fileds": Exit Sub
Next
For j = 0 To 7
    If Trim(Text1(i).Text) = "" Then MsgBox "you have to enter all fileds": Exit Sub
Next

For j = 7 To 0 Step -1
    M = Asc(Text1(j).Text)
    If (M And 2 ^ j) = 0 Then Label1(j).Caption = 0 Else XorV = XorV + (2 ^ j)
Next
Label2.Caption = XorV
For i = 1 To Len(Text2.Text)
    C = Mid(Text2.Text, i, 1)
    Text3.Text = Text3.Text & Chr(Asc(C) Xor XorV)
Next
End Sub

Private Sub Command10_Click()
Text2.Text = "Allah Akbar"
End Sub

Private Sub Command11_Click()
Text3.Text = ""
For i = 1 To Len(Text2.Text)
    M = Asc(Mid(Text2.Text, i, 1))
    For j = 8 To 1 Step -1
        N = Val(Mid(Text4.Text, j, 1))
        If (M And 2 ^ N) = 0 Then CV = CV + (2 ^ (j - 1))
    Next
    Text3.Text = Text3.Text & Chr(CV): CV = 0
Next
End Sub

Private Sub Command12_Click()
If Len(Text4.Text) < 8 Then MsgBox "Must be 8 digit ": Exit Sub
Text3.Text = ""
For i = 1 To Len(Text2.Text)
    M = Asc(Mid(Text2.Text, i, 1))
    For j = 1 To 8
        N = Val(Mid(Text4.Text, j, 1))
        If (M And 2 ^ N) <> 0 Then CV = CV + (2 ^ (j - 1))
    Next
    Text3.Text = Text3.Text & Chr(CV): CV = 0
Next

End Sub

Private Sub Command13_Click()
If Trim(Text5.Text) = "" Or Trim(Text6.Text) = "" Then MsgBox "PLEASE SELECT THE FILE": Exit Sub
Dim M As Byte
If Len(Text4.Text) < 8 Then MsgBox "Must be 8 digit ": Exit Sub

Open Text5.Text For Binary As #1
    For i = 1 To LOF(1)
         Get #1, , M
         For j = 1 To 8
             N = Val(Mid(Text4.Text, j, 1))
             If (M And 2 ^ N) <> 0 Then CV = CV + (2 ^ (j - 1))
         Next
         All = All & Chr(CV): CV = 0
     Next
Close #1

Open Text6.Text For Append As #1
Print #1, All
Close #1

End Sub

Private Sub Command14_Click()
If Trim(Text5.Text) = "" Or Trim(Text6.Text) = "" Then MsgBox "PLEASE SELECT THE FILE": Exit Sub
Dim M As Byte
If Len(Text4.Text) < 8 Then MsgBox "Must be 8 digit ": Exit Sub

Open Text5.Text For Binary As #1
    For i = 1 To LOF(1)
         Get #1, , M
         For j = 1 To 8
             N = Val(Mid(Text4.Text, j, 1))
             If (M And 2 ^ N) = 0 Then CV = CV + (2 ^ (j - 1))
         Next
         All = All & Chr(CV): CV = 0
     Next
Close #1

Open Text6.Text For Append As #1
Print #1, All
Close #1

End Sub

Private Sub Command2_Click()
'On Error Resume Next
'MsgBox "5% the program can't retrive your data"
If Len(Text4.Text) < 8 Then MsgBox "Must be 8 digit ": Exit Sub
Text3.Text = ""
For i = 1 To Len(Text2.Text)
    M = Asc(Mid(Text2.Text, i, 1))
    For j = 1 To 8
        N = Val(Mid(Text4.Text, j, 1))
        If (M And 2 ^ (j - 1)) <> 0 Then CV = CV + (2 ^ N)
    Next
    Text3.Text = Text3.Text & Chr(CV): CV = 0
Next
End Sub

Private Sub Command3_Click()
Text3.Text = ""
For i = 1 To Len(Text2.Text)
    M = Asc(Mid(Text2.Text, i, 1))
    For j = 8 To 1 Step -1
        N = Val(Mid(Text4.Text, j, 1))
        If (M And 2 ^ (j - 1)) = 0 Then CV = CV + (2 ^ N)
    Next
    Text3.Text = Text3.Text & Chr(CV): CV = 0
Next
End Sub

Private Sub Command4_Click()
CD.FileName = ""
CD.ShowOpen
Text5.Text = CD.FileName
Text6.Text = CD.FileName & "_NTS_ENC.txt"
End Sub

Private Sub Command6_Click()
If Trim(Text5.Text) = "" Or Trim(Text6.Text) = "" Then MsgBox "PLEASE SELECT THE FILE": Exit Sub
Dim M As Byte
If Len(Text4.Text) < 8 Then MsgBox "Must be 8 digit ": Exit Sub

Open Text5.Text For Binary As #1
    For i = 1 To LOF(1)
         Get #1, , M
         For j = 1 To 8
             N = Val(Mid(Text4.Text, j, 1))
             If (M And 2 ^ (j - 1)) = 0 Then CV = CV + (2 ^ N)
         Next
         All = All & Chr(CV): CV = 0
     Next
Close #1

Open Text6.Text For Append As #1
Print #1, All
Close #1

End Sub

Private Sub Command7_Click()
If Trim(Text5.Text) = "" Or Trim(Text6.Text) = "" Then MsgBox "PLEASE SELECT THE FILE": Exit Sub
Dim M As Byte
If Len(Text4.Text) < 8 Then MsgBox "Must be 8 digit ": Exit Sub

Open Text5.Text For Binary As #1
    For i = 1 To LOF(1)
         Get #1, , M
         For j = 1 To 8
             N = Val(Mid(Text4.Text, j, 1))
             If (M And 2 ^ (j - 1)) <> 0 Then CV = CV + (2 ^ N)
         Next
         All = All & Chr(CV): CV = 0
     Next
Close #1

Open Text6.Text For Append As #1
Print #1, All
Close #1

End Sub

Private Sub Command8_Click()
If Trim(Text5.Text) = "" Or Trim(Text6.Text) = "" Then MsgBox "PLEASE SELECT THE FILE": Exit Sub
Dim C As Byte
For j = 7 To 0 Step -1
    Label1(j).Caption = "1"
Next

For j = 0 To 7
    If Trim(Text1(i).Text) = "" Then MsgBox "you have to enter all fileds": Exit Sub
Next


For j = 7 To 0 Step -1
    M = Asc(Text1(j).Text)
    If (M And 2 ^ j) = 0 Then Label1(j).Caption = 0 Else XorV = XorV + (2 ^ j)
Next
Label2.Caption = XorV

Open Text5.Text For Binary As #1
    For i = 1 To LOF(1)
        Get #1, , C
        All = All & Chr(C Xor XorV)
    Next
Close #1
Open Text6.Text For Append As #1
Print #1, All
Close #1
End Sub

Private Sub Command9_Click()
Text2.Text = Text3.Text
End Sub

Private Sub Form_Load()
For i = 0 To 7
    Text1(i).Text = ""
Next
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Len(Text1(Index).Text) = 1 Then Text1(Index).Text = Left(Text1(Index).Text, 1): Text1(1 + Index).SetFocus:
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error Resume Next
If InStr(1, Text4.Text, Chr(KeyAscii)) <> 0 Then KeyAscii = 0
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 55 Then KeyAscii = 0: Exit Sub
End Sub
