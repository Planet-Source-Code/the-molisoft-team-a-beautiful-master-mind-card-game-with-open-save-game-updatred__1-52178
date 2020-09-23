VERSION 5.00
Begin VB.Form FrmOpen 
   Caption         =   "Open"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox EA 
      Height          =   285
      Index           =   5
      Left            =   4800
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox EA 
      Height          =   285
      Index           =   4
      Left            =   4800
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox EA 
      Height          =   285
      Index           =   3
      Left            =   4800
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox EA 
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox EA 
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox EA 
      Height          =   285
      Index           =   0
      Left            =   4800
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtOpen1 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   4095
   End
   Begin VB.TextBox TxtOpen 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   4095
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   2280
      Pattern         =   "*.MSMI"
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "FrmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Right$(TxtOpen.Text, 5) = ".MSMI" Then
Dim Mida As String

Dim Input1 As String
FrmMain.List1.Clear
FrmMain.List2.Clear
For i = 0 To 35
FrmMain.ImgToop(i).Picture = Nothing
FrmMain.ImgSmall1(i).Picture = Nothing

Next i

Open TxtOpen.Text For Input As #1
    Input #1, Input1
    TxtOpen1.Text = Input1
Close
In0 = InStr(1, TxtOpen1.Text, "~")
E = Mid$(TxtOpen1.Text, 1, In0 - 1)
In1 = InStr(TxtOpen1.Text, "!")
E1 = Mid$(TxtOpen1.Text, In0 + 1, In1 - In0 - 1)
In2 = InStr(TxtOpen1.Text, "@")
E2 = Mid$(TxtOpen1.Text, In1 + 1, In2 - In1 - 1)
In3 = InStr(TxtOpen1.Text, "#")
E3 = Mid$(TxtOpen1.Text, In2 + 1, In3 - In2 - 1)
In4 = InStr(TxtOpen1.Text, "$")
E4 = Mid$(TxtOpen1.Text, In3 + 1, In4 - In3 - 1)
In5 = InStr(TxtOpen1.Text, "%")
E5 = Mid$(TxtOpen1.Text, In4 + 1, In5 - In4 - 1)

EA(0).Text = E
EA(1).Text = (E4) * 4
EA(2).Text = E2
EA(3).Text = E3
EA(4).Text = E4
EA(5).Text = E5
For i = 0 To E - 1
FrmMain.ImgSmall1(i).Picture = FrmMain.ImageList1.ListImages(Mid$(E1, i + 1, 1) + 9).Picture
Next i
For i = 0 To E - 1
Mida = Mid$(E5, i + 1, 1)
If Mida = 0 Or Mida = "" Then
FrmMain.ImgToop(i).Picture = Nothing
ElseIf Mida = 5 Then
FrmMain.ImgToop(i).Picture = FrmMain.ImagelistKing.ListImages(5).Picture
ElseIf Mida = 6 Then
FrmMain.ImgToop(i).Picture = FrmMain.ImagelistKing.ListImages(6).Picture
End If
Next i

FrmMain.Caption = "MasterMind" & " " & E3 & ".MSMI"
FrmMain.Text1.Text = E2
For i = 0 To 3
    FrmMain.CmdMak(i).Tag = Mid$(E2, i + 1, 1) + 9
Next i
FrmMain.CmdNumber = True
For i = 0 To E - 1
    FrmMain.List1.AddItem Mid$(E1, i + 1, 1)
    FrmMain.List2.AddItem Mid$(E5, i + 1, 1)
Next i
FrmOpen.Hide
Else

End If

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
TxtOpen.Text = Dir1.Path & "\"
If Len(TxtOpen.Text) = 4 Then
TxtOpen.Text = Left(TxtOpen.Text, 3)
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
TxtOpen.Text = Left(TxtOpen.Text, 3)

End Sub

Private Sub File1_Click()
Dim O As String
O = TxtOpen.Text

TxtOpen.Text = ""
If Len(O) = 3 Then
TxtOpen.Text = Dir1.Path & File1.FileName
Else
TxtOpen.Text = Dir1.Path & "\" & File1.FileName
End If

End Sub

Private Sub Form_Load()
Me.Icon = FrmMain.Icon

End Sub
