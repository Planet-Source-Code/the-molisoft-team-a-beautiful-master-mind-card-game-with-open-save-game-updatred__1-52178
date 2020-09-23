VERSION 5.00
Begin VB.Form FrmSave 
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   2295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   2415
      Left            =   2520
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   600
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   720
      MaxLength       =   6
      TabIndex        =   2
      Text            =   "1"
      ToolTipText     =   "Your text'lenght  must be 6"
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Name 
      AutoSize        =   -1  'True
      Caption         =   "Name :"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   510
   End
End
Attribute VB_Name = "FrmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim B As Integer
Dim A As Integer
Dim F As Integer
If FrmMain.List1.List(0) <> "" Then
If txtFilename <> "" Then
For i = 0 To 0
    TxtSave = FrmMain.List1.ListCount & "~"
        For A = 0 To FrmMain.List1.ListCount - 1
            TxtSave = TxtSave & FrmMain.List1.List(A)
        Next A

    TxtSave = TxtSave & "!" & FrmMain.Text1.Text & "@" & txtFilename.Text & "#" & FrmMain.Txtnumber.Text & "$"

Next i
Text1.Text = TxtSave
        For F = 0 To FrmMain.List1.ListCount - 1
        TxtSave = TxtSave & FrmMain.List2.List(F)
        Next F
TxtSave = TxtSave & "%"

Open Dir1.Path & "\" & txtFilename & ".MSMI" For Output As #1
Print #1, TxtSave
Close
Text2.Text = TxtSave
FrmSave.Hide
Else
MsgBox "Please Select the Name of your Game"
End If
Else
MsgBox "You must play game Then You can save your Game", vbOKOnly, "Error Save"

End If

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub

Private Sub EA_Change(Index As Integer)

End Sub

Private Sub Form_Load()
Me.Icon = FrmMain.Icon
End Sub

Private Sub txtFilename_Change()
If Len(txtFilename.Text) = 6 Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If

End Sub
