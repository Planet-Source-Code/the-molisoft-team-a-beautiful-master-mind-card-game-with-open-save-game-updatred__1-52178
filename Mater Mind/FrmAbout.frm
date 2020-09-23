VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   LinkTopic       =   "Form2"
   ScaleHeight     =   4485
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   0
      Width           =   3975
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MasterMind"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   8
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "       Written by"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   -120
         TabIndex        =   11
         Top             =   0
         Width           =   3855
      End
      Begin VB.Label lbl 
         Caption         =   "Mohammad AlianNejadi"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label lbl 
         Caption         =   "Nima Froughi"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label lbl 
         Caption         =   "This product is licensed to:"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label lbl 
         Caption         =   "MOLiSoft Co."
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lbl 
         Caption         =   "Name"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lbl 
         Caption         =   "Company"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lbl 
         Caption         =   "Copyright (C) 2003"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1935
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   2535
         Left            =   0
         Picture         =   "FrmAbout.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2535
         Left            =   0
         Picture         =   "FrmAbout.frx":CBFB
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3480
      Top             =   0
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   2760
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer TmrRound 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   3840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "= A number is Right  and its locality is Right"
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   14
      Top             =   3480
      Width           =   3030
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "= A number is Right  but its locality is Wrong"
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   13
      Top             =   3120
      Width           =   3090
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   1
      Left            =   120
      Top             =   3480
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   3120
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4080
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail: m_alian2003@hotmail.com"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   600
      MouseIcon       =   "FrmAbout.frx":197F6
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblWSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "W-site: www.geocities.com/m_alian14/"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   480
      MouseIcon       =   "FrmAbout.frx":19948
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2880
      Width           =   2790
   End
   Begin VB.Image IMGOK 
      Height          =   555
      Left            =   1320
      Picture         =   "FrmAbout.frx":19A9A
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1155
   End
   Begin VB.Image Image2 
      Height          =   6750
      Left            =   0
      Picture         =   "FrmAbout.frx":1A699
      Top             =   2400
      Width           =   6345
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal Hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

'Private Const NAME = 6
'Private Const COMPANY = 7
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const HKEY_CURRENT_USER = &H80000001

Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_EXPAND_SZ = 2
Private Const REG_SZ = 1

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const HKEY_LOCAL_MACHINE = &H80000002
Dim X As Long
Dim Y As Long




Private Sub Form_Load()
Me.Move FrmMain.Left + 600, (FrmMain.Top + FrmMain.Width) / 2 + 500
    Dim i As Integer
    Dim Hrgn As Long
  Dim hKey As Long, num As Long, strName As String
  Dim strData As String, Retval As Long, RetvalData As Long
  X = 0: Y = 0
  Timer1.Enabled = False
  TmrRound.Enabled = True
  Const Buffer As Long = 255
  Me.AutoRedraw = True
  num = 0
  strName = Space(Buffer)
  strData = Space(Buffer)
  Retval = Buffer
  RetvalData = Buffer
  If RegOpenKey(HKEY_LOCAL_MACHINE, "Software\MOLiSoft\Snakes and Ladders\1.0", hKey) = 0 Then
    While RegEnumValue(hKey, num, strName, Retval, 0, ByVal 0&, ByVal strData, RetvalData) <> ERROR_NO_MORE_ITEMS
       If RetvalData > 0 Then
          List1.AddItem Left$(strData, RetvalData - 1)
       End If
       num = num + 1
       strName = Space(Buffer)
       strData = Space(Buffer)
       Retval = Buffer
       RetvalData = Buffer
    Wend
    RegCloseKey hKey
  Else
    List1.AddItem "Error"
  End If
  'Me.Caption = App.Title & " - About"
  lbl(4).Caption = List1.List(0)
  lbl(5).Caption = List1.List(1)
    Image1.Move 0, 0, Picture1.Width, Picture1.Height
    lbl(7).Top = Picture1.Height
    lbl(0).Top = Picture1.Height + 500
    For i = 1 To lbl.Count - 2
        lbl(i).Top = lbl(i - 1).Top + 500
        lbl(i).Alignment = 2
        lbl(i).ForeColor = vbWhite
        lbl(i).BackStyle = 0
        lbl(i).Width = Picture1.Width
    Next i

End Sub
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
IMGOK.Picture = FrmMain.ImagelistKing.ListImages(9).Picture

End Sub

Private Sub ImgOk_Click()
Unload Me
End Sub

Private Sub ImgOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
IMGOK.Picture = FrmMain.ImagelistKing.ListImages(10).Picture

End Sub

Private Sub lblEmail_Click()
    ShellExecute Me.hWnd, vbNullString, "mailto: m_alian2003@hotmail.com", vbNullString, "c:\", 1
End Sub

Private Sub lblWSite_Click()
    Shell "Explorer http://www.geocities.com/m_alian14/"
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    For i = 0 To lbl.Count - 1
        lbl(i).Top = lbl(i).Top - 50
    Next i
    
End Sub

Private Sub TmrRound_Timer()
    X = X + 10
    Y = Y + 10
    DoEvents
    
        Hrgn = CreateRoundRectRgn(0, 5, Me.ScaleWidth / 15, Me.ScaleHeight / 15, X, Y)
        SetWindowRgn Me.hWnd, Hrgn, True

        
    If X >= 50 Or Y >= 50 Then
        TmrRound.Enabled = False
        Timer1.Enabled = True
    End If
End Sub

