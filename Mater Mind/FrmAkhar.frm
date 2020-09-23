VERSION 5.00
Begin VB.Form FrmAkhar 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   LinkTopic       =   "Form2"
   Picture         =   "FrmAkhar.frx":0000
   ScaleHeight     =   1635
   ScaleWidth      =   2250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image ImgOk 
      Height          =   435
      Left            =   600
      Picture         =   "FrmAkhar.frx":4D1B
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Image ImgAkhar 
      Height          =   1155
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   810
   End
   Begin VB.Image ImgAkhar 
      Height          =   1155
      Index           =   1
      Left            =   480
      Top             =   0
      Width           =   810
   End
   Begin VB.Image ImgAkhar 
      Height          =   1155
      Index           =   2
      Left            =   960
      Top             =   0
      Width           =   810
   End
   Begin VB.Image ImgAkhar 
      Height          =   1155
      Index           =   3
      Left            =   1440
      Top             =   0
      Width           =   810
   End
End
Attribute VB_Name = "FrmAkhar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub Form_Load()
Me.Move FrmMain.Left + 1500, (FrmMain.Top + FrmMain.Width) / 2 + 500

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgOk.Picture = FrmMain.ImagelistKing.ListImages(9).Picture

End Sub

Private Sub ImgOk_Click()
Unload FrmAkhar
FrmMain.CmdStart = True

End Sub

Private Sub ImgOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgOk.Picture = FrmMain.ImagelistKing.ListImages(10).Picture

End Sub
