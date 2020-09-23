VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MasterMind"
   ClientHeight    =   7425
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5010
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":164A
   ScaleHeight     =   7425
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txt1 
      Height          =   285
      Left            =   1080
      TabIndex        =   41
      Text            =   "Text4"
      Top             =   7080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton CmdNumber 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6840
      TabIndex        =   40
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Txtnumber 
      Height          =   495
      Left            =   6840
      TabIndex        =   39
      Top             =   840
      Width           =   495
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   5880
      TabIndex        =   38
      Top             =   120
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   5400
      TabIndex        =   37
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      Height          =   255
      Left            =   1800
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer TmrCharkh 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   840
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   580
      Left            =   220
      ScaleHeight     =   525
      ScaleWidth      =   540
      TabIndex        =   35
      Top             =   600
      Width           =   600
      Begin VB.Image Image2 
         Height          =   525
         Left            =   0
         Picture         =   "Form1.frx":8995
         Top             =   0
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList ImagelistKing 
      Left            =   7560
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   54
      ImageHeight     =   77
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":989D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CA43
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":FBE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12D8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15F35
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16437
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16939
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17B5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18D7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1998C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A6F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D899
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":20A3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":23BE5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5280
      TabIndex        =   34
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Timer TmrMove 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6720
      Top             =   2760
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6120
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton CmdMak 
      Height          =   1335
      Index           =   3
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton CmdMak 
      Height          =   1335
      Index           =   2
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton CmdMak 
      Height          =   1335
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton CmdMak 
      Height          =   1335
      Index           =   0
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox CardRed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Index           =   9
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   810
      TabIndex        =   13
      Top             =   4920
      Width           =   810
   End
   Begin VB.PictureBox CardRed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Index           =   8
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   810
      TabIndex        =   12
      Top             =   4560
      Width           =   810
   End
   Begin VB.PictureBox CardRed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Index           =   7
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   810
      TabIndex        =   11
      Top             =   4200
      Width           =   810
   End
   Begin VB.PictureBox CardRed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Index           =   6
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   810
      TabIndex        =   10
      Top             =   3840
      Width           =   810
   End
   Begin VB.PictureBox CardRed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Index           =   5
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   810
      TabIndex        =   9
      Top             =   3480
      Width           =   810
   End
   Begin VB.PictureBox CardRed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Index           =   4
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   810
      TabIndex        =   8
      Top             =   3120
      Width           =   810
   End
   Begin VB.PictureBox CardRed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Index           =   3
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   810
      TabIndex        =   7
      Top             =   2760
      Width           =   810
   End
   Begin VB.PictureBox CardRed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Index           =   2
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   810
      TabIndex        =   6
      Top             =   2400
      Width           =   810
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8040
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   54
      ImageHeight     =   77
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   36
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":26D8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":29F31
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2D0D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3027D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":33423
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":365C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3976F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3C915
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3FA17
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":42BBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":45D63
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":48F09
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4C0AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4F255
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":523FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":555A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":58747
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5B849
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5E9EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":61B95
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":64D3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":67EE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6B087
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6E22D
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":713D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":74579
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7771F
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7A8C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7DA6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":80C11
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":83DB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":86F5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8A103
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8D2A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9044F
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":935F5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox CardRed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Index           =   1
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   810
      TabIndex        =   4
      Top             =   2040
      Width           =   810
   End
   Begin VB.CommandButton CmdAsl 
      Enabled         =   0   'False
      Height          =   1335
      Index           =   3
      Left            =   6120
      TabIndex        =   3
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton CmdAsl 
      Enabled         =   0   'False
      Height          =   1335
      Index           =   2
      Left            =   5160
      TabIndex        =   2
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton CmdAsl 
      Enabled         =   0   'False
      Height          =   1335
      Index           =   1
      Left            =   6000
      TabIndex        =   1
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton CmdAsl 
      Enabled         =   0   'False
      Height          =   1335
      Index           =   0
      Left            =   5040
      TabIndex        =   0
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      Picture         =   "Form1.frx":9679B
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImageCharkh 
      Left            =   1320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   36
      ImageHeight     =   35
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9780D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":98725
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9963D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9A555
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9B46D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9C385
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9D29D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9E1B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9F0CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9FFE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A0EFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A1E15
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A2D2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A3C45
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A4B5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A5A75
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A698D
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A78A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A87BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A96D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AA5ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AB505
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image ImgLock 
      Height          =   480
      Left            =   360
      Picture         =   "Form1.frx":AC41D
      ToolTipText     =   "You saw the Password"
      Top             =   30
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   4800
      Top             =   7200
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   0
      Top             =   7200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image ImgNesh 
      Height          =   135
      Index           =   8
      Left            =   1080
      Picture         =   "Form1.frx":ACCE7
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   135
   End
   Begin VB.Image ImgNesh 
      Height          =   135
      Index           =   7
      Left            =   1080
      Picture         =   "Form1.frx":ADBED
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   135
   End
   Begin VB.Image ImgNesh 
      Height          =   135
      Index           =   6
      Left            =   1080
      Picture         =   "Form1.frx":AEAF3
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   135
   End
   Begin VB.Image ImgNesh 
      Height          =   135
      Index           =   5
      Left            =   1080
      Picture         =   "Form1.frx":AF9F9
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   135
   End
   Begin VB.Image ImgNesh 
      Height          =   135
      Index           =   4
      Left            =   1080
      Picture         =   "Form1.frx":B08FF
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   135
   End
   Begin VB.Image ImgNesh 
      Height          =   135
      Index           =   3
      Left            =   1080
      Picture         =   "Form1.frx":B1805
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   135
   End
   Begin VB.Image ImgNesh 
      Height          =   135
      Index           =   2
      Left            =   1080
      Picture         =   "Form1.frx":B270B
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   135
   End
   Begin VB.Image ImgNesh 
      Height          =   135
      Index           =   1
      Left            =   1080
      Picture         =   "Form1.frx":B3611
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   135
   End
   Begin VB.Image ImgNesh 
      Height          =   135
      Index           =   0
      Left            =   1080
      Picture         =   "Form1.frx":B4517
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   135
   End
   Begin VB.Image ImgAbout 
      Height          =   375
      Left            =   240
      Top             =   6480
      Width           =   375
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   35
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   34
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   33
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   32
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   31
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   30
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   29
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   28
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   255
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   35
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   34
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   33
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   32
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   31
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   30
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   29
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   28
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   27
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   26
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   25
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   24
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   27
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   26
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   25
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   24
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   23
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   22
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   21
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   20
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   19
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   18
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   17
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   16
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   15
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   14
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   13
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   12
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   11
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   10
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   9
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   8
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   7
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   6
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   5
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   4
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   3
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   2
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   1
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   255
   End
   Begin VB.Image ImgToop 
      Height          =   255
      Index           =   0
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   255
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   23
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   22
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   21
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   20
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   19
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   18
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   17
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   16
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   15
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   14
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   13
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   12
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   11
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   10
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   9
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   8
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   7
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   6
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   5
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   4
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   3
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   2
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   1
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   375
   End
   Begin VB.Image ImgSmall1 
      Height          =   495
      Index           =   0
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label LblTar 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   12
      Left            =   4680
      TabIndex        =   33
      Top             =   4080
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblTar 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   11
      Left            =   4680
      TabIndex        =   32
      Top             =   4320
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblTar 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   10
      Left            =   4680
      TabIndex        =   31
      Top             =   4560
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblTar 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   9
      Left            =   4680
      TabIndex        =   30
      Top             =   4800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblTar 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   8
      Left            =   4680
      TabIndex        =   29
      Top             =   5040
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblTar 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   7
      Left            =   4680
      TabIndex        =   28
      Top             =   5280
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblTar 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   6
      Left            =   4680
      TabIndex        =   27
      Top             =   5520
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblTar 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   5
      Left            =   4680
      TabIndex        =   26
      Top             =   5760
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblTar 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   4
      Left            =   4680
      TabIndex        =   25
      Top             =   6000
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblTar 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   3
      Left            =   4680
      TabIndex        =   24
      Top             =   6240
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblTar 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   2
      Left            =   4680
      TabIndex        =   23
      Top             =   6480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblTar 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   1
      Left            =   4680
      TabIndex        =   22
      Top             =   6720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblTar 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   0
      Left            =   4680
      TabIndex        =   21
      Top             =   6960
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label LblIndexCard 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2520
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   7455
      Left            =   0
      Picture         =   "Form1.frx":B541D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5055
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu MnuNew 
         Caption         =   "NewGame"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuPass 
         Caption         =   "ShowPassword"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "About MasterMind"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Dim i As Long
Dim I1 As Long
Dim Bolmove As Boolean
Dim BolTrue As Boolean
Dim Find1 As Long
Dim Find2 As Long
Dim A As Long
Dim B As Long
Dim C As Long
Dim OldInd As Integer
Dim G As Integer
Dim TextOld As String
Dim Number As Integer

Private Sub CardRed_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If BolTrue = False Then
        If Bolmove = False Then
            LblIndexCard.Caption = CardRed(Index).Index
            CardRed(LblIndexCard.Caption).Left = CardRed(LblIndexCard.Caption).Left + 100
            CardRed(LblIndexCard).Picture = ImageList1.ListImages(LblIndexCard.Caption + 18).Picture
            Bolmove = True
            sndPlaySound App.Path & "\sound\blip2.wav", &H1 Or &H2 Or &H2000
        ElseIf Bolmove = True Then
        
        End If
        BolTrue = True
    End If

End Sub

Private Sub CmdAsl_Click(Index As Integer)
Randomize
Randomize
Randomize
Text1.Text = ""

    For i = 0 To 3
        CmdAsl(i).Caption = Int(Rnd * 9) + 1
            Text1.Text = Text1.Text & CmdAsl(i).Caption
             
    Next i
If Mid$(Text1.Text, 1, 1) = Mid$(Text1.Text, 2, 1) Or _
    Mid$(Text1.Text, 1, 1) = Mid$(Text1.Text, 3, 1) Or _
    Mid$(Text1.Text, 1, 1) = Mid$(Text1.Text, 4, 1) Or _
    Mid$(Text1.Text, 2, 1) = Mid$(Text1.Text, 1, 1) Or _
    Mid$(Text1.Text, 2, 1) = Mid$(Text1.Text, 3, 1) Or _
    Mid$(Text1.Text, 2, 1) = Mid$(Text1.Text, 4, 1) Or _
    Mid$(Text1.Text, 3, 1) = Mid$(Text1.Text, 4, 1) Or _
    Mid$(Text1.Text, 4, 1) = Mid$(Text1.Text, 3, 1) Then
CmdAsl(0) = True
End If

End Sub

Private Sub CmdMak_Click(Index As Integer)

Text3.Text = ""
If BolTrue = True Then
CmdMak(Index).Tag = CardRed(LblIndexCard).Index
CmdMak(Index).Picture = ImageList1.ListImages(CardRed(LblIndexCard).Index + 27).Picture
            sndPlaySound App.Path & "\sound\boing.wav", &H1 Or &H2 Or &H2000


Bolmove = False
BolTrue = False
CardRed(LblIndexCard.Caption).Left = CardRed(LblIndexCard.Caption).Left - 100
CardRed(LblIndexCard).Picture = ImageList1.ListImages(LblIndexCard.Caption + 1 - 1).Picture

End If
For i = 0 To 3
Text3.Text = Text3.Text & CmdMak(i).Tag
Next i
If Len(Text3.Text) = 4 Then
Command2.Enabled = True
Else
Command2.Enabled = False
End If
End Sub

Private Sub CmdNumber_Click()
Number = FrmOpen.EA(4).Text
C = FrmOpen.EA(1).Text
End Sub

Private Sub CmdStart_Click()
CmdAsl(0) = True
For i = 0 To 3
Text3.Text = Text3.Text & CmdMak(i).Tag
CmdMak(i).Picture = ImagelistKing.ListImages(i + 1).Picture

Next i
I1 = 0
C = 0
B = 0
Number = 0
G = 0
For i = 0 To 35
ImgSmall1(i).Picture = Nothing
ImgToop(i).Picture = Nothing
Next i
End Sub

Private Sub Command1_Click()
Text11.Text = FrmOpen.EA(0).Text


End Sub

Private Sub Command2_Click()

Number = Number + 1
Txtnumber = Number
            sndPlaySound App.Path & "\sound\click22.wav", &H1 Or &H2 Or &H2000

Dim IMG1 As String
Dim Mid1 As String

Text2.Text = ""
    For i = 0 To 3
        If InStr(1, Text1.Text, CmdMak(i).Tag) = i + 1 Then
                Text2.Text = Text2.Text & "+"
        ElseIf InStr(Text1.Text, CmdMak(i).Tag) = 0 Then
        Else
                Text2.Text = Text2.Text & "#"
        End If
    Next i
If Mid$(Text3.Text, 1, 1) = Mid$(Text3.Text, 2, 1) Or _
    Mid$(Text3.Text, 1, 1) = Mid$(Text3.Text, 3, 1) Or _
    Mid$(Text3.Text, 1, 1) = Mid$(Text3.Text, 4, 1) Or _
    Mid$(Text3.Text, 2, 1) = Mid$(Text3.Text, 1, 1) Or _
    Mid$(Text3.Text, 2, 1) = Mid$(Text3.Text, 3, 1) Or _
    Mid$(Text3.Text, 2, 1) = Mid$(Text3.Text, 4, 1) Or _
    Mid$(Text3.Text, 3, 1) = Mid$(Text3.Text, 4, 1) Or _
    Mid$(Text3.Text, 4, 1) = Mid$(Text3.Text, 3, 1) Then
    MsgBox "Your Number is Duplacate", vbOKOnly, "Repeat"
    
Else
LblTar(I1).Caption = Text2.Text
I1 = I1 + 1
B = 0
For i = C To C + 3
B = B + 1
TextOld = Mid$(Text3.Text, B, 1)
ImgSmall1(i).Picture = ImageList1.ListImages(TextOld + 9).Picture
 List1.AddItem TextOld
 If Mid$(LblTar(I1 - 1).Caption, B, 1) = "#" Then
 ImgToop(i).Picture = ImagelistKing.ListImages(5).Picture
 List2.AddItem "5"
 ElseIf Mid$(LblTar(I1 - 1).Caption, B, 1) = "+" Then
  ImgToop(i).Picture = ImagelistKing.ListImages(6).Picture
List2.AddItem "6"
Else
List2.AddItem "0"

 End If
 
Next i
End If
    If Text2.Text = "++++" Then
       FrmAkhar.Show
                   sndPlaySound App.Path & "\sound\START.wav", &H1 Or &H2 Or &H2000

    For i = 3 To 0 Step -1
    FrmAkhar.ImgAkhar(i).Picture = ImagelistKing.ListImages(i + 11).Picture
    Next i
    End If

For i = 0 To 3
CmdMak(i).Picture = ImagelistKing.ListImages(i + 1).Picture
CmdMak(i).Tag = ""
Text3.Text = ""
Command2.Enabled = False
Next i
C = C + 4

If Number = 9 Then
Command2.Enabled = False
For i = 1 To 9
CardRed(i).Enabled = False
Next i
Image3_Click
End If


End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TmrCharkh.Enabled = False
End Sub

Private Sub Form_Load()
Dim Path111 As String
Dim In10 As Long
Dim EAE As String

Path111 = Command$
If Path111 <> "" Then
In10 = InStr(4, Path111, "~1")
If In10 <> 0 Then
EAE = Mid$(Path111, 1, In10 - 1) & ".MSMI"
Else
EAE = Path111
End If

FrmOpen.TxtOpen.Text = EAE
FrmOpen.Command1 = True

End If

Call Reg
ImgAbout.Picture = ImagelistKing.ListImages(7).Picture
For i = 0 To 8
ImgNesh(i).Top = ImgNesh(i).Top + 50
ImgNesh(i).Left = ImgNesh(i).Left + 3700
Next i
CmdAsl(0) = True
For i = 0 To 3
Text3.Text = Text3.Text & CmdMak(i).Tag
CmdMak(i).Picture = ImagelistKing.ListImages(i + 1).Picture

Next i
    For i = 1 To 9
            CardRed(i).Picture = ImageList1.ListImages(i).Picture
    Next i

I1 = 0
C = 0
B = 0
'For I = 0 To 35
'ImgSmall1(I).Move ImgSmall1(I).Left + 200, ImgSmall1(I).Top + 150
'ImgToop(I).Move ImgToop(I).Left + 300, ImgToop(I).Top + 150
'Next I
Number = 0
G = 0
End Sub

Private Sub ImgSmall_Click(Index As Integer)

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
Unload FrmAbout
Unload FrmAkhar
Unload FrmMain
Unload FrmOpen
Unload FrmSave

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgAbout.Picture = ImagelistKing.ListImages(7).Picture
TmrCharkh.Enabled = False

End Sub

Private Sub Image2_Click()
If Command2.Enabled = True Then
Command2 = True
End If

End Sub


Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TmrCharkh.Enabled = True

End Sub

Private Sub Image3_Click()
 For i = 0 To 3
    CmdMak(i).Picture = ImageList1.ListImages(Mid$(Text1.Text, i + 1, 1) + 18).Picture
Next i
Image3.Visible = False

End Sub

Private Sub Image4_Click()
Image3.Visible = True
End Sub

Private Sub Image5_Click()
For i = 0 To 3
    CmdMak(i).Picture = ImagelistKing.ListImages(i + 1).Picture
Next i
End Sub

Private Sub ImgAbout_Click()
sndPlaySound App.Path & "\sound\sheep.wav", &H1 Or &H2 Or &H2000
FrmAbout.ImgToop(0).Picture = ImagelistKing.ListImages(5).Picture
FrmAbout.ImgToop(1).Picture = ImagelistKing.ListImages(6).Picture

FrmAbout.Show

End Sub

Private Sub ImgAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgAbout.Picture = ImagelistKing.ListImages(8).Picture

End Sub

Private Sub ImgSmall1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgSmall1(Index).ToolTipText = List1.List(Index)

End Sub

Private Sub MnuAbout_Click()
ImgAbout_Click

End Sub

Private Sub MnuNew_Click()

Dim H As String
H = MsgBox("Do you want to play again", vbOKCancel, "New Game")
If H = vbOK Then
FrmMain.ImgLock.Visible = False

CmdStart = True
Else
End If

End Sub

Private Sub MnuOpen_Click()
FrmOpen.Show
End Sub

Private Sub MnuPass_Click()
Image3_Click
ImgLock.Visible = True

End Sub

Private Sub MnuSave_Click()
FrmSave.Show
End Sub

Private Sub Picture1_Click()
If Command2.Enabled = True Then
Command2 = True
End If
End Sub

Private Sub TmrCharkh_Timer()
G = G + 1
Image2.Picture = ImageCharkh.ListImages(G).Picture
If G >= 22 Then
G = 0
End If

End Sub
Sub Reg()
  Dim hKey As Long
  Dim subkey As String
  Dim Retval As Long
  Dim Retval1 As Long
  Dim Buffer As String
  Dim secattr As SECURITY_ATTRIBUTES
  Dim Path As String
  Dim Path1 As String
  
  Path = App.Path & "\Bekr.exe"
  Path1 = Path & " " & "%1"
  subkey = ".MSMI"
  secattr.nLength = Len(secattr)
  secattr.lpSecurityDescriptor = 0
  secattr.bInheritHandle = 1
Retval = RegCreateKeyEx(HKEY_CLASSES_ROOT, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, hKey, ByVal 0&)
If Retval = 0 Then
    Buffer = "MasterMind"
    RegSetValueEx hKey, "", 0, REG_SZ, ByVal Buffer, Len(Buffer)
End If


'*************************************************************
  secattr.nLength = Len(secattr)
  secattr.lpSecurityDescriptor = 0
  secattr.bInheritHandle = 1
Retval = RegCreateKeyEx(HKEY_CLASSES_ROOT, "MasterMind", 0, "", 0, KEY_ALL_ACCESS, secattr, hKey, ByVal 0&)
    If Retval = 0 Then
            Retval = RegCreateKeyEx(HKEY_CLASSES_ROOT, "MasterMind\DefaultIcon", 0, "", 0, KEY_ALL_ACCESS, secattr, hKey, ByVal 0&)
                If Retval = 0 Then
                    RegSetValueEx hKey, "", 0, REG_SZ, ByVal Path, Len(Path)
                End If
                    Retval = RegCreateKeyEx(HKEY_CLASSES_ROOT, "MasterMind\shell\open\command", 0, "", 0, KEY_ALL_ACCESS, secattr, hKey, ByVal 0&)
                        If Retval = 0 Then
                            RegSetValueEx hKey, "", 0, REG_SZ, ByVal Path1, Len(Path1)
                        End If
    End If
    
End Sub

