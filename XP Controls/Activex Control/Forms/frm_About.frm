VERSION 5.00
Begin VB.Form Frm_About 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Created by Osen Kusnadi"
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   Icon            =   "frm_About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin OsenXPCntrl.OsenXPForm XP 
      Height          =   3435
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   6059
      Icon            =   "frm_About.frx":058A
      Caption         =   "About Osen XP Controls"
      ShowMinimizeButton=   0   'False
      ShowMaximizeButton=   0   'False
      ShowTitleIcon   =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton command2 
      Height          =   375
      Left            =   4230
      TabIndex        =   2
      Top             =   1875
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "System Info"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frm_About.frx":0B24
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton Command1 
      Height          =   375
      Left            =   4230
      TabIndex        =   0
      Top             =   1470
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frm_About.frx":0B40
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   225
      X2              =   5450
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frm_About.frx":0B5C
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   945
      Index           =   0
      Left            =   210
      TabIndex        =   9
      Top             =   2385
      Width           =   5190
   End
   Begin VB.Image Image12 
      Height          =   915
      Left            =   0
      Picture         =   "frm_About.frx":0C7E
      Top             =   495
      Width           =   1995
   End
   Begin VB.Label LblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Windows XP Controls"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1755
      TabIndex        =   8
      Top             =   705
      Width           =   3090
   End
   Begin VB.Label LbVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   1755
      TabIndex        =   7
      Top             =   1005
      Width           =   3090
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   4830
      Picture         =   "frm_About.frx":1BB2
      Stretch         =   -1  'True
      Top             =   675
      Width           =   600
   End
   Begin VB.Label LbDevelover 
      BackStyle       =   0  'Transparent
      Caption         =   "Osen Kusnadi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   210
      TabIndex        =   6
      Top             =   1665
      Width           =   1935
   End
   Begin VB.Label LbPhone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile : +6281310722162"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   210
      TabIndex        =   5
      Top             =   2085
      Width           =   2115
   End
   Begin VB.Label LbEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail : okusnadi@cikarang.actaris.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   210
      TabIndex        =   4
      Top             =   1900
      Width           =   3045
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Develover's info :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   1425
      Width           =   1935
   End
End
Attribute VB_Name = "Frm_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub command2_Click()
    StartSysInfo
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
         LbVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
       Me.LbDevelover.Caption = IsDevelover '= IpDevelover
       If UCase(IsVersion) <> "NOVERSION" Then Me.LbVersion.Caption = IsVersion      '= IpVersion
       Me.LbEmail.Caption = IsEmail '= IpEmail
       Me.LblTitle.Caption = IsTitle '= IpTitle
       Me.LbPhone.Caption = IsPhone '= IpPhone
       XP.Caption = "About " & IsTitle
    XP.Left = 0: XP.Top = 0
    XP.LoadXP True
End Sub
