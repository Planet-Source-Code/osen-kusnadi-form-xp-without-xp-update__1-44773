VERSION 5.00
Begin VB.Form frm_Message 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Created by Osen Kusnadi"
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   Icon            =   "frm_message.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin OsenXPCntrl.XPF XPM 
      Height          =   1875
      Left            =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3307
   End
   Begin OsenXPCntrl.OsenXPButton CmdRespon 
      Height          =   375
      Index           =   0
      Left            =   2700
      TabIndex        =   2
      Top             =   1260
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&No"
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
      MICON           =   "frm_message.frx":058C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton CmdOK 
      Height          =   375
      Left            =   1575
      TabIndex        =   1
      Top             =   1245
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&OK"
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
      MICON           =   "frm_message.frx":05A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton CmdRespon 
      Height          =   375
      Index           =   1
      Left            =   450
      TabIndex        =   3
      Top             =   1260
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Yes"
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
      MICON           =   "frm_message.frx":05C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LbText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message Value"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   750
      Width           =   1275
   End
   Begin VB.Image IMG 
      Height          =   480
      Left            =   270
      Picture         =   "frm_message.frx":05E0
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frm_Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub CmdOK_Click()
    Unload Me
    ResponMsg = 0
End Sub

Private Sub CmdRespon_Click(Index As Integer)
    ResponMsg = Index
    Unload Me
End Sub

Private Sub Form_Load()
Me.Hide
Me.Caption = s_Title
DoEvents
Dim ILen As Integer
If IMGS Then
    IMG.Left = 300
    IMG.Top = 600
    IMG.Visible = True
    LbText.Caption = s_Message
    LbText.Left = IMG.Left + IMG.Width + 150
    LbText.Top = 600
    If LbText.Width > 5800 Then
        LbText.Autosize = False
        ILen = (Len(LbText) \ 60) + 1
        LbText.Width = 5800
        LbText.Height = ILen * 300
    End If
    XPM.Width = LbText.Width + 750 + IMG.Width
    If XPM.Width < 3000 Then XPM.Width = 3200
    If MyType = 0 Then
        CmdOK.Visible = True
        CmdOK.Top = LbText.Top + LbText.Height + 170
        If CmdOK.Top < IMG.Height + 600 Then
            CmdOK.Top = IMG.Height + 600 + 270
        End If
        CmdOK.Left = (XPM.Width - CmdOK.Width) / 2
        XPM.Height = CmdOK.Top + CmdOK.Height + 150
    Else
        CmdRespon(0).Visible = True
        CmdRespon(1).Visible = True
        If XPM.Width < (CmdRespon(1).Width + CmdRespon(0).Width + 150) Then
            XPM.Width = (CmdRespon(1).Width + CmdRespon(0).Width + 150) + 600
        End If
        CmdRespon(1).Left = (XPM.Width - (CmdRespon(1).Width + CmdRespon(0).Width + 90)) / 2
        CmdRespon(0).Left = CmdRespon(1).Left + CmdRespon(1).Width + 90
        CmdRespon(0).Top = LbText.Top + LbText.Height + 170
        If CmdRespon(0).Top < IMG.Height + 600 Then
            CmdRespon(0).Top = IMG.Height + 600 + 200
        End If
        CmdRespon(1).Top = CmdRespon(0).Top
        XPM.Height = CmdRespon(1).Top + CmdRespon(1).Height + 150
    End If
Else
    LbText.Caption = s_Message
    LbText.Left = 300
    LbText.Top = 600
    If LbText.Width > 5800 Then
        LbText.Autosize = False
        ILen = (Len(LbText) \ 60) + 1
        LbText.Width = 5800
        LbText.Height = ILen * 300
    End If
    XPM.Width = LbText.Width + 600
    If MyType = 0 Then
        CmdOK.Visible = True
        CmdOK.Top = LbText.Top + LbText.Height + 170
        CmdOK.Left = (XPM.Width - CmdOK.Width) / 2
        XPM.Height = CmdOK.Top + CmdOK.Height + 150
    Else
        CmdRespon(0).Visible = True
        CmdRespon(1).Visible = True
        If XPM.Width < (CmdRespon(1).Width + CmdRespon(0).Width + 150) Then
            XPM.Width = (CmdRespon(1).Width + CmdRespon(0).Width + 150) + 600
        End If
        CmdRespon(1).Left = (XPM.Width - (CmdRespon(1).Width + CmdRespon(0).Width + 150)) / 2
        CmdRespon(0).Left = CmdRespon(1).Left + CmdRespon(1).Width + 150
        CmdRespon(0).Top = LbText.Top + LbText.Height + 170
        CmdRespon(1).Top = CmdRespon(0).Top
        XPM.Height = CmdRespon(1).Top + CmdRespon(1).Height + 150
    End If
End If
With XPM
    .Left = 0
    .Top = 0
    .LoadXP
End With
Me.BackColor = msgBackColor
DoEvents
End Sub

Private Sub xpm_Closed()
    Unload Me
End Sub


