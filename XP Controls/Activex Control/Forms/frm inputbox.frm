VERSION 5.00
Begin VB.Form FrmInput 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "frmInputBox"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin OsenXPCntrl.XPF XPI 
      Height          =   2025
      Left            =   0
      Top             =   0
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   3572
   End
   Begin OsenXPCntrl.OsenXPText Text1 
      Height          =   300
      Left            =   180
      TabIndex        =   4
      Top             =   1575
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   529
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
   End
   Begin VB.TextBox txtClose 
      Height          =   345
      Left            =   2430
      TabIndex        =   1
      Text            =   "0"
      Top             =   -660
      Width           =   495
   End
   Begin OsenXPCntrl.OsenXPButton CmdRespon 
      Height          =   390
      Index           =   0
      Left            =   4410
      TabIndex        =   2
      Top             =   1000
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   688
      BTYPE           =   3
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      MICON           =   "frm inputbox.frx":0000
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
      Default         =   -1  'True
      Height          =   390
      Index           =   1
      Left            =   4410
      TabIndex        =   3
      Top             =   550
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   688
      BTYPE           =   3
      TX              =   "&OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      MICON           =   "frm inputbox.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   180
      TabIndex        =   0
      Top             =   585
      Width           =   555
   End
End
Attribute VB_Name = "FrmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public bSTand As Boolean

Private Sub CmdRespon_Click(Index As Integer)
    If Index = 1 Then
        ResponInput = Text1.Text
        txtClose.Text = 1
    Else
        ResponInput = ""
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    Text1.Focus
End Sub

Private Sub Form_Load()
On Error GoTo EEE
Dim ILen As Integer
    Me.Hide
    bSTand = False
    DoEvents
    Me.Caption = s_Title
    Me.BackColor = msgBackColor
    Label1.Caption = s_Message
    If Label1.Width > 4100 Then
        With Label1
            .Autosize = False
            ILen = (Len(.Caption) \ 60) + 1
            .Width = 4000
            .Height = ILen * 300
            .Caption = s_Message
        End With
    End If
    XPI.Left = 0: XPI.Top = 0
    XPI.LoadXP
    DoEvents
    bSTand = True
EEE:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If txtClose = 0 Then ResponInput = ""
    bSTand = False
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text1.Text <> "" Then
            CmdRespon_Click 1
        Else
            CmdRespon_Click 0
        End If
    ElseIf KeyAscii = 27 Then
        CmdRespon_Click 0
    End If
End Sub
