VERSION 5.00
Object = "*\A..\..\Activex Control\OsenXPCntrl.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Sample MessageBox"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin OsenXPCntrl.OsenXPText Text3 
      Height          =   315
      Left            =   3900
      TabIndex        =   16
      Top             =   2820
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "65"
   End
   Begin OsenXPCntrl.OsenXPText Text1 
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   720
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Title"
   End
   Begin OsenXPCntrl.OsenXPForm OsenXPForm1 
      Height          =   5145
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   9075
      Icon            =   "Form1.frx":058C
      Caption         =   "Sample MessageBox"
      ShowHelpButton  =   0   'False
      AutoLoad        =   -1  'True
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      ItemData        =   "Form1.frx":0B28
      Left            =   3900
      List            =   "Form1.frx":0B32
      TabIndex        =   12
      Text            =   "0"
      Top             =   2460
      Width           =   1305
   End
   Begin OsenXPCntrl.OsenXPButton OsenXPButton6 
      Height          =   435
      Left            =   2580
      TabIndex        =   8
      Top             =   4530
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Show Form2"
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
      MICON           =   "Form1.frx":0B3C
      PICN            =   "Form1.frx":0B58
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton OsenXPButton5 
      Height          =   405
      Left            =   2580
      TabIndex        =   7
      Top             =   3630
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Inputbox"
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
      MICON           =   "Form1.frx":10F2
      PICN            =   "Form1.frx":110E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton OsenXPButton4 
      Height          =   585
      Left            =   210
      TabIndex        =   6
      Top             =   4380
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1032
      BTYPE           =   3
      TX              =   "MsgQuestion    "
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
      MICON           =   "Form1.frx":16AA
      PICN            =   "Form1.frx":16C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton OsenXPButton3 
      Height          =   585
      Left            =   210
      TabIndex        =   5
      Top             =   3750
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1032
      BTYPE           =   3
      TX              =   "MsgInformation"
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
      MICON           =   "Form1.frx":23A2
      PICN            =   "Form1.frx":23BE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton OsenXPButton2 
      Height          =   585
      Left            =   210
      TabIndex        =   4
      Top             =   3120
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1032
      BTYPE           =   3
      TX              =   "MsgCritical      "
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
      MICON           =   "Form1.frx":309A
      PICN            =   "Form1.frx":30B6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton OsenXPButton1 
      Height          =   585
      Left            =   210
      TabIndex        =   3
      Top             =   2490
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1032
      BTYPE           =   3
      TX              =   "MsgExlamation"
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
      MICON           =   "Form1.frx":3D92
      PICN            =   "Form1.frx":3DAE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1065
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":4A8A
      Top             =   1320
      Width           =   4995
   End
   Begin OsenXPCntrl.OsenXPButton OsenXPButton7 
      Height          =   405
      Left            =   2580
      TabIndex        =   9
      Top             =   3180
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Show Message"
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
      MICON           =   "Form1.frx":4B05
      PICN            =   "Form1.frx":4B21
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton OsenXPButton8 
      Height          =   435
      Left            =   2580
      TabIndex        =   10
      Top             =   4080
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "About"
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
      MICON           =   "Form1.frx":50BD
      PICN            =   "Form1.frx":50D9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon index"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2640
      TabIndex        =   14
      Top             =   2820
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2610
      TabIndex        =   13
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   210
      TabIndex        =   2
      Top             =   1050
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim osen As New cls_Osen_WinXPCntrl
'// Osen XP Components Description ///////////////////////////
'/* if you want to see current icons in design project,insert OsenXPButton in your project
'/* Right click OsenXPButton and select properties
'/* Icon index in property page , can use to change or set icon in messagebox
'////////////////////////////////////////////////////////////



Private Sub OsenXPButton1_Click()
    osen.MsgExclamation Text2, Text1.Text
End Sub

Private Sub OsenXPButton2_Click()
    osen.MsgCritical Text2, Text1.Text
End Sub

Private Sub OsenXPButton3_Click()
    osen.MsgInformation Text2, Text1.Text
End Sub

Private Sub OsenXPButton4_Click()
    osen.MsgQuestion Text2, Text1.Text
End Sub

Private Sub OsenXPButton5_Click()
    Text2 = osen.InputBoxXP("Enter your name : ", "Hello")
End Sub

Private Sub OsenXPButton6_Click()
    Form2.Show
End Sub

Private Sub OsenXPButton7_Click()
    osen.ShowMessage Text2, Text1.Text, Combo1, Val(Text3.Text)
End Sub

Private Sub OsenXPButton8_Click()
    osen.ShowAboutMe , , , , " "
End Sub
