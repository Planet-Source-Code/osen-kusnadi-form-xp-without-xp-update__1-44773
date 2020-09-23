VERSION 5.00
Begin VB.UserControl OsenXPForm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8E9EC&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   2580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4890
   ClipControls    =   0   'False
   EditAtDesignTime=   -1  'True
   ForwardFocus    =   -1  'True
   LockControls    =   -1  'True
   PropertyPages   =   "Osen XP Form.ctx":0000
   ScaleHeight     =   172
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   326
   ToolboxBitmap   =   "Osen XP Form.ctx":000F
   Begin VB.Label lb_ok 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "osen"
      Height          =   195
      Left            =   2010
      TabIndex        =   2
      Top             =   -750
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image CbHelp 
      Height          =   315
      Index           =   3
      Left            =   1740
      Picture         =   "Osen XP Form.ctx":0321
      ToolTipText     =   "close"
      Top             =   3120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   4
      Left            =   2040
      Picture         =   "Osen XP Form.ctx":08A3
      Top             =   4170
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   3
      Left            =   1710
      Picture         =   "Osen XP Form.ctx":0E25
      Top             =   3840
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   4
      Left            =   2040
      Picture         =   "Osen XP Form.ctx":13A7
      Top             =   3450
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   3
      Left            =   1710
      Picture         =   "Osen XP Form.ctx":1929
      Top             =   2760
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   2340
      Picture         =   "Osen XP Form.ctx":1EAB
      Stretch         =   -1  'True
      Top             =   2790
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   90
      Picture         =   "Osen XP Form.ctx":21EF
      Top             =   2820
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   4680
      Picture         =   "Osen XP Form.ctx":25E6
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Image4 
      Height          =   2085
      Left            =   420
      Picture         =   "Osen XP Form.ctx":29E6
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Image5 
      Height          =   2085
      Left            =   4950
      Picture         =   "Osen XP Form.ctx":2D14
      Stretch         =   -1  'True
      Top             =   3030
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Image6 
      Height          =   60
      Left            =   360
      Picture         =   "Osen XP Form.ctx":3042
      Stretch         =   -1  'True
      Top             =   5160
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.Image Image7 
      Height          =   60
      Left            =   4800
      Picture         =   "Osen XP Form.ctx":3370
      Top             =   4860
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Image8 
      Height          =   60
      Left            =   270
      Picture         =   "Osen XP Form.ctx":36AA
      Top             =   4980
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   6990
      Picture         =   "Osen XP Form.ctx":39E3
      Stretch         =   -1  'True
      Top             =   2790
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   5580
      Picture         =   "Osen XP Form.ctx":3D45
      Top             =   2850
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Image11 
      Height          =   450
      Left            =   9270
      Picture         =   "Osen XP Form.ctx":4021
      Top             =   2880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Image12 
      Height          =   2085
      Left            =   5400
      Picture         =   "Osen XP Form.ctx":4303
      Stretch         =   -1  'True
      Top             =   2910
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Image13 
      Height          =   2085
      Left            =   9510
      Picture         =   "Osen XP Form.ctx":4679
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Image14 
      Height          =   60
      Left            =   5310
      Picture         =   "Osen XP Form.ctx":49E1
      Stretch         =   -1  'True
      Top             =   5070
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.Image Image15 
      Height          =   60
      Left            =   9690
      Picture         =   "Osen XP Form.ctx":4D13
      Top             =   4710
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Image16 
      Height          =   60
      Left            =   5280
      Picture         =   "Osen XP Form.ctx":4FA5
      Top             =   4770
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image TitleIcon 
      Height          =   240
      Left            =   120
      Picture         =   "Osen XP Form.ctx":523B
      Top             =   90
      Width           =   240
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   3
      Left            =   1710
      Picture         =   "Osen XP Form.ctx":57C5
      ToolTipText     =   "close"
      Top             =   3465
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   3
      Left            =   1710
      Picture         =   "Osen XP Form.ctx":5D47
      ToolTipText     =   "close"
      Top             =   4185
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image MaximizeButton 
      Height          =   315
      Left            =   4110
      Picture         =   "Osen XP Form.ctx":62C9
      ToolTipText     =   "Maximize"
      Top             =   60
      Width           =   315
   End
   Begin VB.Image HelpButton 
      Height          =   315
      Left            =   4110
      Picture         =   "Osen XP Form.ctx":684B
      ToolTipText     =   "Help"
      Top             =   60
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   2
      Left            =   1350
      Picture         =   "Osen XP Form.ctx":6DCD
      ToolTipText     =   "close"
      Top             =   4185
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   1
      Left            =   990
      Picture         =   "Osen XP Form.ctx":734F
      Top             =   4185
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   0
      Left            =   630
      Picture         =   "Osen XP Form.ctx":78D1
      Top             =   4185
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   2
      Left            =   1350
      Picture         =   "Osen XP Form.ctx":7E53
      ToolTipText     =   "close"
      Top             =   3825
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   1
      Left            =   990
      Picture         =   "Osen XP Form.ctx":83D5
      Top             =   3825
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   0
      Left            =   630
      Picture         =   "Osen XP Form.ctx":8957
      Top             =   3825
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   2
      Left            =   1350
      Picture         =   "Osen XP Form.ctx":8ED9
      ToolTipText     =   "close"
      Top             =   3465
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   1
      Left            =   990
      Picture         =   "Osen XP Form.ctx":945B
      Top             =   3465
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   0
      Left            =   630
      Picture         =   "Osen XP Form.ctx":99DD
      Top             =   3465
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbHelp 
      Height          =   315
      Index           =   2
      Left            =   1350
      Picture         =   "Osen XP Form.ctx":9F5F
      ToolTipText     =   "close"
      Top             =   3105
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbHelp 
      Height          =   315
      Index           =   1
      Left            =   990
      Picture         =   "Osen XP Form.ctx":A4E1
      Top             =   3105
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbHelp 
      Height          =   315
      Index           =   0
      Left            =   630
      Picture         =   "Osen XP Form.ctx":AA63
      Top             =   3105
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   2
      Left            =   1350
      Picture         =   "Osen XP Form.ctx":AFE5
      ToolTipText     =   "close"
      Top             =   2745
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   1
      Left            =   990
      Picture         =   "Osen XP Form.ctx":B567
      Top             =   2745
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   0
      Left            =   630
      Picture         =   "Osen XP Form.ctx":BAE9
      Top             =   2745
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Caption1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SEN MASTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   390
      TabIndex        =   0
      Top             =   90
      Width           =   1215
   End
   Begin VB.Label Caption2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SEN MASTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0083180A&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   90
      Width           =   1335
   End
   Begin VB.Image CloseButton 
      Height          =   315
      Left            =   4440
      Picture         =   "Osen XP Form.ctx":C06B
      ToolTipText     =   "Close"
      Top             =   60
      Width           =   315
   End
   Begin VB.Image MinimizeButton 
      Height          =   315
      Left            =   3780
      Picture         =   "Osen XP Form.ctx":C5ED
      ToolTipText     =   "Minimize"
      Top             =   60
      Width           =   315
   End
   Begin VB.Image Title 
      Height          =   450
      Left            =   180
      Picture         =   "Osen XP Form.ctx":CB6F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image TitleLeft 
      Height          =   450
      Left            =   30
      Picture         =   "Osen XP Form.ctx":CEB3
      Top             =   0
      Width           =   150
   End
   Begin VB.Image TitleRight 
      Height          =   450
      Left            =   4740
      Picture         =   "Osen XP Form.ctx":D2AA
      Top             =   0
      Width           =   150
   End
   Begin VB.Image Left 
      Height          =   2085
      Left            =   30
      Picture         =   "Osen XP Form.ctx":D6AA
      Stretch         =   -1  'True
      Top             =   450
      Width           =   60
   End
   Begin VB.Image Right 
      Height          =   2085
      Left            =   4830
      Picture         =   "Osen XP Form.ctx":D9D8
      Stretch         =   -1  'True
      Top             =   450
      Width           =   60
   End
   Begin VB.Image Bottom 
      Height          =   60
      Left            =   90
      Picture         =   "Osen XP Form.ctx":DD06
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   4755
   End
   Begin VB.Image BottomRight 
      Height          =   60
      Left            =   4830
      Picture         =   "Osen XP Form.ctx":E034
      Top             =   2520
      Width           =   60
   End
   Begin VB.Image BottomLeft 
      Height          =   60
      Left            =   30
      Picture         =   "Osen XP Form.ctx":E36E
      Top             =   2520
      Width           =   60
   End
End
Attribute VB_Name = "OsenXPForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'// My Improvement :
'// 1. AutoLoad
'/  2. Change Skin if form is deactivate
'/  3. Change controlbox when mouse over
'/      Easy to use, no added code,just add this control to your form
'/      easy to change icon, and i includes more than 100 icons
' The following was cut & pasted from original project

' ### ### #####         ###### ### ###### ###    ######
'  #####  ## ###        ###### ### ###### ###    ###
'  #####  #####  ######   ##   ###   ##   ###### ######
' ### ### ###    ######   ##   ###   ##   ###### ######

'     Copyright © 2002 by Doug Sheffer
'
'     Distributed freely so long that this notice stays at the top
'
'     Please include authors name in your resulting application
'/**************** Declare API Function **************************************************************************
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvPara As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Public Event Help()
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Dim bTransparent As Boolean

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private WithEvents MYFORM As Form
Attribute MYFORM.VB_VarHelpID = -1

 Const MF_BYPOSITION = &H400&
 Const MF_BYCOMMAND = 0
 Const SC_RESTORE = &HF120
 Const SC_MOVE = &HF010
 Const SC_SIZE = &HF000
 Const SC_MINIMIZE = &HF020
 Const SC_MAXIMIZE = &HF030
 Const SC_CLOSE = &HF060
 Const WM_GETSYSMENU = &H313
 Const HWND_TOPMOST = -1
 Const HWND_NOTOPMOST = -2

Public My_MDI As PictureBox
Public IsMDI As Boolean
Const GWL_STYLE = (-16)
Const WS_SYSMENU = &H80000

Private MyTitleIcon As Image
'Dim m_ShowHelpButton As Boolean
'Default Property Values:
Const m_def_AutoLoad = False
Const m_def_ShowHelpButton = 0
'Const m_def_ShowHelp = 0
'Property Variables:
Dim m_AutoLoad As Boolean
Dim m_ShowHelpButton As Boolean
'Dim m_ShowHelp As Boolean





'/************************************************************************
'=================== OSEN FORM =============================================

Public Sub RePos()
    'This repositions the different controls on the form when it is resized
    On Error Resume Next
    Dim X As Single
    Dim Y As Single
    
    If UserControl.Height < 615 Then UserControl.Height = 615   'Checks that form
    If UserControl.Width < 1695 Then UserControl.Width = 1695   'is not too small
    
    X = UserControl.Width / Screen.TwipsPerPixelX   'Registers the size of the
    Y = UserControl.Height / Screen.TwipsPerPixelY  'form in pixels
    
    'Titlebar
    With TitleLeft
        .Left = 0
        .Top = 0
    End With
    
    With Title
        .Left = TitleLeft.Width
        .Top = 0
        .Width = X - TitleLeft.Width - TitleRight.Width
    End With
    
    With TitleRight
        .Left = Title.Left + Title.Width
        .Top = 0
    End With
    
    'Borders
    With BottomLeft
        .Left = 0
        .Top = Y - .Height
    End With
    
    With BottomRight
        .Left = X - .Width
        .Top = Y - .Height
    End With
    
    With Left
        .Left = 0
        .Top = TitleLeft.Top + TitleLeft.Height
        .Height = BottomLeft.Top - .Top
    End With
    
    With Right
        .Left = X - .Width
        .Top = TitleRight.Top + TitleRight.Height
        .Height = BottomRight.Top - .Top
    End With
    
    With Bottom
        .Left = BottomLeft.Width
        .Top = Y - Bottom.Height
        .Width = X - BottomLeft.Width - BottomRight.Width
    End With
    
    'Buttons
    With CloseButton
        .Left = Right.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
    End With
    
    With MaximizeButton
        .Left = CloseButton.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
    End With
    
    'Help Button
    With HelpButton
        .Left = CloseButton.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
    End With
    
    With MinimizeButton
        .Left = MaximizeButton.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
    End With
    
    'Icon
    With TitleIcon
        .Left = Left.Left + Left.Width + 2
        .Top = (Title.Height - .Height) / 2
    End With
    
    'Titlebar Caption
    With Caption1
        If TitleIcon.Visible = True Then
        .Left = TitleIcon.Left + TitleIcon.Width + 3
        Else
        .Left = Left.Left + Left.Width + 2.5
        End If
        .Top = ((Title.Height - 13) / 2) - 1
        .Width = MinimizeButton.Left - TitleIcon.Left - TitleIcon.Width - 10
        If MinimizeButton.Visible = False Then
            .Width = MaximizeButton.Left - TitleIcon.Left - TitleIcon.Width - 10
        End If
        If MinimizeButton.Visible = False And TitleIcon.Visible = False Then
            .Width = MaximizeButton.Left - Left.Left - Left.Width - 10
        End If
        If MinimizeButton.Visible = False And MaximizeButton.Visible = False Then
            .Width = CloseButton.Left - TitleIcon.Left - TitleIcon.Width - 10
        End If
        If MinimizeButton.Visible = False And MaximizeButton.Visible = False And TitleIcon.Visible = False Then
            .Width = CloseButton.Left - Left.Left - Left.Width - 10
        End If
        
        .Height = 13
    End With
    
    With Caption2
        If TitleIcon.Visible = True Then
            .Left = TitleIcon.Left + TitleIcon.Width + 2
        Else
            .Left = Left.Left + Left.Width + 1.5
        End If
        .Top = ((Title.Height - 13) / 2) + 1
        .Width = Caption1.Width
        .Height = 13
    End With
    
    'Checks if it should have transparent corners
    If bTransparent = True Then
        ReTrans
    End If
End Sub

Public Sub TransparentEdges()
    'This is used as a safe guard set when the application starts,
    'otherwise the control would set the corners transparent at design time
    bTransparent = True
    RePos
End Sub

Public Sub ReTrans()
    Dim Add As Long
    Dim Sum As Long
    
    Dim X As Single
    Dim Y As Single
    
    If UserControl.Height < 615 Then UserControl.Height = 615   'Checks that form
    If UserControl.Width < 1695 Then UserControl.Width = 1695   'is not too small
    
    X = UserControl.Width / Screen.TwipsPerPixelX   'Registers the size of the
    Y = UserControl.Height / Screen.TwipsPerPixelY  'form in pixels
    
    Sum = CreateRectRgn(5, 0, X - 5, 1)
    CombineRgn Sum, Sum, CreateRectRgn(3, 1, X - 3, 2), 2
    CombineRgn Sum, Sum, CreateRectRgn(2, 2, X - 2, 3), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 3, X - 1, 4), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 4, X - 1, 5), 2
    CombineRgn Sum, Sum, CreateRectRgn(0, 5, X, Y), 2
    SetWindowRgn UserControl.ContainerHwnd, Sum, True   'Sets corners transparent
End Sub

Private Sub Caption1_DblClick()
    If Not IsMissing(MYFORM) Then
        If MYFORM.BorderStyle = 2 Then MaximizeButton_Click
    End If
End Sub

Private Sub CloseButton_Click()
On Error GoTo EF
    Unload MYFORM
EF:
End Sub

Private Sub CloseButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then CloseButton.Picture = CbClose(2).Picture
End Sub

Private Sub CloseButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error GoTo XZCV
    If Title.Picture = Image1.Picture Then
        HelpButton.Picture = CbHelp(0).Picture
    Else
        HelpButton.Picture = CbHelp(3).Picture
    End If
    If MinimizeButton.Enabled Then
        If Title.Picture = Image1.Picture Then
            MinimizeButton.Picture = CbMin(0).Picture
        Else
            MinimizeButton.Picture = CbMin(4).Picture
        End If
    End If
    If MaximizeButton.Enabled Then
        If Title.Picture = Image1.Picture Then
            If MYFORM.WindowState = 0 Then
                MaximizeButton.Picture = CbMax(0).Picture
            Else
                MaximizeButton.Picture = CbRestore(0).Picture
            End If
        Else
            If MYFORM.WindowState = 0 Then
                MaximizeButton.Picture = CbMax(4).Picture
            Else
                MaximizeButton.Picture = CbRestore(3).Picture
            End If
        End If
    End If
    If Button = vbLeftButton Then
        CloseButton.Picture = CbClose(2).Picture
    Else
        CloseButton.Picture = CbClose(1).Picture
    End If
XZCV:
End Sub

Private Sub HelpButton_Click()
    RaiseEvent Help
End Sub

Private Sub HelpButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then HelpButton.Picture = CbHelp(2).Picture
End Sub

Private Sub HelpButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Title.Picture = Image1.Picture Then
        CloseButton.Picture = CbClose(0).Picture
    Else
        CloseButton.Picture = CbClose(3).Picture
    End If
    If Button = vbLeftButton Then
        HelpButton.Picture = CbHelp(2).Picture
    Else
        HelpButton.Picture = CbHelp(1).Picture
    End If
End Sub

Private Sub MaximizeButton_Click()
On Error GoTo xc
    If MYFORM.WindowState = 0 Then
        MYFORM.WindowState = 2
    Else
        MYFORM.WindowState = 0
    End If
    
    UserControl.Width = MYFORM.Width
    UserControl.Height = MYFORM.Height
    ReTransObj MYFORM
    If MYFORM.WindowState = 0 Then
        MaximizeButton.Picture = CbMax(0).Picture
        MaximizeButton.ToolTipText = "Maximize"
    Else
        MaximizeButton.Picture = CbRestore(0).Picture
        MaximizeButton.ToolTipText = "Restore"
    End If
xc:
End Sub

Private Sub MaximizeButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then MaximizeButton.Picture = CbMax(2).Picture
End Sub

Private Sub MaximizeButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
On Error GoTo Asd
    If Title.Picture = Image1.Picture Then
        HelpButton.Picture = CbHelp(0).Picture
    Else
        HelpButton.Picture = CbHelp(3).Picture
    End If
    If MinimizeButton.Enabled Then
        If Title.Picture = Image1.Picture Then
            MinimizeButton.Picture = CbMin(0).Picture
            CloseButton.Picture = CbClose(0).Picture
        Else
            MinimizeButton.Picture = CbMin(4).Picture
            CloseButton.Picture = CbClose(3).Picture
        End If
    End If
    If MaximizeButton.Enabled Then
        If Button = vbLeftButton Then
            If MYFORM.WindowState = 0 Then
                MaximizeButton.Picture = CbMax(2).Picture
            Else
                MaximizeButton.Picture = CbRestore(2).Picture
            End If
        Else
            If MYFORM.WindowState = 0 Then
                MaximizeButton.Picture = CbMax(1).Picture
            Else
                MaximizeButton.Picture = CbRestore(1).Picture
            End If
        End If
    End If
Asd:
End Sub

Private Sub MinimizeButton_Click()
On Error GoTo DDF
    MYFORM.WindowState = 1
DDF:
End Sub

Private Sub MinimizeButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then MinimizeButton.Picture = CbMin(2).Picture
End Sub

Private Sub MinimizeButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo JKL
    If Title.Picture = Image1.Picture Then
        HelpButton.Picture = CbHelp(0).Picture
    Else
        HelpButton.Picture = CbHelp(3).Picture
    End If
    If MaximizeButton.Enabled Then
        If Title.Picture = Image1.Picture Then
            If MYFORM.WindowState = 0 Then
                MaximizeButton.Picture = CbMax(0).Picture
            Else
                MaximizeButton.Picture = CbRestore(0).Picture
            End If
            CloseButton.Picture = CbClose(0).Picture
        Else
            If MYFORM.WindowState = 0 Then
                MaximizeButton.Picture = CbMax(4).Picture
            Else
                MaximizeButton.Picture = CbRestore(3).Picture
            End If
            CloseButton.Picture = CbClose(3).Picture
        End If
    End If
    If MinimizeButton.Enabled Then
        If Button = vbLeftButton Then
            MinimizeButton.Picture = CbMin(2).Picture
        Else
            MinimizeButton.Picture = CbMin(1).Picture
        End If
    End If
JKL:
End Sub

Private Sub MYFORM_Activate()
    SetFormActiveStyle True
End Sub
Private Sub MYFORM_Deactivate()
    SetFormActiveStyle False
End Sub
Private Sub MYFORM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetNormalButtonControl
End Sub
Private Sub Right_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetNormalButtonControl
End Sub

Private Sub Title_DblClick()
    If Not IsMissing(MYFORM) Then
        If MYFORM.BorderStyle = 2 Then MaximizeButton_Click
    End If
End Sub

Private Sub Title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo HErr
    SetNormalButtonControl
HErr:
End Sub

Private Sub TitleRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo HGErr:
    HelpButton.Picture = CbHelp(0).Picture
    SetNormalButtonControl
HGErr:
End Sub

Private Sub UserControl_Initialize()
    On Error GoTo FGR
    bTransparent = False  'So we do not set the corners transparent while still in design mode
    RePos   'Reposition
    Set MyTitleIcon = TitleIcon
    SetHelpButton m_ShowHelpButton
    If m_ShowHelpButton Then
        MinimizeButton.Visible = False
        MaximizeButton.Visible = False 'Not IsVisible
    End If
    Caption1.Caption = "Hello " & GetCompName
    Caption2.Caption = "Hello " & GetCompName
FGR:
End Sub

Private Sub UserControl_InitProperties()
On Error GoTo GHE
    UserControl.Parent.BackColor = DefaultBackgroundColor
    Set UserControl.Parent.Icon = frmImages.IMG.ListImages(123).Picture
    TitleIcon.Picture = frmImages.IMG.ListImages(123).Picture
    IsMDI = False
    m_ShowHelpButton = m_def_ShowHelpButton
    If lb_ok.Caption = "osen" Then
        ShowAboutMe
        lb_ok.Caption = "Sen Master"
    End If
    m_AutoLoad = True
GHE:
End Sub

Private Sub UserControl_Resize()
On Error GoTo XFG
    RePos   'Reposition
XFG:
End Sub

Private Sub Title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub TitleLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub TitleRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub Caption1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub Caption2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Public Function DefaultBackgroundColor() As String
    DefaultBackgroundColor = &HD8E9EC   '&HEAF1F1   'Returns a common off-white Windows XP color
End Function
Public Sub LoadXP( _
      Optional OptShowModal As Boolean = False _
    , Optional OptFormOnTop As Boolean = False)
On Error GoTo H
'/******* Hidden Object in procccess **********************
    Dim IpForm As Form
    Set MYFORM = UserControl.Parent
    Set IpForm = MYFORM
    If OptDefaultBackColor Then MYFORM.BackColor = DefaultBackgroundColor
    IpForm.Hide
    '/* Set My_MDI Picture BOX ------------------
    
    If IsMDI Then
        Set My_MDI = MYFORM.Controls.Add("VB.PictureBox", "picOsenKusnadi", MYFORM)
    End If
    TitleIcon.Picture = MYFORM.Icon
        IpForm.Width = UserControl.Width
        IpForm.Height = UserControl.Height
        If IpForm.BorderStyle <> 0 Then
            IpForm.Height = UserControl.Height + 375
            If IpForm.BorderStyle <> 2 Then
                MinimizeButton.Visible = False
                MaximizeButton.Visible = False
            End If
        End If
    SetStyle IpForm
    UserControl.Width = IpForm.Width        ':   usercontrol.Top = 0
    UserControl.Height = IpForm.Height      ': usercontrol.Left = 0
    '/**************** Set Transparant ************************
    ReTransObj IpForm
    DoEvents
    FormOnTop IpForm.hWnd, OptFormOnTop
    If OptShowModal Then
        IpForm.Hide
    End If
H:
End Sub
Public Sub ReTransObj(IpObject As Object)
    Dim Add As Long
    Dim Sum As Long
    Dim X As Single
    Dim Y As Single
    If IpObject.Height < 615 Then IpObject.Height = 615   'Checks that form
    If IpObject.Width < 1695 Then IpObject.Width = 1695   'is not too small
    X = IpObject.Width / Screen.TwipsPerPixelX   'Registers the size of the
    Y = IpObject.Height / Screen.TwipsPerPixelY  'form in pixels
    Sum = CreateRectRgn(5, 0, X - 5, 1)
    CombineRgn Sum, Sum, CreateRectRgn(3, 1, X - 3, 2), 2
    CombineRgn Sum, Sum, CreateRectRgn(2, 2, X - 2, 3), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 3, X - 1, 4), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 4, X - 1, 5), 2
    CombineRgn Sum, Sum, CreateRectRgn(0, 5, X, Y), 2
    SetWindowRgn IpObject.hWnd, Sum, True   'Sets corners transparent
End Sub
Public Sub SetStyle(IpForm As Object, Optional IsHide As Boolean = False)
On Error GoTo XPF
    IpForm.Hide
    Dim lCurrentSettings As Long
    Const WS_MINIMIZEBOX = &H20000
    Const WS_MAXIMIZEBOX = &H10000
    Const WS_THICKFRAME = &H40000
    Const WS_DLGFRAME = &H400000
    Const WS_CAPTION = &HC00000
    lCurrentSettings = GetWindowLong(IpForm.hWnd, GWL_STYLE)
    lCurrentSettings = lCurrentSettings And Not WS_THICKFRAME
    lCurrentSettings = lCurrentSettings And Not WS_DLGFRAME
    lCurrentSettings = lCurrentSettings And Not WS_CAPTION
    lCurrentSettings = lCurrentSettings And Not WS_MINIMIZEBOX
    lCurrentSettings = lCurrentSettings And Not WS_MAXIMIZEBOX
    lCurrentSettings = lCurrentSettings Or WS_SYSMENU
    IpForm.Hide
    SetWindowLong IpForm.hWnd, GWL_STYLE, lCurrentSettings
    If Not IsHide Then
        SetWindowPos IpForm.hWnd, 0, IpForm.Left / 15, IpForm.Top / 15, (IpForm.Width / 15), (IpForm.Height / 15), &H40
    Else
        SetWindowPos IpForm.hWnd, 0, 0, 0, 0, 0, &H40
        IpForm.Hide
    End If
    If IpForm.BorderStyle <> 0 Then
        IpForm.Height = IpForm.Height - 365
    End If
    IpForm.Left = (Screen.Width - IpForm.Width) / 2
    IpForm.Top = (Screen.Height - IpForm.Height) / 2
XPF:
End Sub
Public Sub FormOnTop(hWindow As Long, bTopMost As Boolean)
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    Select Case bTopMost
    Case True
        Placement = HWND_TOPMOST
    Case False
        Placement = HWND_NOTOPMOST
    End Select
    SetWindowPos hWindow, Placement, 0, 0, 0, 0, wFlags
End Sub
Public Function GetCompName() As String
    Dim Commstr As String, nErr As Long
    Commstr = Space(255)
    nErr = GetComputerName(Commstr, 255)
    GetCompName = Commstr
End Function

Private Sub UserControl_Show()
    On Error GoTo H
        If m_AutoLoad Then
            Dim oxc As Object
            For Each oxc In UserControl.Parent
                If TypeOf oxc Is OsenXPForm Then
                    oxc.Left = 0
                    oxc.Top = 0
                    oxc.LoadXP
                    Exit For
                End If
            Next
        End If
H:
End Sub

Private Sub UserControl_Terminate()
    Set MYFORM = Nothing
    TaskBarShow
End Sub
Function TaskBarHide()
    Dim rtn
    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Function
Function TaskBarShow()
    Dim rtn As Long
    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Function
Public Sub SetNormalButtonControl()
On Error Resume Next
    If Title.Picture = Image1.Picture Then
       CloseButton.Picture = CbClose(0).Picture
       HelpButton.Picture = CbHelp(0).Picture
       If MinimizeButton.Enabled Then MinimizeButton.Picture = CbMin(0).Picture
       If Not MaximizeButton.Enabled Then Exit Sub
       If IsMissing(MYFORM) Then Exit Sub
        If MYFORM.WindowState = 0 Then
            MaximizeButton.Picture = CbMax(0).Picture
        Else
            MaximizeButton.Picture = CbRestore(0).Picture
        End If
    Else
       CloseButton.Picture = CbClose(3).Picture
       HelpButton.Picture = CbHelp(3).Picture
       If MinimizeButton.Enabled Then MinimizeButton.Picture = CbMin(4).Picture
       If Not MaximizeButton.Enabled Then Exit Sub
       If IsMissing(MYFORM) Then Exit Sub
        If MYFORM.WindowState = 0 Then
            MaximizeButton.Picture = CbMax(4).Picture
        Else
            MaximizeButton.Picture = CbRestore(3).Picture
        End If
    End If
End Sub
Public Sub ChangeTitle(NewTitle As String)
    Caption1.Caption = NewTitle
End Sub
Public Sub ChangeIcon(newIcon As StdPicture)
    TitleIcon.Picture = newIcon
End Sub
Public Sub ChangeForeColor(NewColour As Long)
    Caption1.ForeColor = NewColour
End Sub

Public Sub SetMDIForm(XPS As Object)
On Error GoTo HJKL
    If IsMDI Then
        With My_MDI
            If XPS.Left < XPS.Width Then
                .Left = XPS.Width + 15 + XPS.Left
            Else
                .Left = 75
            End If
            .Top = XPS.Top
            .Width = MYFORM.Width - XPS.Width - 135
            .Height = XPS.Height
            .Visible = True
        End With
    End If
HJKL:
End Sub
Public Sub SetControlButton(IsControl As Integer, Optional IsEnable As Boolean = True, _
            Optional IsVisible As Boolean = True)
     On Error GoTo JRT
     Select Case IsControl
        Case 0
            MinimizeButton.Enabled = IsEnable
            MinimizeButton.Visible = IsVisible
        Case 1
            MaximizeButton.Enabled = IsEnable
            MaximizeButton.Visible = IsVisible
        Case 2
            CloseButton.Enabled = IsEnable
            CloseButton.Visible = IsVisible
    End Select
JRT:
End Sub

Public Sub SetHelpButton(IsVisible As Boolean)
    MinimizeButton.Visible = Not IsVisible
    MaximizeButton.Visible = Not IsVisible
    HelpButton.Visible = IsVisible
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TitleIcon,TitleIcon,-1,Picture
Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Icon = TitleIcon.Picture
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
On Error GoTo hjk
    Set TitleIcon.Picture = New_Icon
    Set MyTitleIcon.Picture = New_Icon
    Set UserControl.Parent.Icon = TitleIcon.Picture
    PropertyChanged "Icon"
hjk:
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set MyTitleIcon.Picture = PropBag.ReadProperty("Icon", Nothing)
    Caption1.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    Caption1.Caption = PropBag.ReadProperty("Caption", "Osen Kusnadi")
    Caption2.Caption = Caption1.Caption
    m_ShowCloseButton = PropBag.ReadProperty("ShowCloseButton", m_def_ShowCloseButton)
    m_ShowMinimizeButton = PropBag.ReadProperty("ShowMinimizeButton", m_def_ShowMinimizeButton)
    m_ShowMaximizeButton = PropBag.ReadProperty("ShowMaximizeButton", m_def_ShowMaximizeButton)
    SetHelpButton PropBag.ReadProperty("ShowHelpButton", True)
    MinimizeButton.Visible = PropBag.ReadProperty("ShowMinimizeButton", True)
    MaximizeButton.Visible = PropBag.ReadProperty("ShowMaximizeButton", True)
    If MinimizeButton.Enabled Then
        MinimizeButton.Picture = CbMin(0).Picture
    Else
        MinimizeButton.Picture = CbMin(3).Picture
    End If
    If MaximizeButton.Enabled Then
        MaximizeButton.Picture = CbMax(0).Picture
    Else
        MaximizeButton.Picture = CbMax(3).Picture
    End If
    TitleIcon.Visible = PropBag.ReadProperty("ShowTitleIcon", True)
    RePos
    CloseButton.Visible = PropBag.ReadProperty("ShowCloseButton", True)
    m_AutoLoad = PropBag.ReadProperty("AutoLoad", m_def_AutoLoad)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Icon", MyTitleIcon.Picture, Nothing)
    Call PropBag.WriteProperty("ForeColor", Caption1.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("Caption", Caption1.Caption, "Osen Kusnadi")
    Call PropBag.WriteProperty("ShowCloseButton", m_ShowCloseButton, m_def_ShowCloseButton)
    Call PropBag.WriteProperty("ShowMinimizeButton", m_ShowMinimizeButton, m_def_ShowMinimizeButton)
    Call PropBag.WriteProperty("ShowMaximizeButton", m_ShowMaximizeButton, m_def_ShowMaximizeButton)
    Call PropBag.WriteProperty("ShowHelpButton", HelpButton.Visible, True)
    Call PropBag.WriteProperty("ShowMinimizeButton", MinimizeButton.Visible, True)
    Call PropBag.WriteProperty("ShowMaximizeButton", MaximizeButton.Visible, True)
    If MinimizeButton.Enabled Then
        MinimizeButton.Picture = CbMin(0).Picture
    Else
        MinimizeButton.Picture = CbMin(3).Picture
    End If
    If MaximizeButton.Enabled Then
        MaximizeButton.Picture = CbMax(0).Picture
    Else
        MaximizeButton.Picture = CbMax(3).Picture
    End If
    Call PropBag.WriteProperty("ShowTitleIcon", TitleIcon.Visible, True)
    Call PropBag.WriteProperty("ShowCloseButton", CloseButton.Visible, True)
    Call PropBag.WriteProperty("AutoLoad", m_AutoLoad, m_def_AutoLoad)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Caption1,Caption1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Caption1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Caption1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Caption1,Caption1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Caption1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Caption1.Caption() = New_Caption
    Caption2.Caption() = New_Caption
    UserControl.Parent.Caption = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=HelpButton,HelpButton,-1,Enabled
Public Property Get ShowHelpButton() As Boolean
    ShowHelpButton = HelpButton.Visible
End Property

Public Property Let ShowHelpButton(ByVal New_ShoeHelpButton As Boolean)
    SetHelpButton New_ShoeHelpButton
    PropertyChanged "ShoeHelpButton"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MinimizeButton,MinimizeButton,-1,Enabled
Public Property Get ShowMinimizeButton() As Boolean
Attribute ShowMinimizeButton.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    ShowMinimizeButton = MinimizeButton.Visible
End Property

Public Property Let ShowMinimizeButton(ByVal New_ShowMinimizeButton As Boolean)
    MinimizeButton.Visible = New_ShowMinimizeButton
    PropertyChanged "ShowMinimizeButton"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaximizeButton,MaximizeButton,-1,Enabled
Public Property Get ShowMaximizeButton() As Boolean
Attribute ShowMaximizeButton.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    ShowMaximizeButton = MaximizeButton.Visible
End Property

Public Property Let ShowMaximizeButton(ByVal New_ShowMaximizeButton As Boolean)
    MaximizeButton.Visible() = New_ShowMaximizeButton
    PropertyChanged "ShowMaximizeButton"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TitleIcon,TitleIcon,-1,Enabled
Public Property Get ShowTitleIcon() As Boolean
Attribute ShowTitleIcon.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    ShowTitleIcon = TitleIcon.Visible
End Property

Public Property Let ShowTitleIcon(ByVal New_ShowTitleIcon As Boolean)
    TitleIcon.Visible = New_ShowTitleIcon
    RePos
    PropertyChanged "ShowTitleIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=CloseButton,CloseButton,-1,Enabled
Public Property Get ShowCloseButton() As Boolean
Attribute ShowCloseButton.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    ShowCloseButton = CloseButton.Visible
End Property

Public Property Let ShowCloseButton(ByVal New_ShowCloseButton As Boolean)
    CloseButton.Visible = New_ShowCloseButton
    PropertyChanged "ShowCloseButton"
End Property
Public Sub MaxBtnClick()
    MaximizeButton_Click
End Sub


Public Sub SetFormActiveStyle(ActiveForm As Boolean)
    If ActiveForm Then
        Set Title.Picture = Image1.Picture
        Set TitleLeft.Picture = Image2.Picture
        Set TitleRight.Picture = Image3.Picture
        Set Left.Picture = Image4.Picture
        Set Right.Picture = Image5.Picture
        Set BottomLeft.Picture = Image8.Picture
        Set BottomRight.Picture = Image7.Picture
        Set Bottom.Picture = Image6.Picture
        Caption1.ForeColor = vbWhite
        CloseButton.Picture = CbClose(0).Picture
        MinimizeButton.Picture = CbMin(0).Picture
        MaximizeButton.Picture = CbMax(0).Picture
        HelpButton.Picture = CbHelp(0).Picture
        Caption2.Visible = True
    Else
        Caption2.Visible = False
        Set Title.Picture = Image9.Picture
        Set TitleLeft.Picture = Image10.Picture
        Set TitleRight.Picture = Image11.Picture
        Set Left.Picture = Image12.Picture
        Set Right.Picture = Image13.Picture
        Set BottomLeft.Picture = Image16.Picture
        Set BottomRight.Picture = Image15.Picture
        Set Bottom.Picture = Image14.Picture
        CloseButton.Picture = CbClose(3).Picture
        MinimizeButton.Picture = CbMin(4).Picture
        MaximizeButton.Picture = CbMax(4).Picture
        Caption1.ForeColor = vbWhite ' &HE0E0E0   'vbRed '&HE0E0E0
        HelpButton.Picture = CbHelp(3).Picture
    End If
End Sub

Public Sub LoadOsenXPForm(XPName As Object, Optional IsModal As Boolean = False, Optional IsTop As Boolean = False)
    XPName.Left = 0
    XPName.Top = 0
    XPName.LoadXP IsModal, IsTop
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get AutoLoad() As Boolean
    AutoLoad = m_AutoLoad
End Property

Public Property Let AutoLoad(ByVal New_AutoLoad As Boolean)
    m_AutoLoad = New_AutoLoad
    PropertyChanged "AutoLoad"
End Property

