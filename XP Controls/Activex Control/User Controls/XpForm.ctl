VERSION 5.00
Begin VB.UserControl XPF 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   2595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4890
   ClipControls    =   0   'False
   EditAtDesignTime=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   173
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   326
   ToolboxBitmap   =   "XpForm.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5100
      Top             =   1830
   End
   Begin VB.Image Cb_CLose 
      Height          =   315
      Index           =   1
      Left            =   2640
      Picture         =   "XpForm.ctx":0312
      Top             =   2790
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Cb_CLose 
      Height          =   315
      Index           =   2
      Left            =   2970
      Picture         =   "XpForm.ctx":0894
      Top             =   2790
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Cb_CLose 
      Height          =   315
      Index           =   3
      Left            =   3300
      Picture         =   "XpForm.ctx":0E16
      Top             =   2790
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image8 
      Height          =   315
      Left            =   5700
      Picture         =   "XpForm.ctx":1398
      Top             =   1140
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   5700
      Picture         =   "XpForm.ctx":19D2
      Top             =   1470
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image7 
      Height          =   315
      Left            =   5340
      Picture         =   "XpForm.ctx":200C
      Top             =   810
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image6 
      Height          =   315
      Left            =   5700
      Picture         =   "XpForm.ctx":2646
      Top             =   810
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image5 
      Height          =   315
      Left            =   6060
      Picture         =   "XpForm.ctx":2C80
      ToolTipText     =   "close"
      Top             =   810
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image4 
      Height          =   315
      Left            =   5340
      Picture         =   "XpForm.ctx":32BA
      Top             =   480
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   5700
      Picture         =   "XpForm.ctx":372D
      Top             =   480
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   6060
      Picture         =   "XpForm.ctx":3BA1
      ToolTipText     =   "close"
      Top             =   480
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Caption1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Osen Kusnadi"
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
      Left            =   420
      TabIndex        =   0
      Top             =   90
      Width           =   1185
   End
   Begin VB.Label Caption2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Osen Kusnadi"
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
      Left            =   390
      TabIndex        =   1
      Top             =   90
      Width           =   1335
   End
   Begin VB.Image TitleIcon 
      Height          =   255
      Left            =   120
      Picture         =   "XpForm.ctx":401F
      Stretch         =   -1  'True
      Top             =   90
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image CloseButton 
      Height          =   315
      Left            =   4440
      Picture         =   "XpForm.ctx":43A9
      ToolTipText     =   "Close"
      Top             =   60
      Width           =   315
   End
   Begin VB.Image MaximizeButton 
      Height          =   315
      Left            =   4110
      Picture         =   "XpForm.ctx":492B
      ToolTipText     =   "Maximize"
      Top             =   60
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image MinimizeButton 
      Height          =   315
      Left            =   3780
      Picture         =   "XpForm.ctx":4F65
      ToolTipText     =   "Minimize"
      Top             =   60
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Title 
      Height          =   450
      Left            =   180
      Picture         =   "XpForm.ctx":559F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image TitleLeft 
      Height          =   450
      Left            =   30
      Picture         =   "XpForm.ctx":58E3
      Top             =   0
      Width           =   150
   End
   Begin VB.Image TitleRight 
      Height          =   450
      Left            =   4740
      Picture         =   "XpForm.ctx":5CDA
      Top             =   0
      Width           =   150
   End
   Begin VB.Image Left 
      Height          =   2085
      Left            =   30
      Picture         =   "XpForm.ctx":60DA
      Stretch         =   -1  'True
      Top             =   450
      Width           =   60
   End
   Begin VB.Image Right 
      Height          =   2085
      Left            =   4830
      Picture         =   "XpForm.ctx":6408
      Stretch         =   -1  'True
      Top             =   450
      Width           =   60
   End
   Begin VB.Image Bottom 
      Height          =   60
      Left            =   90
      Picture         =   "XpForm.ctx":6736
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   4755
   End
   Begin VB.Image BottomRight 
      Height          =   60
      Left            =   4830
      Picture         =   "XpForm.ctx":6A64
      Top             =   2520
      Width           =   60
   End
   Begin VB.Image BottomLeft 
      Height          =   60
      Left            =   30
      Picture         =   "XpForm.ctx":6D9E
      Top             =   2520
      Width           =   60
   End
End
Attribute VB_Name = "XPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'/**************** Declare API Function **************************************************************************
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

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

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Dim bTransparent As Boolean
Private WithEvents MYFORM As Form
Attribute MYFORM.VB_VarHelpID = -1
Public Event Closed()
 Const MF_BYPOSITION = &H400&
 Const MF_BYCOMMAND = 0
 Const SC_RESTORE = &HF120
 Const SC_MOVE = &HF010
 Const SC_SIZE = &HF000
 Const SC_MINIMIZE = &HF020
 Const SC_MAXIMIZE = &HF030
 Const SC_CLOSE = &HF060
 Const WM_GETSYSMENU = &H313


Const GWL_STYLE = (-16)
Const WS_SYSMENU = &H80000
Public Sub RePos()
    'This repositions the different controls on the form when it is resized
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
    
    With MinimizeButton
        .Left = MaximizeButton.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
    End With
    
    'Icon
    With TitleIcon
        .Left = Left.Left + Left.Width + 2
        .Top = (Title.Height - .Height) / 2.5
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

Private Sub Caption1_Change()
    Caption2.Caption = Caption1.Caption
End Sub

Private Sub CloseButton_Click()
On Error GoTo EF
    Unload UserControl.Parent
EF:
End Sub

Private Sub CloseButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        CloseButton.Picture = Cb_CLose(3).Picture
    Else
        CloseButton.Picture = Cb_CLose(2).Picture
    End If
End Sub

Private Sub MYFORM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CloseButton.Picture = Cb_CLose(1).Picture
    
End Sub

Private Sub Title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error GoTo H
    CloseButton.Picture = Cb_CLose(1).Picture
H:
End Sub

Private Sub UserControl_Initialize()
    bTransparent = False  'So we do not set the corners transparent while still in design mode
    RePos   'Reposition
End Sub

Private Sub UserControl_Resize()
    RePos   'Reposition
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
    DefaultBackgroundColor = &HEAF1F1   'Returns a common off-white Windows XP color
End Function
Public Sub LoadXP()
Dim IpForm As Form
    TitleIcon.Visible = False
    Set IpForm = UserControl.Parent
    Set MYFORM = IpForm
    Caption1.Caption = IpForm.Caption
    IpForm.Width = UserControl.Width
    IpForm.Height = UserControl.Height
    If IpForm.BorderStyle <> 0 Then
        IpForm.Height = UserControl.Height + 375
    End If
    SetStyle IpForm
    UserControl.Width = IpForm.Width    ':   XP_Name.Top = 0
    UserControl.Height = IpForm.Height  ': XP_Name.Left = 0
    ReTransObj IpForm
    DoEvents
    IpForm.Hide
End Sub
Private Sub ReTransObj(IpObject As Object)
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
Private Sub SetStyle(IpForm As Object)
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
    SetWindowLong IpForm.hWnd, GWL_STYLE, lCurrentSettings
    SetWindowPos IpForm.hWnd, 0, IpForm.Left / 15, IpForm.Top / 15, (IpForm.Width / 15), (IpForm.Height / 15), &H40
    If IpForm.BorderStyle <> 0 Then
    IpForm.Height = IpForm.Height - 365
    End If
    IpForm.Left = (Screen.Width - IpForm.Width) / 2
    IpForm.Top = (Screen.Height - IpForm.Height) / 2
End Sub

Private Sub UserControl_Terminate()
    Set MYFORM = Nothing
End Sub
