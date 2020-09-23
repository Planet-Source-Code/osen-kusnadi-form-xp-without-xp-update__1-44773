VERSION 5.00
Begin VB.UserControl OsenXPText 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   ScaleHeight     =   375
   ScaleWidth      =   1800
   ToolboxBitmap   =   "osenXPText.ctx":0000
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   600
      ScaleHeight     =   1995
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   210
      Width           =   3855
      Begin VB.TextBox MyTxt 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   1050
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   810
         Width           =   1875
      End
   End
End
Attribute VB_Name = "OsenXPText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum states
    Normal = 0
    Disable = 1
    ReadOnly = 2
End Enum
'Default Property Values:
Const m_def_DataFields = ""
'Property Variables:
Private WithEvents m_DataSource As ADODB.Recordset
Attribute m_DataSource.VB_VarHelpID = -1
Dim m_DataFields As String
'Event Declarations:
Event Change() 'MappingInfo=MyTxt,MyTxt,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event Click() 'MappingInfo=MyTxt,MyTxt,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=MyTxt,MyTxt,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=MyTxt,MyTxt,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=MyTxt,MyTxt,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."

Sub RePos()
On Error Resume Next
UserControl.ScaleMode = 1
    With Pic
        .ScaleMode = 1
        .Left = 0
        .Top = 0
        .Width = UserControl.Width
        .Height = UserControl.Height
        MyTxt.Width = .Width - 90
        MyTxt.Height = .Height - 90
        MyTxt.Left = 45
        MyTxt.Top = 45
    End With
End Sub


Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    MyTxt.SetFocus

End Sub


Private Sub UserControl_Resize()
    RePos
    MyXPtxt MyTxt, Pic, RGB(240, 232, 224), Normal
End Sub

Private Function MyXPtxt(txt As TextBox, Pic As PictureBox, BackColor As ColorConstants, State As states)
    Pic.BackColor = BackColor
    Pic.ScaleMode = 1
    txt.Appearance = 0
    txt.BorderStyle = 0
    Pic.AutoRedraw = True
    Pic.DrawWidth = 1
    Pic.Line (0, 0)-(Pic.Width, 0), RGB(127, 157, 185)
    Pic.Line (0, 0)-(0, Pic.Height), RGB(127, 157, 185)
    Pic.Line (Pic.Width - 15, 0)-(Pic.Width - 15, Pic.Height), RGB(127, 157, 185)
    Pic.Line (0, Pic.Height - 15)-(Pic.Width, Pic.Height - 15), RGB(127, 157, 185)
    
    If State = Normal Then
        txt.BackColor = vbWhite
        txt.Enabled = True
        txt.Locked = False
    ElseIf State = Disable Then
        txt.Enabled = False
        txt.BackColor = RGB(235, 235, 228)
        txt.ForeColor = RGB(161, 161, 146)
    ElseIf State = ReadOnly Then
        txt.Enabled = True
        txt.Locked = True
    End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyTxt,MyTxt,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = MyTxt.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    If New_Alignment > 2 Then New_Alignment = 0
    MyTxt.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

Private Sub MyTxt_Change()
    RaiseEvent Change
End Sub

Private Sub MyTxt_Click()
    RaiseEvent Click
End Sub

Private Sub MyTxt_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyTxt,MyTxt,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = MyTxt.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    MyTxt.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyTxt,MyTxt,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = MyTxt.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set MyTxt.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyTxt,MyTxt,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = MyTxt.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    MyTxt.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Private Sub MyTxt_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyTxt,MyTxt,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = MyTxt.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    MyTxt.Locked() = New_Locked
    Pic.Enabled = Not New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyTxt,MyTxt,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = MyTxt.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    MyTxt.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

Private Sub MyTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyTxt,MyTxt,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
    PasswordChar = MyTxt.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    MyTxt.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyTxt,MyTxt,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = MyTxt.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    MyTxt.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyTxt,MyTxt,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = MyTxt.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    MyTxt.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyTxt,MyTxt,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = MyTxt.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    MyTxt.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyTxt,MyTxt,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = MyTxt.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    MyTxt.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyTxt,MyTxt,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = MyTxt.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    MyTxt.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get DataFields() As String
    DataFields = m_DataFields
End Property

Public Property Let DataFields(ByVal New_DataFields As String)
    m_DataFields = New_DataFields
    MyTxt.DataField = New_DataFields
    PropertyChanged "DataFields"
End Property
'
'Public Sub DataSource(NewRst As ADODB.Recordset)
'    Set MyTxt.DataSource = NewRst
'End Sub
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_DataFields = m_def_DataFields
    MyTxt.Text = Ambient.DisplayName
    UserControl.Height = 330
    MyTxt.FontName = "Verdana"
    UserControl_Resize
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    MyTxt.Alignment = PropBag.ReadProperty("Alignment", 0)
    MyTxt.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    MyTxt.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MyTxt.Font = PropBag.ReadProperty("Font", Ambient.Font)
    MyTxt.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    MyTxt.Locked = PropBag.ReadProperty("Locked", False)
    MyTxt.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    MyTxt.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    MyTxt.SelStart = PropBag.ReadProperty("SelStart", 0)
    MyTxt.SelText = PropBag.ReadProperty("SelText", "")
    MyTxt.SelLength = PropBag.ReadProperty("SelLength", 0)
    MyTxt.Text = PropBag.ReadProperty("Text", "Text1")
    MyTxt.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_DataFields = PropBag.ReadProperty("DataFields", m_def_DataFields)
    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    Set m_DataSource = PropBag.ReadProperty("DataSource", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Alignment", MyTxt.Alignment, 0)
    Call PropBag.WriteProperty("BackColor", MyTxt.BackColor, &H80000005)
    Call PropBag.WriteProperty("Enabled", MyTxt.Enabled, True)
    Call PropBag.WriteProperty("Font", MyTxt.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", MyTxt.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Locked", MyTxt.Locked, False)
    Call PropBag.WriteProperty("MaxLength", MyTxt.MaxLength, 0)
    Call PropBag.WriteProperty("PasswordChar", MyTxt.PasswordChar, "")
    Call PropBag.WriteProperty("SelStart", MyTxt.SelStart, 0)
    Call PropBag.WriteProperty("SelText", MyTxt.SelText, "")
    Call PropBag.WriteProperty("SelLength", MyTxt.SelLength, 0)
    Call PropBag.WriteProperty("Text", MyTxt.Text, "Text1")
    Call PropBag.WriteProperty("ToolTipText", MyTxt.ToolTipText, "")
    Call PropBag.WriteProperty("DataFields", m_DataFields, m_def_DataFields)
    Call PropBag.WriteProperty("DataSource", m_DataSource, Nothing)
End Sub
Public Property Set DataSource(ByVal New_DataSource As Recordset)
    Set m_DataSource = New ADODB.Recordset
    Set m_DataSource = New_DataSource
    Set MyTxt.DataSource = m_DataSource
    MyTxt.DataField = DataFields
    PropertyChanged "DataSource"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Focus() As Boolean
    MyTxt.SelStart = 0
    MyTxt.SelLength = Len(MyTxt)
    MyTxt.SetFocus
End Function

Public Sub RecordPosition(Rs As Recordset)
    If Not (Rs.EOF And Rs.BOF) Then
        MyTxt.Text = "Record : " & Rs.AbsolutePosition & " of " & Rs.RecordCount
    Else
        MyTxt.Text = "No current record"
    End If
End Sub

Public Function GetValue() As Double
    GetValue = Val(MyTxt)
End Function
