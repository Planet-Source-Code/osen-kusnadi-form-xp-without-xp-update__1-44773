VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Osen_WinXPCntrl"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub Class_Initialize()
    msgBackColor = &HD8E9EC
End Sub
Public Function InputBoxXP(strMessage As String, Optional strTitle As String) As String
On Error GoTo EDF
    If strTitle = "" Then
        strTitle = My_Project_Name
    End If
    If strTitle = "" Then
        strTitle = App.Title
    End If
    s_Message = strMessage
    s_Title = strTitle
    FrmInput.Show 1
    InputBoxXP = ResponInput
    Exit Function
EDF:
    InputBoxXP = InputBox(strMessage, strTitle)
End Function

Public Function ShowMessage(strMessage As String, Optional strTitle As String, _
    Optional IpType As Integer = 0, _
    Optional IpIcon As Integer = 0, _
    Optional IpSet As Boolean = False, _
    Optional IpObj As Object = Nothing) As Boolean
On Error GoTo S
    If strTitle = "" Then
        strTitle = My_Project_Name
    End If
    If strTitle = "" Then
        strTitle = App.Title
    End If

    s_Title = strTitle
    s_Message = strMessage
    MyType = IpType
    If IpIcon > 0 Then
        IMGS = True
    Else
        IMGS = False
    End If
    If IpIcon > 72 Then IpIcon = 1
    If IpIcon > 0 Then frm_Message.IMG.Picture = frmImages.IMG.ListImages(IpIcon).Picture
    SetDefault = IpSet
    Set IsDefPic = IpObj
    frm_Message.Show 1
    ShowMessage = ResponMsg
    Exit Function
S:
End Function
Public Sub MsgInformation(StrInfo As String, Optional strTitle As String)
    On Error GoTo H
    If strTitle = "" Then
        strTitle = My_Project_Name
    End If
    If strTitle = "" Then
        strTitle = "Information"
    End If
    s_Title = strTitle
    s_Message = StrInfo
    MyType = 0
    IMGS = True
    frm_Message.IMG.Picture = frmImages.IMG.ListImages(63).Picture
    frm_Message.Show 1
    Exit Sub
H:
    MsgBox StrInfo, vbInformation, strTitle
End Sub
Public Sub MsgCritical(StrInfo As String, Optional strTitle As String)
   On Error GoTo H
    If strTitle = "" Then
        strTitle = My_Project_Name
    End If
    If strTitle = "" Then
        strTitle = "Error"
    End If
   s_Title = strTitle
    s_Message = StrInfo
    MyType = 0
    IMGS = True
    frm_Message.IMG.Picture = frmImages.IMG.ListImages(62).Picture
    frm_Message.Show 1
    Exit Sub
H:
    MsgBox StrInfo, vbCritical, strTitle
End Sub
Public Sub MsgExclamation(StrInfo As String, Optional strTitle As String)
    On Error GoTo H
    If strTitle = "" Then
        strTitle = My_Project_Name
    End If
    If strTitle = "" Then
        strTitle = "Exlamation"
    End If
    s_Title = strTitle
    s_Message = StrInfo
    MyType = 0
    IMGS = True
    frm_Message.IMG.Picture = frmImages.IMG.ListImages(61).Picture
    frm_Message.Show 1
    Exit Sub
H:
    MsgBox StrInfo, vbExclamation, strTitle
End Sub
Public Function MsgQuestion(StrInfo As String, Optional strTitle As String) As Boolean
  On Error GoTo H
    If strTitle = "" Then
        strTitle = My_Project_Name
    End If
    If strTitle = "" Then
        strTitle = "Question"
    End If
    s_Title = strTitle
    s_Message = StrInfo
    MyType = 1
    IMGS = True
    frm_Message.IMG.Picture = frmImages.IMG.ListImages(64).Picture
    frm_Message.Show 1
    MsgQuestion = ResponMsg
    Exit Function
H:
    If MsgBox(s_Message, vbQuestion + vbYesNo, s_Title) = vbYes Then
        MsgQuestion = True
    Else
        MsgQuestion = False
    End If
End Function
Public Function MsgDeleteRecord() As Boolean
On Error GoTo H
    s_Title = "Delete record"
    s_Message = "Are you sure you want to " & vbLf & "delete current record ?"
    MyType = 1
    IMGS = True
    frm_Message.IMG.Picture = frmImages.IMG.ListImages(64).Picture
    frm_Message.Show 1
    MsgDeleteRecord = ResponMsg
    Exit Function
H:
    If MsgBox(s_Message, vbQuestion + vbYesNo, s_Title) = vbYes Then
        MsgDeleteRecord = True
    Else
        MsgDeleteRecord = False
    End If
End Function
Public Function MsgExitDatabase() As Boolean
On Error GoTo H
    s_Title = "Exit system"
    s_Message = "Really Exit Database ?"
    MyType = 1
    IMGS = True
    frm_Message.IMG.Picture = frmImages.IMG.ListImages(64).Picture
    frm_Message.Show 1
    MsgExitDatabase = ResponMsg
    Exit Function
H:
    If MsgBox(s_Message, vbQuestion + vbYesNo, s_Title) = vbYes Then
        MsgExitDatabase = True
    Else
        MsgExitDatabase = False
    End If
End Function

Public Sub ChangeMsgBackColor(NewColor As Long)
    msgBackColor = NewColor
End Sub
Public Sub ShowAboutMe(Optional IpTitle As String = "Osen XP Components", _
            Optional IpDevelover As String = "Osen Kusnadi", _
            Optional IpVersion As String = "NoVersion", _
            Optional IpEmail As String = "Email : okusnadi@cikarang.actaris.com", _
            Optional IpPhone As String = "Contact Number : 62-21-8970270 ext : 254 ")
On Error Resume Next
    With Frm_About
        IsDevelover = IpDevelover
        IsVersion = IpVersion
        IsEmail = IpEmail
        IsTitle = IpTitle
        IsPhone = IpPhone
        .Show 1
    End With
End Sub

'/= Get Standard Picture ======================================================
Public Function GetMyIcon(ImageIndex) As StdPicture
On Error Resume Next
    If ImageIndex > 0 And ImageIndex <= frmImages.IMG.ListImages.Count Then
        Set GetMyIcon = frmImages.IMG.ListImages(ImageIndex).Picture
    Else
        Set GetMyIcon = frmImages.IMG.ListImages(64).Picture
    End If
End Function
Public Sub MsgError(ErrNumber As Long, ErrDesc As String)
    MsgCritical "Error number : " & ErrNumber & vbLf & _
                ErrDesc, "Error Description"
End Sub
Public Function GetMyBackgroundXP() As StdPicture
        Set GetMyBackgroundXP = frmImages.IMG.ListImages("bgxp1").Picture
End Function



