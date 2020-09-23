VERSION 5.00
Begin VB.UserControl Touch2 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1080
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   PropertyPages   =   "Touch2.ctx":0000
   ScaleHeight     =   810
   ScaleWidth      =   1080
   ToolboxBitmap   =   "Touch2.ctx":0023
   Begin VB.PictureBox picBut 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   1680
      Picture         =   "Touch2.ctx":0335
      ScaleHeight     =   420
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   525
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picBut 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   1080
      Picture         =   "Touch2.ctx":0553
      ScaleHeight     =   420
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   525
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picBut 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   1680
      Picture         =   "Touch2.ctx":0772
      ScaleHeight     =   420
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picBut 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   1095
      Picture         =   "Touch2.ctx":0987
      ScaleHeight     =   420
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H8000000A&
      Caption         =   "cmdSelect"
      Height          =   795
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
End
Attribute VB_Name = "Touch2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Option Compare Text

'////////////////////////////////////////////////////
'Touch Controls
'
'References Microsoft activeX data objects 2.0
'
'Author: Samuel Sunadh ; sunadh@yahoo.com
'////////////////////////////////////////////////////


Dim dbCon As New adodb.Connection
Dim rsGbl As New adodb.Recordset
Dim intPgCount As Integer
Dim intPgNo As Integer
Dim UWidth As Single
Dim UHeight As Single

Public Enum Touch2Styles
    Horizontal = 0
    Vertical = 1
End Enum

'Default Property Values:
Const m_def_Defaultclick = False
Const m_def_DefaultClickButtonNo = 0
Const m_def_ButDistance = 80
Const m_def_ButtonStyle = 0
Const m_def_TotalButtons = 0
Const m_def_ButHeight = 795
Const m_def_ButWidth = 1065
Const m_def_Code = ""
Const m_def_ConString = ""
Const m_def_SqlStr = ""
Const m_def_Description = ""
Const m_ButStyle = 0
Const m_def_ButtonsTotal = 0

'Property Variables:
Dim m_ButDistance As Long
Dim m_Defaultclick As Boolean
Dim m_DefaultClickButtonNo As Integer
Dim m_ButtonStyle As Byte
Dim m_TotalButtons As Integer
Dim m_ButHeight As Long
Dim m_ButWidth As Long
Dim m_Code As Variant
Dim m_Constring As String
Dim m_sqlStr As String
Dim m_Description As String
'Event Declarations:
Event Click() 'MappingInfo=cmdSelect(0),cmdSelect,0,Click
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

Private Sub UserControl_AmbientChanged(PropertyName As String)
    Refresh
End Sub


Private Sub UserControl_Resize()
    With UserControl
        If .Height < UHeight Then .Height = UHeight
        If .Width < UWidth Then .Width = UWidth
    End With
End Sub

Private Sub UserControl_Show()
    If m_Constring = vbNullString And m_sqlStr = vbNullString Then
    Call FillButtons(m_TotalButtons, m_ButtonStyle, m_ButDistance)
    End If
End Sub

Private Function initConnection(ConStr As String, LSqlstr As String)
Dim i As Long
On Error GoTo Err_InitConnection
    If Not ConStr = vbNullString And Not LSqlstr = vbNullString Then
    Call FillButtons(m_TotalButtons, m_ButtonStyle, m_ButDistance)
    With dbCon
           If .State = 1 Then .Close
            .Open ConStr
        End With
        With rsGbl
            .CursorLocation = adUseClient
            .Open LSqlstr, dbCon, adOpenKeyset, adLockReadOnly, adCmdText
            If Not (.EOF And .BOF) Then
                i = m_TotalButtons
                .PageSize = IIf(i = 0, 1, i)
                intPgCount = .PageCount
                intPgNo = 1
            End If
        End With
        Call fillForm
    End If
Exit Function
Err_InitConnection:
End Function

Private Function UpdPageNo(inNum As Integer)
    If inNum = 0 Then
        intPgNo = intPgNo - 1
        If intPgNo < 1 Then intPgNo = 1
    Else
        intPgNo = intPgNo + 1
        If intPgNo > intPgCount Then intPgNo = intPgCount
    End If
End Function


Private Sub unloadObj(C As Object)
Dim i As Long
    If C.Count > 1 Then
    For i = 1 To C.Count - 1
        Unload C(i)
    Next i
    End If
End Sub

Private Function fillForm() As Long
Dim k As Long, i As Integer, intPgSize As Integer, TotFields As Integer
On Error Resume Next
With rsGbl
    If .State = 1 Then
        If Not (.EOF And .BOF) Then
            .AbsolutePage = intPgNo
            intPgSize = .PageSize
            k = 1
            For i = k To m_TotalButtons
                cmdSelect(i).Visible = False
            Next i
            TotFields = .Fields.Count - 1
            Do While Not .EOF And intPgSize > 0
                cmdSelect(k).Picture = LoadPicture()
                cmdSelect(k).Caption = Trim(.Fields(1)) & ""
                cmdSelect(k).Tag = .Fields(0)
                If TotFields = 2 Then
                    cmdSelect(k).BackColor = IIf(.Fields(2) > 0, .Fields(2), &H8000000F)
                ElseIf TotFields = 3 Then
                    cmdSelect(k).BackColor = IIf(.Fields(2) > 0, .Fields(2), &H8000000F)
                    If Not Dir$(.Fields(3) & "") = "" Then
                        Select Case Right(.Fields(3), 3)
                            Case Is = "bmp", "dib", "gif", "jpg", "wmf", "emf", "ico", "cur"
                            cmdSelect(k).Picture = LoadPicture(.Fields(3))
                        End Select
                    End If
                End If
                cmdSelect(k).Visible = True
                intPgSize = intPgSize - 1
                k = k + 1
                .MoveNext
                If .EOF Then Exit Do
            Loop
        Else
            For i = 0 To m_TotalButtons + 1
                cmdSelect(i).Visible = False
            Next i
        End If
    End If
End With

If Defaultclick = True Then
    If m_DefaultClickButtonNo > 0 And m_DefaultClickButtonNo <= m_TotalButtons Then
        cmdSelect(m_DefaultClickButtonNo).Value = True
    End If
End If

On Error GoTo 0

End Function


Private Sub FillButtons(ByVal mButTotal As Integer, Ctlstyle As Variant, ButGape As Long)
Dim i As Integer, Tot As Integer
UWidth = 0: UHeight = 0
UserControl.Width = cmdSelect(0).Width: UserControl.Height = cmdSelect(0).Height
Call unloadObj(cmdSelect)
    For i = 0 To mButTotal + 1
        If i = 0 Then
            cmdSelect(i).Width = m_ButWidth
            cmdSelect(i).Height = m_ButHeight
            cmdSelect(i).BackColor = &H8000000A
            cmdSelect(i).Caption = IIf(Ctlstyle = 0, "Prev Page", "Page Up")
            cmdSelect(i).Picture = IIf(Ctlstyle = 0, picBut(0).Picture, picBut(2).Picture)
        ElseIf i = mButTotal + 1 Then
            Load cmdSelect(i)
            cmdSelect(i).Caption = IIf(Ctlstyle = 0, "Next Page", "Page Down")
            cmdSelect(i).Picture = IIf(Ctlstyle = 0, picBut(1).Picture, picBut(3).Picture)
            If Ctlstyle = 0 Then 'Horizontal
                cmdSelect(i).Left = cmdSelect(i - 1).Left + cmdSelect(i - 1).Width + ButGape
                UserControl.Width = cmdSelect(i).Left + cmdSelect(i).Width
            Else 'Verticle
                cmdSelect(i).Left = cmdSelect(i - 1).Left
                cmdSelect(i).Top = cmdSelect(i - 1).Top + cmdSelect(i - 1).Height + ButGape
                UserControl.Height = cmdSelect(i).Top + cmdSelect(i).Height
                UserControl.Width = cmdSelect(i).Left + cmdSelect(i).Width
            End If
        Else
            Load cmdSelect(i)
            If m_Constring = vbNullString And m_sqlStr = vbNullString Then cmdSelect(i).Caption = "Button" & i
            cmdSelect(i).Picture = UserControl.Picture
            If Ctlstyle = 0 Then 'Horizontal
                cmdSelect(i).Left = cmdSelect(i - 1).Left + cmdSelect(i - 1).Width + ButGape
                UserControl.Width = cmdSelect(i).Left + cmdSelect(i).Width
            Else 'Verticle
                cmdSelect(i).Left = cmdSelect(i - 1).Left
                cmdSelect(i).Top = cmdSelect(i - 1).Top + cmdSelect(i - 1).Height + ButGape
                UserControl.Height = cmdSelect(i - 1).Top + cmdSelect(i - 1).Height
                UserControl.Width = cmdSelect(i).Left + cmdSelect(i).Width
            End If
        End If
        cmdSelect(i).Visible = True
    Next i
        UWidth = UserControl.Width
        UHeight = cmdSelect(i - 1).Top + cmdSelect(i - 1).Height + 10
End Sub


Private Sub UserControl_Terminate()
    Call unloadObj(cmdSelect)
    If rsGbl.State = 1 Then rsGbl.Close
    Set rsGbl = Nothing
    If dbCon.State = 1 Then dbCon.Close
    Set dbCon = Nothing
End Sub


Private Sub cmdSelect_Click(Index As Integer)
    If Index = 0 Or Index = m_TotalButtons + 1 Then
        Call UpdPageNo(Index)
        Call fillForm
        If Defaultclick = False Then
            Code = ""
            Description = ""
        End If
    Else
        Code = cmdSelect(Index).Tag
        Description = cmdSelect(Index).Caption
    End If
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get ButtonStyle() As Touch2Styles
    ButtonStyle = m_ButtonStyle
End Property

Public Property Let ButtonStyle(ByVal New_ButtonStyle As Touch2Styles)
    m_ButtonStyle = New_ButtonStyle
    PropertyChanged "ButtonStyle"
    Call FillButtons(m_TotalButtons, m_ButtonStyle, m_ButDistance)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get TotalButtons() As Integer
    TotalButtons = m_TotalButtons
End Property

Public Property Let TotalButtons(ByVal New_TotalButtons As Integer)
    m_TotalButtons = New_TotalButtons
    PropertyChanged "TotalButtons"
    Call FillButtons(m_TotalButtons, m_ButtonStyle, m_ButDistance)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,795
Public Property Get ButHeight() As Long
    ButHeight = m_ButHeight
End Property

Public Property Let ButHeight(ByVal New_ButHeight As Long)
    m_ButHeight = New_ButHeight
    PropertyChanged "ButHeight"
    Call FillButtons(m_TotalButtons, m_ButtonStyle, m_ButDistance)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0, 1065
Public Property Get ButWidth() As Long
    ButWidth = m_ButWidth
End Property

Public Property Let ButWidth(ByVal New_ButWidth As Long)
    m_ButWidth = New_ButWidth
    PropertyChanged "ButWidth"
    Call FillButtons(m_TotalButtons, m_ButtonStyle, m_ButDistance)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Code() As Variant
    Code = m_Code
End Property

Public Property Let Code(ByVal New_Code As Variant)
    m_Code = New_Code
    PropertyChanged "Code"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ConString() As String
    ConString = m_Constring
End Property

Public Property Let ConString(ByVal New_Constring As String)
    m_Constring = New_Constring
    PropertyChanged "Constring"
    Call initConnection(m_Constring, m_sqlStr)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get sqlStr() As String
    sqlStr = m_sqlStr
End Property

Public Property Let sqlStr(ByVal New_sqlStr As String)
    m_sqlStr = New_sqlStr
    PropertyChanged "Sqlstr"
    Call initConnection(m_Constring, m_sqlStr)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(ByVal New_Description As String)
    m_Description = New_Description
    PropertyChanged "Description"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ButtonStyle = m_def_ButtonStyle
    m_TotalButtons = m_def_TotalButtons
    m_ButHeight = m_def_ButHeight
    m_ButWidth = m_def_ButWidth
    m_Code = m_def_Code
    m_Constring = m_def_ConString
    m_sqlStr = m_def_SqlStr
    m_Description = m_def_Description
    m_Defaultclick = m_def_Defaultclick
    m_DefaultClickButtonNo = m_def_DefaultClickButtonNo
    m_ButDistance = m_def_ButDistance
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set cmdSelect(0).Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ButtonStyle = PropBag.ReadProperty("ButtonStyle", m_def_ButtonStyle)
    m_TotalButtons = PropBag.ReadProperty("TotalButtons", m_def_TotalButtons)
    m_ButHeight = PropBag.ReadProperty("ButHeight", m_def_ButHeight)
    cmdSelect(0).Height = m_ButHeight
    m_ButWidth = PropBag.ReadProperty("ButWidth", m_def_ButWidth)
    cmdSelect(0).Width = m_ButWidth
    m_Code = PropBag.ReadProperty("Code", m_def_Code)
    m_Constring = PropBag.ReadProperty("Constring", m_def_ConString)
    m_sqlStr = PropBag.ReadProperty("Sqlstr", m_def_SqlStr)
    m_Description = PropBag.ReadProperty("Description", m_def_Description)
    m_Defaultclick = PropBag.ReadProperty("Defaultclick", m_def_Defaultclick)
    m_DefaultClickButtonNo = PropBag.ReadProperty("DefaultClickButtonNo", m_def_DefaultClickButtonNo)
    m_ButDistance = PropBag.ReadProperty("ButDistance", m_def_ButDistance)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("FontSize", cmdSelect(0).FontSize, 0)
    Call PropBag.WriteProperty("FontItalic", cmdSelect(0).FontItalic, 0)
    Call PropBag.WriteProperty("FontBold", cmdSelect(0).FontBold, 0)
    Call PropBag.WriteProperty("Font", cmdSelect(0).Font, Ambient.Font)
    Call PropBag.WriteProperty("ButtonStyle", m_ButtonStyle, m_def_ButtonStyle)
    Call PropBag.WriteProperty("TotalButtons", m_TotalButtons, m_def_TotalButtons)
    Call PropBag.WriteProperty("ButHeight", m_ButHeight, m_def_ButHeight)
    Call PropBag.WriteProperty("ButWidth", m_ButWidth, m_def_ButWidth)
    Call PropBag.WriteProperty("Code", m_Code, m_def_Code)
    Call PropBag.WriteProperty("Constring", m_Constring, m_def_ConString)
    Call PropBag.WriteProperty("Sqlstr", m_sqlStr, m_def_SqlStr)
    Call PropBag.WriteProperty("Description", m_Description, m_def_Description)
    Call PropBag.WriteProperty("Defaultclick", m_Defaultclick, m_def_Defaultclick)
    Call PropBag.WriteProperty("DefaultClickButtonNo", m_DefaultClickButtonNo, m_def_DefaultClickButtonNo)
    Call PropBag.WriteProperty("ButDistance", m_ButDistance, m_def_ButDistance)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get Defaultclick() As Boolean
    Defaultclick = m_Defaultclick
End Property

Public Property Let Defaultclick(ByVal New_Defaultclick As Boolean)
    m_Defaultclick = New_Defaultclick
    PropertyChanged "Defaultclick"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get DefaultClickButtonNo() As Integer
    DefaultClickButtonNo = m_DefaultClickButtonNo
End Property

Public Property Let DefaultClickButtonNo(ByVal New_DefaultClickButtonNo As Integer)
    m_DefaultClickButtonNo = New_DefaultClickButtonNo
    PropertyChanged "DefaultClickButtonNo"
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdSelect(0),cmdSelect,0,Font
Public Property Get Font() As Font
    Set Font = cmdSelect(0).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set cmdSelect(0).Font = New_Font
    PropertyChanged "Font"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ButDistance() As Long
    ButDistance = m_ButDistance
End Property

Public Property Let ButDistance(ByVal New_ButDistance As Long)
    m_ButDistance = New_ButDistance
    PropertyChanged "ButDistance"
    Call FillButtons(m_TotalButtons, m_ButtonStyle, m_ButDistance)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

