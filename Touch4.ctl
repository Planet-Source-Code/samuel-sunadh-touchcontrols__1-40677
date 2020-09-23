VERSION 5.00
Begin VB.UserControl Touch4 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1065
   LockControls    =   -1  'True
   PropertyPages   =   "Touch4.ctx":0000
   ScaleHeight     =   1710
   ScaleWidth      =   1065
   ToolboxBitmap   =   "Touch4.ctx":002B
   Begin VB.PictureBox picBut 
      Height          =   480
      Index           =   3
      Left            =   1740
      Picture         =   "Touch4.ctx":033D
      ScaleHeight     =   420
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   465
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picBut 
      Height          =   480
      Index           =   2
      Left            =   1155
      Picture         =   "Touch4.ctx":05DA
      ScaleHeight     =   420
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   465
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picBut 
      Height          =   480
      Index           =   1
      Left            =   1740
      Picture         =   "Touch4.ctx":07EF
      ScaleHeight     =   420
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picBut 
      Height          =   480
      Index           =   0
      Left            =   1155
      Picture         =   "Touch4.ctx":0A03
      ScaleHeight     =   420
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdNavigate 
      BackColor       =   &H8000000A&
      Caption         =   "&First"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "Prev"
      Top             =   915
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H8000000A&
      Caption         =   "cmdSelect"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
Attribute VB_Name = "Touch4"
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

'Default Property Values:
Const m_def_ButDistance = 80
Const m_def_ButHeight = 795
Const m_def_ButWidth = 1065
Const m_def_Code = ""
Const m_def_ConString = ""
Const m_def_ExtraBut = 0
Const m_def_Rows = 0
Const m_def_Columns = 0
Const m_def_SqlStr = ""
Const m_def_Description = ""

'Property Variables:
Dim m_ButDistance As Long
Dim m_ButHeight As Long
Dim m_ButWidth As Long
Dim m_Code As Variant
Dim m_Constring As String
Dim m_ExtraBut As Integer
Dim m_Rows As Integer
Dim m_Columns As Integer
Dim m_sqlStr As String
Dim m_Description As String
'Event Declarations:
Event Click() 'MappingInfo=cmdSelect(0),cmdSelect,0,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp


Private Sub cmdNavigate_Click(Index As Integer)
    Call UpdPageNo(Index)
    Call fillForm
End Sub

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
        Call FillButtons(m_Rows, m_Columns, m_ExtraBut, m_ButDistance)
    End If
End Sub

Private Function initConnection(ConStr As String, LSqlstr As String)
Dim i As Long
On Error GoTo Err_InitConnection
    If Not ConStr = vbNullString And Not LSqlstr = vbNullString Then
        Call FillButtons(m_Rows, m_Columns, m_ExtraBut, m_ButDistance)
    With dbCon
           If .State = 1 Then .Close
            .Open ConStr
        End With
        With rsGbl
            .CursorLocation = adUseClient
            .Open LSqlstr, dbCon, adOpenKeyset, adLockReadOnly, adCmdText
            If Not (.EOF And .BOF) Then
                i = (m_Rows * m_Columns) + m_ExtraBut
                .PageSize = IIf(i = 0, 1, i)
                intPgCount = .PageCount
                intPgNo = 1
            End If
        End With
        Call fillForm
    End If
Exit Function
Err_InitConnection:
Exit Function
End Function


Private Function UpdPageNo(inNum As Integer)
      If inNum = 0 Then
          intPgNo = 1
      ElseIf inNum = 1 Then
          intPgNo = intPgNo - 1
          If intPgNo < 1 Then intPgNo = 1
      ElseIf inNum = 2 Then
          intPgNo = intPgNo + 1
          If intPgNo > intPgCount Then intPgNo = intPgCount
      ElseIf inNum = 3 Then
          intPgNo = intPgCount
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
With rsGbl
    If .State = 1 Then
        If Not (.EOF And .BOF) Then
            .AbsolutePage = intPgNo
            intPgSize = .PageSize
            
            For i = k To (m_Columns * m_Rows) + m_ExtraBut - 1
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
            For i = 0 To 3
                cmdNavigate(i).Visible = False
            Next i
            For i = k To (m_Columns * m_Rows) + m_ExtraBut - 1
                cmdSelect(i).Visible = False
            Next i
        End If
    End If
End With
End Function


Private Sub FillButtons(ByVal mRows As Integer, ByVal mCols As Integer, ByVal mEBut As Integer, ButGape As Long)
Dim i As Integer, j As Integer, k As Long, l As Integer, M As Long
Dim Tot As Integer
Call unloadObj(cmdSelect)
UWidth = 0: UHeight = 0
UserControl.Width = cmdSelect(0).Width
UserControl.Height = cmdSelect(0).Height
mRows = IIf(mRows = 0, 1, mRows)
mCols = IIf(mCols = 0, 1, mCols)
Tot = ((mRows * mCols) + mEBut) - 1
    For i = 0 To Tot
        k = j
        If l = 0 Then
            cmdSelect(k).Width = m_ButWidth
            cmdSelect(k).Height = m_ButHeight
            cmdSelect(k).BackColor = &H8000000A
        Else
            Load cmdSelect(k)
            If l > mCols - 1 Then
                l = 0
                cmdSelect(k).Left = cmdSelect(l).Left
                cmdSelect(k).Top = M + cmdSelect(l).Height + ButGape
                UserControl.Width = cmdSelect(k).Left + cmdSelect(k).Width
            Else
                cmdSelect(k).Left = cmdSelect(k - 1).Left + cmdSelect(k - 1).Width + ButGape
                cmdSelect(k).Top = cmdSelect(k - 1).Top
                UWidth = (cmdSelect(k).Width * mCols) + (ButGape * mCols)
                UserControl.Width = (cmdSelect(k).Width * mCols) + (ButGape * mCols)
            End If
            'cmdSelect(K).Caption = "Button" & i + 1
        End If
        If m_Constring = vbNullString And m_sqlStr = vbNullString Then cmdSelect(k).Caption = "Button" & i + 1
        cmdSelect(k).Visible = True
        M = cmdSelect(k).Top: l = l + 1: j = j + 1
    Next i
    If mEBut > 0 Then
         Call addNavigationButtons(cmdSelect(k).Left + cmdSelect(k).Width + ButGape, cmdSelect(k).Top, m_ButDistance)
    Else
        m_Columns = IIf(m_Columns = 0, 1, m_Columns)
        Call addNavigationButtons(cmdSelect(i - m_Columns).Left, cmdSelect(i - m_Columns).Top + cmdSelect(i - m_Columns).Height + ButGape, m_ButDistance)
    End If
End Sub

Private Function addNavigationButtons(LastButLeft As Long, LastButTop, ButGape As Long)
Dim i As Integer
Call unloadObj(cmdNavigate)
For i = 0 To 3
    If i = 0 Then
        cmdNavigate(i).Width = m_ButWidth
        cmdNavigate(i).Height = m_ButHeight
        cmdNavigate(i).Left = LastButLeft
        cmdNavigate(i).Top = LastButTop
    Else
        Load cmdNavigate(i)
        cmdNavigate(i).Left = cmdNavigate(i - 1).Left + cmdNavigate(i - 1).Width + ButGape
    End If
    cmdNavigate(i).Visible = True
    If i = 0 Then
        cmdNavigate(i).Caption = "&First Page"
    ElseIf i = 1 Then
        cmdNavigate(i).Caption = "&Prev Page"
    ElseIf i = 2 Then
        cmdNavigate(i).Caption = "&Next Page"
    ElseIf i = 3 Then
        cmdNavigate(i).Caption = "&Last Page"
    End If
    cmdNavigate(i).Picture = picBut(i).Picture
Next i
    UserControl.Height = cmdNavigate(i - 1).Top + cmdNavigate(i - 1).Height + 10
    UHeight = UserControl.Height
    UserControl.Width = cmdNavigate(i - 1).Left + cmdNavigate(i - 1).Width
    UWidth = UserControl.Width - 40
End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub


Private Sub cmdSelect_Click(Index As Integer)
    Code = cmdSelect(Index).Tag
    Description = cmdSelect(Index).Caption
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,795
Public Property Get ButHeight() As Long
    ButHeight = m_ButHeight
End Property

Public Property Let ButHeight(ByVal New_ButHeight As Long)
    m_ButHeight = New_ButHeight
    PropertyChanged "ButHeight"
    Call FillButtons(m_Rows, m_Columns, m_ExtraBut, m_ButDistance)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1065
Public Property Get ButWidth() As Long
    ButWidth = m_ButWidth
End Property

Public Property Let ButWidth(ByVal New_ButWidth As Long)
    m_ButWidth = New_ButWidth
    PropertyChanged "ButWidth"
    Call FillButtons(m_Rows, m_Columns, m_ExtraBut, m_ButDistance)
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
    PropertyChanged "ConString"
    Call initConnection(m_Constring, m_sqlStr)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ExtraBut() As Integer
    ExtraBut = m_ExtraBut
End Property

Public Property Let ExtraBut(ByVal New_ExtraBut As Integer)
    m_ExtraBut = New_ExtraBut
    PropertyChanged "ExtraBut"
    Call FillButtons(m_Rows, m_Columns, m_ExtraBut, m_ButDistance)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Rows() As Integer
    Rows = m_Rows
End Property

Public Property Let Rows(ByVal New_Rows As Integer)
    m_Rows = New_Rows
    PropertyChanged "Rows"
    Call FillButtons(m_Rows, m_Columns, m_ExtraBut, m_ButDistance)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Columns() As Integer
    Columns = m_Columns
End Property

Public Property Let Columns(ByVal New_Columns As Integer)
    m_Columns = New_Columns
    PropertyChanged "Columns"
    Call FillButtons(m_Rows, m_Columns, m_ExtraBut, m_ButDistance)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get sqlStr() As String
    sqlStr = m_sqlStr
End Property

Public Property Let sqlStr(ByVal New_sqlStr As String)
    m_sqlStr = New_sqlStr
    PropertyChanged "SqlStr"
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
    m_Code = m_def_Code
    m_Constring = m_def_ConString
    m_ExtraBut = m_def_ExtraBut
    m_Rows = m_def_Rows
    m_Columns = m_def_Columns
    m_sqlStr = m_def_SqlStr
    m_Description = m_def_Description
    m_ButHeight = m_def_ButHeight
    m_ButWidth = m_def_ButWidth
    m_ButDistance = m_def_ButDistance
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_Code = PropBag.ReadProperty("Code", m_def_Code)
    m_Constring = PropBag.ReadProperty("ConString", m_def_ConString)
    m_ExtraBut = PropBag.ReadProperty("ExtraBut", m_def_ExtraBut)
    m_Rows = PropBag.ReadProperty("Rows", m_def_Rows)
    m_Columns = PropBag.ReadProperty("Columns", m_def_Columns)
    m_sqlStr = PropBag.ReadProperty("SqlStr", m_def_SqlStr)
    m_Description = PropBag.ReadProperty("Description", m_def_Description)
    Set cmdSelect(0).Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set cmdNavigate(0).Font = PropBag.ReadProperty("Font", Ambient.Font)
    cmdSelect(0).Height = m_ButHeight
    cmdNavigate(0).Height = m_ButHeight
    cmdSelect(0).Width = m_ButWidth
    cmdNavigate(0).Width = m_ButWidth
    cmdSelect(0).FontBold = PropBag.ReadProperty("FontBold", 0)
    cmdNavigate(0).FontBold = PropBag.ReadProperty("FontBold", 0)
    cmdSelect(0).FontItalic = PropBag.ReadProperty("FontItalic", 0)
    cmdNavigate(0).FontItalic = PropBag.ReadProperty("FontItalic", 0)
    m_ButHeight = PropBag.ReadProperty("ButHeight", m_def_ButHeight)
    m_ButWidth = PropBag.ReadProperty("ButWidth", m_def_ButWidth)
    m_ButDistance = PropBag.ReadProperty("ButDistance", m_def_ButDistance)
End Sub

Private Sub UserControl_Terminate()
    Call unloadObj(cmdSelect)
    Call unloadObj(cmdNavigate)
    If rsGbl.State = 1 Then rsGbl.Close
    Set rsGbl = Nothing
    If dbCon.State = 1 Then dbCon.Close
    Set dbCon = Nothing
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Code", m_Code, m_def_Code)
    Call PropBag.WriteProperty("ConString", m_Constring, m_def_ConString)
    Call PropBag.WriteProperty("ExtraBut", m_ExtraBut, m_def_ExtraBut)
    Call PropBag.WriteProperty("Rows", m_Rows, m_def_Rows)
    Call PropBag.WriteProperty("Columns", m_Columns, m_def_Columns)
    Call PropBag.WriteProperty("SqlStr", m_sqlStr, m_def_SqlStr)
    Call PropBag.WriteProperty("Description", m_Description, m_def_Description)
    Call PropBag.WriteProperty("Font", cmdSelect(0).Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", cmdSelect(0).FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", cmdSelect(0).FontItalic, 0)
    Call PropBag.WriteProperty("FontName", cmdSelect(0).FontName, "")
    Call PropBag.WriteProperty("FontSize", cmdSelect(0).FontSize, 0)
    Call PropBag.WriteProperty("FontName", cmdSelect(0).FontName, "")
    Call PropBag.WriteProperty("ButHeight", m_ButHeight, m_def_ButHeight)
    Call PropBag.WriteProperty("ButWidth", m_ButWidth, m_def_ButWidth)
    Call PropBag.WriteProperty("ButDistance", m_ButDistance, m_def_ButDistance)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,80
Public Property Get ButDistance() As Long
    ButDistance = m_ButDistance
End Property

Public Property Let ButDistance(ByVal New_ButDistance As Long)
    m_ButDistance = New_ButDistance
    PropertyChanged "ButDistance"
    Call FillButtons(m_Rows, m_Columns, m_ExtraBut, m_ButDistance)
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdSelect(0),cmdSelect,0,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = cmdSelect(0).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set cmdSelect(0).Font = New_Font
    PropertyChanged "Font"
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

