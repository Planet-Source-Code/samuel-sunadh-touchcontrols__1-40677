VERSION 5.00
Object = "*\ATouch.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin Touch.Touch4 Touch4Item 
      Height          =   4305
      Left            =   1665
      TabIndex        =   1
      Top             =   600
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   7594
      ExtraBut        =   1
      Rows            =   4
      Columns         =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Arial"
      FontSize        =   8.25
      FontName        =   "Arial"
   End
   Begin Touch.Touch2 Touch2Menu 
      Height          =   5175
      Left            =   390
      TabIndex        =   0
      Top             =   585
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   9128
      FontSize        =   8.25
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      TotalButtons    =   4
      Defaultclick    =   -1  'True
      DefaultClickButtonNo=   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   1665
      TabIndex        =   3
      Top             =   150
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   405
      TabIndex        =   2
      Top             =   165
      Width           =   810
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conStr As String


Private Sub Form_Load()
    conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Test.mdb" & ";Jet OLEDB:Database Password="
    Touch2Menu.conString = conStr
    Touch4Item.conString = conStr
    Touch2Menu.sqlstr = "Select Menucode,Menudesc,Bcolor,Menupic from menu"
End Sub

Private Sub Touch2Menu_Click()
    Touch4Item.sqlstr = "select Itemcode,Itemdesc,Items.bcolor,Items.Itempic" & _
                        " from Items,Submenu,Menu where Items.submenucode=" & _
                        " Submenu.submenucode and submenu.menucode=menu.menucode" & _
                        " and Menu.menucode='" & Touch2Menu.Code & "' order by Itemdesc"
End Sub

Private Sub Touch4Item_Click()
    MsgBox "Code " & Touch4Item.Code & "   Description  " & Touch4Item.Description
End Sub
