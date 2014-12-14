VERSION 5.00
Begin VB.UserControl bluebutton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   2160
      TabIndex        =   0
      Top             =   1440
      Width           =   90
   End
End
Attribute VB_Name = "bluebutton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'缺省属性值:
Const m_def_DownColor = 0
'Const m_def_Caption = ""
'Const m_def_ForeColor = 0
'Const m_def_DownColor = 0
'Const m_def_Caption = "0"
'属性变量:
Dim m_DownColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
'事件声明:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Click()
Event DblClick()

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
lblxy
End Property

Private Sub Label1_Click()
    RaiseEvent Click
    lblxy
End Sub

Private Sub Label1_DblClick()
    RaiseEvent DblClick
    lblxy
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    m_BackColor = Me.BackColor
    Me.BackColor = m_DownColor
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
        Me.BackColor = m_BackColor
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
    lblxy
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
    lblxy
End Sub
'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    Set Label1.Font = Ambient.Font
    m_DownColor = m_def_DownColor
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_DownColor = PropBag.ReadProperty("DownColor", m_def_DownColor)
    Label1.Caption = PropBag.ReadProperty("Caption", "Label1")
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    lblxy
End Sub

Private Sub UserControl_Resize()
    lblxy
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("DownColor", m_DownColor, m_def_DownColor)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "Label1")
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    lblxy
End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "返回/设置对象中文本和图形的前景色。"
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
    lblxy
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    m_BackColor = Me.BackColor
    Me.BackColor = m_DownColor
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    Me.BackColor = m_BackColor
End Sub

Public Property Get DownColor() As OLE_COLOR
    DownColor = m_DownColor
End Property

Public Property Let DownColor(ByVal New_DownColor As OLE_COLOR)
    m_DownColor = New_DownColor
    PropertyChanged "DownColor"
End Property

Public Property Get Caption() As String
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
    lblxy
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "返回/设置对象的边框样式。"
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Sub lblxy()
    Label1.Left = ScaleWidth / 2 - Label1.Width / 2
    Label1.Top = ScaleHeight / 2 - Label1.Height / 2
End Sub
