VERSION 5.00
Begin VB.UserControl RunCodeAtDesignTime 
   BackColor       =   &H00095ACA&
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   3900
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "4.如果 TPRunAnIndexCode设置为 1 ，那么label2的显示文字会变成""b""。"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   180
      TabIndex        =   12
      Top             =   5340
      Width           =   3435
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "4.如果 TPRunAnIndexCode设置为 0 ，那么label1的显示文字会变成""a""。"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   180
      TabIndex        =   11
      Top             =   4800
      Width           =   3435
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "3.设置Spliter为 # ."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   180
      TabIndex        =   10
      Top             =   4500
      Width           =   3135
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "2.设置TPCodes为    Label1.caption=""a""#Label2.caption=""b""."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   180
      TabIndex        =   9
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "1.设置TPObjects为 Label1#Label2"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   3660
      Width           =   3135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "示例"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   3360
      Width           =   1395
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "4.在TPRunAnIndexCode中设置索引号（索引号为从 0 开始的 TPCodes.成功设置后，你的代码将会执行."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   180
      TabIndex        =   6
      Top             =   2580
      Width           =   3435
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "3.在TPCodes中预设代码（VBScript）。用 Spliter(默认为#)隔开每句代码."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   180
      TabIndex        =   5
      Top             =   2070
      Width           =   3435
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "2.在TPObjects中确定要操控的控件，用 Spliter(默认为#)隔开每个名称."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   180
      TabIndex        =   4
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "1.将该控件放在要操控的控件的容器上."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "使用教程"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   1020
      Width           =   1395
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4020
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Powered By DeseCity Studio"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   530
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "设计界面代码执行工具 V1.0"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   420
      TabIndex        =   0
      Top             =   150
      Width           =   3075
   End
End
Attribute VB_Name = "RunCodeAtDesignTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Designed by lkfiuean(天天那么得瑟)
Const m_def_TPRunAnIndexCode = -1
Const m_def_TPCodes = ""
Const m_def_TPObjects = ""

Dim m_TPRunAnIndexCode As Integer
Dim m_TPCodes As String
Dim m_TPObjects As String
'缺省属性值:
Const m_def_BackColor = &H95ACA
Const m_def_TPSpliter = "#"
'属性变量:
Dim m_BackColor As OLE_COLOR
Dim m_TPSpliter As String



'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_TPCodes = PropBag.ReadProperty("TPCodes", m_def_TPCodes)
    m_TPObjects = PropBag.ReadProperty("TPObjects", m_def_TPObjects)
    m_TPRunAnIndexCode = PropBag.ReadProperty("TPRunAnIndexCode", m_def_TPRunAnIndexCode)
    m_TPSpliter = PropBag.ReadProperty("TPSpliter", m_def_TPSpliter)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    Label2.Caption = PropBag.ReadProperty("Caption", "Powered By DeseCity Studio")
End Sub

Private Sub UserControl_Resize()
    If Width > 3900 Then Width = 3900
    If Height > 6000 Then Height = 6000
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("TPCodes", m_TPCodes)
    Call PropBag.WriteProperty("TPObjects", m_TPObjects, m_def_TPObjects)
    Call PropBag.WriteProperty("TPRunAnIndexCode", m_TPRunAnIndexCode, m_def_TPRunAnIndexCode)
    Call PropBag.WriteProperty("TPSpliter", m_TPSpliter, m_def_TPSpliter)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Caption", Label2.Caption, "Powered By DeseCity Studio")
End Sub

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    'm_TPCodes = m_def_TPCodes
    m_TPObjects = m_def_TPObjects
    m_TPRunAnIndexCode = m_def_TPRunAnIndexCode
    m_TPSpliter = m_def_TPSpliter
End Sub
'
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = UserControl.BackColor
    'm_BackColor = m_def_BackColor
'End Property

Public Property Get TPCodes() As String
Attribute TPCodes.VB_Description = "代码列表，每句代码之间必须要用TPSpliter隔开，如Label1.caption=""a""#Label2.caption=""b""."
    If Ambient.UserMode Then Err.Raise 393
    TPCodes = m_TPCodes
End Property

Public Property Let TPCodes(ByVal New_TPCodes As String)
    If Ambient.UserMode Then Err.Raise 382
    m_TPCodes = New_TPCodes
    PropertyChanged "TPCodes"
End Property

Public Property Get TPObjects() As String
Attribute TPObjects.VB_Description = "定义TPCodes中操作的对象.每个对象之间用TPSpliter隔开.如label1#label2."
    If Ambient.UserMode Then Err.Raise 393
    TPObjects = m_TPObjects
End Property

Public Property Let TPObjects(ByVal New_TPObjects As String)
    If Ambient.UserMode Then Err.Raise 382
    m_TPObjects = New_TPObjects
    PropertyChanged "TPObjects"
End Property

Public Property Get TPRunAnIndexCode() As Integer
Attribute TPRunAnIndexCode.VB_Description = "在设计界面运行一行TPCodes中的代码.值可以指定运行代码的索引号，索引号即TPCodes中的第x个代码(0开始)。如TPCodes为label1.caption=""a""#label2.caption=""b""，那么如果TPRunAnIndexCode为1，就会使label2的显示文字变成b."
    If Ambient.UserMode Then Err.Raise 393
    TPRunAnIndexCode = m_TPRunAnIndexCode
End Property

Public Property Let TPRunAnIndexCode(ByVal New_TPRunAnIndexCode As Integer)
    If Ambient.UserMode Then Err.Raise 382
    m_TPRunAnIndexCode = New_TPRunAnIndexCode
    PropertyChanged "TPRunAnIndexCode"
    '*********运行一行代码********
    On Error GoTo exitprop
    Dim sc As Object
    Set sc = CreateObject("scriptcontrol")
    'sc.Reset
    sc.Language = "vbscript"
    Dim CObjects$(), CCodes$()
    CObjects = Split(m_TPObjects, m_TPSpliter)
    CCodes = Split(m_TPCodes, m_TPSpliter)
    Dim objs, parentobj As Object
    Set parentobj = UserControl.Parent
    For Each objs In CObjects
        sc.AddObject objs, parentobj.Controls(objs)
        'MsgBox parentobj.Controls(objs).Name
    Next
    sc.AddCode CCodes(m_TPRunAnIndexCode)
    Set sc = Nothing
    Exit Property
exitprop:
    If Err.Description = "下标越界" Then
        MsgBox "索引号不存在."
        Err.Clear
    Else
        MsgBox Err.Description
        Err.Clear
    End If
Exit Property
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,3,0,0
Public Property Get TPSpliter() As String
Attribute TPSpliter.VB_Description = "分隔符"
    If Ambient.UserMode Then Err.Raise 393
    TPSpliter = m_TPSpliter
End Property

Public Property Let TPSpliter(ByVal New_TPSpliter As String)
    If Ambient.UserMode Then Err.Raise 382
    m_TPSpliter = New_TPSpliter
    PropertyChanged "TPSpliter"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=10,0,0,&H95ACA&
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Label2,Label2,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "返回/设置对象的标题栏中或图标下面的文本。"
    Caption = Label2.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label2.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

