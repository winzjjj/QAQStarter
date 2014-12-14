VERSION 5.00
Begin VB.UserControl 下载控件 
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8355
   ScaleHeight     =   3855
   ScaleWidth      =   8355
   ToolboxBitmap   =   "下载控件.ctx":0000
   Begin QAQStarter.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "状态"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin QAQStarter.WinHttpDown WinHttpDown1 
      Left            =   7560
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1935
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "下载控件"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Event 下载失败()
Event 下载完毕()
Event 下载错误()
Public Function 下载文件(URL As String, 路径和名称 As String) As String
    
'    arr = Split(路径和名称, "\")
'
'
'    '  If UBound(Split(路径和名称, "\")) = 1 Then
'    '   Label2 = arr(UBound(Split(路径和名称, "\")))
'    '  Else
'    '    Label2 = arr(UBound(Split(路径和名称, "\")))
'    ' End If
'
'    Label2 = arr(UBound(Split(路径和名称, "\")))
    tmp1 = Split(路径和名称, "\")
    tmp2 = Split(URL, "/")
    If tmp1(UBound(tmp1)) <> tmp2(UBound(tmp2)) Then
    If Right(路径和名称, 1) = "\" Then
    路径和名称 = 路径和名称 & tmp2(UBound(tmp2))
    Else
    路径和名称 = 路径和名称 & "\" & tmp2(UBound(tmp2))
    End If
    End If
    WinHttpDown1.FileName = 路径和名称
    WinHttpDown1.URL = URL
    WinHttpDown1.GetStart
End Function
Public Function 暂停下载() As String
    WinHttpDown1.GetPause
End Function
Public Function 停止下载() As String
    WinHttpDown1.GetStop
End Function



Private Sub Label4_Click()

End Sub

Private Sub WinHttpDown1_HttpState(ByVal 文件大小 As String, ByVal 进度 As Single, ByVal 下载速度 As String, ByVal 已下载大小 As String)
    ProgressBar1.Text = 已下载大小 & "/" & 文件大小
    Label5 = 下载速度 & "/s"
    Label6 = Format(进度, "0.00") & "%"
    ProgressBar1.Value = 进度
End Sub
Private Sub WinHttpDown1_StateChanged(ByVal State As Integer)
    On Error Resume Next
    Dim msg As String
    Select Case State
    Case 1
        msg = "发送请求..."
    Case 2
        msg = "获取远程文件信息..."
    Case 3
        
    Case 4
        
        msg = "下载被终止..."
    Case 5
        msg = "停止下载"
    Case 6
        msg = "暂停下载"
    Case 7
        
        msg = "连接服务器失败"
        RaiseEvent 下载失败
    Case 8
        msg = "发送请求失败"
        RaiseEvent 下载失败
    Case 9
        
        msg = "下载完毕"
        RaiseEvent 下载完毕
        
    Case 10
        msg = "保存路径出错"
        RaiseEvent 下载失败
    End Select
    
    ProgressBar1.Text = msg
    
End Sub

