VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form serveradfrm 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QAQStarter 推荐服务器"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   Icon            =   "serverad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8265
   StartUpPosition =   3  '窗口缺省
   Begin QAQStarter.bluebutton bluebutton1 
      Height          =   315
      Left            =   6600
      TabIndex        =   13
      Top             =   5640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DownColor       =   12632256
      Caption         =   "刷新列表"
      BorderStyle     =   1
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      ScaleHeight     =   855
      ScaleWidth      =   7815
      TabIndex        =   11
      Top             =   4680
      Width           =   7815
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   600
         TabIndex        =   21
         Top             =   120
         Width           =   5655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   20
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   180
         TabIndex        =   12
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      ScaleHeight     =   855
      ScaleWidth      =   7815
      TabIndex        =   9
      Top             =   3660
      Width           =   7815
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   600
         TabIndex        =   19
         Top             =   120
         Width           =   5655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   18
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   180
         TabIndex        =   10
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      ScaleHeight     =   855
      ScaleWidth      =   7815
      TabIndex        =   7
      Top             =   2640
      Width           =   7815
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   600
         TabIndex        =   17
         Top             =   120
         Width           =   5655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   16
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   180
         TabIndex        =   8
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      ScaleHeight     =   855
      ScaleWidth      =   7815
      TabIndex        =   5
      Top             =   1620
      Width           =   7815
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   600
         TabIndex        =   15
         Top             =   120
         Width           =   5655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   14
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   180
         TabIndex        =   6
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      ScaleHeight     =   855
      ScaleWidth      =   7815
      TabIndex        =   1
      Top             =   600
      Width           =   7815
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   4
         Top             =   450
         Width           =   4335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   90
         Width           =   5655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   180
         TabIndex        =   2
         Top             =   60
         Width           =   375
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5040
      Top             =   5580
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "推荐服务器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3555
   End
End
Attribute VB_Name = "serveradfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'我实在想不出服务器状态跟启动器有什么关系...
'而且我觉得这个功能比较累赘。mc服务器状态毕竟能在游戏中查得到。
'好吧，我把测试状态的函数集成到了一个函数中。
'这个还是弄服务器广告吧..233...
'（得瑟留）

Sub RequestServerState172(serverip As String)
    Dim tmp1()
    serverip = InputBox("请输入服务器 IP:端口", "添加服务器")
    tmp1 = Split(serverip, ":")
    Winsock1.Connect tmp1(0), tmp1(1)
End Sub

Private Sub Winsock1_Connect()
    Dim Ping172(1 To 4) As Variant
    Ping172(1) = &H904000F
    Ping172(2) = "localhostc"
    Ping172(3) = &H101D8
    Ping172(4) = &H1 '这个是备用的，发过了上面三个包，再次查询只要发这个包就可以
    '开始发送
    With Winsock1
        .SendData Ping172(1)
        .SendData Ping172(2)
        .SendData Ping172(3)
    End With
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
   Dim tmp1 As String
    Dim WasteHand As Variant
    Winsock1.GetData WasteHand, , 3 '把前面三个没用的报头去掉
    Winsock1.GetData tmp1, vbString
    tmp2 = InStr(tmp1, "online"":")
    tmp3 = InStr(tmp2 + 8, tmp1, "}")
    tmp4 = Mid(tmp1, tmp2 + 8, tmp3 - tmp2 - 8)
    If InStr(tmp4, ",") <> 0 Then
    tmp2 = InStr(tmp1, "online"":")
    tmp3 = InStr(tmp2 + 8, tmp1, ",")
    tmp4 = Mid(tmp1, tmp2 + 8, tmp3 - tmp2 - 8)
    End If
    tmp5 = InStr(tmp1, "max"":")
    tmp6 = InStr(tmp5 + 5, tmp1, ",")
    tmp7 = Mid(tmp1, tmp5 + 5, tmp6 - tmp5 - 5)
    MsgBox tmp4 & " " & tmp7, vbInformation, "服务器状态查询"
End Sub

Function GetServerAD()
On Error GoTo err1
    Dim rtn
    rtn = HtmlStr("http://desecity.com/projects/qaqstarter/server-list.txt")
    If rtn = "" Then GetServerAD = False: Exit Function
    Erase sl
    Dim sp1$(), sp2$()
    sp1 = Split(rtn, "#")
    For I = 0 To 4
        sp2 = Split(sp1(I), "|")
        ReDim Preserve sl(I)
        sl(I).servername = sp2(0)
        sl(I).serverip = sp2(1)
        Label3(I).Caption = sp2(0)
        Label4(I).Caption = sp2(1)
    Next
    GetServerAD = True
Exit Function
err1:
geterverad = False
End Function
