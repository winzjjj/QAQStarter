VERSION 5.00
Begin VB.Form frmsetting 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "配置"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "frmsetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   4695
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "启动器设置"
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
      Height          =   1815
      Left            =   180
      TabIndex        =   18
      Top             =   3480
      Width           =   4335
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "使用背景图片（.QAQStarter_Data 目录）"
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
         TabIndex        =   20
         Top             =   720
         Value           =   -1  'True
         Width           =   4035
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "默认颜色"
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
         TabIndex        =   19
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Notice: QAQStarter 只会将 .QAQStarter_Data 目录的 BMP 和 JPG 图片随机作为背景."
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   555
         Left            =   180
         TabIndex        =   21
         Top             =   1080
         Width           =   3975
      End
   End
   Begin QAQStarter.bluebutton btnApply 
      Height          =   435
      Left            =   3600
      TabIndex        =   17
      Top             =   5400
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   767
      BackColor       =   13470983
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DownColor       =   11959559
      Caption         =   "确定"
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "高级设置"
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
      Height          =   1155
      Left            =   180
      TabIndex        =   12
      Top             =   2100
      Width           =   4335
      Begin VB.TextBox txtpassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   660
         Width           =   2355
      End
      Begin VB.CheckBox chkEna_zlg 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "启用正版登录"
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
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Width           =   1455
      End
      Begin QAQStarter.jcbutton btntstlogin 
         Height          =   330
         Left            =   3180
         TabIndex        =   16
         Top             =   660
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         ButtonStyle     =   4
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "测试登陆"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "密码:"
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
         Height          =   315
         Left            =   180
         TabIndex        =   15
         Top             =   660
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "基本设置"
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
      Height          =   1815
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4335
      Begin VB.TextBox txtjavahome 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1020
         TabIndex        =   8
         Top             =   1260
         Width           =   1215
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1020
         TabIndex        =   3
         Top             =   330
         Width           =   2055
      End
      Begin VB.TextBox txtMemory 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1020
         TabIndex        =   1
         Text            =   "1024"
         Top             =   780
         Width           =   1575
      End
      Begin QAQStarter.jcbutton jcbutton1 
         Height          =   330
         Left            =   3180
         TabIndex        =   2
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "随机生成"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin QAQStarter.jcbutton jcbutton2 
         Height          =   330
         Left            =   3180
         TabIndex        =   4
         Top             =   780
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "系统判断"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin QAQStarter.jcbutton jcbutton3 
         Height          =   330
         Left            =   2340
         TabIndex        =   10
         Top             =   1260
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "浏览..."
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin QAQStarter.jcbutton jcbutton4 
         Height          =   330
         Left            =   3180
         TabIndex        =   11
         Top             =   1260
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "自动获取"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Jre 路径:"
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
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "用户名:"
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
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Jre 内存:"
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
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   795
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "MB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2700
         TabIndex        =   5
         Top             =   780
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmsetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function RandomUserNamer()
    Dim nameseed1(), nameseed2()
    nameseed1 = Array("creeper", "steve", "demaxiya", "gay", "him", "hero", "qaq", "smart", "pro", "night", "boom", "tnt", "dc", "simple", "hey", "micro", "star", "stream")
    nameseed2 = Array("555", "win", "aaa", "233", "98765432", "321", "II", "er", "mmm", "edd", "ess", "ive", "less", "ant", "III", "IV", "V")
    Dim shunxu As Boolean
    Randomize
    shunxu = IIf(Rnd() >= 0.5, True, False)
    Dim cse1$, cse2$
    Randomize
    cs1 = nameseed1(CInt((UBound(nameseed1))) * Rnd())
    Randomize
    cs2 = nameseed2(CInt((UBound(nameseed2))) * Rnd())
    If shunxu Then
        RandomUserNamer = cs1 & cs2
    Else
        RandomUserNamer = cs2 & cs1
    End If
    Randomize
    Dim addnumyn As Boolean
    addnumyn = IIf(Rnd() >= 0.2, True, False)
    If addnumyn Then
        Randomize
        Dim addnum&
        addnum = CInt(Rnd() * 100)
        RandomUserNamer = RandomUserNamer & CStr(addnum)
    End If
End Function

Private Sub btnApply_Click()

    If txtUserName.Text = "" Then MsgBox "请输入用户名.", vbExclamation, "补充你的凭据": Exit Sub
    If txtMemory.Text = "" Then MsgBox "请输入内存大小.", vbExclamation: Exit Sub
    If IsNumeric(txtMemory.Text) = True Then
        If txtMemory.Text > 0 Then GoTo passmemory
    End If
        MsgBox "内存大小不是一个正整数。", vbExclamation, "修改你的凭据": Exit Sub
passmemory:
    If txtjavahome.Text = "" Then MsgBox "请输入 JRE 路径.", vbExclamation, "补充你的凭据": Exit Sub
    If Not JreCheck(txtjavahome.Text) Then MsgBox "JRE 路径 指向了一个并不存在的路径。", vbExclamation, "修改你的凭据": Exit Sub
    Open App.Path & "\.QAQStarter_Data\config.ini" For Output As #1
    Print #1, txtUserName.Text '输出用户名
    Print #1, txtMemory.Text '输出内存大小
    Print #1, txtjavahome.Text '输出JAVA路径
    Print #1, txtpassword.Text '输出密码
    Print #1, IIf(Option1.Value = True, "normal", "showback")  '输出皮肤
    Dim changedskin As Boolean
    If strColor <> IIf(Option1.Value = True, "normal", "showback") Then
    changedskin = True
    End If
    LoadSetting
    MsgBox "设置保存成功!" & IIf(changedskin = True, "皮肤设置将在下一次启动 QAQStarter 时生效.", ""), vbInformation
    If ForcingSetting Then ForcedSetReturn = True: ForcingSetting = False '强制设置完成
    Unload Me
End Sub

Private Sub btntstlogin_Click()
tmps1 = OnlineCheck(txtUserName, txtpassword)
tmps2 = InStr(tmps1, "{""accessToken"":""")
tmps3 = InStr(tmps2 + 17, tmps1, """")
accessToken = Mid(tmps1, tmps2 + 17, tmps3 - tmps2 - 17)
tmps2 = InStr(tmps1, """id"":""")
tmps3 = InStr(tmps2 + 6, tmps1, """")
id = Mid(tmps1, tmps2 + 6, tmps3 - tmps2 - 6)
tmps2 = InStr(tmps1, """twitch_access_token"",""value"":""")
tmps3 = InStr(tmps2 + 31, tmps1, """")
twitch_access_token = Mid(tmps1, tmps2 + 31, tmps3 - tmps2 - 31)
MsgBox accessToken & vbCrLf & "id=" & id & vbCrLf & "twitch_access_token=" & twitch_access_token, vbInformation, "正版验证成功"
End Sub



Private Sub chkEna_zlg_Click()
If txtpassword.Enabled = True Then
    txtpassword.Enabled = False
    txtpassword = ""
    btntstlogin.Enabled = False
Else
    txtpassword.Enabled = True
    btntstlogin.Enabled = True
End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Shell "cmd.exe /c md """ & App.Path & "\.QAQStarter_Data""", vbHide
    txtUserName.Text = strUserName
    txtMemory.Text = IIf(varMemory = "", 1024, CLng(varMemory))
    txtjavahome.Text = strJREPath
    Controls(IIf(strColor = "showback", "option2", "option1")).Value = True
    If strPassword <> "" Then
        chkEna_zlg.Value = 1
        txtpassword.Text = strPassword
    End If
    Me.Controls("opt" & strColor).Value = True
    If ForcingSetting Then
    txtjavahome.Text = sMid(GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\jarfile\shell\open\command", ""), """", """", , , 1)
    ZCNCDX = GetMemoryInfo(总共的物理内存, M)
    Me.txtMemory = Int(ZCNCDX / 2)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    If ForcingSetting Then '是否正在强制设置
        If MsgBox("必须完成初始配置。您可以按“否”关闭 QAQStarter，在以后的时间完成配置，也可以按“是”继续配置。", vbExclamation + vbYesNo, "是否继续配置？") = vbNo Then
            Cancel = 0
            ForcedSetReturn = False
        End If
    Else
        Cancel = 0
    End If
End Sub

Private Sub jcbutton1_Click()
    txtUserName.Text = RandomUserNamer
End Sub

Private Sub jcbutton2_Click()
ZCNCDX = GetMemoryInfo(总共的物理内存, M)
Me.txtMemory = Int(ZCNCDX / 2)


End Sub

Private Function JreCheck(strPath$) As Boolean
Dim fso As Object
Set fso = CreateObject("scripting.filesystemobject")
With fso
If .FolderExists(strPath) Then
JreCheck = False
Else
If Dir(strPath, vbSystem + vbHidden) = "" Then
JreCheck = False
Else
JreCheck = True
End If
End If
End With
Set fso = Nothing
End Function

Private Sub jcbutton3_Click()
    aa = mdlFileDlg.FileDialog(Me, False, "选择 Java 的路径", "javaw.exe|javaw.exe|java.exe|java.exe")
    txtjavahome.Text = IIf(aa = "", txtjavahome.Text, aa)
    
End Sub

Private Sub jcbutton4_Click()
txtjavahome.Text = sMid(GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\jarfile\shell\open\command", ""), """", """", , , 1)
End Sub
