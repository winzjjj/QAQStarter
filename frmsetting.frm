VERSION 5.00
Begin VB.Form frmsetting 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Caption         =   "ʹ�ñ���ͼƬ��.QAQStarter_Data Ŀ¼��"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "Ĭ����ɫ"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "Notice: QAQStarter ֻ�Ὣ .QAQStarter_Data Ŀ¼�� BMP �� JPG ͼƬ�����Ϊ����."
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DownColor       =   11959559
      Caption         =   "ȷ��"
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "�߼�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
            Name            =   "΢���ź�"
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
         Caption         =   "���������¼"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "���Ե�½"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
            Name            =   "΢���ź�"
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
            Name            =   "΢���ź�"
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
            Name            =   "΢���ź�"
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
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "�������"
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
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ϵͳ�ж�"
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
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "���..."
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
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "�Զ���ȡ"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Jre ·��:"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "�û���:"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "Jre �ڴ�:"
         BeginProperty Font 
            Name            =   "΢���ź�"
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

    If txtUserName.Text = "" Then MsgBox "�������û���.", vbExclamation, "�������ƾ��": Exit Sub
    If txtMemory.Text = "" Then MsgBox "�������ڴ��С.", vbExclamation: Exit Sub
    If IsNumeric(txtMemory.Text) = True Then
        If txtMemory.Text > 0 Then GoTo passmemory
    End If
        MsgBox "�ڴ��С����һ����������", vbExclamation, "�޸����ƾ��": Exit Sub
passmemory:
    If txtjavahome.Text = "" Then MsgBox "������ JRE ·��.", vbExclamation, "�������ƾ��": Exit Sub
    If Not JreCheck(txtjavahome.Text) Then MsgBox "JRE ·�� ָ����һ���������ڵ�·����", vbExclamation, "�޸����ƾ��": Exit Sub
    Open App.Path & "\.QAQStarter_Data\config.ini" For Output As #1
    Print #1, txtUserName.Text '����û���
    Print #1, txtMemory.Text '����ڴ��С
    Print #1, txtjavahome.Text '���JAVA·��
    Print #1, txtpassword.Text '�������
    Print #1, IIf(Option1.Value = True, "normal", "showback")  '���Ƥ��
    Dim changedskin As Boolean
    If strColor <> IIf(Option1.Value = True, "normal", "showback") Then
    changedskin = True
    End If
    LoadSetting
    MsgBox "���ñ���ɹ�!" & IIf(changedskin = True, "Ƥ�����ý�����һ������ QAQStarter ʱ��Ч.", ""), vbInformation
    If ForcingSetting Then ForcedSetReturn = True: ForcingSetting = False 'ǿ���������
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
MsgBox accessToken & vbCrLf & "id=" & id & vbCrLf & "twitch_access_token=" & twitch_access_token, vbInformation, "������֤�ɹ�"
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
    ZCNCDX = GetMemoryInfo(�ܹ��������ڴ�, M)
    Me.txtMemory = Int(ZCNCDX / 2)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    If ForcingSetting Then '�Ƿ�����ǿ������
        If MsgBox("������ɳ�ʼ���á������԰����񡱹ر� QAQStarter�����Ժ��ʱ��������ã�Ҳ���԰����ǡ��������á�", vbExclamation + vbYesNo, "�Ƿ�������ã�") = vbNo Then
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
ZCNCDX = GetMemoryInfo(�ܹ��������ڴ�, M)
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
    aa = mdlFileDlg.FileDialog(Me, False, "ѡ�� Java ��·��", "javaw.exe|javaw.exe|java.exe|java.exe")
    txtjavahome.Text = IIf(aa = "", txtjavahome.Text, aa)
    
End Sub

Private Sub jcbutton4_Click()
txtjavahome.Text = sMid(GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\jarfile\shell\open\command", ""), """", """", , , 1)
End Sub
