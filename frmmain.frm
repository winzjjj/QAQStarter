VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00525252&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7200
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10200
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10200
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox gbdx 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   180
      ScaleHeight     =   675
      ScaleWidth      =   1095
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox bootitem 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmmain.frx":57E2
      Left            =   7740
      List            =   "frmmain.frx":57E4
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   6240
      Width           =   2355
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2280
      Top             =   4920
   End
   Begin VB.FileListBox File1 
      Height          =   270
      Left            =   720
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin QAQStarter.jcbutton jcbutton1 
      Height          =   375
      Left            =   8100
      TabIndex        =   2
      Top             =   6720
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Minecraft Ӧ���г�"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin QAQStarter.jcbutton jcbutton2 
      Height          =   375
      Left            =   6660
      TabIndex        =   5
      Top             =   6720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Caption         =   "���˷��Ƽ�"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin QAQStarter.jcbutton btnsetting 
      Height          =   375
      Left            =   5220
      TabIndex        =   3
      Top             =   6720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Caption         =   "����������"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   6480
      Width           =   10215
   End
   Begin VB.Label btnstartgame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   3120
      TabIndex        =   7
      Top             =   3720
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   135
      Picture         =   "frmmain.frx":57E6
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9120
      TabIndex        =   9
      ToolTipText     =   "����"
      Top             =   50
      Width           =   255
   End
   Begin QAQStarter.TrkenderPlus clickxger 
      Height          =   1455
      Left            =   3120
      Top             =   3720
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2566
      TMD             =   50
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3780
      TabIndex        =   8
      Top             =   2580
      Width           =   4635
   End
   Begin VB.Label btnCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ��Ϸ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4020
      TabIndex        =   6
      Top             =   4080
      Width           =   2055
   End
   Begin QAQStarter.TrkenderPlus TMDer 
      Height          =   1455
      Left            =   3120
      Top             =   3720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2566
      TMD             =   120
   End
   Begin VB.Label btnmin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   4
      ToolTipText     =   "��С��"
      Top             =   0
      Width           =   255
   End
   Begin VB.Label btnclose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9840
      TabIndex        =   1
      ToolTipText     =   "�رմ���"
      Top             =   45
      Width           =   255
   End
   Begin VB.Label lbltitle 
      BackStyle       =   0  'Transparent
      Caption         =   "QAQStarter - �����������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   60
      Width           =   8415
   End
   Begin VB.Menu mnuQuestion 
      Caption         =   "mnuQuestion"
      Visible         =   0   'False
      Begin VB.Menu mnuDevelopers 
         Caption         =   "������Ա��..."
      End
      Begin VB.Menu mnuAboutUs 
         Caption         =   "��������..."
      End
      Begin VB.Menu mnuReportBug 
         Caption         =   "�ύ����..."
      End
      Begin VB.Menu mnuEnterQAQGame 
         Caption         =   "���� QAQ��Ϸ��̳(QAQGame)..."
      End
      Begin VB.Menu mnuEnterDSDN 
         Caption         =   "���� DeseCity ����������(DSDN)..."
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const HTCAPTION = 2
Dim XiangShang As Boolean

Private Sub btnclose_Click()
    End
End Sub

Private Sub btnclose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        btnclose.FontBold = True
    End If
End Sub

Private Sub btnclose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        btnclose.FontBold = False
    End If
End Sub

Private Sub btnmin_Click()
    Me.WindowState = 1
End Sub
Private Sub btnmin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        btnmin.FontBold = True
    End If
End Sub

Private Sub btnmin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        btnmin.FontBold = False
    End If
End Sub

Private Sub btnsetting_Click()
    frmsetting.Show 1
End Sub

Private Sub btnstartgame_Click()
banben = bootitem.Text
If banben = "<ѡ��������>" Then MsgBox "����ѡ�������", vbExclamation: Me.Show: Exit Sub
btnCaption.Caption = "������...."
btnCaption.ForeColor = &H80000011
btnstartgame.Enabled = False
clickxger.Visible = False
'�ȸ����Ӽ��Libraries��˵
lblStatus.Caption = "��������������п�..."
If CheckConnect = True Then
frmmarket.Show
Me.Hide
frmmarket.CheckLibraries App.Path & "\.minecraft\versions\" & bootitem.Text, bootitem.Text
Else
    MsgBox "�����쳣�����п���¼��ʧ�ܣ�", vbExclamation
End If
DoEvents
Dim cmdarg
Me.Show
If frmsetting.chkEna_zlg.Value = 0 Then
'������֤
    DoEvents
    lblStatus.Caption = "��������������������..."
    DoEvents
    ArgSettingfor16x CLng(varMemory), strUserName, banben, ".minecraft\native", ".minecraft\assets"
    cmdarg = mdlMakeArg.OutputArg4Command
Else
'������֤
    lblStatus.Caption = "��������������������..."
    DoEvents
    tmps1 = OnlineCheck(strUserName, strPassword)
    tmps2 = InStr(tmps1, "{""accessToken"":""")
    tmps3 = InStr(tmps2 + 16, tmps1, """")
    accessToken = Mid(tmps1, tmps2 + 16, tmps3 - tmps2 - 16)
    tmps2 = InStr(tmps1, """id"":""")
    tmps3 = InStr(tmps2 + 6, tmps1, """")
    id = Mid(tmps1, tmps2 + 6, tmps3 - tmps2 - 6)
    tmps2 = InStr(tmps1, """name"":""")
    tmps3 = InStr(tmps2 + 8, tmps1, """")
    Username = Mid(tmps1, tmps2 + 8, tmps3 - tmps2 - 8)
    ArgSettingfor16x CLng(varMemory), Username, banben, ".minecraft\native", ".minecraft\assets", , , , , accessToken, id
    cmdarg = mdlMakeArg.OutputArg4Command
End If
'Debug.Print cmdarg
'Open App.Path & "\run.bat" For Output As #1
'Print #1, "cd """ & App.Path & """"
'Print #1, """" & frmsetting.txtjavahome & """ " & cmdarg
'Print #1, "del %0"
'Close #1
'Shell App.Path & "\run.bat", vbHide
DoEvents

lblStatus.Caption = "��ʼ���� Minecraft..."
DoEvents
Shell """" & frmsetting.txtjavahome & """ " & cmdarg, vbNormalFocus
End
End Sub

Private Sub btnstartgame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
    'btnstartgame.BorderStyle = 1
    clickxger.Visible = True
    End If
End Sub

Private Sub btnstartgame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

clickxger.Visible = False

End Sub

Private Sub Form_DblClick()
'���뱳��
File1.Path = App.Path & "\.QAQStarter_Data"
File1.Pattern = "*.jpg;*.bmp"
If File1.ListCount <> 0 And strColor = "showback" Then
Randomize
ѡ��ͼƬ = File1.List(Int((File1.ListCount - 1 - 0 + 1) * Rnd + 0))
LoadFormPicture App.Path & "\.QAQStarter_Data\" & ѡ��ͼƬ
End If
End Sub

Private Sub Form_Paint()

     Dim pImg As Long, pImg2 As Long
     Dim pGraphics As Long
     Dim w As Long, h As Long, w2 As Long, h2 As Long

    Call PaintPng(App.Path & "\.QAQStarter_Data\black.png", Me.hDC, 0, 0)
    Call PaintPng(App.Path & "\.QAQStarter_Data\minecraft.png", Me.hDC, -20, 33)
End Sub

Private Sub Form_Load()

If CheckConnect = True Then
frmmarket.CheckUpdate
frmmarket.CheckNative App.Path & "\.minecraft\Native\"
Else
    MsgBox "�����쳣��������������...", vbExclamation
End If
lbltitle = "QAQStarter V" & App.Major & "." & App.Minor & "." & App.Revision & " " & VerName & " - ����������磡"
CheckOldVersion App.Path & "\.minecraft"
ShowFolderList App.Path & "\.minecraft\versions\"

'���뱳��
File1.Path = App.Path & "\.QAQStarter_Data"
File1.Pattern = "*.jpg;*.bmp"
If File1.ListCount <> 0 And strColor = "showback" Then
Randomize
ѡ��ͼƬ = File1.List(Int((File1.ListCount - 1 - 0 + 1) * Rnd + 0))
LoadFormPicture App.Path & "\.QAQStarter_Data\" & ѡ��ͼƬ
End If
Me.Height = 6705
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
XiangShang = True
Timer1.Enabled = True
End Sub

Private Sub jcbutton1_Click()
    jcbutton1.Enabled = False
    If CheckConnect = False Then MsgBox "�����쳣���޷��� Minecraft Ӧ���г�...", vbExclamation: jcbutton1.Enabled = True: Exit Sub
    jcbutton1.Enabled = True
    frmmarket.Show 1
End Sub

Private Sub jcbutton2_Click()
jcbutton2.Enabled = False
If serveradfrm.GetServerAD Then
Load serveradfrm
jcbutton2.Enabled = True
serveradfrm.Show 1
Else
MsgBox "�����쳣���޷���÷������б�", vbExclamation
End If
jcbutton2.Enabled = True
End Sub

Private Sub Label2_Click()
    'MsgBox "QAQStarter ������Ա����" & vbCrLf & _
                  "* ������" & vbCrLf & _
                  "������ô��ɪ" & vbCrLf & "winzjjj" & vbCrLf & "�ǳ�" & vbCrLf & _
                  "* �������" & vbCrLf & "������ô��ɪ" & vbCrLf & _
                  "* ����֧��" & vbCrLf & "С��" & vbCrLf & _
                  "* �ر���л" & vbCrLf & "BMCL ������" & vbCrLf & "Minecraft QAQ������" & vbCrLf & "Mojang��Microsoft��CCTV��CCAV" & vbCrLf & _
                  "======QAQStarter ��", vbInformation
                frmThanks.Show
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
XiangShang = False
Timer1.Enabled = True
End Sub

Private Sub Label5_Click()
    PopupMenu mnuQuestion, , Label5.Left, Label5.Height + Label5.Top
End Sub

Private Sub lbltitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
    End If
End Sub

Public Sub ShowFolderList(folderspec)
bootitem.Clear
bootitem.AddItem "<ѡ��������>"
On Error Resume Next
     Dim fs, f, f1, S, sf
     Dim hs, h, h1, hf
     Set fs = CreateObject("Scripting.FileSystemObject")
     Set f = fs.GetFolder(folderspec)
     Set sf = f.SubFolders
     For Each f1 In sf
        bootitem.AddItem f1.Name
        bootitem.ListIndex = 0
     Next
End Sub

Private Sub mnuAboutUs_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuDevelopers_Click()
    frmThanks.Show 1
End Sub

Private Sub mnuEnterDSDN_Click()
    Shell "cmd /c start explorer.exe ""http://desecity.com""", vbHide
End Sub

Private Sub mnuReportBug_Click()
 Shell "cmd /c start explorer.exe ""http://www.qaqgame.com/forum.php?mod=viewthread&tid=4&extra=""", vbHide
End Sub

Private Sub Timer1_Timer()
'���ӣ��Ҿ�ϲ�����ο��������֣�
If XiangShang = True Then ' ����
If Me.Height > 6705 Then Me.Height = Me.Height - 30 Else Timer1.Enabled = False
Else
If Me.Height < 7200 Then Me.Height = Me.Height + 30 Else Timer1.Enabled = False
End If
End Sub

Sub LoadFormPicture(PicturePath As String)
    Dim zz As IPictureDisp
    Set zz = LoadPicture(PicturePath)
    With gbdx
        .Move 0, 0, 10230, 7320
        .PaintPicture zz, 0, 0, 10230, 7320   '����ͼƬ��С
        SavePicture .image, App.Path & "\.QAQStarter_Data\back.tmp"  '������ļ�
        Me.Picture = LoadPicture(App.Path & "\.QAQStarter_Data\back.tmp")  '������ر��κ�ͼƬ
        Kill App.Path & "\.QAQStarter_Data\back.tmp"   'ɾ������ͼƬ����
    End With
End Sub


Private Sub mnuEnterQAQGame_click()
Shell "cmd /c start explorer.exe ""http://www.qaqgame.com/""", vbHide

End Sub
