VERSION 5.00
Begin VB.Form frmmarket 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QAQStarter �������� - ����������磡"
   ClientHeight    =   6210
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   10380
   Icon            =   "frmmarket.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   10380
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox ptab 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10380
      TabIndex        =   11
      Top             =   0
      Width           =   10380
      Begin QAQStarter.jcbutton jcbutton1 
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   450
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
         BackColor       =   16777215
         Caption         =   "�汾�ļ�"
         Mode            =   2
         Value           =   -1  'True
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin QAQStarter.jcbutton jcbutton2 
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   60
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   450
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
         BackColor       =   16777215
         Caption         =   "����Mod"
         Mode            =   2
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin QAQStarter.jcbutton jcbutton3 
         Height          =   255
         Left            =   2340
         TabIndex        =   14
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
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
         BackColor       =   16777215
         Caption         =   "���ò��ʰ�"
         Mode            =   2
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "QAQGame.com"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009D7A11&
         Height          =   255
         Left            =   8880
         TabIndex        =   19
         Top             =   60
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "������Դ�����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5460
         TabIndex        =   17
         Top             =   60
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Minecraft��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009D7A11&
         Height          =   255
         Left            =   6780
         TabIndex        =   16
         Top             =   60
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "MCBBS.net"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009D7A11&
         Height          =   255
         Left            =   7860
         TabIndex        =   15
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F5F5F5&
      Caption         =   "�����б�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5715
      Left            =   6540
      TabIndex        =   8
      Top             =   420
      Width           =   3735
      Begin QAQStarter.���ؿؼ� ���ؿؼ�1 
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   5160
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   873
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   1920
         Top             =   2160
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   180
         TabIndex        =   9
         Top             =   420
         Width           =   3375
      End
   End
   Begin VB.PictureBox pnlVersions 
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      Height          =   5835
      Left            =   0
      ScaleHeight     =   5835
      ScaleWidth      =   6495
      TabIndex        =   0
      Top             =   420
      Width           =   6495
      Begin VB.ListBox List6 
         Height          =   420
         Left            =   4920
         TabIndex        =   10
         Top             =   4740
         Visible         =   0   'False
         Width           =   1215
      End
      Begin QAQStarter.PButton PButton2 
         Height          =   735
         Left            =   3720
         TabIndex        =   7
         Top             =   2100
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   1296
         Text            =   "������Դ"
      End
      Begin QAQStarter.PButton PButton1 
         Height          =   735
         Left            =   3720
         TabIndex        =   6
         Top             =   1200
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   1296
         Text            =   "���������б�"
      End
      Begin VB.ListBox List2 
         Height          =   240
         Left            =   5160
         TabIndex        =   2
         Top             =   4440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ListBox List1 
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
         Height          =   5640
         Left            =   60
         TabIndex        =   1
         Top             =   0
         Width           =   3435
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   6960
         TabIndex        =   5
         Top             =   4680
         Width           =   5415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "�汾���ͣ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "�汾���ƣ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   120
         Width           =   5535
      End
   End
   Begin QAQStarter.RunCodeAtDesignTime RunCodeAtDesignTime1 
      Left            =   -660
      Top             =   420
      _ExtentX        =   6800
      _ExtentY        =   1614
      TPCodes         =   "pnlmods.zorder#pnlversions.zorder#frame1.zorder"
      TPObjects       =   "pnlMods#pnlVersions#frame1"
      TPRunAnIndexCode=   2
      BackColor       =   0
   End
End
Attribute VB_Name = "frmmarket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tags1, tags2, tags3, tags4, tags5, tags6, tags7 As String
Dim yans As Integer
Dim �����ش�С As Long



Private Sub Form_Load()
'��Ҫ�õĶ�׼����
tags1 = """latest"": {"
tags2 = """"
tags3 = """: """
tags4 = """id"": """
tags5 = """type"": """
tags6 = """name"": "
tags7 = """windows"": """
wlywj = HtmlStr("http://bmclapi.bangbang93.com/versions/versions.json")
If wlywj = "" Then MsgBox "����ʧ�ܣ��޷���ð汾�б�", vbExclamation: Exit Sub
'��ȡ���а汾
tmp1 = 1
Do
    tmp1 = InStr(tmp1 + 1, wlywj, tags4)
    If tmp1 = 0 Then Exit Do
    tmp2 = InStr(tmp1 + 8, wlywj, tags2)
    �汾�� = Mid(wlywj, tmp1 + 7, tmp2 - tmp1 - 7)
    tmp3 = InStr(tmp1, wlywj, tags5)
    tmp4 = InStr(tmp3 + 10, wlywj, tags2)
    �汾���� = Mid(wlywj, tmp3 + 9, tmp4 - tmp3 - 9)
    �汾���� = Replace(�汾����, "snapshot", "���հ�")
    �汾���� = Replace(�汾����, "release", "��ʽ��")
    �汾���� = Replace(�汾����, "old_alpha", "Alpha��")
    �汾���� = Replace(�汾����, "old_beta", "Beta��")
    List1.AddItem "�汾���ƣ�" & �汾�� & " �汾���ͣ�" & �汾����
    List2.AddItem �汾��
    DoEvents
Loop
List1.ListIndex = 0
End Sub




Private Sub jcbutton2_Click()
    MsgBox "��������,�����ڴ�...", vbInformation
    jcbutton1.Value = True
    jcbutton2.Value = False
End Sub

Private Sub jcbutton3_Click()
    MsgBox "�������ߣ������ڴ�...", vbInformation
    jcbutton1.Value = True
    jcbutton3.Value = False
End Sub

Private Sub Label1_Click()
Shell "cmd.exe /c start http://www.qaqgame.com/", vbHide
End Sub

Private Sub Label10_Click()
    Shell "cmd.exe /c start http://www.mcbbs.net/", vbHide
End Sub

Private Sub Label11_Click()
    Shell "cmd.exe /c start http://tieba.baidu.com/f?kw=minecraft", vbHide
End Sub

Private Sub List1_Click()
Label2.Caption = "�汾���ƣ�" & Mid(List1.List(List1.ListIndex), 6, InStr(List1.List(List1.ListIndex), " �汾���ͣ�") - 5)
Label3.Caption = "�汾���ͣ�" & Mid(List1.List(List1.ListIndex), InStr(List1.List(List1.ListIndex), " �汾���ͣ�") + 6)
End Sub

Private Sub List3_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText List3.List(List3.ListIndex)
List3.ListIndex = 0
End Sub



Private Sub PButton1_Click()
On Error Resume Next
Ҫ���صİ汾 = List2.List(List1.ListIndex)
���ص�ַ = "http://bmclapi.bangbang93.com/versions/" & Ҫ���صİ汾 & "/" & Ҫ���صİ汾 & ".jar"
�����ַ = App.Path & "\.minecraft\versions\" & Ҫ���صİ汾 & "\"
ml = "cmd.exe /c md """ & App.Path & "\.minecraft\versions\" & Ҫ���صİ汾 & """"
Shell ml
List3.AddItem ���ص�ַ & "[���ص�]" & �����ַ
If List3.ListCount = 1 Then ���ؿؼ�1.�����ļ� Mid(List3.List(0), 1, InStr(List3.List(0), "[���ص�]") - 1), Mid(List3.List(0), InStr(List3.List(0), "[���ص�]") + 5)
���ص�ַ = "http://s3.amazonaws.com/Minecraft.Download/versions/" & Ҫ���صİ汾 & "/" & Ҫ���صİ汾 & ".json"
�����ַ = App.Path & "\.minecraft\versions\" & Ҫ���صİ汾 & "\"
ml = "cmd.exe /c md """ & App.Path & "\.minecraft\versions\" & Ҫ���صİ汾 & """"
Shell ml
List3.AddItem ���ص�ַ & "[���ص�]" & �����ַ
List3.ListIndex = 0
End Sub

Private Sub PButton2_Click()
If List1.ListIndex = -1 Then MsgBox "����ѡ��汾": Exit Sub
Ҫ���صİ汾 = List2.List(List1.ListIndex)

    '��ȡ�б�

On Error Resume Next
temp1 = InStr(Ҫ���صİ汾, ".")
temp2 = InStr(temp1 + 1, Ҫ���صİ汾, ".")
temp3 = Mid(Ҫ���صİ汾, temp1 + 1)
If Val(temp3) > 7.9 Then '1.8.xͳͳ��1.8.json
�����б� = HtmlStr("http://bmclapi.bangbang93.com/indexes/1.8.json")
tmp1 = 1
Do
    tmp1 = InStr(tmp1 + 1, �����б�, "    """)
    If tmp1 = 0 Then Exit Do
    tmp2 = InStr(tmp1 + 5, �����б�, """")
    tmp3 = Mid(�����б�, tmp1 + 5, tmp2 - tmp1 - 5)
    If tmp3 = "hash" Then
    tmp4 = InStr(tmp2 + 5, �����б�, """")
    tmp5 = Mid(�����б�, tmp2 + 4, tmp4 - tmp2 - 4)
    List6.AddItem tmp5
    End If
Loop


    '����ļ�
For I = 0 To List6.ListCount - 1
    If Dir(App.Path & "\.minecraft\assets\objects\" & Left(List6.List(I), 2) & "/" & List6.List(I)) = "" Then
    ���ص�ַ = "http://bmclapi.bangbang93.com/assets/" & Left(List6.List(I), 2) & "/" & List6.List(I)
    ����λ�� = App.Path & "\.minecraft\assets\objects\" & Left(List6.List(I), 2)
    Shell "cmd.exe /c md """ & App.Path & "\.minecraft\assets\objects\" & Left(List6.List(I), 2) & "\" & """", vbHide
    List3.AddItem ���ص�ַ & "[���ص�]" & ����λ��
    End If
Next I
If Dir(App.Path & "\.minecraft\assets\indexes\1.8.json") = "" Then 'jsonҲҪ��
    Shell "cmd.exe /c md """ & App.Path & "\.minecraft\assets\indexes\1.8.json" & """", vbHide
    ���ص�ַ = "http://bmclapi.bangbang93.com/indexes/1.8.json"
    ����λ�� = App.Path & "\.minecraft\assets\indexes"
    List3.AddItem ���ص�ַ & "[���ص�]" & ����λ��
End If
If List3.ListCount >= 1 Then ���ؿؼ�1.�����ļ� Mid(List3.List(0), 1, InStr(List3.List(0), "[���ص�]") - 1), Mid(List3.List(0), InStr(List3.List(0), "[���ص�]") + 5)


ElseIf McVersion Like "??w???" Then


�����б� = HtmlStr("http://bmclapi.bangbang93.com/indexes/" & McVersion & ".json")
tmp1 = 1
Do
    tmp1 = InStr(tmp1 + 1, �����б�, "    """)
    If tmp1 = 0 Then Exit Do
    tmp2 = InStr(tmp1 + 5, �����б�, """")
    tmp3 = Mid(�����б�, tmp1 + 5, tmp2 - tmp1 - 5)
    If tmp3 = "hash" Then
    tmp4 = InStr(tmp2 + 5, �����б�, """")
    tmp5 = Mid(�����б�, tmp2 + 4, tmp4 - tmp2 - 4)
    List6.AddItem tmp5
    End If
Loop


    '����ļ�
For I = 0 To List6.ListCount - 1
    If Dir(App.Path & "\.minecraft\assets\objects\" & Left(List6.List(I), 2) & "/" & List6.List(I)) = "" Then
    ���ص�ַ = "http://bmclapi.bangbang93.com/assets/" & Left(List6.List(I), 2) & "/" & List6.List(I)
    ����λ�� = App.Path & "\.minecraft\assets\objects\" & Left(List6.List(I), 2)
    Shell "cmd.exe /c md """ & App.Path & "\.minecraft\assets\objects\" & Left(List6.List(I), 2) & "\" & """", vbHide
    List3.AddItem ���ص�ַ & "[���ص�]" & ����λ��
    End If
Next I
If Dir(App.Path & "\.minecraft\assets\indexes\" & McVersion & ".json") = "" Then 'jsonҲҪ��
    ���ص�ַ = "http://bmclapi.bangbang93.com/indexes/" & McVersion & ".json"
    ����λ�� = App.Path & "\.minecraft\assets\indexes"
    List3.AddItem ���ص�ַ & "[���ص�]" & ����λ��
End If
If List3.ListCount >= 1 Then ���ؿؼ�1.�����ļ� Mid(List3.List(0), 1, InStr(List3.List(0), "[���ص�]") - 1), Mid(List3.List(0), InStr(List3.List(0), "[���ص�]") + 5)





Else  '����ȫ��legacy
�����б� = HtmlStr("http://s3.amazonaws.com/Minecraft.Download/indexes/legacy.json")


tmp1 = 1
Do
    tmp1 = InStr(tmp1 + 1, �����б�, "    """)
    If tmp1 = 0 Then Exit Do
    tmp2 = InStr(tmp1 + 5, �����б�, """")
    tmp3 = Mid(�����б�, tmp1 + 5, tmp2 - tmp1 - 5)
    If tmp3 <> "hash" And tmp3 <> "size" Then
    List6.AddItem tmp3
    End If
Loop


    '����ļ�
For I = 0 To List6.ListCount - 1
    If Dir(App.Path & "\.minecraft\assets\" & List6.List(I)) = "" Then
    ���ص�ַ = "http://bmclapi.bangbang93.com/resources/" & List6.List(I)
    ����λ�� = App.Path & "\.minecraft\assets\" & Left(List6.List(I), InStrRev(List6.List(I), "/"))
    Shell "cmd.exe /c md """ & ����λ�� & """", vbHide
    List3.AddItem ���ص�ַ & "[���ص�]" & ����λ��
    End If
Next I
Sleep 200
If List3.ListCount >= 1 Then ���ؿؼ�1.�����ļ� Mid(List3.List(0), 1, InStr(List3.List(0), "[���ص�]") - 1), Mid(List3.List(0), InStr(List3.List(0), "[���ص�]") + 5)


End If


If List3.ListCount = 0 Then MsgBox "û����Դ��Ҫ����"
End Sub




'Private Sub tabs_Click(Index As Integer)
'        For i = 0 To 2: tabs(i).ForeColor = vbBlack: Next
'    tabs(Index).ForeColor = &HD5741C
'    Select Case Index
'    Case 0
'        pnlVersions.ZOrder
'    Case 1
'        pnlMods.ZOrder
'
'    Case 2
'        'pnlTextures.zorder
'    End Select
'    Frame1.ZOrder
'End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub ���ؿؼ�1_���ش���()
MsgBox "���ش���"
On Error Resume Next
If List3.ListCount = 1 Then
    List3.RemoveItem 0
    If Me.Caption = "Libraries������" Or Me.Caption = "Native������" Then Unload Me
Else
    List3.RemoveItem 0
    frmmain.ShowFolderList App.Path & "\.minecraft\versions\"
    Sleep 100
���ؿؼ�1.�����ļ� Replace(Mid(List3.List(0), 1, InStr(List3.List(0), "[���ص�]") - 1), "\", "/"), Mid(List3.List(0), InStr(List3.List(0), "[���ص�]") + 5)
End If
End Sub



Private Sub ���ؿؼ�1_����ʧ��()
MsgBox "����ʧ��"
On Error Resume Next
If List3.ListCount = 1 Then
    List3.RemoveItem 0
    If Me.Caption = "Libraries������" Or Me.Caption = "Native������" Then Unload Me
Else
    List3.RemoveItem 0
    frmmain.ShowFolderList App.Path & "\.minecraft\versions\"
    Sleep 100
���ؿؼ�1.�����ļ� Replace(Mid(List3.List(0), 1, InStr(List3.List(0), "[���ص�]") - 1), "\", "/"), Mid(List3.List(0), InStr(List3.List(0), "[���ص�]") + 5)
End If
End Sub

Private Sub ���ؿؼ�1_�������()
On Error Resume Next
If List3.ListCount = 1 Then
    If Me.Caption = "Libraries������" Or Me.Caption = "Native������" Then Unload Me
    If Me.Caption = "�Զ�������" Then
    Open App.Path & "\upd.bat" For Output As #1
    Print #1, "taskkill /f /im " & App.EXEName & ".exe" & vbCrLf
    Print #1, "taskkill /f /im " & App.EXEName & ".exe" & vbCrLf
    Print #1, "taskkill /f /im " & App.EXEName & ".exe" & vbCrLf
    Print #1, "del " & App.EXEName & ".exe" & vbCrLf
    Dim S() As String
    Path = Mid(List3.List(0), 1, InStr(List3.List(0), "[���ص�]") - 1)
    S = Split(Path, "/")
    FileName = S(UBound(S))
    Print #1, "ren " & FileName & " " & App.EXEName & ".exe" & vbCrLf
    Print #1, "start " & App.EXEName & ".exe" & vbCrLf
    Print #1, "del %0"
    Close
    Shell App.Path & "\upd.bat", vbHide
    End
    End If
    List3.RemoveItem 0
Else
    List3.RemoveItem 0
    frmmain.ShowFolderList App.Path & "\.minecraft\versions\"
    Sleep 100
���ؿؼ�1.�����ļ� Replace(Mid(List3.List(0), 1, InStr(List3.List(0), "[���ص�]") - 1), "\", "/"), Mid(List3.List(0), InStr(List3.List(0), "[���ص�]") + 5)
End If
End Sub


Public Sub CheckLibraries(ByVal VersionPath As String, ByVal VersionName As String)
Dim NeiR, Lip As String
If Dir(VersionPath & "\minecraft.jar") <> "" Then Unload Me: Exit Sub
Me.Caption = "Libraries������"
pnlVersions.Visible = False
Frame1.Top = 50
Frame1.Left = 50
Me.Width = Frame1.Width + 200
Me.Height = Frame1.Height + 500
ptab.Visible = False
NeiR = ""
Close
Open VersionPath & "\" & VersionName & ".json" For Input As #1
Do Until EOF(1)
    Line Input #1, Lip
    NeiR = NeiR & vbCrLf & Lip
Loop

temp1 = 1
Do
    temp1 = InStr(temp1 + 1, NeiR, "{")
    If temp1 = 0 Then Exit Do
    temp2 = InStr(temp1, NeiR, "}")
    If Mid(NeiR, temp1, temp2 - temp1) Like "*""name"": ""*:*""*" Then
    '������п�����
        temp3 = Mid(NeiR, temp1, temp2 - temp1)
        temp4 = InStr(temp3, """name"": """)
        temp5 = InStr(temp4 + 9, temp3, """")
        temp5 = Mid(temp3, temp4 + 9, temp5 - temp4 - 9)
                temp5x1 = InStr(temp5, ":")
        temp5x2 = Mid(temp5, 1, temp5x1 - 1)
        temp5x3 = Mid(temp5, temp5x1)
        temp5 = Replace(temp5x2, ".", "\") & temp5x3
        temp6 = Replace(temp5, ":", "\")
        temp7 = InStr(temp5, ":")
        temp8 = InStr(temp7 + 1, temp5, ":")
        temp9 = Mid(temp5, temp7 + 1, temp8 - temp7 - 1)
        temp10 = Mid(temp5, temp8 + 1)

        '��������temp3����
        temp12 = InStr(temp3, tags7)
        If temp12 <> 0 Then 'Ҫ����natives
            temp13 = InStr(temp12 + 13, temp3, tags2)
            temp14 = Mid(temp3, temp12 + 12, temp13 - temp12 - 12)
            If InStr(temp14, "${arch}") <> 0 Then
            '���arch Ҫȫ������(32 64��Ҫ)
                temp14x1 = Replace(temp14, "${arch}", "32")
                temp14x2 = Replace(temp14, "${arch}", "64")
                temp15 = temp6 & "\" & temp9 & "-" & temp10 & "-" & temp14x1 & ".jar"
                luj = App.Path & "\.minecraft\libraries\" & temp15
                mul = App.Path & "\.minecraft\libraries\" & temp6 & "\"
                temp16 = Dir(mul)
                If Dir(luj) = "" And temp16 = "" Then
                    List3.AddItem "http://bmclapi.bangbang93.com/libraries/" & temp15 & "[���ص�]" & App.Path & "\.minecraft\libraries\" & temp6
                    Shell "cmd.exe /c md """ & App.Path & "\.minecraft\libraries\" & temp6 & """", vbHide
                End If

                temp15 = temp6 & "\" & temp9 & "-" & temp10 & "-" & temp14x2 & ".jar"
                luj = App.Path & "\.minecraft\libraries\" & temp15
                mul = App.Path & "\.minecraft\libraries\" & temp6 & "\"
                temp16 = Dir(mul)
                If Dir(luj) = "" And temp16 = "" Then
                    List3.AddItem "http://bmclapi.bangbang93.com/libraries/" & temp15 & "[���ص�]" & App.Path & "\.minecraft\libraries\" & temp6
                    Shell "cmd.exe /c md """ & App.Path & "\.minecraft\libraries\" & temp6 & """", vbHide
                End If
                
                
            Else

            temp15 = temp6 & "\" & temp9 & "-" & temp10 & "-" & temp14 & ".jar"
            luj = App.Path & "\.minecraft\libraries\" & temp15
            mul = App.Path & "\.minecraft\libraries\" & temp6 & "\"
            temp16 = Dir(mul)
            If Dir(luj) = "" And temp16 = "" Then
            List3.AddItem "http://bmclapi.bangbang93.com/libraries/" & temp15 & "[���ص�]" & App.Path & "\.minecraft\libraries\" & temp6
            Shell "cmd.exe /c md """ & App.Path & "\.minecraft\libraries\" & temp6 & """", vbHide
            End If
        End If
        End If

        'ûnative
        temp15 = temp6 & "\" & temp9 & "-" & temp10 & ".jar"
        luj = App.Path & "\.minecraft\libraries\" & temp15
        mul = App.Path & "\.minecraft\libraries\" & temp6 & "\"
        temp16 = Dir(mul)
        If Dir(luj) = "" And temp16 = "" Then
        �����ص�ַ = "http://bmclapi.bangbang93.com/libraries/" '������˵ �����
        If InStr(temp3, """url"": """) <> 0 Then '������ �����forge�Ļ�����
        temp17 = InStr(temp3, """url"": """)
        temp18 = InStr(temp17 + 8, temp3, """")
        �����ص�ַ = Mid(temp3, temp17 + 8, temp18 - temp17 - 8)
        End If
        List3.AddItem �����ص�ַ & temp15 & "[���ص�]" & App.Path & "\.minecraft\libraries\" & temp6
        Shell "cmd.exe /c md """ & App.Path & "\.minecraft\libraries\" & temp6 & """", vbHide
        End If
    End If
Loop
If List3.ListCount = 0 Then Unload Me: Exit Sub
Dim XzDz As String
Dim BcWz As String
XzDz = Mid(List3.List(0), 1, InStr(List3.List(0), "[���ص�]") - 1)
BcWz = Mid(List3.List(0), InStr(List3.List(0), "[���ص�]") + 5)
XzDz = Replace(XzDz, "\", "/")
Sleep 1000 '����Ŀ¼�ȴ�ʱ��
���ؿؼ�1.�����ļ� XzDz, BcWz
yans = 0
Me.Hide
Me.Show 1
Exit Sub
errline:
If Err.Number = 79 Then
    MsgBox "δ�ҵ�Minecraft�汾��Json�ļ�.Libraries���ʧ��.", vbExclamation
    End
    Exit Sub
End If
End Sub

Public Sub CheckNative(ByVal NativePath As String)
If Dir(NativePath & "\avutil-ttv-51.dll") = "" Then List3.AddItem "http://qaqstarter-winzjjj.qiniudn.com/avutil-ttv-51.dll" & "[���ص�]" & NativePath
If Dir(NativePath & "\jinput-dx8.dll") = "" Then List3.AddItem "http://qaqstarter-winzjjj.qiniudn.com/jinput-dx8.dll" & "[���ص�]" & NativePath
If Dir(NativePath & "\jinput-raw.dll") = "" Then List3.AddItem "http://qaqstarter-winzjjj.qiniudn.com/jinput-raw.dll" & "[���ص�]" & NativePath
If Dir(NativePath & "\jinput-wintab.dll") = "" Then List3.AddItem "http://qaqstarter-winzjjj.qiniudn.com/jinput-wintab.dll" & "[���ص�]" & NativePath
If Dir(NativePath & "\libmp3lame-ttv.dll") = "" Then List3.AddItem "http://qaqstarter-winzjjj.qiniudn.com/libmp3lame-ttv.dll" & "[���ص�]" & NativePath
If Dir(NativePath & "\lwjgl.dll") = "" Then List3.AddItem "http://qaqstarter-winzjjj.qiniudn.com/lwjgl.dll" & "[���ص�]" & NativePath
If Dir(NativePath & "\OpenAL32.dll") = "" Then List3.AddItem "http://qaqstarter-winzjjj.qiniudn.com/OpenAL32.dll" & "[���ص�]" & NativePath
If Dir(NativePath & "\swresample-ttv-0.dll") = "" Then List3.AddItem "http://qaqstarter-winzjjj.qiniudn.com/swresample-ttv-0.dll" & "[���ص�]" & NativePath
If Dir(NativePath & "\twitchsdk.dll") = "" Then List3.AddItem "http://qaqstarter-winzjjj.qiniudn.com/twitchsdk.dll" & "[���ص�]" & NativePath
    If Dir(NativePath & "\jinput-dx8_64.dll") = "" Then List3.AddItem "http://qaqstarter-winzjjj.qiniudn.com/jinput-dx8_64.dll" & "[���ص�]" & NativePath
    If Dir(NativePath & "\jinput-raw_64.dll") = "" Then List3.AddItem "http://qaqstarter-winzjjj.qiniudn.com/jinput-raw_64.dll" & "[���ص�]" & NativePath
    If Dir(NativePath & "\lwjgl64.dll") = "" Then List3.AddItem "http://qaqstarter-winzjjj.qiniudn.com/lwjgl64.dll" & "[���ص�]" & NativePath
    If Dir(NativePath & "\OpenAL64.dll") = "" Then List3.AddItem "http://qaqstarter-winzjjj.qiniudn.com/OpenAL64.dll" & "[���ص�]" & NativePath

If List3.ListCount = 0 Then Exit Sub
Me.Caption = "Native������"
pnlVersions.Visible = False
Frame1.Top = 50
Frame1.Left = 50
Me.Width = Frame1.Width + 200
Me.Height = Frame1.Height + 500
ptab.Visible = False
Shell "cmd.exe /c md """ & NativePath & """", vbHide
Dim XzDz As String
Dim BcWz As String
XzDz = Mid(List3.List(0), 1, InStr(List3.List(0), "[���ص�]") - 1)
BcWz = Mid(List3.List(0), InStr(List3.List(0), "[���ص�]") + 5)
XzDz = Replace(XzDz, "\", "/")
Sleep 1000 '����Ŀ¼�ȴ�ʱ��
���ؿؼ�1.�����ļ� XzDz, BcWz
yans = 0
Me.Hide
Me.Show 1
End Sub



Public Sub CheckUpdate()
On Error Resume Next
Dim tmp5() As String
tmp1 = HtmlStr("http://qaqupd.jobidc.com/")
tmp2 = InStr(tmp1, "��")
tmp3 = InStr(tmp1, "��")
tmp4 = Mid(tmp1, tmp2 + 1, tmp3 - tmp2 - 1)
tmp5 = Split(tmp4, "|")
If tmp5(0) <> App.Major & "." & App.Minor & "." & App.Revision Then
    If MsgBox(Replace(tmp5(2), "[vbcrlf]", vbCrLf), vbQuestion + vbYesNo, "����QAQStarter�°汾") = vbYes Then
    List3.AddItem tmp5(1) & "[���ص�]" & App.Path
    Me.Caption = "�Զ�������"
    pnlVersions.Visible = False
    ptab.Visible = False
    Frame1.Top = 50
    Frame1.Left = 50
    Me.Width = Frame1.Width + 200
    Me.Height = Frame1.Height + 500
    ���ؿؼ�1.�����ļ� Mid(List3.List(0), 1, InStr(List3.List(0), "[���ص�]") - 1), Mid(List3.List(0), InStr(List3.List(0), "[���ص�]") + 5)
    Me.Hide
    Me.Show 1
    End If
End If

End Sub

Public Sub GetModsList(ByVal Version As String)
List4.Clear
List5.Clear
If Dir(App.Path & "\.QAQStarter_Data\modlist.tmp") = "" Then
tmp1 = HtmlStr("http://qaqupd.jobidc.com/")
tmp2 = InStr(tmp1, "-MODLISTBEGIN")
tmp3 = InStr(tmp1, "-MODLISTEND")
tmp4 = Mid(tmp1, tmp2 + 13, tmp3 - tmp2 - 13)
tmp4 = Replace(tmp4, "</p>", "")
tmp4 = Replace(tmp4, "<p>", "")
tmp4 = Replace(tmp4, "<br />", "")
Open App.Path & "\.QAQStarter_Data\modlist.tmp" For Output As #1
Print #1, tmp4
Close
End If
'��ʼ����
Close
Open App.Path & "\.QAQStarter_Data\modlist.tmp" For Input As #1
Dim temp2() As String
Dim KYDQ As Boolean
Do Until EOF(1)
Line Input #1, temp1
temp2 = Split(temp1, "|")
If temp1 = "[" & Version & "]" Then
KYDQ = True
ElseIf KYDQ = True And UBound(temp2) = 2 Then
List4.AddItem temp2(0)
List5.AddItem temp2(1) & "|" & temp2(2)
ElseIf InStr(temp1, "[") <> 0 Then
KYDQ = False
End If
Loop
Close
End Sub



