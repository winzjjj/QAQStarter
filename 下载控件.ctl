VERSION 5.00
Begin VB.UserControl ���ؿؼ� 
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8355
   ScaleHeight     =   3855
   ScaleWidth      =   8355
   ToolboxBitmap   =   "���ؿؼ�.ctx":0000
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
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "״̬"
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
         Name            =   "����"
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
         Name            =   "����"
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
Attribute VB_Name = "���ؿؼ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Event ����ʧ��()
Event �������()
Event ���ش���()
Public Function �����ļ�(URL As String, ·�������� As String) As String
    
'    arr = Split(·��������, "\")
'
'
'    '  If UBound(Split(·��������, "\")) = 1 Then
'    '   Label2 = arr(UBound(Split(·��������, "\")))
'    '  Else
'    '    Label2 = arr(UBound(Split(·��������, "\")))
'    ' End If
'
'    Label2 = arr(UBound(Split(·��������, "\")))
    tmp1 = Split(·��������, "\")
    tmp2 = Split(URL, "/")
    If tmp1(UBound(tmp1)) <> tmp2(UBound(tmp2)) Then
    If Right(·��������, 1) = "\" Then
    ·�������� = ·�������� & tmp2(UBound(tmp2))
    Else
    ·�������� = ·�������� & "\" & tmp2(UBound(tmp2))
    End If
    End If
    WinHttpDown1.FileName = ·��������
    WinHttpDown1.URL = URL
    WinHttpDown1.GetStart
End Function
Public Function ��ͣ����() As String
    WinHttpDown1.GetPause
End Function
Public Function ֹͣ����() As String
    WinHttpDown1.GetStop
End Function



Private Sub Label4_Click()

End Sub

Private Sub WinHttpDown1_HttpState(ByVal �ļ���С As String, ByVal ���� As Single, ByVal �����ٶ� As String, ByVal �����ش�С As String)
    ProgressBar1.Text = �����ش�С & "/" & �ļ���С
    Label5 = �����ٶ� & "/s"
    Label6 = Format(����, "0.00") & "%"
    ProgressBar1.Value = ����
End Sub
Private Sub WinHttpDown1_StateChanged(ByVal State As Integer)
    On Error Resume Next
    Dim msg As String
    Select Case State
    Case 1
        msg = "��������..."
    Case 2
        msg = "��ȡԶ���ļ���Ϣ..."
    Case 3
        
    Case 4
        
        msg = "���ر���ֹ..."
    Case 5
        msg = "ֹͣ����"
    Case 6
        msg = "��ͣ����"
    Case 7
        
        msg = "���ӷ�����ʧ��"
        RaiseEvent ����ʧ��
    Case 8
        msg = "��������ʧ��"
        RaiseEvent ����ʧ��
    Case 9
        
        msg = "�������"
        RaiseEvent �������
        
    Case 10
        msg = "����·������"
        RaiseEvent ����ʧ��
    End Select
    
    ProgressBar1.Text = msg
    
End Sub

