VERSION 5.00
Begin VB.UserControl WinHttpDown 
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   InvisibleAtRuntime=   -1  'True
   Picture         =   "WinHttpDown.ctx":0000
   ScaleHeight     =   2715
   ScaleWidth      =   4590
   ToolboxBitmap   =   "WinHttpDown.ctx":0974
   Begin VB.Image Image1 
      Height          =   240
      Left            =   90
      Picture         =   "WinHttpDown.ctx":0C86
      Top             =   90
      Width           =   240
   End
End
Attribute VB_Name = "WinHttpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
'**ģ �� ����vbWinHttpDown - WinHttpDown
'**˵    ����KO ��Ȩ����2012 - 2013(C) �뱣��ԭ��
'**�� �� �ˣ�KO��461029730��
'**��    �ڣ�2012-04-25 17:06:21
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����
'**��    ����V1.0.0
'*************************************************************************

Option Explicit
Private WithEvents HttpRequest As WinHttpRequest
Attribute HttpRequest.VB_VarHelpID = -1
Dim ko As Double
Public Event HttpState(ByVal �ļ���С As String, ByVal ���� As Single, ByVal �����ٶ� As String, ByVal �����ش�С As String)
Public Event StateChanged(ByVal State As Integer)

Public FileName As String                                                       '����·��
Dim FileNamedl As String, FileNamecfg As String
Public URL As String                                                            'URL

Dim TotalLength As Double                                                       '���������صĳ���
Dim DataLenght As Double                                                        'Ҫ��ȡ���ݵ��ܳ���
Dim dtTimerStart As Double                                                      '����ʱ��,���������ٶ�
Dim Length As Double                                                            '���浱ǰ���������ĳ���

Dim FileHandle As Long                                                          '�ļ��������
Dim Down As Boolean                                                             '�ж��Ƿ������ع����� �����������Ҫ �����Ը�Ϊ����Public�ڳ����������  Down=True  '�������ع�����

Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Private Const INVALID_HANDLE_VALUE = -1
' ============================================
' ��������
' ============================================
Private Function VBStrFormatByteSize(ByVal LngSize As Double) As String
    On Error Resume Next
    Dim strSize As String * 128
    Dim strData As String
    Dim lPos        As Long
    StrFormatByteSize LngSize, strSize, 128
    lPos = InStr(1, strSize, Chr$(0))
    strData = Left$(strSize, lPos - 1)
    VBStrFormatByteSize = strData
End Function
Public Sub GetStart()
    Dim FileHandle2 As Long
    Dim dd As Double
    Dim StrUrlString As String
    If Down = True Then Exit Sub
    Down = True

    Length = 0 '��ʼ��

    RaiseEvent StateChanged(1) '��������
    
    FileNamedl = FileName & ".dl"
    FileNamecfg = FileName & ".dl.cfg" '���з�360���ص��ļ�

    If Dir(FileNamedl) <> "" And Dir(FileNamecfg) <> "" Then
        Open FileNamecfg For Binary As #1
        dd = LOF(1)
        StrUrlString = Input(dd, #1)
        Close #1
        If Trim(StrUrlString) = Trim(URL) Then '�ж���ʱ�ļ������URL �Ƿ�͵�ǰURL�Ƿ���һ��
            TotalLength = FileLen(FileNamedl) '��ȡ�����ش�С
        Else '�������ͬ ���ⴴ��cfg dl �ļ�
            TotalLength = 0
            Call GetFileDlcfg
        End If
    Else
        TotalLength = 0
        If Dir(FileNamedl) = "" And Dir(FileNamecfg) = "" Then '�������ж��� ��ֹ��һ��Ϊ�� ��һ����Ϊ��
            '************************��������ڿ�
        Else
            Call GetFileDlcfg
        End If
    End If

    On Error Resume Next
    Set HttpRequest = New WinHttpRequest
    HttpRequest.Open "GET", StrHttp(URL), True
    HttpRequest.SetRequestHeader "Range", "bytes=" & CStr(TotalLength) & "-" '�Ӷ����ֽڿ�ʼ����
    HttpRequest.Send '����Ϳ�ʼ������
    If Err.Number <> 0 Then '������ص�ַʲô����д  �ͻ���� �����������������ж�
        Down = False
        RaiseEvent StateChanged(8) '����״̬ ��������ʧ��
    End If
    On Error GoTo 0
End Sub
Private Sub GetFileDlcfg() '����FileNamedl �� FileNamecfg��ֵ
    Dim DbName As Boolean
    Dim n As Integer
    n = 1
    DbName = False
    Do While DbName = False
        If Dir(GetFileName(FileName, n) & ".dl") = "" And Dir(GetFileName(FileName, n) & ".dl.cfg") = "" Then
            FileNamedl = GetFileName(FileName, n) & ".dl"
            FileNamecfg = GetFileName(FileName, n) & ".dl.cfg"
            DbName = True
        Else
            n = n + 1
        End If
    Loop
End Sub
Private Function StrHttp(ByVal URL As String) As String                         '����ַ�Ƿ���http��ͷ
    If UCase(Mid(Trim(URL), 1, 7)) <> UCase("http://") Then
        StrHttp = "http://" & URL
    Else
        StrHttp = URL
    End If
End Function
Private Sub HttpRequest_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String) '���¼� �������ع����� �����ͻȻ�Ͽ������� �������ӳ�ʱ
    CloseHandle FileHandle '�ر��ļ�
    Down = False
    RaiseEvent StateChanged(4) '����״̬ ���ر���ֹ
End Sub
Private Sub HttpRequest_OnResponseDataAvailable(Data() As Byte)
On Error Resume Next
    Dim LoadPercent As Single
    Dim xhTrmer As Double
    Dim TimeConsuming As Double   '���غ�ʱ
    Dim dbSpeed  As Double   '�����ٶ�
    Dim dblenght As Double '��¼�����ض����ֽ�

    WriteFile FileHandle, Data(0), UBound(Data) + 1, 0, ByVal 0& 'д���ļ�

    Length = Length + UBound(Data) + 1 '��¼��ǰ�����ֽ�
    dblenght = TotalLength + Length '��¼�������ֽ�

    LoadPercent = dblenght / DataLenght * 100 '����
    If LoadPercent > 100 Then LoadPercent = 100 '����ȡ�ļ���Сʧ��ʱ DataLenght=1 �������������ж� ��Ȼ���Ȼ���ʾ������

    TimeConsuming = Timer() - dtTimerStart
    
    If TimeConsuming > 0 Then
        dbSpeed = Length / TimeConsuming '�����ٶ�
    End If

    RaiseEvent HttpState(VBStrFormatByteSize(DataLenght), LoadPercent, VBStrFormatByteSize(dbSpeed), VBStrFormatByteSize(dblenght))
    DoEvents
End Sub
Private Sub HttpRequest_OnResponseFinished() '�������
    On Error Resume Next
    Dim n As Integer
    Dim DbName As Boolean
    Dim DbFileName As String
    Call CloseHandle(FileHandle) '�ر��ļ�
    If Dir(FileNamecfg) <> "" Then
        Kill FileNamecfg 'ɾ����ʱ�ļ�
    End If
    If Dir(FileName) <> "" Then '���ж� �ļ��Ƿ���� ������� ����ԭ��·�����ļ����� +(n)
        n = 1
        DbName = False
        Do While DbName = False
            If Dir(GetFileName(FileName, n)) = "" Then
                DbFileName = GetFileName(FileName, n)
                DbName = True
            Else
                n = n + 1
            End If
        Loop
    Else
        DbFileName = FileName
    End If
    Name FileNamedl As DbFileName
    Down = False
    RaiseEvent StateChanged(9) '����״̬ �������
End Sub
Private Sub HttpRequest_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
    RaiseEvent StateChanged(2) '��ȡ�ļ���Ϣ

    If Status <> 302 And Status <> 200 And Status <> 206 And Status <> 416 Then
        CloseHandle FileHandle '�ر��ļ�
        Down = False
        RaiseEvent StateChanged(7) '����״̬ ���ӷ�����ʧ��
        Exit Sub
    End If

    If InStr(LCase(HttpRequest.GetAllResponseHeaders), LCase("Content-Length")) > 0 Then
        DataLenght = CDbl(HttpRequest.GetResponseHeader("Content-Length")) '��ȡ�ļ���С
    Else
        DataLenght = 1
    End If

    DataLenght = DataLenght + TotalLength '�ļ��ܴ�С ��ǰ�����صĴ�С+���ڻ�ȡ�Ĵ�С

    FileHandle = CreateFile(FileNamedl, &H40000000, 0, ByVal 0&, 4, 0, ByVal 0&) '�����ļ�
    If FileHandle = INVALID_HANDLE_VALUE Then '����ʧ����Ŷ
        Down = False
        RaiseEvent StateChanged(10) '����״̬ �����ļ�ʧ��
        Set HttpRequest = Nothing '�˳�����
        Exit Sub '�˳���������
    End If

    Dim FileHandle1 As Long
    Dim UrlData() As Byte
    UrlData = StrConv(URL, vbFromUnicode)
    FileHandle1 = CreateFile(FileNamecfg, &H40000000, 0, ByVal 0&, 4, 0, ByVal 0&) '������ʱ�ļ�
    If FileHandle1 = INVALID_HANDLE_VALUE Then '������ʱ�ļ�ʧ��
        Down = False
        RaiseEvent StateChanged(10) '����״̬ �����ļ�ʧ��
        Set HttpRequest = Nothing '�˳�����
        Exit Sub '�˳���������
    End If
    '���� ��ʵ���Զ�URL������ ���� ����������Ҫרҵ�� ^o^
    WriteFile FileHandle1, UrlData(0), UBound(UrlData) + 1, 0, ByVal 0& '�ѵ�ǰURL д����ʱ�ļ��б�������
    Call CloseHandle(FileHandle1) '�ر���ʱ�ļ����

    SetFilePointer FileHandle, TotalLength, ByVal 0&, 0 '��ָ���ֽڿ�ʼд���ļ�

    dtTimerStart = Timer() '��ʼ��ʱ
    RaiseEvent StateChanged(3) '����״̬ ��ʼ��������
End Sub
Public Sub GetStop() 'ֹͣ����
    If Down = True Then
        Set HttpRequest = Nothing '�˳�����
        Call CloseHandle(FileHandle) '�ر��ļ�
        If Dir(FileNamecfg) <> "" Then
            Kill FileNamecfg 'ɾ����ʱ�ļ�
        End If
        Down = False
        RaiseEvent StateChanged(5) '����״̬ ֹͣ����
    End If
End Sub
Public Sub GetPause() '��ͣ����
    If Down = True Then
        Set HttpRequest = Nothing
        CloseHandle FileHandle '�ر��ļ�
        Down = False
        RaiseEvent StateChanged(6) '����״̬ ��ͣ����
    End If
End Sub
Private Sub UserControl_Initialize()
    Down = False
End Sub
Private Sub UserControl_Resize()
    UserControl.Width = 420
    UserControl.Height = 420
End Sub
Private Sub UserControl_Terminate()
    Set HttpRequest = Nothing
    CloseHandle FileHandle '�ر��ļ�
End Sub
Private Function GetFileName(Text As String, Index As Integer) As String
    Dim I As Integer, n As Integer
    Dim A As String, B As String, C As String
    If Trim(Text) = "" Then
        GetFileName = ""
        Exit Function
    End If
    I = InStrRev(Text, "\")
    If I = 0 Then
        GetFileName = ""
        Exit Function
    End If
    n = InStrRev(Text, ".")
    A = Mid(Text, 1, I) '��ȡ�ļ�·��ǰ����
    If n = 0 Then
        B = "" 'û��ʽ
    Else
        B = Mid(Text, n) '��ȡ��ʽ
    End If
    If n > I Then '��ȡ����
        C = Mid(Text, I + 1, n - (I + 1))
    Else
        C = Mid(Text, I + 1)
    End If
    GetFileName = A & C & "(" & Index & ")" & B
End Function
