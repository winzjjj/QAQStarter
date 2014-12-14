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
'**模 块 名：vbWinHttpDown - WinHttpDown
'**说    明：KO 版权所有2012 - 2013(C) 请保存原著
'**创 建 人：KO【461029730】
'**日    期：2012-04-25 17:06:21
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V1.0.0
'*************************************************************************

Option Explicit
Private WithEvents HttpRequest As WinHttpRequest
Attribute HttpRequest.VB_VarHelpID = -1
Dim ko As Double
Public Event HttpState(ByVal 文件大小 As String, ByVal 进度 As Single, ByVal 下载速度 As String, ByVal 已下载大小 As String)
Public Event StateChanged(ByVal State As Integer)

Public FileName As String                                                       '保存路径
Dim FileNamedl As String, FileNamecfg As String
Public URL As String                                                            'URL

Dim TotalLength As Double                                                       '保存已下载的长度
Dim DataLenght As Double                                                        '要读取数据的总长度
Dim dtTimerStart As Double                                                      '下载时间,用来计算速度
Dim Length As Double                                                            '保存当前下载总量的长度

Dim FileHandle As Long                                                          '文件操作句柄
Dim Down As Boolean                                                             '判断是否在下载过程中 这里如程序需要 可以自改为公用Public在程序里面调用  Down=True  '正在下载过程中

Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Private Const INVALID_HANDLE_VALUE = -1
' ============================================
' 计算数据
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

    Length = 0 '初始化

    RaiseEvent StateChanged(1) '发送请求
    
    FileNamedl = FileName & ".dl"
    FileNamecfg = FileName & ".dl.cfg" '这有仿360下载的文件

    If Dir(FileNamedl) <> "" And Dir(FileNamecfg) <> "" Then
        Open FileNamecfg For Binary As #1
        dd = LOF(1)
        StrUrlString = Input(dd, #1)
        Close #1
        If Trim(StrUrlString) = Trim(URL) Then '判断临时文件里面的URL 是否和当前URL是否是一个
            TotalLength = FileLen(FileNamedl) '获取已下载大小
        Else '如果不相同 另外创建cfg dl 文件
            TotalLength = 0
            Call GetFileDlcfg
        End If
    Else
        TotalLength = 0
        If Dir(FileNamedl) = "" And Dir(FileNamecfg) = "" Then '这里再判断下 防止有一个为空 有一个不为空
            '************************如果都等于空
        Else
            Call GetFileDlcfg
        End If
    End If

    On Error Resume Next
    Set HttpRequest = New WinHttpRequest
    HttpRequest.Open "GET", StrHttp(URL), True
    HttpRequest.SetRequestHeader "Range", "bytes=" & CStr(TotalLength) & "-" '从多少字节开始下载
    HttpRequest.Send '这里就开始连接了
    If Err.Number <> 0 Then '如果下载地址什么的乱写  就会出错 所以这里做个错误判断
        Down = False
        RaiseEvent StateChanged(8) '发送状态 发送请求失败
    End If
    On Error GoTo 0
End Sub
Private Sub GetFileDlcfg() '重设FileNamedl 和 FileNamecfg的值
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
Private Function StrHttp(ByVal URL As String) As String                         '检测地址是否是http开头
    If UCase(Mid(Trim(URL), 1, 7)) <> UCase("http://") Then
        StrHttp = "http://" & URL
    Else
        StrHttp = URL
    End If
End Function
Private Sub HttpRequest_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String) '此事件 是在下载过程中 如果是突然断开了连接 或者连接超时
    CloseHandle FileHandle '关闭文件
    Down = False
    RaiseEvent StateChanged(4) '发送状态 下载被终止
End Sub
Private Sub HttpRequest_OnResponseDataAvailable(Data() As Byte)
On Error Resume Next
    Dim LoadPercent As Single
    Dim xhTrmer As Double
    Dim TimeConsuming As Double   '下载耗时
    Dim dbSpeed  As Double   '下载速度
    Dim dblenght As Double '记录总下载多少字节

    WriteFile FileHandle, Data(0), UBound(Data) + 1, 0, ByVal 0& '写入文件

    Length = Length + UBound(Data) + 1 '记录当前下载字节
    dblenght = TotalLength + Length '记录总下载字节

    LoadPercent = dblenght / DataLenght * 100 '进度
    If LoadPercent > 100 Then LoadPercent = 100 '当获取文件大小失败时 DataLenght=1 所以这里做个判断 不然进度会显示不正常

    TimeConsuming = Timer() - dtTimerStart
    
    If TimeConsuming > 0 Then
        dbSpeed = Length / TimeConsuming '下载速度
    End If

    RaiseEvent HttpState(VBStrFormatByteSize(DataLenght), LoadPercent, VBStrFormatByteSize(dbSpeed), VBStrFormatByteSize(dblenght))
    DoEvents
End Sub
Private Sub HttpRequest_OnResponseFinished() '下载完毕
    On Error Resume Next
    Dim n As Integer
    Dim DbName As Boolean
    Dim DbFileName As String
    Call CloseHandle(FileHandle) '关闭文件
    If Dir(FileNamecfg) <> "" Then
        Kill FileNamecfg '删除临时文件
    End If
    If Dir(FileName) <> "" Then '判判断 文件是否存在 如果存在 就在原来路径的文件名称 +(n)
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
    RaiseEvent StateChanged(9) '发送状态 下载完毕
End Sub
Private Sub HttpRequest_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
    RaiseEvent StateChanged(2) '获取文件信息

    If Status <> 302 And Status <> 200 And Status <> 206 And Status <> 416 Then
        CloseHandle FileHandle '关闭文件
        Down = False
        RaiseEvent StateChanged(7) '发送状态 连接服务器失败
        Exit Sub
    End If

    If InStr(LCase(HttpRequest.GetAllResponseHeaders), LCase("Content-Length")) > 0 Then
        DataLenght = CDbl(HttpRequest.GetResponseHeader("Content-Length")) '获取文件大小
    Else
        DataLenght = 1
    End If

    DataLenght = DataLenght + TotalLength '文件总大小 是前已下载的大小+现在获取的大小

    FileHandle = CreateFile(FileNamedl, &H40000000, 0, ByVal 0&, 4, 0, ByVal 0&) '创建文件
    If FileHandle = INVALID_HANDLE_VALUE Then '创建失败了哦
        Down = False
        RaiseEvent StateChanged(10) '发送状态 创建文件失败
        Set HttpRequest = Nothing '退出连接
        Exit Sub '退出整个过程
    End If

    Dim FileHandle1 As Long
    Dim UrlData() As Byte
    UrlData = StrConv(URL, vbFromUnicode)
    FileHandle1 = CreateFile(FileNamecfg, &H40000000, 0, ByVal 0&, 4, 0, ByVal 0&) '创建临时文件
    If FileHandle1 = INVALID_HANDLE_VALUE Then '创建临时文件失败
        Down = False
        RaiseEvent StateChanged(10) '发送状态 创建文件失败
        Set HttpRequest = Nothing '退出连接
        Exit Sub '退出整个过程
    End If
    '这里 其实可以对URL进行下 加密 这样看起来要专业点 ^o^
    WriteFile FileHandle1, UrlData(0), UBound(UrlData) + 1, 0, ByVal 0& '把当前URL 写进临时文件中保存起来
    Call CloseHandle(FileHandle1) '关闭临时文件句柄

    SetFilePointer FileHandle, TotalLength, ByVal 0&, 0 '从指定字节开始写入文件

    dtTimerStart = Timer() '开始计时
    RaiseEvent StateChanged(3) '发送状态 开始接收数据
End Sub
Public Sub GetStop() '停止下载
    If Down = True Then
        Set HttpRequest = Nothing '退出连接
        Call CloseHandle(FileHandle) '关闭文件
        If Dir(FileNamecfg) <> "" Then
            Kill FileNamecfg '删除临时文件
        End If
        Down = False
        RaiseEvent StateChanged(5) '发送状态 停止下载
    End If
End Sub
Public Sub GetPause() '暂停下载
    If Down = True Then
        Set HttpRequest = Nothing
        CloseHandle FileHandle '关闭文件
        Down = False
        RaiseEvent StateChanged(6) '发送状态 暂停下载
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
    CloseHandle FileHandle '关闭文件
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
    A = Mid(Text, 1, I) '获取文件路径前部分
    If n = 0 Then
        B = "" '没格式
    Else
        B = Mid(Text, n) '获取格式
    End If
    If n > I Then '获取名称
        C = Mid(Text, I + 1, n - (I + 1))
    Else
        C = Mid(Text, I + 1)
    End If
    GetFileName = A & C & "(" & Index & ")" & B
End Function
