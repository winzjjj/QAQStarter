Attribute VB_Name = "mdlMain"
Public Const VerName = "(Beta)"
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Declare Function GetProcAddress Lib "kernel32" _
 (ByVal hModule As Long, _
 ByVal lpProcName As String) As Long

Private Declare Function GetModuleHandle Lib "kernel32" _
 Alias "GetModuleHandleA" _
 (ByVal lpModuleName As String) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" _
 () As Long
Private Declare Function IsWow64Process Lib "kernel32" _
 (ByVal hProc As Long, _
 bWow64Process As Boolean) As Long
 Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public ForcingSetting As Boolean  '是否正在强制设置
Public ForcedSetReturn As Boolean '返回的设置变量

Public strUserName As String
Public varMemory As Variant
Public strJREPath As String
Public strPassword As String
Public strColor As String
Public sl() As ServerList
Public NotOnLine As Boolean  '若设置该参数为 true，默认为离线状态

Public Type ServerList
servername As String
serverip As String
End Type

Public Sub ZXSCWJ(DZ As String, BCDZ As String, Optional WJM As String)
Dim dx As Long, DX2 As Long
Dim TempFile As Long
Dim LoadBytes() As Byte
Dim LoadBytes2() As Byte
Dim LoadBytes3() As Byte
Dim ZCDZ As String, ZCBCDZ As String
ZCDZ = DZ
ZCBCDZ = BCDZ
If Right(ZCDZ, 1) = "\" Then ZCDZ = Left(ZCDZ, Len(ZCDZ) - 1)
If Right(ZCBCDZ, 1) = "\" Then ZCBCDZ = Left(ZCBCDZ, Len(ZCBCDZ) - 1)
TempFile = FreeFile
Open ZCDZ & "\temp" For Binary As #TempFile
dx = LOF(TempFile)
ReDim LoadBytes(0 To LOF(TempFile) - 1) As Byte
Get #TempFile, , LoadBytes
Close TempFile
Dim ZC As String
ZC = StrConv(LoadBytes, vbUnicode)
DX2 = Val(sMid(ZC, "Content-Length: ", vbCrLf, , , 1))
TempFile = FreeFile
ReDim LoadBytes2(0 To (dx - DX2)) As Byte
Open ZCDZ & "\temp" For Binary As #TempFile
Get #TempFile, , LoadBytes2
Close TempFile
TempFile = FreeFile
Open ZCDZ & "\temp2" For Output As #TempFile
Print #TempFile, BytesToBstr(LoadBytes2, "UTF-8")
Close #TempFile
Dim MZ As String
MZ = HQMZ(ZCDZ & "\temp2", WJM)
TempFile = FreeFile
Open ZCDZ & "\temp" For Binary As #TempFile
ReDim LoadBytes3(0 To DX2 - 1) As Byte
Get #TempFile, dx - DX2 + 1, LoadBytes3
Close TempFile
TempFile = FreeFile
Open ZCBCDZ & "\" & MZ For Binary As #TempFile
Put #TempFile, , LoadBytes3
Close #TempFile
Kill ZCDZ & "\temp"
Kill ZCDZ & "\temp2"
End Sub

Private Function HQMZ(DZ As String, Optional BYMZ As String) As String
Dim TempFile As Long
Dim LoadBytes() As Byte
TempFile = FreeFile
Open DZ For Binary As #TempFile
ReDim LoadBytes(0 To LOF(TempFile) - 1) As Byte
Get #TempFile, , LoadBytes
Close TempFile
Dim FHZ As Integer
HQMZ = sMid(StrConv(LoadBytes, vbUnicode), "Content-Disposition: attachment;filename=""", """" & vbCrLf, , , 1, FHZ)
If FHZ = 1 Or FHZ = 2 Then HQMZ = BYMZ
End Function

Public Function BytesToBstr(strBody, CodeBase As String) As String
On Error Resume Next
Dim ObjStream
Set ObjStream = CreateObject("Adodb.Stream")
ObjStream.Type = 1
ObjStream.Mode = 3
ObjStream.Open
ObjStream.Write strBody
ObjStream.Position = 0
ObjStream.Type = 2
ObjStream.Charset = CodeBase
BytesToBstr = ObjStream.ReadText
ObjStream.Close
Set ObjStream = Nothing
End Function

Public Function sMid(zhong As String, Optional qian As String, Optional hou As String, Optional QnH As Integer = 0, Optional QHJ As Integer = 0, Optional QK As Integer = 0, Optional FHZ As Integer) As String
On Error Resume Next
DoEvents
Dim P1 As Double, P2 As Double
If zhong = "" Then sMid = "0": FHZ = 0: Exit Function
If qian <> "" And QHJ = 0 Then P1 = InStr(zhong, qian)
If qian <> "" And QHJ = 1 Then P1 = InStrRev(zhong, qian)
If qian = "" Then P1 = 1
If P1 = 0 And qian <> "" Then sMid = "1": FHZ = 1: Exit Function
If QnH = 0 And QK = 0 And hou <> "" Then P2 = InStr(zhong, hou)
If QnH = 0 And QK = 1 And hou <> "" Then P2 = InStr(P1 + Len(qian), zhong, hou)
If QnH = 1 And hou <> "" Then P2 = InStrRev(zhong, hou)
If P2 = 0 And hou <> "" Then sMid = "2": FHZ = 2: Exit Function
If P2 < P1 + Len(qian) And hou <> "" Then sMid = "0": FHZ = 0: Exit Function
If hou <> "" Then sMid = Mid(zhong, P1 + Len(qian), P2 - (P1 + Len(qian)))
If hou = "" Then sMid = Mid(zhong, P1 + Len(qian))
End Function

Public Function HQWJDX(DZ As String) As Long
Dim TempFile As Long
Dim LoadBytes() As Byte
Dim ZCDZ As String
ZCDZ = DZ
If Right(ZCDZ, 1) = "\" Then ZCDZ = Left(ZCDZ, Len(ZCDZ) - 1)
TempFile = FreeFile
Open ZCDZ & "\temp" For Binary As #TempFile
ReDim LoadBytes(0 To LOF(TempFile) - 1) As Byte
Get #TempFile, , LoadBytes
Close TempFile
Dim ZC As String
ZC = StrConv(LoadBytes, vbUnicode)
HQWJDX = Val(sMid(ZC, "Content-Length: ", vbCrLf, , , 1))
End Function


Public Function Is64bit() As Boolean
 Dim handle As Long, bolFunc As Boolean
 ' Assume initially that this is not a Wow64 process
 bolFunc = False
 ' Now check to see if IsWow64Process function exists
 handle = GetProcAddress(GetModuleHandle("kernel32"), _
 "IsWow64Process")
 If handle > 0 Then ' IsWow64Process function exists
 ' Now use the function to determine if
 ' we are running under Wow64
 IsWow64Process GetCurrentProcess(), bolFunc
 End If
 Is64bit = bolFunc
End Function


Function HttpPost(ByVal URL As String, ByVal Postmsg As String) As String
     '自己写的发送Post的函数
     '函数返回值是获得的返回信息(HTML)
     '第一个参数是要发送的Url地址
     '第二个参数是要发送的消息(键值对应，不必编码)
     Dim XmlHttp As Object
     Set XmlHttp = CreateObject("Msxml2.XMLHTTP")
     If Not IsObject(XmlHttp) Then
         Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
         If Not IsObject(XmlHttp) Then Exit Function
     End If
     XmlHttp.Open "POST", URL, False
     XmlHttp.SetRequestHeader "CONTENT-TYPE", "application/json"
     XmlHttp.Send EnUrl(Postmsg)
     Do While XmlHttp.ReadyState <> 4
         DoEvents
     Loop
     '如果把下面一行(以及后面的End IF)的注释去除，即设置为仅当返回码是200时才返回页面内容
     'If XmlHttp.Status = 200 Then
         HttpPost = XmlHttp.ResponseText
     'End If
End Function

Function EnUrl(Str)
     EnUrl = ""
     For I = 1 To Len(Str)
         ThisChr = Mid(Str, I, 1)
         If Abs(Asc(ThisChr)) < &HFF Then
             EnUrl = EnUrl & ThisChr
         Else
             innerCode = Asc(ThisChr)
             If innerCode < 0 Then
                 innerCode = innerCode + &H10000
             End If
             Hight8 = (innerCode And &HFF00) \ &HFF
             Low8 = innerCode And &HFF
             EnUrl = EnUrl & "%" & Hex(Hight8) & "%" & Hex(Low8)
         End If
     Next
End Function
'Public Function HtmlStr(Url As String)
'On Error GoTo errline
'Set xmlHTTP1 = CreateObject("Microsoft.XMLHTTP")
'xmlHTTP1.Open "get", Url, True
'xmlHTTP1.Send
'While xmlHTTP1.readyState <> 4
'DoEvents
'Wend
'HtmlStr = xmlHTTP1.responseText
'Set xmlHTTP1 = Nothing
'Exit Function
'errline:
'MsgBox "信息获取失败，请重试...", vbCritical
'End Function
Function HtmlStr$(URL$) '提取网页源码函数
On Error Resume Next
    Dim h As Object
    Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.SetTimeouts 3000, 3000, 3000, 3000
    h.Open "GET", URL, True
    h.Send
    h.WaitForResponse
    HtmlStr = h.ResponseText '这是结果
End Function
'获得系统内存(MB)
Public Function GetSystemMemoryforMB() As Long
    Dim MemStat As MEMORYSTATUS
    GlobalMemoryStatus MemStat
    GetSystemMemoryforMB = MemStat.dwTotalPhys / 1024 / 1024
End Function

'强制设置
Function ForceSetting()
    ForcingSetting = True
    frmsetting.Show 1
    If Not ForcedSetReturn Then End
End Function

Public Sub LoadSetting()
Close
    Open App.Path & "\.QAQStarter_Data\config.ini" For Input As #1
    Dim I&
    Line Input #1, strUserName
    Line Input #1, varMemory
    Line Input #1, strJREPath
    Line Input #1, strPassword
    Line Input #1, strColor
    Close #1
End Sub

Sub Main()
    If LCase(Command$) = "--notonline" Then NotOnLine = True: MsgBox "已经开启离线模式。所有涉及网络的操作都会被拒绝。"
    If Dir(App.Path & "\.QAQStarter_Data\config.ini") = "" Then
        MsgBox "欢迎使用 QAQStarter！在启动 Minecraft 之前，我们先进行一些配置。", vbInformation
        ForceSetting  '强制设置
    End If
    frmSplash.Show
    DoEvents
    LoadSetting
    Load frmmain
    DoEvents
    Unload frmSplash
    frmmain.Show
End Sub

Public Sub CheckOldVersion(ByVal MinecraftPath As String)
If Dir(MinecraftPath & "\bin\minecraft.jar") <> "" Then
If MsgBox("QAQStarter检测到有旧版本的启动文件 是否迁移？", vbYesNo + vbQuestion) = vbYes Then
NewVersionPathName = InputBox("请输入要保存为的版本名")
Open App.Path & "\MoveOldVersion.bat" For Output As #1
Print #1, "cmd.exe /c xcopy """ & App.Path & "\.minecraft\bin"" """ & App.Path & "\.minecraft\versions\" & NewVersionPathName & "\""  /f /s /q"
Print #1, "cmd.exe /c rd """ & App.Path & "\.minecraft\bin"" /s/q"
Print #1, "del %0"
Close
Shell App.Path & "\MoveOldVersion.bat", vbHide
End If
End If
End Sub



Public Function OnlineCheck(ByVal Username As String, ByVal Password As String)
验证内容 = "{""agent"": {""name"": ""Minecraft"",""version"": 1},""username"": """ & Username & """,""password"": """ & Password & """,""requestUser"":true}"
验证地址 = "https://authserver.mojang.com/authenticate"
OnlineCheck = HttpPostForJson(验证地址, 验证内容)
End Function

Public Function HttpPostForJson(ByVal URL As String, ByVal Postmsg As String) As String
'for json by winzjjj
     Dim XmlHttp As Object
     Set XmlHttp = CreateObject("Msxml2.XMLHTTP")
     If Not IsObject(XmlHttp) Then
         Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
         If Not IsObject(XmlHttp) Then Exit Function
     End If
     XmlHttp.Open "POST", URL, False
     XmlHttp.SetRequestHeader "CONTENT-TYPE", "application/json"
     XmlHttp.Send EnUrl(Postmsg)
     Do While XmlHttp.ReadyState <> 4
         DoEvents
     Loop
         HttpPostForJson = XmlHttp.ResponseText
End Function
Public Function GetJsonv(ByVal Data As String, ByVal key As String) As String
'读json函数
    Dim js
    Set js = CreateObject("scriptcontrol")
    js.Language = "javascript"
    GetJsonv = js.eval("(function(){return " & Data & "." & key & ";})()")
End Function

Public Function CheckConnect() As Boolean
    If NotOnLine Then
        CheckConnect = False
        Exit Function
    End If
    If HtmlStr("http://www.baidu.com") <> "" Then CheckConnect = True
End Function
