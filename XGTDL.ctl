VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl XGTDL 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "XGTDL.ctx":0000
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   1560
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   960
      Top             =   1560
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   480
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "XGTDL.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "XGTDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Type SYSTEMTIME
wYear As Integer
wMonth As Integer
wDayOfWeek As Integer
wDay As Integer
wHour As Integer
wMinute As Integer
wSecond As Integer
wMilliseconds As Integer
End Type
Dim FSSJB As String
Dim TempFile As Long
Dim YC As Boolean
Dim SFXZ As Boolean
Dim MAXDX As Long
Dim DQDX As Long
Dim XZBCDZ As String
Dim ZCXZDZ As String
Dim SFDL As Boolean
Dim SFDLDZ As String
Dim lpSystemTime As SYSTEMTIME
Dim BCDZMZ As String
Dim XZJS As Integer
Dim ZDYMMWJ As String
Dim DQXZSD As Long
Dim DQXZSD2 As Long
Event 下载进度(已下载大小 As Long, 总大小 As Long, 下载速度 As Long)
Event 下载完毕()
Event 下载失败()
Event 下载错误()

Private Function Sj(Sjs As Integer) As String
Dim i As Integer, j As Integer
Dim a As String, b As String
a = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
j = Sjs
For i = 1 To j
Randomize
b = b & Mid(a, Int((Len(a)) * Rnd + 1), 1)
Next i
Sj = b
End Function

Private Sub ZXSCWJ(Dz As String, BCDZ As String, Optional WJM As String)
On Error GoTo CuoWu
    Close #TempFile
    Winsock1.Close
    Dim DX As Long, DX2 As Long
    Dim TempFile2 As Long
    Dim LoadBytes() As Byte
    Dim LoadBytes2() As Byte
    Dim LoadBytes3() As Byte
    Dim ZCDZ As String, ZCBCDZ As String
    ZCDZ = Dz
    ZCBCDZ = BCDZ
    If Right(ZCDZ, 1) = "\" Then ZCDZ = Left(ZCDZ, Len(ZCDZ) - 1)
    If Right(ZCBCDZ, 1) = "\" Then ZCBCDZ = Left(ZCBCDZ, Len(ZCBCDZ) - 1)
    TempFile2 = FreeFile
    Open ZCDZ & "\" & BCDZMZ For Binary As #TempFile2
    DX = LOF(TempFile2)
    ReDim LoadBytes(0 To LOF(TempFile2) - 1) As Byte
    Get #TempFile2, , LoadBytes
    Close TempFile2
    Dim ZC As String
    ZC = StrConv(LoadBytes, vbUnicode)
    DX2 = Val(sMid(ZC, "Content-Length: ", vbCrLf, , , 1))
    TempFile2 = FreeFile
    ReDim LoadBytes2(0 To (DX - DX2)) As Byte
    Open ZCDZ & "\" & BCDZMZ For Binary As #TempFile2
    Get #TempFile2, , LoadBytes2
    Close TempFile2
    TempFile2 = FreeFile
    Open ZCDZ & "\" & BCDZMZ & "-2" For Output As #TempFile2
    Print #TempFile2, BytesToBstr(LoadBytes2, "UTF-8")
    Close #TempFile2
    Dim MZ As String
    If ZDYMMWJ = "" Then
    MZ = HQMZ(ZCDZ & "\" & BCDZMZ & "-2", WJM)
    Else
    MZ = ZDYMMWJ
    End If
    TempFile2 = FreeFile
    Open ZCDZ & "\" & BCDZMZ For Binary As #TempFile2
    ReDim LoadBytes3(0 To DX2 - 1) As Byte
    Get #TempFile2, DX - DX2 + 1, LoadBytes3
    Close TempFile2
    TempFile2 = FreeFile
    Open ZCBCDZ & "\" & MZ For Binary As #TempFile2
    Put #TempFile2, , LoadBytes3
    Close #TempFile2
    Kill ZCDZ & "\" & BCDZMZ
    Kill ZCDZ & "\" & BCDZMZ & "-2"
    Exit Sub
CuoWu:
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Winsock1.Close
    Close #TempFile
    RaiseEvent 下载错误
End Sub

Private Function HQMZ(Dz As String, Optional BYMZ As String) As String
On Error GoTo CuoWu
    Dim TempFile As Long
    Dim LoadBytes() As Byte
    TempFile = FreeFile
    Open Dz For Binary As #TempFile
    ReDim LoadBytes(0 To LOF(TempFile) - 1) As Byte
    Get #TempFile, , LoadBytes
    Close TempFile
    Dim FHZ As Integer
    HQMZ = sMid(StrConv(LoadBytes, vbUnicode), "Content-Disposition: attachment;filename=""", """" & vbCrLf, , , 1, FHZ)
    If FHZ = 1 Or FHZ = 2 Then FHZ = 0: HQMZ = sMid(StrConv(LoadBytes, vbUnicode), "Content-Disposition: attachment; filename*=""", """" & vbCrLf, , , 1, FHZ)
    If FHZ = 1 Or FHZ = 2 Then FHZ = 0: HQMZ = BYMZ
    Exit Function
CuoWu:
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Winsock1.Close
    Close #TempFile
    RaiseEvent 下载错误
End Function

Private Function HQMZ2(WZ As String, Optional BYMZ As String) As String
On Error GoTo CuoWu
    Dim FHZ As Integer
    HQMZ2 = sMid(WZ, "Content-Disposition: attachment;filename=""", """" & vbCrLf, , , 1, FHZ)
    If FHZ = 1 Or FHZ = 2 Then FHZ = 0: HQMZ2 = sMid(WZ, "Content-Disposition: attachment; filename*=""", """" & vbCrLf, , , 1, FHZ)
    If FHZ = 1 Or FHZ = 2 Then FHZ = 0: HQMZ2 = BYMZ
    Exit Function
CuoWu:
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Winsock1.Close
    Close #TempFile
    RaiseEvent 下载错误
End Function

Private Function BytesToBstr(strBody, CodeBase As String) As String
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

Private Function sMid(zhong As String, Optional qian As String, Optional hou As String, Optional QnH As Integer = 0, Optional QHJ As Integer = 0, Optional QK As Integer = 0, Optional FHZ As Integer) As String
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

Private Function HQWJDX(Dz As String) As Long
On Error GoTo CuoWu
    Dim TempFile As Long
    Dim LoadBytes() As Byte
    Dim ZCDZ As String
    ZCDZ = Dz
    If Right(ZCDZ, 1) = "\" Then ZCDZ = Left(ZCDZ, Len(ZCDZ) - 1)
    TempFile = FreeFile
    Open ZCDZ & "\" & BCDZMZ For Binary As #TempFile
    ReDim LoadBytes(0 To LOF(TempFile) - 1) As Byte
    Get #TempFile, , LoadBytes
    Close TempFile
    Dim ZC As String
    ZC = StrConv(LoadBytes, vbUnicode)
    HQWJDX = Val(sMid(ZC, "Content-Length: ", vbCrLf, , , 1))
    Exit Function
CuoWu:
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Winsock1.Close
    Close #TempFile
    RaiseEvent 下载错误
End Function

Public Sub 重新下载()
    SFXZ = False
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Winsock1.Close
    Close #TempFile
    If Dir(XZBCDZ & "\" & BCDZMZ) <> "" Then Kill XZBCDZ & "\" & BCDZMZ
    If Dir(XZBCDZ & "\" & BCDZMZ & "-2") <> "" Then Kill XZBCDZ & "\" & BCDZMZ & "-2"
    下载文件 ZCXZDZ, XZBCDZ, SFDL, SFDLDZ, ZDYMMWJ
End Sub

Public Sub 下载文件(地址 As String, 保存地址 As String, Optional 是否开启代理 As Boolean = False, Optional 代理IP端口 As String, Optional 自定义文件名字 As String)
On Error GoTo CuoWu
    FSSJB = ""
    TempFile = 0
    YC = False
    SFXZ = False
    MAXDX = 0
    DQDX = 0
    XZBCDZ = 保存地址
    ZCXZDZ = 地址
    SFDL = 是否开启代理
    SFDLDZ = 代理IP端口
    ZDYMMWJ = 自定义文件名字
    
    If ZDYMMWJ = "" Then
    Dim ZCAAA() As String
    ZCAAA = Split(ZCXZDZ, "/")
    ZDYMMWJ = ZCAAA(UBound(ZCAAA))
    End If
    
    If Right(XZBCDZ, 1) = "\" Then XZBCDZ = Left(XZBCDZ, Len(XZBCDZ) - 1)
    Dim Dz As String
    Dz = ZCXZDZ
    Dz = Replace(Dz, "http://", "")
    If 是否开启代理 = False Then
    FSSJB = "GET " & "/" & sMid(Dz, "/") & " HTTP/1.1" & vbCrLf
    Else
    FSSJB = "GET http://" & Dz & " HTTP/1.1" & vbCrLf
    End If
    FSSJB = FSSJB & "Host: " & sMid(Dz, , "/", , , 1) & vbCrLf
    FSSJB = FSSJB & "Connection: keep-alive" & vbCrLf
    FSSJB = FSSJB & "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8" & vbCrLf
    FSSJB = FSSJB & "User-Agent: Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/37.0.2041.4 Safari/537.36" & vbCrLf
    FSSJB = FSSJB & "Accept-Language: zh-CN,zh;q=0.8" & vbCrLf & vbCrLf
    YC = False
    SFXZ = True
    If 是否开启代理 = False Then
    Winsock1.Connect sMid(Dz, , "/", , , 1), 80
    Else
    Winsock1.Connect Split(代理IP端口, ":")(0), Split(代理IP端口, ":")(1)
    End If
    Exit Sub
CuoWu:
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Winsock1.Close
    Close #TempFile
    RaiseEvent 下载错误
End Sub

Public Sub 停止下载()
On Error GoTo CuoWu
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Winsock1.Close
    Close #TempFile
    If Dir(XZBCDZ & "\" & BCDZMZ) <> "" Then Kill XZBCDZ & "\" & BCDZMZ
    If Dir(XZBCDZ & "\" & BCDZMZ & "-2") <> "" Then Kill XZBCDZ & "\" & BCDZMZ & "-2"
    Exit Sub
CuoWu:
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Winsock1.Close
    Close #TempFile
    RaiseEvent 下载错误
End Sub

Private Sub Timer2_Timer()
XZJS = XZJS - 1
If XZJS = 0 Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Close #TempFile
ZXSCWJ XZBCDZ, XZBCDZ, ZDYMMWJ
SFXZ = False
Winsock1.Close
RaiseEvent 下载进度(MAXDX, MAXDX, DQXZSD2)
RaiseEvent 下载完毕
End If
End Sub

Private Sub Timer3_Timer()
DQXZSD2 = Int((Int(DQDX) - DQXZSD) / 1024)
DQXZSD = Int(DQDX)
End Sub

Private Sub UserControl_Resize()
UserControl.Width = Image1.Width
UserControl.Height = Image1.Height
End Sub

Private Sub Winsock1_Close()
If Winsock1.State <> 8 And SFXZ = True And Timer1.Enabled = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Winsock1.Close
Close #TempFile
RaiseEvent 下载失败
End If
End Sub

Private Sub Winsock1_Connect()
On Error GoTo CuoWu
    If SFXZ = True Then
    TempFile = FreeFile
    GetSystemTime lpSystemTime
    BCDZMZ = Replace(Replace(Replace(Now, "/", ""), " ", ""), ":", "") & Format(lpSystemTime.wMilliseconds, "000")
    BCDZMZ = BCDZMZ & "-" & Sj(32)
    Open XZBCDZ & "\" & BCDZMZ For Binary As #TempFile
    Winsock1.SendData FSSJB
    Timer1.Enabled = True
    XZJS = 5
    Timer2.Enabled = True
    Timer3.Enabled = True
    End If
    Exit Sub
CuoWu:
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Winsock1.Close
    Close #TempFile
    RaiseEvent 下载错误
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error GoTo CuoWu
    Dim ZC() As Byte
    Winsock1.GetData ZC, vbByte
    Dim ZCFHZ As String
    ZCFHZ = StrConv(ZC, vbUnicode)
    If ZDYMMWJ = "" Then
    ZDYMMWJ = HQMZ2(ZCFHZ)
    End If
    If InStr(ZCFHZ, " 302 ") <> 0 And (InStr(ZCFHZ, "Location: ") <> 0 Or InStr(ZCFHZ, "location: ") <> 0) Then
    ZCFHZ = Replace(Replace(ZCFHZ, "Location: ", "location: "), " ", "")
    SFXZ = False
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Winsock1.Close
    Close #TempFile
    If ZDYMMWJ = "" Then
    ZDYMMWJ = HQMZ2(ZCFHZ)
    End If
    If PDWJMGG(ZDYMMWJ) = True Then
    Dim ZCAAA() As String
    ZCAAA = Split(ZCXZDZ, "/")
    ZDYMMWJ = ZCAAA(UBound(ZCAAA))
    End If
    If Dir(XZBCDZ & "\" & BCDZMZ) <> "" Then Kill XZBCDZ & "\" & BCDZMZ
    If Dir(XZBCDZ & "\" & BCDZMZ & "-2") <> "" Then Kill XZBCDZ & "\" & BCDZMZ & "-2"
    下载文件 sMid(ZCFHZ, "location:", vbCrLf, , , 1), XZBCDZ, SFDL, SFDLDZ, ZDYMMWJ
    Exit Sub
    End If
    Put #TempFile, , ZC
    DQDX = DQDX + UBound(ZC)
    If YC = False Then
    YC = True
    MAXDX = HQWJDX(XZBCDZ)
    MAXDX = Int(MAXDX)
    RaiseEvent 下载进度(0, MAXDX, DQXZSD2)
    End If
    If YC = True And MAXDX >= DQDX Then XZJS = 5: RaiseEvent 下载进度(Int(DQDX), MAXDX, DQXZSD2)
    If DQDX >= MAXDX And YC = True And MAXDX <> 0 And SFXZ = True Then
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        Close #TempFile
        ZXSCWJ XZBCDZ, XZBCDZ, ZDYMMWJ
        SFXZ = False
        Winsock1.Close
        RaiseEvent 下载进度(MAXDX, MAXDX, DQXZSD2)
        RaiseEvent 下载完毕
    End If
    Exit Sub
CuoWu:
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Winsock1.Close
    Close #TempFile
    RaiseEvent 下载错误
End Sub

Private Sub Timer1_Timer()
On Error GoTo CuoWu
    If Winsock1.State = 8 And SFXZ = True Then
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Close #TempFile
    ZXSCWJ XZBCDZ, XZBCDZ, ZDYMMWJ
    SFXZ = False
    Winsock1.Close
    RaiseEvent 下载进度(MAXDX, MAXDX, DQXZSD2)
    RaiseEvent 下载完毕
    End If
    Exit Sub
CuoWu:
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Winsock1.Close
    Close #TempFile
    RaiseEvent 下载错误
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Winsock1.State <> 8 And SFXZ = True And Timer1.Enabled = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Winsock1.Close
Close #TempFile
RaiseEvent 下载失败
End If
End Sub

Private Function PDWJMGG(MZ As String) As Boolean
Dim i As Integer
Dim ZC() As String
ZC = Split("\ / : * ? "" < > |", " ")
For i = 0 To UBound(ZC)
If InStr(MZ, ZC(i)) <> 0 Then
PDWJMGG = True
Exit Function
End If
Next i
End Function
