Attribute VB_Name = "mdlOtder"
Option Explicit
Public DQMLDZ As String
Public ZJTCWJ As Boolean
'------------------API�����б�------------------
Private Declare Function GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUSEX) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'------------------ȫ���Զ���-------------------
Public Enum MEMORYS_TATUS
 �ܹ��������ڴ�
 �ܹ��������ڴ�
 �ܹ��Ľ����ڴ�
 ���õ������ڴ�
 ���õ������ڴ�
 ���õĽ����ڴ�
 ���õ������ڴ�
 ���õ������ڴ�
 ���õĽ����ڴ�
 ���õ��ڴ����
End Enum
Public Enum RET_TYPE
 k
 M
End Enum
'------------------�ڲ��Զ���-------------------
Private Type LARGE_INTEGER
 LowPart As Long
 HighPart As Long
End Type
Private Type MEMORYSTATUSEX
 dwLength As Long
 dwMemoryLoad As Long
 ullTotalPhys As LARGE_INTEGER
 ullAvailPhys As LARGE_INTEGER
 ullTotalPageFile As LARGE_INTEGER
 ullAvailPageFile As LARGE_INTEGER
 ullTotalVirtual As LARGE_INTEGER
 ullAvailVirtual As LARGE_INTEGER
 ullAvailExtendedVirtual As LARGE_INTEGER
End Type
Public Type OnePixel
x As Long
y As Long
R As Integer
G As Integer
B As Integer
End Type
'----------------------------------------------

'����������������������������������������������������������������������������������������
'�����������������������������������������ӳ������������������������������������������
'����������������������������������������������������������������������������������������
Public Function GetMemoryInfo(��ȡ��Ŀ As MEMORYS_TATUS, Optional ���ص�λ As RET_TYPE) As Long
 On Error Resume Next
 Dim Memsts As MEMORYSTATUSEX
  
 Memsts.dwLength = Len(Memsts)
 Call GlobalMemoryStatusEx(Memsts)
 Select Case ��ȡ��Ŀ
 Case �ܹ��������ڴ�
 GetMemoryInfo = LargeIntToLong(Memsts.ullTotalPhys) \ (1024 ^ ���ص�λ)
 Case �ܹ��������ڴ�
 GetMemoryInfo = LargeIntToLong(Memsts.ullTotalVirtual) \ (1024 ^ ���ص�λ)
 Case �ܹ��Ľ����ڴ�
 GetMemoryInfo = LargeIntToLong(Memsts.ullTotalPageFile) \ (1024 ^ ���ص�λ)
 Case ���õ������ڴ�
 GetMemoryInfo = LargeIntToLong(Memsts.ullAvailPhys) \ (1024 ^ ���ص�λ)
 Case ���õ������ڴ�
 GetMemoryInfo = LargeIntToLong(Memsts.ullAvailVirtual) \ (1024 ^ ���ص�λ)
 Case ���õĽ����ڴ�
 GetMemoryInfo = LargeIntToLong(Memsts.ullAvailPageFile) \ (1024 ^ ���ص�λ)
 Case ���õ������ڴ�
 GetMemoryInfo = (LargeIntToLong(Memsts.ullTotalPhys) - LargeIntToLong(Memsts.ullAvailPhys)) \ (1024 ^ ���ص�λ)
 Case ���õ������ڴ�
 GetMemoryInfo = (LargeIntToLong(Memsts.ullTotalVirtual) - LargeIntToLong(Memsts.ullAvailVirtual)) \ (1024 ^ ���ص�λ)
 Case ���õĽ����ڴ�
 GetMemoryInfo = (LargeIntToLong(Memsts.ullTotalPageFile) - LargeIntToLong(Memsts.ullAvailPageFile)) \ (1024 ^ ���ص�λ)
 Case ���õ��ڴ����
 GetMemoryInfo = Memsts.dwMemoryLoad
 End Select
End Function

Private Function LargeIntToLong(liInput As LARGE_INTEGER) As Long
 Dim TmpVal As Currency
 Call CopyMemory(TmpVal, liInput, LenB(liInput))
 LargeIntToLong = CLng(Int(TmpVal * 10000 / 1024))
End Function

Public Function OpenTextAPrint(DZ As String, Wb As String)
On Error Resume Next
    Open DZ For Append As #1
    Print #1, Wb & vbCrLf
    Close #1
End Function

Public Function OpenTextAGET(DZ As String) As String
On Error Resume Next
Dim TempFile As Long
Dim LoadBytes() As Byte
TempFile = FreeFile
Open DZ For Binary As #TempFile
ReDim LoadBytes(0 To LOF(TempFile) - 1) As Byte
Get #TempFile, , LoadBytes
Close TempFile
OpenTextAGET = StrConv(LoadBytes, vbUnicode)
End Function

Public Function PhotoColor(ByRef Photo As Object, ByVal CPoints_X As Long, ByVal CPoints_Y As Long)
    Dim ThePoints() As OnePixel
    Dim DensityX As Long, DensityY As Long '�ܶ�
    Dim i&, j&, k&
    Dim ALLr&, ALLg&, ALLb&
    DensityX = Int(Photo.Picture.Width / CPoints_X)
    DensityY = Int(Photo.Picture.Height / CPoints_Y)
    For i = 1 To CPoints_Y
        For j = 1 To CPoints_X
            ReDim Preserve ThePoints(k)
            ThePoints(k).x = j * DensityX
            ThePoints(k).y = i * DensityY
            k = k + 1
        Next j
    Next i
    Dim tmpcolor&
    For k = LBound(ThePoints) To UBound(ThePoints)
        tmpcolor = Photo.POINT(ThePoints(k).x, ThePoints(k).y)
        ThePoints(k).R = (tmpcolor And &HFF&)
        ThePoints(k).G = (tmpcolor And &HFF00&) \ 256&
        ThePoints(k).B = (tmpcolor And &HFF0000) \ 65536
        ALLr = ALLr + ThePoints(k).R
        ALLg = ALLg + ThePoints(k).G
        ALLb = ALLb + ThePoints(k).B
    Next k
    ALLr = Int(ALLr / (CPoints_X / CPoints_Y))
    ALLg = Int(ALLg / (CPoints_X / CPoints_Y))
    ALLb = Int(ALLb / (CPoints_X / CPoints_Y))
    PhotoColor = RGB(ALLr, ALLg, ALLb)
End Function
