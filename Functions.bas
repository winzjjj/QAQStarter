Attribute VB_Name = "Functions"
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Type POINTAPI
X As Long
Y As Long
End Type

Public Function �ж�����Ƿ�ָ��ָ���ؼ���(hwn As Long) As Boolean
    Dim NowPOINT As POINTAPI
    Dim thWnd As Long
    GetCursorPos NowPOINT
    thWnd = WindowFromPoint(NowPOINT.X, NowPOINT.Y)
    If hwn = thWnd Then �ж�����Ƿ�ָ��ָ���ؼ��� = True Else: �ж�����Ƿ�ָ��ָ���ؼ��� = False
End Function
