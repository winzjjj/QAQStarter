Attribute VB_Name = "Functions"
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Type POINTAPI
X As Long
Y As Long
End Type

Public Function 判断鼠标是否指向指定控件上(hwn As Long) As Boolean
    Dim NowPOINT As POINTAPI
    Dim thWnd As Long
    GetCursorPos NowPOINT
    thWnd = WindowFromPoint(NowPOINT.X, NowPOINT.Y)
    If hwn = thWnd Then 判断鼠标是否指向指定控件上 = True Else: 判断鼠标是否指向指定控件上 = False
End Function
