VERSION 5.00
Begin VB.UserControl PButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F2AF00&
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   ScaleHeight     =   615
   ScaleWidth      =   1005
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.Shape Shape1 
      Height          =   618
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "PButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim C_Color_Back As OLE_COLOR
Dim C_Color_Begin As OLE_COLOR
Dim C_Color_End As OLE_COLOR
Dim C_Color_Text As OLE_COLOR
Dim C_Text As String
Dim C_Font_Name As String
Dim C_Font_Size As Integer
Dim C_Font_Bold As Boolean
Dim C_Font_Italic As Boolean
Dim C_Font_Underline As Boolean
Dim C_Enabled As Boolean

Dim MouseMoved As Boolean
Dim MouseDowned As Boolean

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub Refresh()
    UserControl.FontName = C_Font_Name
    UserControl.FontSize = C_Font_Size
    UserControl.FontBold = C_Font_Bold
    UserControl.FontItalic = C_Font_Italic
    UserControl.FontUnderline = C_Font_Underline
    Cls
    If MouseDowned = False Then
        UserControl.CurrentX = (UserControl.Width - Label1.Width) / 2
        UserControl.CurrentY = (UserControl.Height - Label1.Height) / 2
    Else
        UserControl.CurrentX = (UserControl.Width - Label1.Width) / 2 + 30
        UserControl.CurrentY = (UserControl.Height - Label1.Height) / 2 + 30
    End If
    UserControl.Print Label1
End Sub

Public Property Get Enabled() As Boolean
    Enabled = C_Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    C_Enabled = vNewValue
    PropertyChanged "Enabled"
End Property

Public Property Get Font_Name() As String
    Font_Name = C_Font_Name
End Property

Public Property Let Font_Name(ByVal vNewValue As String)
    C_Font_Name = vNewValue
    Label1.FontName = vNewValue
    Refresh
    PropertyChanged "Font_Name"
End Property

Public Property Get Font_Size() As Integer
    Font_Size = C_Font_Size
End Property

Public Property Let Font_Size(ByVal vNewValue As Integer)
    C_Font_Size = vNewValue
    Label1.FontSize = vNewValue
    Refresh
    PropertyChanged "Font_Size"
End Property

Public Property Get Font_Bold() As Boolean
    Font_Bold = C_Font_Bold
End Property

Public Property Let Font_Bold(ByVal vNewValue As Boolean)
    C_Font_Bold = vNewValue
    Label1.FontBold = vNewValue
    Refresh
    PropertyChanged "Font_Bold"
End Property

Public Property Get Font_Italic() As Boolean
    Font_Italic = C_Font_Italic
End Property

Public Property Let Font_Italic(ByVal vNewValue As Boolean)
    C_Font_Italic = vNewValue
    Label1.FontItalic = vNewValue
    Refresh
    PropertyChanged "Font_Italic"
End Property

Public Property Get Font_Underline() As Boolean
    Font_Underline = C_Font_Underline
End Property

Public Property Let Font_Underline(ByVal vNewValue As Boolean)
    C_Font_Underline = vNewValue
    Label1.FontUnderline = vNewValue
    Refresh
    PropertyChanged "Font_Underline"
End Property

Public Property Get Text() As String
    Text = C_Text
End Property

Public Property Let Text(ByVal vNewValue As String)
    C_Text = vNewValue
    Label1 = vNewValue
    Refresh
    PropertyChanged "Text"
End Property

Public Property Get Color_Back() As OLE_COLOR
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    UserControl.BackColor = vNewValue
    C_Color_Begin = C_Color_Back
    PropertyChanged "Color_Begin"
    Refresh
    PropertyChanged "Color_Back"
End Property

Public Property Get Color_Begin() As OLE_COLOR
    Color_Begin = C_Color_Begin
End Property

Public Property Let Color_Begin(ByVal vNewValue As OLE_COLOR)
    C_Color_Begin = C_Color_Back
    PropertyChanged "Color_Begin"
End Property

Public Property Get Color_End() As OLE_COLOR
    Color_End = C_Color_End
End Property

Public Property Let Color_End(ByVal vNewValue As OLE_COLOR)
    C_Color_End = vNewValue
    PropertyChanged "Color_End"
End Property

Public Property Get Color_Text() As OLE_COLOR
    Color_Text = C_Color_Text
End Property

Public Property Let Color_Text(ByVal vNewValue As OLE_COLOR)
    C_Color_Text = vNewValue
    UserControl.ForeColor = vNewValue
    Refresh
    PropertyChanged "Color_Text"
End Property

Private Sub Timer1_Timer()
    If ≈–∂œ Û±Í «∑Ò÷∏œÚ÷∏∂®øÿº˛…œ(UserControl.hwnd) = False Then
        MouseMoved = False
        Timer1.Enabled = False
        Timer2.Enabled = True
        Shape1.Visible = False
    End If
End Sub

Private Sub Timer2_Timer()
    Dim E As Long
    Dim R1 As Integer, G1 As Integer, B1 As Integer
    Dim R2 As Integer, G2 As Integer, B2 As Integer
    If MouseMoved = True Then
        E = C_Color_End
    Else
        E = C_Color_Back
    End If
    R1 = UserControl.BackColor Mod 256
    G1 = (UserControl.BackColor Mod 65536) \ 256
    B1 = UserControl.BackColor \ 65536
    R2 = E Mod 256
    G2 = (E Mod 65536) \ 256
    B2 = E \ 65536
    If R1 < R2 Then
        If R1 + 5 < R2 Then
            R1 = R1 + 5
        Else
            R1 = R2
        End If
    End If
    If R1 > R2 Then
        If R1 - 5 > R2 Then
            R1 = R1 - 5
        Else
            R1 = R2
        End If
    End If
    If G1 < G2 Then
        If G1 + 5 < G2 Then
            G1 = G1 + 5
        Else
            G1 = G2
        End If
    End If
    If G1 > G2 Then
        If G1 - 5 > G2 Then
            G1 = G1 - 5
        Else
            G1 = G2
        End If
    End If
    If B1 < B2 Then
        If B1 + 5 < B2 Then
            B1 = B1 + 5
        Else
            B1 = B2
        End If
    End If
    If B1 > B2 Then
        If B1 - 5 > B2 Then
            B1 = B1 - 5
        Else
            B1 = B2
        End If
    End If
    UserControl.BackColor = RGB(R1, G1, B1)
    Refresh
    If (R1 = R2) And (G1 = G2) And (B1 = B2) Then
        Timer2.Enabled = False
        If MouseMoved = True Then
            UserControl.BackColor = C_Color_End
        Else
            UserControl.BackColor = C_Color_Back
        End If
        Refresh
    End If
End Sub

Private Sub UserControl_Click()
    If Enabled = True Then RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    If Enabled = True Then RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    C_Enabled = True
    C_Color_Back = &HF2AF00
    C_Color_Begin = &HF2AF00
    C_Color_End = &HFF7402
    C_Color_Text = &H0&
    C_Text = "PButton"
    C_Font_Name = "Œ¢»Ì—≈∫⁄"
    C_Font_Size = 11
    C_Font_Bold = False
    C_Font_Italic = False
    C_Font_Underline = False
    Label1 = "PButton"
    Label1.FontName = "Œ¢»Ì—≈∫⁄"
    Label1.FontSize = 11
    Label1.FontBold = False
    Label1.FontItalic = False
    Label1.FontUnderline = False
    Refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Or KeyCode = 13 Then
        If Enabled = True Then
        MouseDowned = True
        Refresh
        End If
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Or KeyCode = 13 Then
        MouseDowned = False
        Refresh
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled = True Then
        MouseDowned = True
        Refresh
        RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled = True Then
        MouseMoved = True
        Timer1.Enabled = True
        Timer2.Enabled = True
        Shape1.Height = UserControl.Height
        Shape1.Width = UserControl.Width
        Shape1.BorderColor = RGB(Abs(255 - C_Color_End Mod 256), Abs(255 - (C_Color_End Mod 65536) \ 256), Abs(255 - C_Color_End \ 65536))
        Shape1.Visible = True
        RaiseEvent MouseMove(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDowned = False
    Refresh
    If Enabled = True Then RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    C_Color_Back = PropBag.ReadProperty("Color_Back", &HF2AF00)
    C_Color_Begin = PropBag.ReadProperty("Color_Begin", &HF2AF00)
    C_Color_End = PropBag.ReadProperty("Color_End", &HFF7402)
    C_Color_Text = PropBag.ReadProperty("Color_Text", &H0&)
    C_Text = PropBag.ReadProperty("Text", "PButton")
    C_Font_Name = PropBag.ReadProperty("Font_Name", "Œ¢»Ì—≈∫⁄")
    C_Font_Size = PropBag.ReadProperty("Font_Size", 11)
    C_Font_Bold = PropBag.ReadProperty("Font_Bold", False)
    C_Font_Italic = PropBag.ReadProperty("Font_Italic", False)
    C_Font_Underline = PropBag.ReadProperty("Font_Underline", False)
    C_Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.BackColor = C_Color_Back
    UserControl.ForeColor = C_Color_Text
    Label1 = C_Text
    Label1.FontName = C_Font_Name
    Label1.FontSize = C_Font_Size
    Label1.FontBold = C_Font_Bold
    Label1.FontItalic = C_Font_Italic
    Label1.FontUnderline = C_Font_Underline
    Refresh
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HF2AF00)
    Call PropBag.WriteProperty("Color_Begin", C_Color_Begin, &HF2AF00)
    Call PropBag.WriteProperty("Color_End", C_Color_End, &HFF7402)
    Call PropBag.WriteProperty("Color_Text", C_Color_Text, &H0&)
    Call PropBag.WriteProperty("Text", C_Text, "PButton")
    Call PropBag.WriteProperty("Font_Name", C_Font_Name, "Œ¢»Ì—≈∫⁄")
    Call PropBag.WriteProperty("Font_Size", C_Font_Size, 11)
    Call PropBag.WriteProperty("Font_Bold", C_Font_Bold, False)
    Call PropBag.WriteProperty("Font_Italic", C_Font_Italic, False)
    Call PropBag.WriteProperty("Font_Underline", C_Font_Underline, False)
    Call PropBag.WriteProperty("Enabled", C_Enabled, True)
End Sub
