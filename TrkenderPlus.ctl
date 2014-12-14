VERSION 5.00
Begin VB.UserControl TrkenderPlus 
   BackColor       =   &H000000FF&
   BackStyle       =   0  '͸��
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  '��
   HitBehavior     =   0  '��
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Windowless      =   -1  'True
   Begin VB.Image Image2 
      Height          =   960
      Left            =   2520
      Top             =   1560
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   6015
      Left            =   1440
      Picture         =   "TrkenderPlus.ctx":0000
      Top             =   1560
      Width           =   10395
   End
End
Attribute VB_Name = "TrkenderPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
'�¼�����:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "���û���һ�������ϰ��²��ͷ���갴ťʱ������"

Dim A As Long
'ȱʡ����ֵ:
Const m_def_TMD = 160
'���Ա���:
Dim m_TMD As Byte

Private Sub UserControl_Paint()
    Dim DC As Long, OldBmp As Long
    DC = CreateCompatibleDC(UserControl.hdc)
        OldBmp = SelectObject(DC, Image1.Picture)
    'BitBlt UserControl.hdc, 0, 0, 64, 64, DC, 0, 0, vbSrcCopy
    AlphaBlend UserControl.hdc, 0, 0, 600, 400, DC, 0, 0, 600, 400, m_TMD * 65536
    SelectObject DC, OldBmp
    DeleteDC DC
    'UserControl.PaintPicture UserControl.Picture, 0, 0
End Sub
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'���ﴦ����Ǳ༭��ʱ����϶�����
Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    HitResult = vbHitResultHit
End Sub

'Ϊ�û��ؼ���ʼ������
Private Sub UserControl_InitProperties()
    m_TMD = m_def_TMD
End Sub

'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_TMD = PropBag.ReadProperty("TMD", m_def_TMD)
End Sub

'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("TMD", m_TMD, m_def_TMD)
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=1,0,0,160
Public Property Get TMD() As Byte
Attribute TMD.VB_Description = "TMD������͸���ȣ��������ˣ�"
    TMD = m_TMD
End Property

Public Property Let TMD(ByVal New_TMD As Byte)
    m_TMD = New_TMD
    PropertyChanged "TMD"
End Property

