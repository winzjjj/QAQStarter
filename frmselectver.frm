VERSION 5.00
Begin VB.Form frmselectver 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "补充你的凭据"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5130
   Icon            =   "frmselectver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5130
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cbomcversion 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmselectver.frx":000C
      Left            =   240
      List            =   "frmselectver.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   4695
   End
   Begin QAQStarter.jcbutton btnsetting 
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1260
      Width           =   915
      _extentx        =   1614
      _extenty        =   661
      buttonstyle     =   10
      font            =   "frmselectver.frx":004B
      backcolor       =   16777215
      caption         =   "确定"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
      colorscheme     =   3
   End
   Begin QAQStarter.jcbutton btncancel 
      Height          =   375
      Left            =   4020
      TabIndex        =   3
      Top             =   1260
      Width           =   915
      _extentx        =   1614
      _extenty        =   661
      buttonstyle     =   10
      font            =   "frmselectver.frx":0073
      backcolor       =   16777215
      caption         =   "取消"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
      colorscheme     =   3
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "不知道版本?点击自动尝试启动"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F8AC0E&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "告诉我们您的 Minecraft 版本"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   4635
   End
End
Attribute VB_Name = "frmselectver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
