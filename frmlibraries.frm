VERSION 5.00
Begin VB.Form frmlibraries 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "运行库管理"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6450
   Icon            =   "frmlibraries.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6450
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox thelibraries 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   2220
      TabIndex        =   7
      Top             =   1440
      Width           =   4095
   End
   Begin VB.ComboBox cmbversion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmlibraries.frx":000C
      Left            =   3900
      List            =   "frmlibraries.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin QAQStarter.jcbutton jcbutton2 
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   960
      Width           =   1815
      _extentx        =   3201
      _extenty        =   661
      buttonstyle     =   10
      font            =   "frmlibraries.frx":0010
      backcolor       =   16777215
      caption         =   "Minecraft Forge"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
      colorscheme     =   3
   End
   Begin QAQStarter.jcbutton jcbutton1 
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
      _extentx        =   3201
      _extenty        =   661
      buttonstyle     =   10
      font            =   "frmlibraries.frx":0038
      backcolor       =   16777215
      caption         =   "SkinMe"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
      colorscheme     =   3
   End
   Begin QAQStarter.jcbutton jcbutton3 
      Height          =   375
      Left            =   180
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
      _extentx        =   3201
      _extenty        =   661
      buttonstyle     =   10
      font            =   "frmlibraries.frx":0060
      backcolor       =   16777215
      caption         =   "其它运行库"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
      colorscheme     =   3
   End
   Begin QAQStarter.jcbutton jcbutton4 
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   5400
      Width           =   1395
      _extentx        =   2461
      _extenty        =   661
      buttonstyle     =   10
      font            =   "frmlibraries.frx":0088
      backcolor       =   16777215
      caption         =   "确定"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
      colorscheme     =   3
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "当前管理版本:"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F8AC0E&
      Height          =   315
      Left            =   2220
      TabIndex        =   6
      Top             =   960
      Width           =   1755
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "管理启动 Minecraft 时加载的运行库."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   540
      Width           =   5835
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "运行库(Libraries)"
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
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   3195
   End
End
Attribute VB_Name = "frmlibraries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

