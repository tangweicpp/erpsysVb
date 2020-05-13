VERSION 5.00
Begin VB.Form Frm_OrderResize 
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   18120
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   11640
      TabIndex        =   5
      Top             =   1080
      Width           =   3615
   End
   Begin VB.CheckBox chkNew 
      Caption         =   "新工单接口无数据"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0080FF80&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox txtSum 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   1920
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "@黑体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3000
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Label lblOrdername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "工单号 :"
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   645
   End
End
Attribute VB_Name = "Frm_OrderResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
