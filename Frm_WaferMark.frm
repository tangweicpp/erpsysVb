VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_WaferMark 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   11145
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   16080
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
   ScaleHeight     =   11145
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin TabDlg.SSTab SST 
      Height          =   12615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20535
      _ExtentX        =   36221
      _ExtentY        =   22251
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "上传"
      TabPicture(0)   =   "Frm_WaferMark.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPATH"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblID(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblPKG_ID(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPath(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Fps(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DTPicker4(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DTPicker3(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CommonDialog1(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtText4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtText3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmd(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdCommand1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Frm_WaferMark.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Frm_WaferMark.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton cmdCommand1 
         Caption         =   "查询"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2640
         TabIndex        =   37
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "上传"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Index           =   3
         Left            =   480
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtText3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   14
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox txtText4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtText1 
         Height          =   375
         Left            =   -73680
         TabIndex        =   12
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtpkg 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   -73680
         TabIndex        =   11
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtdn 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -73680
         TabIndex        =   10
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox txtpo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -73680
         TabIndex        =   9
         Top             =   3000
         Width           =   3255
      End
      Begin VB.TextBox txtpn 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -73680
         TabIndex        =   8
         Top             =   3720
         Width           =   3255
      End
      Begin VB.TextBox txtmpn 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -73680
         TabIndex        =   7
         Top             =   4440
         Width           =   3255
      End
      Begin VB.TextBox txtqty 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -73680
         TabIndex        =   6
         Top             =   5160
         Width           =   3255
      End
      Begin VB.TextBox txtcarton 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -73680
         TabIndex        =   5
         Top             =   5880
         Width           =   3255
      End
      Begin VB.TextBox txtJ 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   -70080
         TabIndex        =   4
         Top             =   7395
         Width           =   8175
      End
      Begin VB.CommandButton cmd 
         Caption         =   "确认"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Index           =   1
         Left            =   -60720
         TabIndex        =   3
         Top             =   7680
         Width           =   1695
      End
      Begin VB.CommandButton cmd 
         Caption         =   "清空当前数据"
         Height          =   600
         Index           =   0
         Left            =   -74640
         TabIndex        =   2
         Top             =   11640
         Width           =   1335
      End
      Begin VB.TextBox txtText2 
         Height          =   615
         Left            =   -57480
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Index           =   1
         Left            =   11760
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   16
         Top             =   3000
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   65280
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16777215
         Format          =   106561537
         CurrentDate     =   43271
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Index           =   0
         Left            =   6000
         TabIndex        =   17
         Top             =   3000
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   65280
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16777215
         Format          =   106561537
         CurrentDate     =   43271
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   8055
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   4320
         Width           =   16095
         _Version        =   524288
         _ExtentX        =   28390
         _ExtentY        =   14208
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   0
         SpreadDesigner  =   "Frm_WaferMark.frx":0054
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   6735
         Index           =   0
         Left            =   -70320
         TabIndex        =   19
         Top             =   480
         Width           =   11295
         _Version        =   524288
         _ExtentX        =   19923
         _ExtentY        =   11880
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   0
         SpreadDesigner  =   "Frm_WaferMark.frx":0536
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   5655
         Index           =   2
         Left            =   -58920
         TabIndex        =   20
         Top             =   1560
         Width           =   4215
         _Version        =   524288
         _ExtentX        =   7435
         _ExtentY        =   9975
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   0
         SpreadDesigner  =   "Frm_WaferMark.frx":0A18
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin MSForms.TextBox txtPath 
         Height          =   315
         Index           =   1
         Left            =   5880
         TabIndex        =   36
         Top             =   1320
         Width           =   5655
         VariousPropertyBits=   746604563
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "9975;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblPKG_ID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOT_ID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   35
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   34
         Top             =   2040
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   3
         Left            =   360
         TabIndex        =   33
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间："
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   4680
         TabIndex        =   32
         Top             =   3000
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblPATH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PATH:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   31
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描框:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -74760
         TabIndex        =   30
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblPKG_ID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PKG_ID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -74880
         TabIndex        =   29
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label lblPN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PN:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74280
         TabIndex        =   28
         Top             =   3720
         Width           =   435
      End
      Begin VB.Label lblQTY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTY:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74520
         TabIndex        =   27
         Top             =   5160
         Width           =   600
      End
      Begin VB.Label lblMPN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MPN:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74520
         TabIndex        =   26
         Top             =   4440
         Width           =   645
      End
      Begin VB.Label lblPO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74280
         TabIndex        =   25
         Top             =   3000
         Width           =   435
      End
      Begin VB.Label lblDN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74280
         TabIndex        =   24
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label lblCARTON 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARTON:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74880
         TabIndex        =   23
         Top             =   5880
         Width           =   1140
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   -73560
         TabIndex        =   22
         Top             =   7560
         Width           =   615
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   1085
         _cy             =   873
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " 核对记录"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -58920
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Frm_WaferMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmd_Click(Index As Integer)


  CommonDialog1(1).Filter = "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
            CommonDialog1(1).ShowOpen
            
            If CommonDialog1(1).filename = "" Then
                Exit Sub

            End If

            txtPath(1).text = CommonDialog1(1).filename
    
            CommonDialog1(1).filename = ""
            
            If txtPath(1).text = "" Then
                MsgBox "请选择要上传的文件", vbInformation, "提示"
                Exit Sub

            End If
            
                    Call Upload_0


End Sub

Private Sub cmdCommand1_Click()

Query1

End Sub

Private Sub Form_Load()
' Me.WindowState = 2

 DTPicker3(1).Value = Format(Now() - 30, "YYYY-MM-DD")
 DTPicker4(0).Value = Format(Now(), "YYYY-MM-DD")
 
 
End Sub




Private Sub Upload_0()




 On Error GoTo ErrHandle

 

    Dim VBExcel     As Excel.Application

    Dim xlBook      As Excel.Workbook

    Dim xlSheet     As Excel.Worksheet
    
    Dim I As Integer
    Dim PO  As String
    Dim Fab_Device As String
    Dim Customer_Device  As String
    Dim Lot_id  As String
    Dim Lot_id_new  As String
    Dim wafer_id  As String
    Dim REMARK As String
    Dim seqnum As Integer
    

    Dim User As String
    Dim rs         As New ADODB.Recordset
    Dim rswafer         As New ADODB.Recordset
    Dim rsseq         As New ADODB.Recordset
    
    Dim strsql    As String
    Dim strSqlin     As String
    Dim strSqlin1     As String
    Dim strseq As String

    User = gUserName
    
    Fps(1).MaxRows = 0
    
    
    strseq = " select seq_mpswmark.nextval from dual "
    If rsseq.State = adStateOpen Then rsseq.Close
    rsseq.Open strseq, Cnn, adOpenStatic, adLockReadOnly, adCmdText
     
    seqnum = rsseq.Fields(0).Value
    
        
    Set VBExcel = CreateObject("excel.application")
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(txtPath(1).text)
    Set xlSheet = xlBook.Worksheets(1)
    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 21 Then
        MsgBox "Excel中的列数和设定的模版列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        GoTo EXITPRO
        Exit Sub

    End If
    
    Fps(1).MaxRows = 0
    
    For I = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
       PO = Replace(Trim(xlSheet.Range("B" & I)), Chr(13) + Chr(10), "")
       Fab_Device = Replace(Trim(xlSheet.Range("E" & I)), Chr(13) + Chr(10), "")
       Customer_Device = Replace(Trim(xlSheet.Range("F" & I)), Chr(13) + Chr(10), "")
       Lot_id_new = Replace(Trim(xlSheet.Range("J" & I)), Chr(13) + Chr(10), "")
       
       If Len(Replace(Trim(xlSheet.Range("K" & I)), Chr(13) + Chr(10), "")) = 1 Then
        wafer_id = Lot_id_new & "0" & Replace(Trim(xlSheet.Range("K" & I)), Chr(13) + Chr(10), "")
        Else
         wafer_id = Lot_id_new & Replace(Trim(xlSheet.Range("K" & I)), Chr(13) + Chr(10), "")
        End If
        
       REMARK = Replace(Trim(xlSheet.Range("U" & I)), Chr(13) + Chr(10), "")
       
       If Len(REMARK) >= 2 Then
          Lot_id = Replace(Lot_id_new, REMARK, "")
          wafer_id = Replace(wafer_id, REMARK, "")
       Else
          Lot_id = Lot_id_new
       End If

       strsql = " select a.substrateid,nvl(b.wafer_id,' ')  from  mappingdatatest a  left join mps_mark b " & _
                " on b.wafer_id = a.substrateid  where a.substrateid = '" & wafer_id & "' "
       
         If rswafer.State = adStateOpen Then rswafer.Close
         
         rswafer.Open strsql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
         
         If Not rswafer.EOF Then
         
         If Len(Trim(rswafer.Fields(0))) > 0 And Trim(rswafer.Fields(1)) <> Trim(rswafer.Fields(0)) Then
         
         strSqlin = " insert into mps_mark (customer,po,fab_device,customer_device,lot,wafer_id,remark,create_by,create_date,REMARK1,flag,REMARK2 )" & _
                  " values('68','" & PO & "','" & Fab_Device & "','" & Customer_Device & "','" & Lot_id & "','" & wafer_id & "','" & REMARK & "','" & User & "',sysdate,'" & seqnum & "',0,'" & Lot_id_new & "')"
       
          strSqlin1 = " insert into erptemp..mps_mark (customer,po,fab_device,customer_device,lot,wafer_id,remark,create_by,create_date,REMARK1,flag,REMARK2 )" & _
                  " values('68','" & PO & "','" & Fab_Device & "','" & Customer_Device & "','" & Lot_id & "','" & wafer_id & "','" & REMARK & "','" & User & "', GETDATE(),'" & seqnum & "',0,'" & Lot_id_new & "')"
       
       
        AddSql (strSqlin)
        AddSql2 (strSqlin1)
        
         Else
            
             MsgBox "wafer_id " & wafer_id & "已存在异常标识，请确认！", vbInformation, "提示"
             Exit Sub
            
         End If
         
         Else

         MsgBox "wafer_id " & wafer_id & "已存在，请确认！", vbInformation, "提示"
         Exit Sub
            
         End If
       
  
    Next
    

 
    
  
       Query (seqnum)
    
      
    
   
   
EXITUPLOAD:

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing
   
    Exit Sub
EXITPRO:

    On Error Resume Next

    MousePointer = 0

    If Not VBExcel Is Nothing Then

        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing

    End If

    Exit Sub
ErrHandle:
    GoTo EXITPRO

End Sub


Private Sub Query(seqnum As Integer)

    Dim rs         As New ADODB.Recordset

    Dim strsql     As String
    
    Dim rs1         As New ADODB.Recordset
    
    Dim strSql1     As String
    
    Dim strflag     As String
  
    Dim strde       As String
  
    strSql1 = "  select mes_dn_pkg.WAFER_REMARK_MPS('" & seqnum & "') from dual "
    
    strflag = getStr2(strSql1)
   
    
      If strflag <> "1" Then
          
        
         strde = " delete from  mps_mark  where remark1 = '" & seqnum & "' "
         AddSql (strde)
         MsgBox "上传失败", vbInformation, "提示"
         
         Exit Sub
       
      End If
  
        strsql = " select a.*  from mps_mark a where a.remark1 = '" & seqnum & "' "
        
    
    Fps(1).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType(rs)
    Else
        Call ListDataType(rs)
        MsgBox "没有数据", vbInformation, "提示"
        Exit Sub

    End If
    
     MsgBox "上传完成", vbInformation, "提示"


End Sub


Private Sub Query1()


    Dim rs         As New ADODB.Recordset

    Dim strsql     As String
    
    If Len(Trim(txtText3.text)) = 0 Then
        
         MsgBox "请输入LOT号", vbInformation, "提示"
         Exit Sub
        
    Else
          strsql = " select a.* from mps_mark a where a.lot = '" & Trim$(txtText3.text) & "' "
        
    End If
       
      
        
    
    Fps(1).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType(rs)
    Else
        Call ListDataType(rs)
        MsgBox "没有数据", vbInformation, "提示"
        Exit Sub

    End If
    

    

End Sub



Private Sub ListDataType(rs As ADODB.Recordset)

    Dim I As Long

    With Fps(1)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    

    

End Sub












