VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_HWAsn 
   BackColor       =   &H00FFFFFF&
   Caption         =   "HW_ASN"
   ClientHeight    =   12795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20595
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
   ScaleHeight     =   12795
   ScaleWidth      =   20595
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SST 
      Height          =   12615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20535
      _ExtentX        =   36221
      _ExtentY        =   22251
      _Version        =   393216
      Tab             =   1
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
      TabPicture(0)   =   "Frm_HWAsn.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtText4"
      Tab(0).Control(1)=   "txtText3"
      Tab(0).Control(2)=   "cmdup(1)"
      Tab(0).Control(3)=   "cmd(1)"
      Tab(0).Control(4)=   "CommonDialog1(1)"
      Tab(0).Control(5)=   "DTPicker3(1)"
      Tab(0).Control(6)=   "DTPicker4(0)"
      Tab(0).Control(7)=   "Fps(1)"
      Tab(0).Control(8)=   "lblPATH"
      Tab(0).Control(9)=   "lbl(1)"
      Tab(0).Control(10)=   "lbl(0)"
      Tab(0).Control(11)=   "lblID(1)"
      Tab(0).Control(12)=   "lblPKG_ID(1)"
      Tab(0).Control(13)=   "txtPath(1)"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "核对"
      TabPicture(1)   =   "Frm_HWAsn.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblscan"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblASN_PKG_ID"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblPN"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblQTY"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblMPN"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblPO"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblDN"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblCARTON"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "media"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lbllog"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Fps(2)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Fps(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtText1"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtpkg"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtdn"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtpo"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtpn"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtmpn"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtqty"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtcarton"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtJ"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "cmdConfirm"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "cmdClean"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtText2"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "打印"
      TabPicture(2)   =   "Frm_HWAsn.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox txtText2 
         Height          =   615
         Left            =   17520
         TabIndex        =   35
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdClean 
         Caption         =   "清空当前数据"
         Height          =   600
         Left            =   360
         TabIndex        =   33
         Top             =   11640
         Width           =   1335
      End
      Begin VB.CommandButton cmdConfirm 
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
         Left            =   14280
         TabIndex        =   32
         Top             =   7680
         Width           =   1695
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
         Left            =   4920
         TabIndex        =   31
         Top             =   7395
         Width           =   8175
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
         Left            =   1320
         TabIndex        =   28
         Top             =   5880
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
         Left            =   1320
         TabIndex        =   27
         Top             =   5160
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
         Left            =   1320
         TabIndex        =   25
         Top             =   4440
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
         Left            =   1320
         TabIndex        =   20
         Top             =   3720
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
         Left            =   1320
         TabIndex        =   19
         Top             =   3000
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
         Left            =   1320
         TabIndex        =   18
         Top             =   2280
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
         Left            =   1320
         TabIndex        =   17
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtText1 
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   600
         Width           =   3255
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
         Left            =   -69000
         TabIndex        =   7
         Top             =   2040
         Width           =   2775
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
         Left            =   -73320
         TabIndex        =   6
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CommandButton cmdup 
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
         Index           =   1
         Left            =   -72480
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmd 
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
         Height          =   840
         Index           =   1
         Left            =   -74520
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Index           =   1
         Left            =   -63240
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Index           =   1
         Left            =   -73320
         TabIndex        =   4
         Top             =   3000
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
         Format          =   323223553
         CurrentDate     =   43271
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Index           =   0
         Left            =   -69000
         TabIndex        =   5
         Top             =   3000
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
         Format          =   323223553
         CurrentDate     =   43271
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   8055
         Index           =   1
         Left            =   -74880
         TabIndex        =   12
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
         SpreadDesigner  =   "Frm_HWAsn.frx":0054
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   6735
         Index           =   0
         Left            =   4680
         TabIndex        =   16
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
         SpreadDesigner  =   "Frm_HWAsn.frx":058A
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   5655
         Index           =   2
         Left            =   16080
         TabIndex        =   36
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
         SpreadDesigner  =   "Frm_HWAsn.frx":0AC0
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lbllog 
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
         Left            =   16080
         TabIndex        =   37
         Top             =   960
         Width           =   1095
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   1440
         TabIndex        =   34
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
         Left            =   120
         TabIndex        =   30
         Top             =   5880
         Width           =   1140
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
         Left            =   720
         TabIndex        =   29
         Top             =   2280
         Width           =   450
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
         Left            =   720
         TabIndex        =   26
         Top             =   3000
         Width           =   435
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
         Left            =   480
         TabIndex        =   24
         Top             =   4440
         Width           =   645
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
         Left            =   480
         TabIndex        =   23
         Top             =   5160
         Width           =   600
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
         Left            =   720
         TabIndex        =   22
         Top             =   3720
         Width           =   435
      End
      Begin VB.Label lblASN_PKG_ID 
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
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label lblscan 
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
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   975
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
         Left            =   -70080
         TabIndex        =   13
         Top             =   1320
         Width           =   780
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
         Index           =   1
         Left            =   -70320
         TabIndex        =   11
         Top             =   3000
         Width           =   1275
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
         Index           =   0
         Left            =   -74640
         TabIndex        =   10
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
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
         Index           =   1
         Left            =   -69720
         TabIndex        =   9
         Top             =   2040
         Width           =   615
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
         Index           =   1
         Left            =   -74520
         TabIndex        =   8
         Top             =   2160
         Width           =   1035
      End
      Begin MSForms.TextBox txtPath 
         Height          =   315
         Index           =   1
         Left            =   -69120
         TabIndex        =   3
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
   End
End
Attribute VB_Name = "Frm_HWAsn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Enum FpsD

    e_ID
    e_PKG
    E_DN
    e_PO
    E_PN
    E_MPN
    E_QTY
    e_CARTON
    e_MCol

End Enum

Private Sub cmd_Click(Index As Integer)
Dim rs        As New ADODB.Recordset
Dim strSql    As String
Dim strSql1   As String
Dim strsql2   As String
Dim strSql3   As String
Dim strSql4   As String
Dim StartDate As String
Dim endDate   As String

strSql = ""
strSql1 = ""
strsql2 = ""
strSql3 = ""
strSql4 = ""
strSql = " select  aa.remark3 as ID, aa.asn_mfg_code, aa.asn_pkg_id,aa.asn_pn,aa.remark1 AS 箱号,aa.remark2 as DN,aa.asn_m_lot " & " ,aa.asn_qty,aa.asn_unit,aa.asn_nw,aa.asn_gw ,aa.asn_code09 ,aa.asn_po,aa.asn_remark,aa.remark5 " & " from  HW_ASN_BOX aa where 1=1 "
If Len(Trim(txtText3.Text)) > 0 Then
    strsql2 = "and aa.asn_pkg_id = '" & txtText3.Text & "' "

End If

If Len(Trim(txtText4.Text)) > 0 Then
    strsql2 = " and aa.remark2 = '" & txtText4.Text & "' "

End If

If Len(Trim(DTPicker3(1).Value)) > 0 Then
    StartDate = Format(Trim(DTPicker3(1).Value), "yyyy-mm-dd")
    strSql3 = " and to_char(aa.create_date,'YYYY-MM-DD') >= '" & StartDate & "' "

End If

If Len(Trim(DTPicker4(0).Value)) > 0 Then
    endDate = Format(Trim(DTPicker4(0).Value), "yyyy-mm-dd")
    strSql4 = " and to_char(aa.create_date,'YYYY-MM-DD') <= '" & endDate & "' "

End If

If Len(Trim(DTPicker3(1).Value)) = 0 And Len(Trim(DTPicker4(0).Value)) = 0 Then
    MsgBox "请选择开始或结束时间", vbInformation, "提示"
    Exit Sub

End If

strSql = strSql + strSql1 + strsql2 + strSql3 + strSql4
If rs.State = adStateOpen Then rs.Close
rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF Then  '表示有数据了
    Call ListDataType(rs)
Else
    Call ListDataType(rs)
    MsgBox "没有数据", vbInformation, "提示"
    Exit Sub

End If

End Sub

Private Sub Upload_0()

On Error GoTo ErrHandle

Dim VBExcel              As Excel.Application
Dim xlBook               As Excel.Workbook
Dim xlSheet              As Excel.Worksheet
Dim i                    As Integer
Dim ASN_PN               As String
Dim ASN_REV              As String
Dim ASN_CE               As String
Dim ASN_FCC              As String
Dim ASN_ROHS             As String
Dim ASN_CI               As String
Dim ASN_P                As String
Dim ASN_LEGAL_INSPECTION As String
Dim ASN_QTY              As String
Dim ASN_UNIT             As String
Dim ASN_SN_TN            As String
Dim ASN_MODEL            As String
Dim ASN_ITEMDESCCN       As String
Dim ASN_ITEMDESCEN       As String
Dim ASN_SN               As String
Dim ASN_PKG_ID           As String
Dim ASN_MPN              As String
Dim ASN_MFG_CODE         As String
Dim ASN_DATE             As String
Dim ASN_M_LOT            As String
Dim ASN_NW               As String
Dim ASN_GW               As String
Dim ASN_CODE09           As String
Dim ASN_PO               As String
Dim ASN_REMARK           As String
Dim ASN_MADE_IN          As String
Dim ASN_SUPPLIER_BARCODE As String
Dim ASN_BARCODE09        As String
Dim ASN_PRO_DATE         As String
Dim REMARK1              As String
Dim REMARK2              As String
Dim seqnum               As Integer
Dim User                 As String
Dim rsseq                As New ADODB.Recordset
Dim rs                   As New ADODB.Recordset
Dim rscarton             As New ADODB.Recordset
Dim strSql               As String
Dim strSqlin             As String
Dim strSqlin1            As String
Dim strseq               As String
Dim strcarton            As String
Dim carton               As String
Dim ASN_FA_CODE As String

User = gUserName
ASN_PN = ""
ASN_REV = ""
ASN_CE = ""
ASN_FCC = ""
ASN_ROHS = ""
ASN_CI = ""
ASN_P = ""
ASN_LEGAL_INSPECTION = ""
ASN_QTY = ""
ASN_UNIT = ""
ASN_SN_TN = ""
ASN_MODEL = ""
ASN_ITEMDESCCN = ""
ASN_ITEMDESCEN = ""
ASN_SN = ""
ASN_PKG_ID = ""
ASN_MPN = ""
ASN_MFG_CODE = ""
ASN_DATE = ""
ASN_M_LOT = ""
ASN_NW = ""
ASN_GW = ""
ASN_CODE09 = ""
ASN_PO = ""
ASN_REMARK = ""
ASN_MADE_IN = ""
ASN_SUPPLIER_BARCODE = ""
ASN_BARCODE09 = ""
ASN_PRO_DATE = ""
ASN_FA_CODE = ""
Dim strprint As String

strprint = ""
Fps(1).MaxRows = 0
Set VBExcel = CreateObject("excel.application")
VBExcel.Visible = False
Set xlBook = VBExcel.Workbooks.Open(txtPath(1).Text)
Set xlSheet = xlBook.Worksheets(1)
'    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 27 Then
'        MsgBox "Excel中的列数和设定的模版列数不一致，请确认Excel是否正确！", vbInformation, "提示"
'        GoTo EXITPRO
'        Exit Sub
'
'    End If
'
strseq = " select ASN_BOX_HW.NEXTVAL from dual "
If rsseq.State = adStateOpen Then rs.Close
rsseq.Open strseq, Cnn, adOpenStatic, adLockReadOnly, adCmdText
seqnum = rsseq.Fields(0).Value
Fps(1).MaxRows = 0
If xlSheet.Range("A1").CurrentRegion.Columns.count = 28 Then

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        ASN_PN = Replace(Trim(xlSheet.Range("A" & i)), Chr(13) + Chr(10), "")
        ASN_REV = Replace(Trim(xlSheet.Range("B" & i)), Chr(13) + Chr(10), "")
        ASN_CE = Replace(Trim(xlSheet.Range("C" & i)), Chr(13) + Chr(10), "")
        ASN_FCC = Replace(Trim(xlSheet.Range("D" & i)), Chr(13) + Chr(10), "")
        ASN_ROHS = Replace(Trim(xlSheet.Range("E" & i)), Chr(13) + Chr(10), "")
        ASN_CI = Replace(Trim(xlSheet.Range("F" & i)), Chr(13) + Chr(10), "")
        ASN_P = Replace(Trim(xlSheet.Range("G" & i)), Chr(13) + Chr(10), "")
        ASN_LEGAL_INSPECTION = Replace(Trim(xlSheet.Range("H" & i)), Chr(13) + Chr(10), "")
        ASN_QTY = Replace(Trim(xlSheet.Range("I" & i)), Chr(13) + Chr(10), "")
        ASN_UNIT = Replace(Trim(xlSheet.Range("J" & i)), Chr(13) + Chr(10), "")
        ASN_SN_TN = Replace(Trim(xlSheet.Range("K" & i)), Chr(13) + Chr(10), "")
        ASN_MODEL = Replace(Trim(xlSheet.Range("L" & i)), Chr(13) + Chr(10), "")
        ASN_ITEMDESCCN = Replace(Trim(xlSheet.Range("M" & i)), Chr(13) + Chr(10), "")
        ASN_ITEMDESCEN = Replace(Trim(xlSheet.Range("N" & i)), Chr(13) + Chr(10), "")
        ASN_SN = Replace(Trim(xlSheet.Range("O" & i)), Chr(13) + Chr(10), "")
        ASN_PKG_ID = Replace(Trim(xlSheet.Range("P" & i)), Chr(13) + Chr(10), "")
        ASN_MPN = Replace(Trim(xlSheet.Range("Q" & i)), Chr(13) + Chr(10), "")
        ASN_MFG_CODE = Replace(Trim(xlSheet.Range("R" & i)), Chr(13) + Chr(10), "")
        ASN_FA_CODE = Replace(Trim(xlSheet.Range("S" & i)), Chr(13) + Chr(10), "")
        ASN_DATE = Replace(Trim(xlSheet.Range("T" & i)), Chr(13) + Chr(10), "")
        ASN_M_LOT = Replace(Trim(xlSheet.Range("U" & i)), Chr(13) + Chr(10), "")
        ASN_NW = Replace(Trim(xlSheet.Range("V" & i)), Chr(13) + Chr(10), "")
        ASN_GW = Replace(Trim(xlSheet.Range("W" & i)), Chr(13) + Chr(10), "")
        ASN_CODE09 = Replace(Trim(xlSheet.Range("X" & i)), Chr(13) + Chr(10), "")
        ASN_PO = Replace(Trim(xlSheet.Range("Y" & i)), Chr(13) + Chr(10), "")
        ASN_REMARK = Replace(Trim(xlSheet.Range("Z" & i)), Chr(13) + Chr(10), "")
        ASN_MADE_IN = Replace(Trim(xlSheet.Range("AA" & i)), Chr(13) + Chr(10), "")
        ASN_SUPPLIER_BARCODE = Replace(Trim(xlSheet.Range("AB" & i)), Chr(13) + Chr(10), "")
        REMARK1 = Mid(ASN_REMARK, InStr(ASN_REMARK, "(") + 1, InStr(ASN_REMARK, ")") - InStr(ASN_REMARK, "(") - 1)
        REMARK2 = Mid(ASN_REMARK, InStr(ASN_REMARK, "DN#") + 3, 10)
        strcarton = " select distinct o.carton from packing_detailed o where o.dn_num = '" & REMARK2 & "' and o.kid = '" & REMARK1 & "' "
        If rscarton.State = adStateOpen Then rscarton.Close
        rscarton.Open strcarton, Cnn, adOpenStatic, adLockReadOnly, adCmdText
        If Not rscarton.EOF Then
            carton = rscarton.Fields(0).Value
        Else
            MsgBox "DN" & REMARK2 & "不存在箱号:" & REMARK1, vbInformation, "提示"
            GoTo EXITPRO
            Exit Sub

        End If

        strSql = " select * from  HW_ASN_BOX aa where aa.asn_pkg_id = '" & ASN_PKG_ID & "'"
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
        If Not rs.EOF Then  '表示有数据了
            MsgBox "pkg_id" & ASN_PKG_ID & "已存在！", vbInformation, "提示"
            GoTo EXITPRO
            Exit Sub

        End If

        strSqlin = "insert into HW_ASN_BOX  values ('" & ASN_PN & "','" & ASN_REV & "','" & ASN_CE & "','" & ASN_FCC & "','" & ASN_ROHS & "','" & ASN_CI & "','" & ASN_P & "','" & ASN_LEGAL_INSPECTION & "'" & "  ,'" & ASN_QTY & "','" & ASN_UNIT & "','" & ASN_SN_TN & "'  ,'" & ASN_MODEL & "','" & ASN_ITEMDESCCN & "' ,'" & ASN_ITEMDESCEN & "'  " & "  ,'" & ASN_SN & "' ,'" & ASN_PKG_ID & "' ,'" & ASN_MPN & "' ,'" & ASN_MFG_CODE & "' ,'" & ASN_DATE & "' ,'" & ASN_M_LOT & "' ,'" & ASN_NW & "' ,'" & ASN_GW & "' ,'" & ASN_CODE09 & "' ,'" & ASN_PO & "' " & "  , '" & ASN_REMARK & "' ,'" & ASN_MADE_IN & "' ,'" & ASN_SUPPLIER_BARCODE & "' ,'" & REMARK1 & "' ,'" & REMARK2 & "' ,'" & seqnum & "' ,'" & carton & "' ,'' ,sysdate ,'" & User & "' ,0 ,'','','" & ASN_FA_CODE & "' ) "
        strSqlin1 = "insert into  erptemp..HW_ASN_BOX  values ('" & ASN_PN & "','" & ASN_REV & "','" & ASN_CE & "','" & ASN_FCC & "','" & ASN_ROHS & "','" & ASN_CI & "','" & ASN_P & "','" & ASN_LEGAL_INSPECTION & "'" & "  ,'" & ASN_QTY & "','" & ASN_UNIT & "','" & ASN_SN_TN & "'  ,'" & ASN_MODEL & "','" & ASN_ITEMDESCCN & "' ,'" & ASN_ITEMDESCEN & "'  " & "  ,'" & ASN_SN & "' ,'" & ASN_PKG_ID & "' ,'" & ASN_MPN & "' ,'" & ASN_MFG_CODE & "' ,'" & ASN_DATE & "' ,'" & ASN_M_LOT & "' ,'" & ASN_NW & "' ,'" & ASN_GW & "' ,'" & ASN_CODE09 & "' ,'" & ASN_PO & "' " & "  , '" & ASN_REMARK & "' ,'" & ASN_MADE_IN & "' ,'" & ASN_SUPPLIER_BARCODE & "' ,'" & REMARK1 & "' ,'" & REMARK2 & "' ,'" & seqnum & "' ,'" & carton & "' ,'' ,GETDATE() ,'" & User & "' ,0 ,'','','" & ASN_FA_CODE & "'  ) "
        AddSql (strSqlin)
        AddSql2 (strSqlin1)
    Next
    MsgBox "上传完成", vbInformation, "提示"
    Query (seqnum)
ElseIf xlSheet.Range("A1").CurrentRegion.Columns.count = 26 Then

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        ASN_PKG_ID = Replace(Trim(xlSheet.Range("A" & i)), Chr(13) + Chr(10), "")
        ASN_PO = Replace(Trim(xlSheet.Range("B" & i)), Chr(13) + Chr(10), "")
        ASN_QTY = Replace(Trim(xlSheet.Range("C" & i)), Chr(13) + Chr(10), "")
        ASN_PN = Replace(Trim(xlSheet.Range("D" & i)), Chr(13) + Chr(10), "")
        ASN_MPN = Replace(Trim(xlSheet.Range("F" & i)), Chr(13) + Chr(10), "")
        ASN_ITEMDESCCN = Replace(Trim(xlSheet.Range("G" & i)), Chr(13) + Chr(10), "")
        ASN_CODE09 = Replace(Trim(xlSheet.Range("H" & i)), Chr(13) + Chr(10), "")
        ASN_BARCODE09 = Replace(Trim(xlSheet.Range("I" & i)), Chr(13) + Chr(10), "")
        ASN_MFG_CODE = Replace(Trim(xlSheet.Range("J" & i)), Chr(13) + Chr(10), "")
        ASN_ROHS = Replace(Trim(xlSheet.Range("N" & i)), Chr(13) + Chr(10), "")
        ASN_M_LOT = Replace(Trim(xlSheet.Range("P" & i)), Chr(13) + Chr(10), "")
        ASN_MADE_IN = Replace(Trim(xlSheet.Range("Q" & i)), Chr(13) + Chr(10), "")
        ASN_PRO_DATE = Replace(Trim(xlSheet.Range("R" & i)), Chr(13) + Chr(10), "")
        ASN_REMARK = Replace(Trim(xlSheet.Range("S" & i)), Chr(13) + Chr(10), "")
        ASN_UNIT = Replace(Trim(xlSheet.Range("T" & i)), Chr(13) + Chr(10), "")
        ASN_ITEMDESCEN = Replace(Trim(xlSheet.Range("W" & i)), Chr(13) + Chr(10), "")
        ASN_NW = Replace(Trim(xlSheet.Range("Y" & i)), Chr(13) + Chr(10), "")
        ASN_SN_TN = Replace(Trim(xlSheet.Range("Z" & i)), Chr(13) + Chr(10), "")
        REMARK1 = Mid(ASN_REMARK, InStr(ASN_REMARK, "(") + 1, InStr(ASN_REMARK, ")") - InStr(ASN_REMARK, "("))
        REMARK2 = Mid(ASN_REMARK, InStr(ASN_REMARK, "DN#") + 3, 10)
        strcarton = " select distinct o.carton from packing_detailed o where o.dn_num = '" & REMARK2 & "' and o.kid = '" & REMARK1 & "' "
        If rscarton.State = adStateOpen Then rscarton.Close
        rscarton.Open strcarton, Cnn, adOpenStatic, adLockReadOnly, adCmdText
        If Not rscarton.EOF Then
            carton = rscarton.Fields(0).Value
        Else
            MsgBox "DN" & REMARK2 & "不存在箱号:" & REMARK1, vbInformation, "提示"
            GoTo EXITPRO
            Exit Sub

        End If

        strSql = " select * from  HW_ASN_BOX aa where aa.asn_pkg_id = '" & ASN_PKG_ID & "'"
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
        If Not rs.EOF Then  '表示有数据了
            MsgBox "pkg_id" & ASN_PKG_ID & "已存在！", vbInformation, "提示"
            GoTo EXITPRO
            Exit Sub

        End If

        strSqlin = "insert into HW_ASN_BOX  values ('" & ASN_PN & "','" & ASN_REV & "','" & ASN_CE & "','" & ASN_FCC & "','" & ASN_ROHS & "','" & ASN_CI & "','" & ASN_P & "','" & ASN_LEGAL_INSPECTION & "'" & "  ,'" & ASN_QTY & "','" & ASN_UNIT & "','" & ASN_SN_TN & "'  ,'" & ASN_MODEL & "','" & ASN_ITEMDESCCN & "' ,'" & ASN_ITEMDESCEN & "'  " & "  ,'" & ASN_SN & "' ,'" & ASN_PKG_ID & "' ,'" & ASN_MPN & "' ,'" & ASN_MFG_CODE & "' ,'" & ASN_DATE & "' ,'" & ASN_M_LOT & "' ,'" & ASN_NW & "' ,'" & ASN_GW & "' ,'" & ASN_CODE09 & "' ,'" & ASN_PO & "' " & "  , '" & ASN_REMARK & "' ,'" & ASN_MADE_IN & "' ,'" & ASN_SUPPLIER_BARCODE & "' ,'" & REMARK1 & "' ,'" & REMARK2 & "' ,'" & seqnum & "' ,'" & carton & "' ,'' ,sysdate ,'" & User & "' ,0, '" & ASN_PRO_DATE & "' ,'" & ASN_BARCODE09 & "'  ) "
        strSqlin1 = "insert into  erptemp..HW_ASN_BOX  values ('" & ASN_PN & "','" & ASN_REV & "','" & ASN_CE & "','" & ASN_FCC & "','" & ASN_ROHS & "','" & ASN_CI & "','" & ASN_P & "','" & ASN_LEGAL_INSPECTION & "'" & "  ,'" & ASN_QTY & "','" & ASN_UNIT & "','" & ASN_SN_TN & "'  ,'" & ASN_MODEL & "','" & ASN_ITEMDESCCN & "' ,'" & ASN_ITEMDESCEN & "'  " & "  ,'" & ASN_SN & "' ,'" & ASN_PKG_ID & "' ,'" & ASN_MPN & "' ,'" & ASN_MFG_CODE & "' ,'" & ASN_DATE & "' ,'" & ASN_M_LOT & "' ,'" & ASN_NW & "' ,'" & ASN_GW & "' ,'" & ASN_CODE09 & "' ,'" & ASN_PO & "' " & "  , '" & ASN_REMARK & "' ,'" & ASN_MADE_IN & "' ,'" & ASN_SUPPLIER_BARCODE & "' ,'" & REMARK1 & "' ,'" & REMARK2 & "' ,'" & seqnum & "' ,'" & carton & "' ,'' ,GETDATE() ,'" & User & "' ,0, '" & ASN_PRO_DATE & "' ,'" & ASN_BARCODE09 & "'   ) "
        AddSql (strSqlin)
        AddSql2 (strSqlin1)
        strprint = ASN_PKG_ID + "@" + ASN_PO + "@" + ASN_MFG_CODE + "@" + ASN_PN + "@" + "@" + ASN_ROHS + "@" + ASN_QTY + "@" + ASN_NW + "@" + ASN_ITEMDESCCN + "@" + ASN_CODE09 + "@" + ASN_MPN + "@" + ASN_M_LOT + "@"
        strprint = strprint + ASN_MADE_IN + "@" + ASN_PRO_DATE + "@" + ASN_REMARK + "@" + ASN_ITEMDESCEN + "@" + ASN_UNIT + "@" + ASN_SN_TN + "@" + ASN_BARCODE09 + "@" + vbCrLf
    Next
    Call addLabelTxt("HW" + Trim(Str(seqnum)), strprint, "\\10.160.1.14\BarCode\37\37HWOUTold\")
    MsgBox "上传完成", vbInformation, "提示"
    Query (seqnum)
Else
    MsgBox "Excel中的列数和设定的模版列数不一致，请确认Excel是否正确！", vbInformation, "提示"
    GoTo EXITPRO
    Exit Sub

End If

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
Dim rs     As New ADODB.Recordset
Dim strSql As String

strSql = " select  aa.remark3 as ID, aa.asn_mfg_code, aa.asn_pkg_id,aa.asn_pn,aa.remark1 AS 箱号 ,aa.remark2 as DN ,aa.asn_m_lot " & " ,aa.asn_qty,aa.asn_unit,aa.asn_nw,aa.asn_gw ,aa.asn_code09 ,aa.asn_po,aa.asn_remark " & " from  HW_ASN_BOX aa where aa.remark3 = '" & seqnum & "' "
Fps(1).MaxRows = 0
If rs.State = adStateOpen Then rs.Close
rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF Then  '表示有数据了
    Call ListDataType(rs)
Else
    Call ListDataType(rs)
    MsgBox "没有数据", vbInformation, "提示"
    Exit Sub

End If

End Sub

Private Sub ListDataType(rs As ADODB.Recordset)
Dim i As Long

With Fps(1)
    .MaxRows = 0
    Set .DataSource = rs

End With

'ForeColor
'     With Fps(1)
'
'        For I = 1 To .MaxRows
'            .Row = I
'            .Col = 3
'            .BackColor = &HFF00&
'        Next
'
'    End With
'
End Sub

Private Sub ListDataType1(rs As ADODB.Recordset, J As Integer)
Dim i As Integer

With Fps(0)

    For i = 1 To rs.RecordCount
        .MaxRows = .MaxRows + 1
        .SetText FpsD.e_ID, .MaxRows, Trim$("" & rs!选择)
        .SetText FpsD.e_PKG, .MaxRows, Trim$("" & rs!pkg_id)
        .SetText FpsD.E_DN, .MaxRows, Trim$("" & rs!dn)
        .SetText FpsD.e_PO, .MaxRows, Trim$("" & rs!PO)
        .SetText FpsD.E_PN, .MaxRows, Trim$("" & rs!PN)
        .SetText FpsD.E_MPN, .MaxRows, Trim$("" & rs!MPN)
        .SetText FpsD.E_QTY, .MaxRows, Trim$("" & rs!QTY)
        .SetText FpsD.e_CARTON, .MaxRows, Trim$("" & rs!carton)
        rs.MoveNext
    Next

End With

With Fps(0)
    .Row = J
    .Col = 1
    .ColWidth(1) = 2
    .CellType = CellTypeCheckBox
    .Text = 0

End With

End Sub

Private Sub ListDataType2(rs As ADODB.Recordset)
Dim i As Long

With Fps(2)
    .MaxRows = 0
    Set .DataSource = rs

End With

'ForeColor
'     With Fps(1)
'
'        For I = 1 To .MaxRows
'            .Row = I
'            .Col = 3
'            .BackColor = &HFF00&
'        Next
'
'    End With
'
End Sub

Private Sub cmdClean_Click()
clean
If Val(txtJ.Text) > 0 Then
    txtJ.Text = Val(txtJ.Text) - 1

End If

Fps(0).DeleteRows Fps(0).MaxRows, 1
Fps(0).MaxRows = Fps(0).MaxRows - 1

End Sub

Private Sub cmdConfirm_Click()
Dim pkg_id As String
Dim i      As Integer
Dim N      As Integer
Dim rs     As New ADODB.Recordset
Dim strSql As String
Dim User   As String

User = gUserName
N = 0

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .Text = "1" Then
            N = N + 1

        End If

    Next

End With

If N > 0 Then

    With Fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .Text = "1" Then
                .Col = 2
                pkg_id = .Text
                AddSql ("update HW_ASN_BOX set flag = 1,remark5 = '" & User & "' || ';' || to_char(sysdate,'YYYY-MM-DD hh24:mi:ss')  where asn_pkg_id = '" & pkg_id & "' ")
                AddSql2 ("update erptemp..HW_ASN_BOX set flag = 1,remark5 = '" & User & "'  + ';' + CONVERT(VARCHAR(100),GETDATE(),21)  where asn_pkg_id = '" & pkg_id & "' ")

            End If

        Next

    End With

    AddSql ("insert into HWLBLMATCHHIS(DN,CREATE_DATE,CREATE_BY,STATUS) values('" & txtText2 & "', sysdate,'" & User & "','ASN PASS') ")
    strSql = " select aa.remark2,aa.asn_pkg_id,aa.asn_qty,remark5 from HW_ASN_BOX aa  where aa.remark2 = '" & txtText2.Text & "'"
    Fps(2).MaxRows = 0
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then  '表示有数据了
        Call ListDataType2(rs)

    End If

Else
    Play ("请扫描需要核对的包装ID")

End If

clean1

End Sub

Private Sub cmdup_Click(Index As Integer)
CommonDialog1(1).Filter = "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
CommonDialog1(1).ShowOpen
If CommonDialog1(1).filename = "" Then
    Exit Sub

End If

txtPath(1).Text = CommonDialog1(1).filename
CommonDialog1(1).filename = ""
If txtPath(1).Text = "" Then
    MsgBox "请选择要上传的文件", vbInformation, "提示"
    Exit Sub

End If

Call Upload_0

End Sub

Private Sub Form_Load()
' Me.WindowState = 2
DTPicker3(1).Value = Format(Now() - 30, "YYYY-MM-DD")
DTPicker4(0).Value = Format(Now(), "YYYY-MM-DD")
txtJ.Text = 0
txtJ.BackColor = &HFFFF&

With Fps(0)
    .MaxCols = 8
    .SetText 1, 0, "选择"
    .SetText 2, 0, "PKG_ID"
    .SetText 3, 0, "DN"
    .SetText 4, 0, "PO"
    .SetText 5, 0, "PN"
    .SetText 6, 0, "MPN"
    .SetText 7, 0, "QTY"
    .SetText 8, 0, "CARTON"

End With

End Sub

Private Sub txtText1_KeyPress(KeyAscii As Integer)
Dim J As String

J = Val(txtJ.Text) + 1
If KeyAscii <> vbKeyReturn Then
    Exit Sub

End If

If txtpkg.Text = "" Then
    ForPKG (J)
Else
    Call ForCarton

End If

txtText1.Text = ""
txtText1.SetFocus

End Sub

Private Sub ForPKG(J As Integer)
Dim sScan As String
Dim sOra  As String
Dim sOra1 As String
Dim rs    As New ADODB.Recordset
Dim rs1   As New ADODB.Recordset

sScan = Mid(UCase(Trim(txtText1.Text)), 16, 18)
If CheckPKG(sScan) = False Then
sScan = Trim(txtText1.Text)
 If CheckPKG(sScan) = False Then
    Play ("包装ID无效")
    txtJ.Text = J - 1
    Exit Sub

End If
End If

If CheckPKG1(sScan) = True Then
    Play ("包装ID已做过核对")
    txtJ.Text = J - 1
    Exit Sub

End If

sOra = " select  aa.asn_pkg_id,aa.remark2,aa.asn_po,aa.asn_pn,aa.asn_mpn,aa.asn_qty" & " ,aa.remark1 from HW_ASN_BOX aa  where aa.asn_pkg_id = '" & sScan & "' and aa.flag = 0"
If rs.State = adStateOpen Then rs.Close
rs.Open sOra, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF Then
    txtpkg.Text = UCase(Trim(rs.Fields(0).Value))
    txtpkg.BackColor = &HFFFF&
    txtdn.Text = UCase(Trim(rs.Fields(1).Value))
    txtdn.BackColor = &HFFFF&
    If Len(Trim(txtText2.Text)) = 0 Then
        txtText2.Text = txtdn.Text

    End If

    If CheckPKG2(sScan) = False And Len(Trim(txtText2.Text)) > 0 Then
        Play ("请扫描同DN产品")
        Exit Sub

    End If

    txtpo.Text = UCase(Trim(rs.Fields(2).Value))
    txtpo.BackColor = &HFFFF&
    txtpn.Text = UCase(Trim(rs.Fields(3).Value))
    txtpn.BackColor = &HFFFF&
    txtmpn.Text = UCase(Trim(rs.Fields(4).Value))
    txtmpn.BackColor = &HFFFF&
    txtqty.Text = UCase(Trim(rs.Fields(5).Value))
    txtqty.BackColor = &HFFFF&
    txtcarton.Text = UCase(Trim(rs.Fields(6).Value))
    txtcarton.BackColor = &HFFFF&
    sOra1 = " select   '' as 选择 ,'' as  pkg_id ,'' as DN,'' as  PO,'' as PN,'' as MPN,'' as QTY,'' as CARTON  from dual  "
    If rs1.State = adStateOpen Then rs1.Close
    rs1.Open sOra1, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs1.EOF Then
        Call ListDataType1(rs1, J)
        txtJ.Text = J

    End If

End If

Play ("包装ID扫描成功")

End Sub

Private Sub ForCarton()
Dim sScan    As String
Dim sKeyWord As String
Dim sOra     As String
Dim sOra1    As String
Dim rs       As New ADODB.Recordset
Dim rs1      As New ADODB.Recordset

sScan = UCase(Trim(txtText1.Text))
sKeyWord = Left$(sScan, 1)
If sKeyWord = "2" Then
    If txtdn.BackColor = &HFF00& Then
        Play ("请勿重复扫描条码")
        Exit Sub

    End If

    If Mid(sScan, 3) = txtdn.Text Then
        txtdn.BackColor = &HFF00&

        With Fps(0)
            .Row = Val(txtJ.Text)
            .Col = 3
            .Text = txtdn.Text
            .BackColor = &HFF00&

        End With

        Play ("DN号核对成功")
    Else
        Play ("TURE_CODE")
        Exit Sub

    End If

ElseIf sKeyWord = "K" Then
    If Len(sScan) > 3 Then
        If txtpo.BackColor = &HFF00& Then
            Play ("请勿重复扫描条码")
            Exit Sub

        End If

        If Mid(sScan, 2) = txtpo.Text Then
            txtpo.BackColor = &HFF00&

            With Fps(0)
                .Row = Val(txtJ.Text)
                .Col = 4
                .Text = txtpo.Text
                .BackColor = &HFF00&

            End With

            Play ("PO核对成功")
        Else
            Play ("TURE_CODE")
            Exit Sub

        End If

    End If

ElseIf sKeyWord = "3" Then
    If txtcarton.BackColor = &HFF00& Then
        Play ("请勿重复扫描条码")
        Exit Sub

    End If

    If Mid(sScan, InStr(sScan, "K")) = txtcarton.Text Then
        txtcarton.BackColor = &HFF00&

        With Fps(0)
            .Row = Val(txtJ.Text)
            .Col = 8
            .Text = txtcarton.Text
            .BackColor = &HFF00&

        End With

        Play ("箱号核对成功")
    Else
        Play ("TURE_CODE")
        Exit Sub

    End If

ElseIf sKeyWord = "P" Then
    If txtpn.BackColor = &HFF00& Then
        Play ("请勿重复扫描条码")
        Exit Sub

    End If

    If Mid(sScan, 2) = txtpn.Text Then
        txtpn.BackColor = &HFF00&

        With Fps(0)
            .Row = Val(txtJ.Text)
            .Col = 5
            .Text = txtpn.Text
            .BackColor = &HFF00&

        End With

        Play ("PN核对成功")
    Else
        Play ("TURE_CODE")
        Exit Sub

    End If

ElseIf sKeyWord = "1" Then
    If txtmpn.BackColor = &HFF00& Then
        Play ("请勿重复扫描条码")
        Exit Sub

    End If

    If Mid(sScan, 3) = txtmpn.Text Then
        txtmpn.BackColor = &HFF00&

        With Fps(0)
            .Row = Val(txtJ.Text)
            .Col = 6
            .Text = txtmpn.Text
            .BackColor = &HFF00&

        End With

        Play ("MPN核对成功")
    Else
        Play ("TURE_CODE")
        Exit Sub

    End If

ElseIf sKeyWord = "Q" Then
    If txtqty.BackColor = &HFF00& Then
        Play ("请勿重复扫描条码")
        Exit Sub

    End If

    If Mid(sScan, 2) = txtqty.Text Then
        txtqty.BackColor = &HFF00&

        With Fps(0)
            .Row = Val(txtJ.Text)
            .Col = 7
            .Text = txtqty.Text
            .BackColor = &HFF00&

        End With

        Play ("数量核对成功")
    Else
        Play ("TURE_CODE")
        Exit Sub

    End If

Else
    Play ("TURE_CODE")

End If

If txtdn.BackColor = &HFF00& And txtpo.BackColor = &HFF00& And txtpn.BackColor = &HFF00& And txtmpn.BackColor = &HFF00& And txtqty.BackColor = &HFF00& And txtcarton.BackColor = &HFF00& Then

    With Fps(0)
        .Row = Val(txtJ.Text)
        .Col = 1
        .Text = 1
        .Lock = True
        .BackColor = &HFF00&
        .Col = 2
        .Text = txtpkg.Text
        .BackColor = &HFF00&

    End With

    clean
    If CheckPKG3(txtText2.Text, txtJ) = True Then
        Play ("DN核对完成")
        cmdConfirm_Click
        Exit Sub

    End If

    Play ("包装ID扫描成功")
    Play ("扫描完成请扫描下包装ID")

End If

End Sub

Private Function CheckPKG(PKG As String) As Boolean
Dim sOra As String

sOra = "select * from HW_ASN_BOX aa  where aa.asn_pkg_id = '" & PKG & "' "
CheckPKG = IsOraRecord(sOra)

End Function

Private Function CheckPKG1(PKG As String) As Boolean
Dim sOra As String

sOra = "select * from HW_ASN_BOX aa  where aa.asn_pkg_id = '" & PKG & "' and aa.flag <> 0 "
CheckPKG1 = IsOraRecord(sOra)

End Function

Private Function CheckPKG2(PKG As String) As Boolean
Dim sOra As String

sOra = "select * from HW_ASN_BOX aa  where  aa.asn_pkg_id = '" & PKG & "'  and aa.remark2 = '" & txtText2.Text & "'"
CheckPKG2 = IsOraRecord(sOra)

End Function

Private Function CheckPKG3(dn As String, NUMJ As String) As Boolean
Dim sOra As String

sOra = " select count(*) from HW_ASN_BOX aa where aa.remark2 = '" & dn & "'  having count(*) = '" & NUMJ & "' "
CheckPKG3 = IsOraRecord(sOra)

End Function

Private Sub clean()
txtdn.Text = ""
txtpo.Text = ""
txtpn.Text = ""
txtmpn.Text = ""
txtqty.Text = ""
txtcarton.Text = ""
txtpkg.Text = ""
txtdn.BackColor = &HFFFFFF
txtpo.BackColor = &HFFFFFF
txtpn.BackColor = &HFFFFFF
txtmpn.BackColor = &HFFFFFF
txtqty.BackColor = &HFFFFFF
txtcarton.BackColor = &HFFFFFF
txtpkg.BackColor = &HFFFFFF

End Sub
 
Private Sub clean1()
txtdn.Text = ""
txtpo.Text = ""
txtpn.Text = ""
txtmpn.Text = ""
txtqty.Text = ""
txtcarton.Text = ""
txtpkg.Text = ""
txtdn.BackColor = &HFFFFFF
txtpo.BackColor = &HFFFFFF
txtpn.BackColor = &HFFFFFF
txtmpn.BackColor = &HFFFFFF
txtqty.BackColor = &HFFFFFF
txtcarton.BackColor = &HFFFFFF
txtpkg.BackColor = &HFFFFFF
txtJ.Text = 0
txtJ.BackColor = &HFFFF&
txtText2.Text = ""
Fps(0).MaxRows = 0

End Sub

Private Sub Play(sFileName As String)
Dim sPath   As String
Dim sSuffix As String

sPath = "D:\HW_ASN\"
sSuffix = ".wav"
media.url = sPath & sFileName & sSuffix
Sleep (200)

End Sub


