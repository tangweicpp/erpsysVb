VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_37LblPrint 
   BackColor       =   &H00C0C0C0&
   Caption         =   "SEMTECH-标签打印"
   ClientHeight    =   12720
   ClientLeft      =   60
   ClientTop       =   405
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
   MDIChild        =   -1  'True
   ScaleHeight     =   12720
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   1482
      ButtonWidth     =   1032
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "打印"
            Key             =   "PRINT"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "删除"
            Key             =   "DEL"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "导出"
            Key             =   "EXPORT"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "EXIT"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTTab0 
      Height          =   11295
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   19845
      _ExtentX        =   35004
      _ExtentY        =   19923
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483637
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "标签扫描打印"
      TabPicture(0)   =   "Frm_37ToHW_LblPrint.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "m1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblScanning"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDN"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblPO"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblQTY"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblReelList"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblMPN"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblJOBList"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblS"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ImageList2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "UpDown1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fps(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fps(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "fps(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtScan"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtDN"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtPO"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtQty"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtTmpQty"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtStatus"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "chkCheck1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "ProgressBar1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtPrintInterval"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "标签扫描补打"
      TabPicture(1)   =   "Frm_37ToHW_LblPrint.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtScan2"
      Tab(1).Control(1)=   "cbLblType"
      Tab(1).Control(2)=   "Command2"
      Tab(1).Control(3)=   "txtPassWd2"
      Tab(1).Control(4)=   "Command1"
      Tab(1).Control(5)=   "txtPassWd"
      Tab(1).Control(6)=   "txtUser"
      Tab(1).Control(7)=   "txtUser2"
      Tab(1).Control(8)=   "lblBarcodeScan2"
      Tab(1).Control(9)=   "lblType"
      Tab(1).Control(10)=   "Label266"
      Tab(1).Control(11)=   "Label1"
      Tab(1).ControlCount=   12
      Begin VB.TextBox txtScan2 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   -73320
         TabIndex        =   31
         Top             =   780
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.ComboBox cbLblType 
         Height          =   315
         ItemData        =   "Frm_37ToHW_LblPrint.frx":0038
         Left            =   -73320
         List            =   "Frm_37ToHW_LblPrint.frx":004B
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1260
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF80FF&
         Caption         =   "导出补打记录"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65880
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   780
         Width           =   2295
      End
      Begin VB.TextBox txtPassWd2 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -71640
         PasswordChar    =   "*"
         TabIndex        =   28
         Top             =   4140
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "验证补打密码"
         Height          =   840
         Left            =   -68640
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3660
         Width           =   1575
      End
      Begin VB.TextBox txtPassWd 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -71640
         PasswordChar    =   "*"
         TabIndex        =   26
         Top             =   3660
         Width           =   2895
      End
      Begin VB.TextBox txtUser 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -72960
         TabIndex        =   25
         Text            =   "10354"
         Top             =   3660
         Width           =   1215
      End
      Begin VB.TextBox txtUser2 
         Height          =   375
         Left            =   -72960
         TabIndex        =   24
         Top             =   4140
         Width           =   1215
      End
      Begin VB.TextBox txtPrintInterval 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "6"
         Top             =   45
         Width           =   300
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   10320
         TabIndex        =   20
         Top             =   480
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   873
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CheckBox chkCheck1 
         Caption         =   "测试用"
         Height          =   255
         Left            =   7200
         TabIndex        =   19
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   6615
         Left            =   7920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   4440
         Width           =   10095
      End
      Begin VB.TextBox txtTmpQty 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtPO 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtDN 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtScan 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   585
         Width           =   5415
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   9015
         Index           =   0
         Left            =   1560
         TabIndex        =   10
         Top             =   1680
         Width           =   5415
         _Version        =   524288
         _ExtentX        =   9551
         _ExtentY        =   15901
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
         MaxCols         =   4
         MaxRows         =   0
         SpreadDesigner  =   "Frm_37ToHW_LblPrint.frx":0094
         Appearance      =   1
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   3015
         Index           =   1
         Left            =   7920
         TabIndex        =   12
         Top             =   1320
         Width           =   4935
         _Version        =   524288
         _ExtentX        =   8705
         _ExtentY        =   5318
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
         SpreadDesigner  =   "Frm_37ToHW_LblPrint.frx":0582
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   3015
         Index           =   2
         Left            =   13080
         TabIndex        =   13
         Top             =   1320
         Width           =   4935
         _Version        =   524288
         _ExtentX        =   8705
         _ExtentY        =   5318
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
         SpreadDesigner  =   "Frm_37ToHW_LblPrint.frx":0A1E
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   405
         Left            =   9540
         TabIndex        =   22
         Top             =   45
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   714
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txtPrintInterval"
         BuddyDispid     =   196617
         OrigLeft        =   5880
         OrigTop         =   6240
         OrigRight       =   6135
         OrigBottom      =   6645
         Max             =   8
         Min             =   2
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   240
         Top             =   2160
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint.frx":0EBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint.frx":2FF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint.frx":5E7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint.frx":8630
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint.frx":A76A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint.frx":CF1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint.frx":F6CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint.frx":12750
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint.frx":14F02
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint.frx":1521C
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint.frx":15EF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint.frx":18F78
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint.frx":1B72A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblBarcodeScan2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode Scan:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -74760
         TabIndex        =   35
         Top             =   780
         Width           =   1395
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标签类型:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74400
         TabIndex        =   34
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label266 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm_37ToHW_LblPrint.frx":1C004
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   -74640
         TabIndex        =   33
         Top             =   4185
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm_37ToHW_LblPrint.frx":1C018
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   -74640
         TabIndex        =   32
         Top             =   3720
         Width           =   1635
      End
      Begin VB.Label lblS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "打印间隔(s):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7800
         TabIndex        =   23
         Top             =   120
         Width           =   1320
      End
      Begin VB.Line Line1 
         X1              =   3000
         X2              =   3480
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblJOBList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JOB List:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   15000
         TabIndex        =   15
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lblMPN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M.P.N List:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9360
         TabIndex        =   14
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lblReelList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ReelList:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   720
         TabIndex        =   11
         Top             =   6120
         Width           =   840
      End
      Begin VB.Label lblQTY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTY:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   8
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label lblPO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P.O:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3075
         TabIndex        =   6
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblDN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D.N:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1140
         TabIndex        =   4
         Top             =   960
         Width           =   390
      End
      Begin VB.Label lblScanning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode Scan:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1395
      End
      Begin WMPLibCtl.WindowsMediaPlayer m1 
         Height          =   495
         Left            =   7200
         TabIndex        =   1
         Top             =   420
         Visible         =   0   'False
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
   End
End
Attribute VB_Name = "Frm_37LblPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilllMicroSeconds As Long)
Dim lMicroSec As Long

Private Sub cmdPrint2_Click()
If Trim(txtPassWd.Text) = "htpkg10354" Then
    txtScan2.Visible = True
Else
    MsgBox "请输入补打密码或密码不正确", vbInformation, "提示"
    Exit Sub
End If
End Sub

Private Sub chkCheck1_Click()

If chkCheck1.Value = 1 Then
    Call setTestPrintPath
Else
    Call setPrintPath

End If
End Sub

Private Sub Command1_Click()
Dim strsql As String

If txtUser2.Text = txtUser.Text Then
    MsgBox "员工不可输入组长的工号", vbCritical, "提示"
    Exit Sub
End If

strsql = "select * from tblOperatorData r where  r.状态标记=1  and r.用户号='10354'and r.密码='" & Replace(Trim(txtPassWd.Text), "'", "") & "'"
If Get_SqlStr(strsql) = "" Then
    MsgBox "组长密码不正确", vbCritical, "提示"
    Exit Sub
End If

strsql = "select * from tblOperatorData r where  r.状态标记=1  and r.用户号='" & Trim(txtUser2.Text) & "'and r.密码='" & Replace(Trim(txtPassWd2.Text), "'", "") & "'"
If Get_SqlStr(strsql) = "" Then
    MsgBox "员工工号或者密码不正确", vbCritical, "提示"
    Exit Sub
End If

txtScan2.Visible = True
End Sub

Private Sub Command2_Click()
ExporToExcel ("select KEYNAME 补打类型,keyvalue 补打值,CREATE_DATE 补打时间,CREATE_BY 补打人员工号,CREATE_TIMES 第几次补打 from TBL_37_PRINT2_LIST order by CREATE_date desc")
End Sub

Private Sub Form_Activate()
txtScan.SetFocus
SSTTab0.Tab = 0
End Sub

Private Sub Form_Load()
InitCtrls
InitData

End Sub

Private Sub InitData()
lMicroSec = CLng(txtPrintInterval.Text) * 500

If chkCheck1.Value = 1 Then
    Call setTestPrintPath
Else
    Call setPrintPath

End If

End Sub

Private Sub InitCtrls()

If gUserName = "07885" Then
    Toolbar1.Buttons(3).Enabled = True
    txtDN.Locked = False
    chkCheck1.Value = 1

ElseIf gUserName = "10354" Then
    Toolbar1.Buttons(3).Enabled = True
    txtDN.Locked = False
End If

InitFps
initMedia

End Sub

Private Sub InitFps()

With fpS(0)
    .ReDraw = False
    .MaxCols = 3
    .MaxRows = 0
    .FontBold = False
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .TypeVAlign = TypeVAlignCenter
    .SelForeColor = &HFF8080
    .SetText 1, 0, "扫描时间"
    .SetText 2, 0, "卷盘ID"
    .SetText 3, 0, "外箱"
    .ColWidth(1) = 16
    .ColWidth(2) = 12
    .ColWidth(3) = 12
    .RowHeight(0) = 15
    .RowHeight(-1) = 12
    .ReDraw = True

End With

With fpS(1)
    .ReDraw = False
    .MaxCols = 3
    .MaxRows = 0
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .TypeVAlign = TypeVAlignCenter
    .SelForeColor = &HFF8080
    .SetText 1, 0, "M.P.N"
    .SetText 2, 0, "目标数量"
    .SetText 3, 0, "累计数量"
    .ColWidth(1) = 14
    .ColWidth(2) = 10
    .ColWidth(3) = 10
    .RowHeight(0) = 20
    .RowHeight(-1) = 15
    .ReDraw = True

End With

With fpS(2)
    .ReDraw = False
    .MaxCols = 3
    .MaxRows = 0
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .TypeVAlign = TypeVAlignCenter
    .SelForeColor = &HFF8080
    .SetText 1, 0, "JOB"
    .SetText 2, 0, "目标数量"
    .SetText 3, 0, "累计数量"
    .ColWidth(1) = 14
    .ColWidth(2) = 10
    .ColWidth(3) = 10
    .RowHeight(0) = 20
    .RowHeight(-1) = 15
    .ReDraw = True

End With

End Sub

Private Sub initMedia()
Dim strDocCfgDir As String, strLocalDocDir As String, strRemoteDocDir As String
Dim strTemp      As String, strArr() As String
strDocCfgDir = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\37\SemtechToHW_Voice.Cfg"

If Dir(strDocCfgDir, vbDirectory) = "" Then
    MsgBox "声音路径配置文件丢失" & vbCrLf & "请联系IT确认", vbExclamation, "警告"
    Exit Sub

End If

Open strDocCfgDir For Input As #1

Do While Not EOF(1)
    Line Input #1, strTemp
    strArr = Split(strTemp, "$")

    Select Case strArr(0)

        Case "LocalDocPath"
            strLocalDocDir = strArr(1)

        Case "RemoteDocPath"
            strRemoteDocDir = strArr(1)

    End Select

Loop
Close #1

If strLocalDocDir <> "" And strRemoteDocDir <> "" Then
    If Dir(strLocalDocDir, vbDirectory) <> "" Then
        strMediaDir = strLocalDocDir
    ElseIf Dir(strRemoteDocDir, vbDirectory) <> "" Then
        strMediaDir = strRemoteDocDir
    Else
        MsgBox "找不到声音文件" & vbCrLf & "请联系IT确认", vbExclamation, "警告"
        Exit Sub

    End If

End If

End Sub

Private Sub txtPrintInterval_Change()
lMicroSec = CLng(txtPrintInterval.Text) * 500

End Sub

Private Sub txtScan_KeyPress(KeyAscii As Integer)
Dim strScan As String
strScan = UCase(Trim(txtScan.Text))

If KeyAscii <> vbKeyReturn Or Len(strScan) = 0 Then Exit Sub
If Len(Trim(txtDN.Text)) = 0 Then
    getDNInfo (strScan)
Else
    getReelInfo (strScan)

End If

txtScan.Text = ""

End Sub

Private Sub getDNInfo(strDN As String)
Dim strsql As String
Dim rs     As New ADODB.Recordset

If Left$(strDN, 1) <> "I" Then
    m1.url = getMediaUrl("DN未获取,请扫描DN条码")
    txtStatus.BackColor = vbRed
    Exit Sub

End If

strDN = Mid(strDN, 2)

If chkDNToHW(strDN) = False Then
    txtStatus.BackColor = vbRed
    Exit Sub

End If

txtDN.Text = strDN
Call updateDNStatus(strDN, "new")
strsql = "select purchasingdocno po,sum(quantity) qty from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "' group by purchasingdocno "
Set rs = Get_OracleRs(strsql)

If rs.RecordCount > 0 Then
    txtPO.Text = "" & rs!PO
    txtQty.Text = "" & rs!qty

End If

'扫描历史
strsql = "select * from ST_TR_SEQ where dn = '" & strDN & "' order by seq"

If Get_OracleCnt(strsql) > 0 Then
    m1.url = getMediaUrl("该DN有历史扫描记录,请确认是否接着上次扫描或清除记录")

    If MsgBox("该DN有历史扫描记录,请确认是否接着上次扫描或清除记录" & vbCrLf & "是:继续 否:清除", vbYesNo, "提示") = vbYes Then
        m1.url = getMediaUrl("扫描记录已刷新,请紧接上次,依次扫描剩余挑料卷盘")
    Else
        AddSql ("delete from ST_TR_SEQ where dn = '" & strDN & "'")
        m1.url = getMediaUrl("扫描记录已删除, 请重新扫描挑料卷盘")

    End If

End If

refreshInfo (strDN)
rs.Close
Set rs = Nothing

End Sub

Private Sub getReelInfo(strReelID As String)
Dim strsql As String
Dim strDN  As String
Dim rs     As New ADODB.Recordset
Dim strMPN As String, strJob As String, strQty As String, lMaxQty As Long
strDN = Trim(txtDN.Text)
lMaxQty = CLng(txtQty.Text)

If Left$(strReelID, 1) <> "S" Then
    m1.url = getMediaUrl("卷盘号未获取,请扫描卷盘号")
    txtStatus.BackColor = vbRed
    Exit Sub

End If

' Job
If Len(strReelID) = 13 Or Len(strReelID) = 14 Then
    strsql = " SELECT RTRIM(SUBSTRING(a.KEY_VALUE,CHARINDEX('|',a.KEY_VALUE)+1,20)) AS JOB ,sum(b.数量) FROM erpdata..tblErpInStockDetailInfo a ,erpdata..tblStockNumSub b  " & " WHERE a.KEY_VALUE LIKE '%|%'  AND a.KEY_NAME = 'CONTAINER_NAME' AND a.KEY_TYPE = 'T' AND SUBSTRING(a.KEY_VALUE,1,CHARINDEX('|',a.KEY_VALUE)-1) = b.箱号 and b.箱号 = '" & strReelID & "' GROUP BY RTRIM(SUBSTRING(a.KEY_VALUE,CHARINDEX('|',a.KEY_VALUE)+1,20)) "
Else
    strsql = "select CUSTOMERLOTID,qty from [erpdata].[dbo].[TblTSV_Tray_details] where TRAYQBOXNUMBER = '" & strReelID & "'"

End If

Set rs = Get_SqlserveRs(strsql)

If rs.RecordCount > 0 Then
    strJob = Trim(rs(0).Value)
    strQty = Trim$(rs(1).Value)

End If

If strJob = "" Then
    m1.url = getMediaUrl("卷盘号不正确")
    txtStatus.BackColor = vbRed
    Exit Sub

End If

' M.P.N
strsql = "select marketingpn as mpn from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "' and batchnumber = '" & strJob & "'"
strMPN = Get_OracleStr(strsql)

' Check
If chkReelID(strReelID, strDN, strJob, strMPN) = False Then
    txtStatus.BackColor = vbRed
    Exit Sub

End If

txtTmpQty.Text = insertReelID(strReelID, strDN, strJob, strMPN, strQty, lMaxQty)
refreshInfo (strDN)

End Sub

Private Sub refreshInfo(strDN As String)
Dim strsql As String
Dim lQty   As Long, lMaxQty As Long
Dim rs     As New ADODB.Recordset
' Reel Info
strsql = "select seqtime,reelid from ST_TR_SEQ where dn = '" & strDN & "' order by seq desc"
Set rs = Get_OracleRs(strsql)

With fpS(0)
    .MaxRows = 0

    If rs.RecordCount > 0 Then
        Set .DataSource = rs

    End If

End With

' M.P.N Info
strsql = " select AA.marketingpn,AA.realqtys, BB.thisqtys from (select marketingpn, sum(quantity) as realqtys from CUSTOMERSHIPPINGUPTBL  where delivery = '" & strDN & "'  group by marketingpn) AA left join (select dev, sum(qty) as thisqtys from ST_TR_SEQ where dn = '" & strDN & "' group by dev) BB on AA.marketingpn = BB.dev "
Set rs = Get_OracleRs(strsql)

With fpS(1)
    .MaxRows = 0

    If rs.RecordCount > 0 Then
        Set .DataSource = rs

    End If
    

End With

' JOB Info
strsql = " select AA.batchnumber,AA.realqtys, BB.thisqtys from (select batchnumber, sum(quantity) as realqtys from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "' group by batchnumber) AA left join (select job, sum(qty) as thisqtys from ST_TR_SEQ where dn = '" & strDN & "' group by job) BB on AA.batchnumber = BB.job "
Set rs = Get_OracleRs(strsql)

With fpS(2)
    .MaxRows = 0

    If rs.RecordCount > 0 Then
        Set .DataSource = rs

    End If

End With

' 累计数量
strsql = "select sum(qty) from ST_TR_SEQ where dn = '" & strDN & "'"
txtTmpQty.Text = Get_OracleNo(strsql)
' 卷盘数量变更
strsql = "select nvl(count(*), 0) from ST_TR_SEQ where dn = '" & strDN & "'"
txtStatus.BackColor = vbWhite
txtStatus.Text = vbCrLf & Get_OracleNo(strsql)
' 当前是否已满
strsql = "select sum(qty) from ST_TR_SEQ where dn = '" & strDN & "'"
lQty = Get_OracleNo(strsql)
strsql = "select sum(quantity) from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "'"
lMaxQty = Get_OracleNo(strsql)

If lQty = lMaxQty Then
    Call updateDNStatus(strDN, "scaned")
    m1.url = getMediaUrl("卷盘已全部扫描完毕,请准备打印工作")
    txtScan.Visible = False
    Toolbar1.Buttons(1).Enabled = True
ElseIf lQty > lMaxQty Then
    m1.url = getMediaUrl("已扫卷盘大于需要挑料的总数, 请确认")
    MsgBox "已扫卷盘大于需要挑料的总数, 请确认", vbExclamation, "警告"
    txtStatus.BackColor = vbRed

End If

rs.Close
Set rs = Nothing

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case SSTTab0.Tab

    Case 0

        Select Case Button.Key

            Case "PRINT"
                printLbl

            Case "DEL"
                delDNInfo

            Case "EXPORT"
                exportInfo

            Case "EXIT"
                Unload Me

        End Select

    Case 1

        Select Case Button.Key

            Case "PRINT"
                printLbl2

            Case "EXIT"
                Unload Me

        End Select

    Case Else

End Select

End Sub

Private Sub exportInfo()
Dim strDN  As String
Dim strsql As String
strDN = Trim$(txtDN.Text)

If Len(strDN) = 0 Then
    MsgBox "请输入要导出的DN", vbInformation, "提示"
    Exit Sub

End If

strsql = "select dn_num dn, trayid, reelid PSN,job_id job, customer_device 客户机种, QTY 数量, KID, DATECODE, CREATE_BY 打印人员, CREATE_DATE 打印时间,'' as 备注 from packing_detailed where dn_num = '" & strDN & "' order by seq desc "
ExporToExcel (strsql)

End Sub

Private Sub delDNInfo()
Dim strDN  As String
Dim strsql As String
Dim rs     As New ADODB.Recordset
Dim i      As Integer

strDN = Trim(txtDN.Text)
If Len(strDN) = 0 Then
    MsgBox "请输入要删除的DN", vbInformation, "提示"
    txtDN.Text = ""
    Exit Sub

End If

strsql = "select * from packing_detailed where dn_num = '" & strDN & "'"
If Get_OracleCnt(strsql) = 0 Then
    MsgBox "查询不到该DN的信息, 不可删除", vbInformation, "提示"
    txtDN.Text = ""
    Exit Sub

End If

If MsgBox("确认要删除吗?", vbYesNoCancel, "提示") = vbNo Then
    Exit Sub

End If

' 1.删除箱号记录
Call DelToErp(strDN)
' 2.删除PACKING_DETAILED
strsql = "insert into packing_detailed_bak select * from packing_detailed where dn_num = '" & strDN & "'"
If AddSql(strsql) Then
    MsgBox "已备份DN数据", vbInformation, "提示"

End If

strsql = "delete from packing_detailed where dn_num = '" & strDN & "'"
AddSql (strsql)
strsql = "delete from PKGIDSEQ_37 where dn = '" & strDN & "'  "
AddSql (strsql)
MsgBox "已删除DN数据", vbInformation, "提示"
strsql = "update PRINT_37FLAG set printed = '0', combined = '0', scaned = '0' where dn = '" & strDN & "'"
AddSql (strsql)
txtDN.Text = ""
txtDN.Locked = True
Toolbar1.Buttons(3).Enabled = False

End Sub

Private Sub printLbl()
Dim strDN As String, strsql As String
Toolbar1.Buttons(1).Enabled = False
strDN = Trim(txtDN.Text)

If strDN = "" Then
    MsgBox "DN不可为空", vbExclamation, "警告"
    Exit Sub

End If

strsql = "select * from PACKING_DETAILED where dn_num = '" & strDN & "' "

If Get_OracleCnt(strsql) = 0 Then
    setPkgID (strDN)    ' 合箱

End If

strsql = "select * from PRINT_37FLAG where dn = '" & strDN & "' and COMBINED = '1'"

If Get_OracleCnt(strsql) = 0 Then
    MsgBox "该DN未完成合箱, 数据有问题, 标签不可打印, 请联系IT删除异常数据", vbCritical, "警告"
    Exit Sub

End If

strsql = "select * from PRINT_37FLAG where dn = '" & strDN & "' and printed = '1'"

If Get_OracleCnt(strsql) > 0 Then
    MsgBox "该DN已打印过全套标签, 不可再次打印, 请使用补打功能", vbCritical, "警告"
    Exit Sub

End If

printStart (strDN)  ' 打印

End Sub

Private Sub setPkgID(strDN As String)

On Error GoTo ERRON

Cnn.BeginTrans
Call InsertPkgID(strDN)
Call UpdatePkgID(strDN)
Cnn.CommitTrans
Call updateDNStatus(strDN, "combined")

' 插入SQl SERVER
AddSql2 ("delete from ERPBASE..PACKING_DETAILED where dn_num = '" & strDN & "'")
AddSql2 ("insert into ERPBASE..PACKING_DETAILED select * from (select * from OPENQUERY(ORACLEDB, 'SELECT * from PACKING_DETAILED' )) X where X.dn_num = '" & strDN & "'")
    

' 更新箱号关系
Call TransToErp(strDN)
    
Exit Sub
ERRON:
MsgBox "执行失败！" + Chr(13) + "原因:" + Err.DESCRIPTION, vbInformation, Me.Caption
Cnn.RollbackTrans

End Sub
Private Sub UpdatePkgID(strDN As String)
    
    Dim rs          As New ADODB.Recordset
    
    Dim strsql      As String

    Dim iOp         As Integer

    Dim iIp         As Integer

    Dim iOpMax      As Integer

    Dim iIpMax      As Integer

    Dim strKey      As String

    Dim sLotBox     As String

    Dim sFirstBoxId As String

    Dim strQID      As String

    Dim lQty        As Long

    Dim strseq      As String

    iOp = 1
    iIp = 1

    strsql = "select max(outbox_num) from PACKING_DETAILED where dn_num = '" & strDN & "'  "
    iOpMax = Get_OracleNo(strsql)

    For iOp = 1 To iOpMax

        strsql = "select max(inbox_num) from PACKING_DETAILED where dn_num = '" & strDN & "' and outbox_num = '" & iOp & "'"
        iIpMax = Get_OracleNo(strsql)
       
        'Set C_ID
        strsql = "select distinct substr(trayid,1, InStr(trayid, '-') - 1) LOTID, job_id  from PACKING_DETAILED where dn_num = '" & strDN & "' and outbox_num = '" & iOp & "' "
        Set rs = Get_OracleRs(strsql)

        If Not rs.BOF Then
        
            rs.MoveFirst

            Do While Not rs.EOF
                strKey = rs!LOTID & "-C"
                
                strsql = " select nvl(max(seq)+1, 1) from PKGIDSEQ_37 where val = '" & strKey & "' "
                strseq = Right("0" & Get_OracleStr(strsql), 2)
                sLotBox = strKey & strseq
                
                strsql = "update PACKING_DETAILED set CARTONID = '" & sLotBox & "' where dn_num = '" & strDN & "' and outbox_num = '" & iOp & "' and job_id = '" & rs.Fields("job_id") & "'"
                AddSql (strsql)
                
                strsql = "insert into PKGIDSEQ_37(VAL,SEQ,CREATE_DATE,DN) values('" & strKey & "', '" & strseq & "', sysdate, '" & strDN & "')"
                AddSql (strsql)
                
                rs.MoveNext
            Loop

        End If
        
        rs.Close

        'Set B_ID
        For iIp = 1 To iIpMax
        
            strsql = "select distinct substr(trayid,1, InStr(trayid, '-') - 1) LOTID, job_id  from PACKING_DETAILED where dn_num = '" & strDN & "' and outbox_num = '" & iOp & "' and inbox_num = '" & iIp & "'"
        
            Set rs = Get_OracleRs(strsql)
        
            If Not rs.BOF Then
        
                rs.MoveFirst

                Do While Not rs.EOF
                    strKey = rs!LOTID & "-B"
                    
                    strsql = " select nvl(max(seq)+1, 1) from PKGIDSEQ_37 where val = '" & strKey & "' "
                    strseq = Right("0" & Get_OracleStr(strsql), 2)
                    sLotBox = strKey & strseq
        
                    strsql = "update PACKING_DETAILED set BOXID = '" & sLotBox & "' where dn_num = '" & strDN & "' and outbox_num = '" & iOp & "' and inbox_num = '" & iIp & "' and job_id = '" & rs.Fields("job_id") & "'"
                    AddSql (strsql)
    
                    strsql = "insert into PKGIDSEQ_37(VAL,SEQ,CREATE_DATE,DN) values('" & strKey & "', '" & strseq & "', sysdate, '" & strDN & "')"
                    AddSql (strsql)
                    
                    rs.MoveNext
                Loop
                
            End If
        
            rs.Close
        Next
    
        ' Set Q_ID
        strsql = "select boxid from PACKING_DETAILED where dn_num = '" & strDN & "' and outbox_num='" & iOp & "' and inbox_num = '1'"
        sFirstBoxId = Get_OracleStr(strsql)

        If sFirstBoxId <> "" Then
            strsql = "select trglabelseq.QTSeq_NotMesQbox('" & sFirstBoxId & "')  from dual"
            strQID = Get_OracleStr(strsql)
            
            strsql = "update PACKING_DETAILED set CARTON = '" & strQID & "' where dn_num = '" & strDN & "' and outbox_num = '" & iOp & "' "
            AddSql (strsql)
        
            lQty = Get_OracleNo("select sum(qty) from PACKING_DETAILED where dn_num = '" & strDN & "' and outbox_num = '" & iOp & "'")

            'Call TransToErp(strQID, lQty)
        End If

    Next
    
    Call updateDNStatus(strDN, "combined")

End Sub

' 打印开始
Private Sub printStart(strDN As String)
Dim iOp    As Integer
Dim strsql As String, strsql2 As String
strsql = "select min(outbox_num) from PACKING_DETAILED where dn_num = '" & strDN & "' and print_flag = '0' order by outbox_num"
iOp = Get_OracleNo(strsql)
strsql2 = "select* from PACKING_DETAILED where dn_num = '" & strDN & "'"

If iOp = 0 And (Get_OracleCnt(strsql2) > 0) Then
    m1.url = getMediaUrl("内盒和卷盘标签已全部打印完毕")

    If MsgBox("内盒卷盘标签已全部打印完毕, 是否打印外箱标签?", vbYesNoCancel, "友情提示:") = vbYes Then
        Call printOPkgLbl(strDN)
    Else
        Toolbar1.Buttons(1).Enabled = True
        Exit Sub

    End If

Else
    Call printIPkgLbl(strDN, iOp)

End If

End Sub

' 打印内箱小标签
Private Sub printIPkgLbl(strDN As String, iOp As Integer)
Dim strsql As String
Dim iOpMax As Integer
Dim iIpMax As Integer
Dim iIp    As Integer
Dim lBar   As Long
ProgressBar1.Value = 0
strsql = "select max(inbox_num) from PACKING_DETAILED where dn_num = '" & strDN & "' and outbox_num = '" & iOp & "'"
iIpMax = Get_OracleNo(strsql)
lBar = 50 / iIpMax

For iIp = 1 To iIpMax

    If (ProgressBar1.Value + lBar) <= 100 Then
        ProgressBar1.Value = ProgressBar1.Value + lBar
    Else
        ProgressBar1.Value = 100

    End If

    ' 内盒 开始
    Call PrintINNERBOXFlag(strDN, iOp, iIp)
    Sleep (lMicroSec)
    ' 37 内盒B标签
    Call PrintSTBoxLbl(strDN, iOp, iIp)
    Sleep (lMicroSec)
    ' 华为 内盒标签
    Call PrintHWBoxLbl(strDN, iOp, iIp)
    Sleep (lMicroSec)
    ' 卷盘 开始
    Call PrintREELFlag(strDN, iOp, iIp)
    Sleep (lMicroSec)
    ' 华为 卷盘PSN标签
    Call PrintHWReelLbl(strDN, iOp, iIp)
    Sleep (lMicroSec)
Next
' C 标签开始 华为取消-C
Call PrintOuterCFlag(strDN, iOp)
'
'    ' 37 外箱C标签
Call PrintSTCartonLbl(strDN, iOp)
' HT Q标签
Call PrintHTCartonLbl(strDN, iOp)
' 更新扫描状态
Call UpdatePrintStatus(strDN, iOp)
ProgressBar1.Value = 100
MsgBox "第" & iOp & "个外箱已经打印完成", vbInformation, "友情提示:"
Toolbar1.Buttons(1).Enabled = True

End Sub

' 打印外箱大标签
Private Sub printOPkgLbl(strDN As String)
Dim strsql As String
Dim iOp    As Integer
Dim iOpMax As Integer
Dim lBar   As Long
Dim sType  As String
ProgressBar1.Value = 0
strsql = "select max(outbox_num) from PACKING_DETAILED where dn_num = '" & strDN & "'  "
iOpMax = Get_OracleNo(strsql)
lBar = 50 / iOpMax

For iOp = 1 To iOpMax

    If (ProgressBar1.Value + lBar) <= 100 Then
        ProgressBar1.Value = ProgressBar1.Value + lBar
    Else
        ProgressBar1.Value = 100

    End If

    Sleep (lMicroSec)

    If sType = "SSE2" Then
        Call PrintCusCartonLbl(strDN, iOp)
        Sleep (lMicroSec)
    Else
        Call PrintSTCartonStanderLbl(strDN, iOp)
        Sleep (lMicroSec)

    End If

Next
MsgBox "外箱大标签打印完成", vbInformation, "友情提示:"
m1.url = getMediaUrl("该D N所有标签已全部打印完毕")
ProgressBar1.Value = 100
Call updateDNStatus(strDN, "printed")

End Sub

'----------------------------------------------------------------------------------------------------'
Rem: 补打标签
Private Sub SSTTab0_Click(PreviousTab As Integer)

Select Case SSTTab0.Tab

    Case 0
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = True

    Case 1
        Toolbar1.Buttons(1).Enabled = True
        Toolbar1.Buttons(3).Enabled = False

End Select

End Sub

Private Sub txtScan2_KeyPress(KeyAscii As Integer)
Dim strScan As String
strScan = UCase(Trim(txtScan2.Text))

If KeyAscii <> vbKeyReturn Or Len(strScan) = 0 Then Exit Sub

Call printLblNew(strScan)
txtScan.Text = ""

End Sub

Private Sub printLbl2()
Dim strKey As String
strKey = UCase(Trim(txtScan2.Text))

If Len(strKey) = 0 Then
    MsgBox "请输入需要补打的条码", vbInformation, "提示"
    Exit Sub

End If

Call printLblNew(strKey)
txtScan.Text = ""

End Sub

Private Sub printLblNew(strKey As String)
Dim iQty As Integer

If cbLblType.Text = "" Then
    MsgBox "请选择补打的标签类型", vbInformation, "提示"
    Exit Sub

End If

Select Case cbLblType.Text

    Case "华为卷盘标签"
        Call PrintHWReelLbl2(strKey)

    Case "华为内盒标签"
        Call PrintHWBoxLbl2(strKey)

    Case "37内盒标签"
        Call PrintSTBoxLbl2(strKey)

    Case "37外箱C标签"
        Call PrintSTCartonLbl2(strKey)

    Case "华为外箱大标签"
        Call PrintSTCartonStanderLbl2(strKey)

End Select

iQty = Get_OracleStr("select nvl(count(*) + 1, 1) from TBL_37_PRINT2_LIST where KEYNAME = '" & cbLblType.Text & "' and KEYVALUE = '" & strKey & "'")
AddSql ("insert into TBL_37_PRINT2_LIST(KEYNAME,KEYVALUE,CREATE_DATE,CREATE_BY,CREATE_TIMES) values('" & cbLblType.Text & "', '" & strKey & "', sysdate, '" & gUserName & "', '" & iQty & "')")

End Sub
