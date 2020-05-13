VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form FrmLblPrint37 
   BackColor       =   &H00C0C0C0&
   Caption         =   "标签打印系统_37(二维码)"
   ClientHeight    =   12195
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
   ScaleHeight     =   12195
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   1535
      ButtonWidth     =   1032
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "打印"
            Key             =   "PRINT"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "删除"
            Key             =   "DEL"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "导出"
            Key             =   "EXPORT"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "EXIT"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4800
         Top             =   120
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
               Picture         =   "FrmLblPrint37.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint37.frx":213A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint37.frx":4FC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint37.frx":7776
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint37.frx":98B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint37.frx":C062
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint37.frx":E814
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint37.frx":11896
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint37.frx":14048
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint37.frx":14362
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint37.frx":1503C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint37.frx":180BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint37.frx":1A870
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTTab0 
      Height          =   13455
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   20325
      _ExtentX        =   35851
      _ExtentY        =   23733
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483637
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "标签打印"
      TabPicture(0)   =   "FrmLblPrint37.frx":1B14A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraMnu"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraScanDetail"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "标签补打"
      TabPicture(1)   =   "FrmLblPrint37.frx":1B166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label266"
      Tab(1).Control(2)=   "lblType"
      Tab(1).Control(3)=   "lblBarcodeScan2"
      Tab(1).Control(4)=   "txtUser2"
      Tab(1).Control(5)=   "txtUser"
      Tab(1).Control(6)=   "txtPassWd"
      Tab(1).Control(7)=   "Command1"
      Tab(1).Control(8)=   "txtPassWd2"
      Tab(1).Control(9)=   "cbLblType"
      Tab(1).Control(10)=   "txtScan2"
      Tab(1).Control(11)=   "txtDN2"
      Tab(1).ControlCount=   12
      Begin VB.TextBox txtDN2 
         Height          =   375
         Left            =   -67680
         TabIndex        =   35
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Frame fraScanDetail 
         Caption         =   "扫描明细"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   10335
         Left            =   240
         TabIndex        =   19
         Top             =   1560
         Width           =   19815
         Begin FPSpreadADO.fpSpread fpS 
            Height          =   9255
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   600
            Width           =   9255
            _Version        =   524288
            _ExtentX        =   16325
            _ExtentY        =   16325
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
            MaxCols         =   6
            MaxRows         =   0
            SpreadDesigner  =   "FrmLblPrint37.frx":1B182
            AppearanceStyle =   0
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
            Height          =   6180
            Left            =   9600
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   3720
            Width           =   9855
         End
         Begin FPSpreadADO.fpSpread fpS 
            Height          =   3015
            Index           =   1
            Left            =   9600
            TabIndex        =   20
            Top             =   600
            Width           =   5055
            _Version        =   524288
            _ExtentX        =   8916
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
            SpreadDesigner  =   "FrmLblPrint37.frx":1B5A4
            TextTip         =   2
            AppearanceStyle =   0
         End
         Begin FPSpreadADO.fpSpread fpS 
            Height          =   3015
            Index           =   2
            Left            =   14760
            TabIndex        =   21
            Top             =   600
            Width           =   4695
            _Version        =   524288
            _ExtentX        =   8281
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
            SpreadDesigner  =   "FrmLblPrint37.frx":1BA16
            TextTip         =   2
            AppearanceStyle =   0
         End
         Begin VB.Label lblReelList 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "卷盘已扫描:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   960
         End
         Begin VB.Label lblMP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "机种已扫描:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9600
            TabIndex        =   23
            Top             =   360
            Width           =   960
         End
         Begin VB.Label lblJOBList 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "JOB已扫描:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   14760
            TabIndex        =   22
            Top             =   360
            Width           =   885
         End
      End
      Begin VB.Frame fraMnu 
         Caption         =   "DN明细"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   19815
         Begin VB.CheckBox Check1 
            Caption         =   "外箱标签测试用"
            Height          =   255
            Left            =   9120
            TabIndex        =   36
            Top             =   690
            Width           =   2055
         End
         Begin VB.TextBox txtShipTo 
            BackColor       =   &H00FFC0FF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtCurOP 
            BackColor       =   &H00FFC0FF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7560
            TabIndex        =   32
            Text            =   "1"
            Top             =   675
            Width           =   975
         End
         Begin VB.TextBox txtMaxOP 
            BackColor       =   &H00FFC0FF&
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   675
            Width           =   1455
         End
         Begin VB.TextBox txtReelID 
            BackColor       =   &H00FFC0FF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   480
            TabIndex        =   27
            Top             =   675
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox txtDN 
            BackColor       =   &H00FFC0FF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   480
            TabIndex        =   15
            Top             =   330
            Width           =   2295
         End
         Begin VB.TextBox txtQty 
            BackColor       =   &H00FFC0FF&
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   330
            Width           =   1455
         End
         Begin VB.Label lblShipTo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出往"
            Height          =   195
            Left            =   7080
            TabIndex        =   33
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lblCurOP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "当前外箱序号"
            Height          =   195
            Left            =   6480
            TabIndex        =   31
            Top             =   720
            Width           =   1080
         End
         Begin VB.Label lblMaxOp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "总箱数"
            Height          =   195
            Left            =   3720
            TabIndex        =   29
            Top             =   720
            Width           =   840
         End
         Begin VB.Label lblReelID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "卷盘"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Visible         =   0   'False
            Width           =   360
         End
         Begin WMPLibCtl.WindowsMediaPlayer player1 
            Height          =   495
            Left            =   14160
            TabIndex        =   18
            Top             =   360
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
         Begin VB.Label lblDN 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DN"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   17
            Top             =   375
            Width           =   210
         End
         Begin VB.Label lblQTY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "总数量(颗)"
            Height          =   195
            Left            =   3720
            TabIndex        =   16
            Top             =   375
            Width           =   840
         End
      End
      Begin VB.TextBox txtScan2 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   -73080
         TabIndex        =   8
         Top             =   2205
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox cbLblType 
         BackColor       =   &H00FFC0FF&
         Height          =   315
         ItemData        =   "FrmLblPrint37.frx":1BE88
         Left            =   -73080
         List            =   "FrmLblPrint37.frx":1BE8A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1830
         Width           =   3735
      End
      Begin VB.TextBox txtPassWd2 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -71760
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   3405
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "验证补打密码"
         Height          =   840
         Left            =   -68640
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2940
         Width           =   1575
      End
      Begin VB.TextBox txtPassWd 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -71760
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2933
         Width           =   2895
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H00FFC0FF&
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -73080
         TabIndex        =   3
         Text            =   "10354"
         Top             =   2933
         Width           =   1215
      End
      Begin VB.TextBox txtUser2 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   -73080
         TabIndex        =   2
         Top             =   3405
         Width           =   1215
      End
      Begin VB.Label lblBarcodeScan2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描标签条码"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -74520
         TabIndex        =   12
         Top             =   2220
         Width           =   1350
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "补打标签类型"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74520
         TabIndex        =   11
         Top             =   1860
         Width           =   1350
      End
      Begin VB.Label Label266 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmLblPrint37.frx":1BE8C
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   -74640
         TabIndex        =   10
         Top             =   3465
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmLblPrint37.frx":1BEA0
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   -74640
         TabIndex        =   9
         Top             =   3000
         Width           =   1500
      End
   End
End
Attribute VB_Name = "FrmLblPrint37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilllMicroSeconds As Long)

Private Const consDNLen = 8

Private Const consReelIDLen = 13

Private Const localSoundDir = "C:\media_source\"

Private Const gSleepMicSec = 2000

Private Const serverSoundDir = "\\10.160.1.84\public\media_source\37HW\"

Private strFlagPath         As String

Private str37BCIDPath       As String

Private str37CartonPath     As String

Private strSSBoxPath        As String

Private strSSBoxPath2       As String

Private strSSBoxPath_Short  As String

Private strSSReelPath       As String

Private strSSReelPath2      As String

Private strSSReelPath_Short As String

Private strSSCartonPath     As String

Private strHTQCartonPath    As String

Private strHWBoxPath        As String

Private strHWReelPath       As String

Private gMediaDir           As String

Private Type CusReel

    PN As String
    lot As String
    DEV As String
    QTY As String
    TRAYID As String

End Type

Private Type CusBox

    DEV As String
    PN As String
    QTY As String

End Type

Private Type HWBox

    CPN As String
    MPN As String
    PODATE As String
    lot As String
    QTY As String
    PSN As String

End Type

Private Type STBox

    JOB As String
    DEV As String
    FactoryFlow As String
    lot As String
    QTY As String
    DATECODE As String
    testdateCode As String

End Type

Private Type STReel

    JOB As String
    DEV As String
    FactoryFlow As String
    lot As String
    QTY As String
    DATECODE As String
    testdateCode As String

End Type

Private Type STCarton

    JOB As String
    DEV As String
    FactoryFlow As String
    lot As String
    QTY As String
    DATECODE As String
    testdateCode As String

End Type

Private Type CUSCARTON

    dn As String
    PO As String
    CPN As String
    FactoryFlow As String
    MPN As String
    JOB As String
    QTY As String
    KID As String
    DATECODE As String

End Type

Private Enum E_REEL_SCAN

    E_REEL_OP_NO = 1
    E_REEL_IP_NO
    E_REEL_RID
    E_REEL_PSN
    E_REEL_JOBID
    E_REEL_SEQ
    E_REEL_SCANTIME
    E_MAX_COL

End Enum

Private Enum E_MPN_SCAN

    E_MPN_ID = 1
    E_MPN_TOTAL_QTY
    E_MPN_CUR_QTY
    E_MAX_COL

End Enum

Private Enum E_JOB_SCAN

    E_JOB_ID = 1
    E_JOB_TOTAL_QTY
    E_JOB_CUR_QTY
    E_MAX_COL

End Enum

Private Type T_REELINFO

    T_TRAYID As String
    T_INBOX_NUM As Long
    T_OUTBOX_NUM As Long
    T_DN_NUM As String
    T_JOB_ID As String
    T_QTY As Long
    T_MPN As String
    T_CREATE_DATE As String
    T_CREATE_BY As String
    T_PRINT_FLAG As String
    T_FLAG As String
    T_CARTON As String
    T_REELID As String
    T_BOXID As String
    T_CARTONID As String
    T_KID As String
    T_SEQ As String
    T_DATECODE As String

End Type

Dim bCheckDC       As Boolean
Dim strLastRightDC As String

Private Sub SSTTab0_Click(PreviousTab As Integer)

Select Case SSTTab0.Tab

    Case 0
        Toolbar1.Buttons("PRINT").Enabled = False
        Toolbar1.Buttons("PRINT").Caption = "打印标签"
        Toolbar1.Buttons("DEL").Enabled = True
        Toolbar1.Buttons("EXPORT").Enabled = True
        Toolbar1.Buttons("EXPORT").Caption = "导出打印记录"

    Case 1
        Toolbar1.Buttons("PRINT").Enabled = True
        Toolbar1.Buttons("PRINT").Caption = "补打标签"
        Toolbar1.Buttons("DEL").Enabled = False
        Toolbar1.Buttons("EXPORT").Enabled = True
        Toolbar1.Buttons("EXPORT").Caption = "导出补打记录"

End Select

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       Toolbar1_ButtonClick
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/8-13:24:18
'
' Parameters :       Button (MSComctlLib.Button)
'--------------------------------------------------------------------------------
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case SSTTab0.Tab

    Case 0

        Select Case Button.Key

            Case "PRINT"
                Call PrintHandler

            Case "DEL"
                Call DeleteHandler

            Case "EXPORT"
                Call ExportHandler

            Case "EXIT"
                Unload Me

        End Select

    Case 1

        Select Case Button.Key

            Case "PRINT"
                Call PrintHandler2

            Case "EXPORT"
                Call ExportHandler2

            Case "EXIT"
                Unload Me

        End Select

    Case Else

End Select

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       DeleteHandler
' Description:       删除DN
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-12:15:17
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub DeleteHandler()
DialogDNDel.Show 1

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ExportHandler
' Description:       导出DN纪录
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-12:15:05
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub ExportHandler()
Dim strDN  As String
Dim strSql As String

strDN = Trim$(txtDN.text)
If Len(strDN) = 0 Then
    MsgBox "请输入要导出的DN", vbInformation, "提示"
    Exit Sub

End If

strSql = "select dn_num dn, OUTBOX_NUM 外箱,INBOX_NUM 内箱,trayid, reelid PSN,boxid,cartonid,job_id job, customer_device 客户机种, QTY 数量, KID, DATECODE, CREATE_BY 打印人员, CREATE_DATE 打印时间,'' as 备注 from packing_detailed where dn_num = '" & strDN & "' order by seq  "
ExporToExcel (strSql)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ExportHandler2
' Description:       导出补打记录
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-12:14:44
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub ExportHandler2()
Dim strSql As String

strSql = "select KEYNAME 补打类型,keyvalue 补打值,CREATE_DATE 补打时间,CREATE_BY 补打人员工号,CREATE_TIMES 第几次补打 from TBL_37_PRINT2_LIST order by CREATE_date desc"
Call ExporToExcel(strSql)

End Sub

Private Sub setPrintPath()
' 标志
strFlagPath = "\\10.160.1.84\public\BarCode\37\37Flag\"
' 37 BID, CID小标签
'str37BCIDPath = "\\10.160.1.84\public\BarCode\37\37内箱\"        ' 37B,C,R小标签
str37BCIDPath = "\\10.160.1.84\public\BarCode\37\37内盒带二维码\"   'QR
'str37CartonPath = "\\10.160.1.84\public\BarCode\37\37外箱\"      ' 37自家外箱大标签
str37CartonPath = "\\10.160.1.84\public\BarCode\37\37外箱带二维码\"
' 出三星标签
strSSBoxPath = "\\10.160.1.84\public\BarCode\37\37BoxNH\"      ' 三星内盒小标签E2
strSSBoxPath2 = "\\10.160.1.84\public\BarCode\37\37BoxNH-新\"  ' 三星内盒小标签特定机种
strSSBoxPath_Short = "\\10.160.1.84\public\BarCode\37\37NH2\"  ' 三星内盒小标签SHORT
strSSReelPath = "\\10.160.1.84\public\BarCode\37\37BoxJP\"     ' 三星卷盘小标签E2
strSSReelPath2 = "\\10.160.1.84\public\BarCode\37\37BoxJP-新\" ' 三星卷盘小标签特定机种
strSSReelPath_Short = "\\10.160.1.84\public\BarCode\37\37JP2\" ' 三星卷盘小标签SHORT
strSSCartonPath = "\\10.160.1.84\public\BarCode\37\37BoxOut\"  ' 三星外箱大标签E2
' 华天Q标签
strHTQCartonPath = "\\10.160.1.84\public\BarCode\37\37Box\"    ' 华天Q箱号小标签
' 出华为
strHWBoxPath = "\\10.160.1.84\public\BarCode\37\37HW\HW内盒\"
strHWReelPath = "\\10.160.1.84\public\BarCode\37\37HW\HW卷盘\"

End Sub

Private Sub setTestPrintPath()
' 标志
strFlagPath = "C:\test\"
' 37 BID, CID小标签
str37BCIDPath = "C:\test\"      ' 37B,C,R小标签
str37CartonPath = "C:\test\"     ' 37自家外箱大标签
' 出三星标签
strSSBoxPath = "C:\test\"      ' 三星内盒小标签
strSSBoxPath2 = "C:\test\"  ' 三星内盒小标签特定机种
strSSBoxPath_Short = "C:\test\"  ' 三星内盒小标签SHORT
strSSReelPath = "C:\test\"    ' 三星卷盘小标签
strSSReelPath2 = "C:\test\"  ' 三星卷盘小标签特定机种
strSSReelPath_Short = "C:\test\" ' 三星卷盘小标签SHORT
strSSCartonPath = "C:\test\" ' 三星外箱大标签
' 华天Q标签
strHTQCartonPath = "C:\test\"   ' 华天Q箱号小标签
' 出华为
strHWBoxPath = "C:\test\"
strHWReelPath = "C:\test\"

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       Form_Activate
' Description:       窗体开始
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-7-5-15:06:41
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Activate()
SSTTab0.Tab = 0
If txtDN.Enabled Then
    txtDN.SetFocus

End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       Form_Load
' Description:       窗体加载
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-7-5-15:07:18
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Load()
Call InitCtrls
Call InitData

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       initCtrls
' Description:       初始化控件
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-7-5-15:07:40
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitCtrls()
Call InitFps
cbLblType.AddItem "37内盒-B标签"
cbLblType.AddItem "37外箱-C标签"
cbLblType.AddItem "37外箱标准大标签"
cbLblType.AddItem "三星卷盘标签"
cbLblType.AddItem "三星内盒标签"
cbLblType.AddItem "三星外箱大标签"
cbLblType.AddItem "华为卷盘标签"
cbLblType.AddItem "华为内盒标签"
cbLblType.AddItem "华为外箱标准大标签"
cbLblType.AddItem "卷盘二维码标签转换"

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       initFps
' Description:       初始化Fps
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-9:43:31
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitFps()

'REEL Fps
With fpS(0)
    .ReDraw = False
    .MaxCols = E_REEL_SCAN.E_MAX_COL - 1
    .MaxRows = 0
    .FontBold = False
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .SelForeColor = &HFF8080
    .SetText E_REEL_SCAN.E_REEL_OP_NO, 0, "外箱"
    .SetText E_REEL_SCAN.E_REEL_IP_NO, 0, "内箱"
    .SetText E_REEL_SCAN.E_REEL_RID, 0, "卷盘ID"
    .SetText E_REEL_SCAN.E_REEL_PSN, 0, "卷盘PSN"
    .SetText E_REEL_SCAN.E_REEL_JOBID, 0, "当前JOB"
    .SetText E_REEL_SCAN.E_REEL_SEQ, 0, "第几卷"
    .SetText E_REEL_SCAN.E_REEL_SCANTIME, 0, "扫描时间"
    .ColWidth(E_REEL_SCAN.E_REEL_OP_NO) = 4
    .ColWidth(E_REEL_SCAN.E_REEL_IP_NO) = 4
    .ColWidth(E_REEL_SCAN.E_REEL_RID) = 14
    .ColWidth(E_REEL_SCAN.E_REEL_PSN) = 16
    .ColWidth(E_REEL_SCAN.E_REEL_JOBID) = 8
    .ColWidth(E_REEL_SCAN.E_REEL_SEQ) = 6
    .ColWidth(E_REEL_SCAN.E_REEL_SCANTIME) = 16
    .ReDraw = True

End With

'MPN Fps
With fpS(1)
    .ReDraw = False
    .MaxCols = E_MPN_SCAN.E_MAX_COL - 1
    .MaxRows = 0
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .SelForeColor = &HFF8080
    .SetText E_MPN_SCAN.E_MPN_ID, 0, "机种名"
    .SetText E_MPN_SCAN.E_MPN_TOTAL_QTY, 0, "总数量"
    .SetText E_MPN_SCAN.E_MPN_CUR_QTY, 0, "已扫描数量"
    .ColWidth(E_MPN_SCAN.E_MPN_ID) = 14
    .ColWidth(E_MPN_SCAN.E_MPN_TOTAL_QTY) = 8
    .ColWidth(E_MPN_SCAN.E_MPN_CUR_QTY) = 8
    .ReDraw = True

End With

'JOB Fps
With fpS(2)
    .ReDraw = False
    .MaxCols = E_JOB_SCAN.E_MAX_COL - 1
    .MaxRows = 0
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .SelForeColor = &HFF8080
    .SetText E_JOB_SCAN.E_JOB_ID, 0, "JOBID"
    .SetText E_JOB_SCAN.E_JOB_TOTAL_QTY, 0, "总数量"
    .SetText E_JOB_SCAN.E_JOB_CUR_QTY, 0, "已扫描数量"
    .ColWidth(E_JOB_SCAN.E_JOB_ID) = 14
    .ColWidth(E_JOB_SCAN.E_JOB_TOTAL_QTY) = 8
    .ColWidth(E_JOB_SCAN.E_JOB_CUR_QTY) = 8
    .ReDraw = True

End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       InitData
' Description:       初始化数据
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-7-5-15:25:15
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitData()
gMediaDir = localSoundDir
bCheckDC = True

Select Case gUserName

    Case "07885"
        Call setTestPrintPath

    Case Else
        Call setPrintPath

End Select

If Dir(gMediaDir, vbDirectory) = "" Then
    gMediaDir = serverSoundDir

End If

If Dir(gMediaDir, vbDirectory) = "" Then
    MsgBox "找不到声音文件,请联系IT处理", vbInformation, "警告"

End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       txtDN_KeyPress
' Description:       扫描DN
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-7-5-15:15:59
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub txtDN_KeyPress(KeyAscii As Integer)
Dim strDN As String

If KeyAscii <> vbKeyReturn Then Exit Sub
strDN = Right$(Trim(txtDN.text), consDNLen)
If Len(strDN) <> consDNLen Then
    MsgBox "请扫描正确的DN", vbInformation, "DN扫描"
    txtDN.text = ""
    Exit Sub

End If

If CheckDN(strDN) = False Then
    txtDN.text = ""
    Exit Sub

End If

Call ShowDNInfo(strDN)
Call ShowScanInfo(strDN)
PlaySound ("D N已获取,请依次扫描挑料卷盘")
txtDN.Enabled = False
lblReelID.Visible = True
txtReelID.Visible = True
txtReelID.SetFocus
Call CheckScanningComplate(strDN)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckDN
' Description:       检查DN是否正确
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-7-5-15:35:05
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function CheckDN(strDN As String) As Boolean
Dim strSql As String

CheckDN = False
strSql = "SELECT * FROM CUSTOMERSHIPPINGUPTBL WHERE DELIVERY = '" & strDN & "'"
If Get_OracleCnt(strSql) = 0 Then
    MsgBox "DN:" & strDN & " 不正确或市场部未上传该DN", vbExclamation, "DN检查"
    Exit Function

End If

CheckDN = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       showDNInfo
' Description:       获取DN信息
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-7-2-16:58:14
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub ShowDNInfo(strDN As String)
Dim strSql    As String
Dim strShipTo As String

strSql = "select labelrequirement from customershippinguptbl where delivery = '" & strDN & "'"
strShipTo = UCase(Get_OracleStr(strSql))
If InStr(strShipTo, "HUAWEI") Then
    txtShipTo.text = "HUAWEI"

End If

If InStr(strShipTo, "E2") Then
    txtShipTo.text = "SSE2"

End If

If InStr(strShipTo, "SEMTECH") Then
    txtShipTo.text = "ST"

End If

If InStr(strShipTo, "SHORT") Then
    txtShipTo.text = "SSSHORT"

End If

txtDN.text = strDN
strSql = "select sum(quantity) from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "'"
txtQty.text = Get_OracleStr(strSql)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ShowScanInfo
' Description:       刷新已扫描状态
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-7-2-17:06:25
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub ShowScanInfo(strDN As String)
Call ShowScanningDetailByReels(strDN)
Call ShowScanningDetailByMPN(strDN)
Call ShowScanningDetailByJob(strDN)
Call ShowScannedReels(strDN)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ShowScanningDetailByReels
' Description:       卷盘视图
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/8-14:18:38
'
' Parameters :       strDN As String
'--------------------------------------------------------------------------------
Private Sub ShowScanningDetailByReels(strDN As String)
Dim strSql  As String
Dim rsReels As New ADODB.Recordset

strSql = "SELECT OUTBOX_NUM 外箱,INBOX_NUM 内箱,TRAYID 卷盘ID,REELID PSN,JOB_ID JOBID,SEQ 第几卷,CREATE_DATE 扫描时间 FROM PACKING_DETAILED WHERE DN_NUM = '" & strDN & "' ORDER BY SEQ DESC"
Set rsReels = Get_OracleRs(strSql)

With fpS(0)
    .MaxRows = 0
    If rsReels.RecordCount > 0 Then
        Set .DataSource = rsReels

    End If

End With

Set rsReels = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ShowScanningDetailByMPN
' Description:       机种视图
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/8-14:19:17
'
' Parameters :       strDN As String
'--------------------------------------------------------------------------------
Private Sub ShowScanningDetailByMPN(strDN As String)
Dim strSql As String
Dim rsMPN  As New ADODB.Recordset

strSql = "SELECT AA.MARKETINGPN,AA.REALQTYS, BB.THISQTYS FROM (SELECT MARKETINGPN, SUM(QUANTITY) AS REALQTYS FROM CUSTOMERSHIPPINGUPTBL " & " WHERE DELIVERY = '" & strDN & "' GROUP BY MARKETINGPN) AA " & " LEFT JOIN (SELECT CUSTOMER_DEVICE,SUM(QTY) AS THISQTYS FROM PACKING_DETAILED WHERE DN_NUM = '" & strDN & "' GROUP BY CUSTOMER_DEVICE) BB ON AA.MARKETINGPN = BB.CUSTOMER_DEVICE "
Set rsMPN = Get_OracleRs(strSql)

With fpS(1)
    .MaxRows = 0
    If rsMPN.RecordCount > 0 Then
        Set .DataSource = rsMPN

    End If

End With

Set rsMPN = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ShowScanningDetailByJob
' Description:       Job视图
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/8-14:19:43
'
' Parameters :       strDN As String
'--------------------------------------------------------------------------------
Private Sub ShowScanningDetailByJob(strDN As String)
Dim strSql As String
Dim rsJob  As New ADODB.Recordset

strSql = " SELECT AA.BATCHNUMBER,AA.REALQTYS, BB.THISQTYS FROM (SELECT BATCHNUMBER, SUM(QUANTITY) AS REALQTYS FROM CUSTOMERSHIPPINGUPTBL WHERE DELIVERY = '" & strDN & "' GROUP BY BATCHNUMBER) AA LEFT JOIN (SELECT JOB_ID, SUM(QTY) AS THISQTYS FROM PACKING_DETAILED WHERE DN_NUM = '" & strDN & "' GROUP BY JOB_ID) BB ON AA.BATCHNUMBER = BB.JOB_ID "
Set rsJob = Get_OracleRs(strSql)

With fpS(2)
    .MaxRows = 0
    If rsJob.RecordCount > 0 Then
        Set .DataSource = rsJob

    End If

End With

Set rsJob = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ShowScannedReels
' Description:       显示已扫描的卷盘数量
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/8-14:29:02
'
' Parameters :       strDN as String
'--------------------------------------------------------------------------------
Private Sub ShowScannedReels(strDN As String)
Dim strSql As String

strSql = "select nvl(count(*), 0) from PACKING_DETAILED where DN_NUM = '" & strDN & "'"
txtStatus.BackColor = vbWhite
txtStatus.text = vbCrLf & Get_OracleNo(strSql)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       txtReelID_Change
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-7-5-17:15:45
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub txtReelID_KeyPress(KeyAscii As Integer)
Dim tReel      As T_REELINFO
Dim strBarcode As String
Dim strQrCode  As String

If KeyAscii <> vbKeyReturn Then Exit Sub
If bCheckDC = False Then
    If CheckReelDC(Right(UCase(Trim(txtReelID.text)), 4)) = True Then
        bCheckDC = True
        Call PlaySound("D C正确,请继续扫描卷盘")
        If CheckScanningComplate(txtDN.text) Then
            Call PlaySound("该DN所有卷盘已全部扫描完毕,请点击打印按钮,开始打印标签")

        End If

    End If

Else
    If UCase(Left$(Trim(txtReelID.text), 3)) = "[)>" Then
        MsgBox "请扫描卷盘的条码", vbInformation, "提示"
        Exit Sub
        strQrCode = UCase(Trim(txtReelID.text))    '二维码
    ElseIf UCase(Left$(Trim(txtReelID.text), 1)) = "S" Then
        strBarcode = UCase(Trim(txtReelID.text))   '条码
    Else
        MsgBox "请扫描正确的卷盘号", vbInformation, "卷盘扫描"
        txtReelID.text = ""
        Exit Sub

    End If

    If strQrCode <> "" Then
        Call GetReelInfoByQrCode(tReel, strQrCode)
    Else
        Call GetReelInfoByBarCode(tReel, strBarcode)

    End If

    If CheckReelID(tReel) = False Then
        txtStatus.BackColor = vbRed
        txtReelID.text = ""
        Exit Sub
    Else
        txtStatus.BackColor = vbWhite

    End If

    Call GetOtherData(tReel)
    Call SavePackingDetail(tReel)
    Call ShowScanInfo(tReel.T_DN_NUM)
    Call PlaySound("卷盘扫描正确,请扫描D C")
    strLastRightDC = GetLastRightDC(tReel.T_TRAYID)
    bCheckDC = False

End If

txtReelID.text = ""

End Sub

Private Function CheckReelDC(strDC As String) As Boolean
CheckReelDC = False
If strDC <> strLastRightDC Then
    MsgBox "扫描的卷盘DC:" & strDC & ",DN上的是:" & strLastRightDC & ",不一致,请确认是否有误", vbCritical, "警告"
    Exit Function

End If

CheckReelDC = True

End Function

Private Function GetLastRightDC(strReelID As String) As String
Dim strJobID As String
Dim strSql   As String

strJobID = GetJobID(strReelID)
strSql = "select distinct date_code from customershippinguptbl where  delivery = '" & Trim(txtDN.text) & "' and batchnumber = '" & strJobID & "' "
GetLastRightDC = Get_OracleStr(strSql)

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetReelInfoByBarCode
' Description:       根据条码获取信息
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/28-16:00:57
'
' Parameters :       tREEL (T_REELINFO)
'                    strBarCode (String)
'--------------------------------------------------------------------------------
Private Sub GetReelInfoByBarCode(ByRef tReel As T_REELINFO, strBarcode As String)
tReel.T_DN_NUM = Trim$(txtDN.text)
tReel.T_TRAYID = strBarcode
tReel.T_JOB_ID = GetJobID(tReel.T_TRAYID)
tReel.T_QTY = GetReelQty(tReel.T_TRAYID)
tReel.T_DATECODE = Get37TestDC(tReel.T_DN_NUM, tReel.T_JOB_ID)
tReel.T_MPN = GetCustPN(tReel.T_DN_NUM, tReel.T_JOB_ID)
tReel.T_SEQ = GetSeq(tReel.T_DN_NUM)
tReel.T_CREATE_DATE = Now
tReel.T_CREATE_BY = gUserName
tReel.T_PRINT_FLAG = "0"
tReel.T_FLAG = "0"
tReel.T_REELID = GetCustReelID(tReel.T_DN_NUM, tReel.T_JOB_ID, tReel.T_TRAYID)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetReelInfoByQrCode
' Description:       根据二维码获取信息
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/28-15:59:01
'
' Parameters :       tREEL (T_REELINFO)
'                    strQrCode (String)
'--------------------------------------------------------------------------------
Private Sub GetReelInfoByQrCode(ByRef tReel As T_REELINFO, strQrCode As String)
tReel.T_DN_NUM = Trim$(txtDN.text)
tReel.T_TRAYID = Mid(strQrCode, InStr(strQrCode, "S"), InStr(strQrCode, "-R") - InStr(strQrCode, "S")) & Mid(strQrCode, InStr(strQrCode, "-R"), InStr(Mid(strQrCode, InStr(strQrCode, "-R")), "Q") - 1)
tReel.T_JOB_ID = Mid(strQrCode, InStr(strQrCode, "1T") + 2, InStr(strQrCode, "1P") - InStr(strQrCode, "1T") - 2)
tReel.T_QTY = Mid(Mid(strQrCode, InStr(strQrCode, "-R")), InStr(Mid(strQrCode, InStr(strQrCode, "-R")), "Q") + 1, InStr(Mid(strQrCode, InStr(strQrCode, "-R")), "6P") - InStr(Mid(strQrCode, InStr(strQrCode, "-R")), "Q") - 1)
tReel.T_DATECODE = Right$(strQrCode, 4)
tReel.T_MPN = Mid(strQrCode, InStr(strQrCode, "1P") + 2, InStr(Mid$(strQrCode, InStr(strQrCode, "1P")), "S") - 3)
tReel.T_SEQ = GetSeq(tReel.T_DN_NUM)
tReel.T_CREATE_DATE = Now
tReel.T_CREATE_BY = gUserName
tReel.T_PRINT_FLAG = "0"
tReel.T_FLAG = "0"
tReel.T_REELID = GetCustReelID(tReel.T_DN_NUM, tReel.T_JOB_ID, tReel.T_TRAYID)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckReelID
' Description:       检查reelID
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/8-14:52:40
'
' Parameters :       strReelID (String)
'--------------------------------------------------------------------------------
Private Function CheckReelID(tReel As T_REELINFO) As Boolean
Dim strSql As String

CheckReelID = False
If tReel.T_JOB_ID = "" Then
    Exit Function

End If

strSql = "select * from packing_detailed where trayid = '" & tReel.T_TRAYID & "' and dn_num <> '" & tReel.T_DN_NUM & "'"
If Get_OracleCnt(strSql) > 0 Then

    MsgBox "该卷盘: " & tReel.T_TRAYID & " 有扫描历史,请确认是否有误", vbCritical, "警告"
    Exit Function
End If

strSql = "select * from packing_detailed where dn_num = '" & tReel.T_DN_NUM & "' and trayid = '" & tReel.T_TRAYID & "'"
If Get_OracleCnt(strSql) > 0 Then
    Call PlaySound("该卷盘已经扫描过, 请勿重复扫描")
    Exit Function

End If

strSql = "SELECT * FROM erpdata..tblStockNumSub where 箱号 = '" & tReel.T_TRAYID & "' "
If Get_SqlserverCnt(strSql) = 0 Then
    MsgBox "该卷盘还没有入库,请先入库,否则无法扫描打印", vbCritical, "提示"
    
    Exit Function

End If

strSql = "select * from customershippinguptbl where delivery =  '" & tReel.T_DN_NUM & "' and batchnumber = '" & tReel.T_JOB_ID & "'"
If Get_OracleCnt(strSql) = 0 Then
    MsgBox "该卷盘: " & tReel.T_TRAYID & " 的JobID: " & tReel.T_JOB_ID & " 不属于本次DN: " & tReel.T_DN_NUM, vbCritical, "警告"
    Exit Function

End If

If CheckSamgJob_PN(tReel.T_DN_NUM, tReel.T_JOB_ID, tReel.T_QTY, tReel.T_MPN) = False Then
    Exit Function

End If

Call PlaySound("卷盘号正确")
CheckReelID = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckReelDC
' Description:       比对卷盘 DC和DN DC是否一致
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/9/26-10:04:19
'
' Parameters :       strReelDC (String)
'--------------------------------------------------------------------------------
'Private Function CheckReelDC(tReel As T_REELINFO) As Boolean
'Dim strSql As String
'Dim strDNDC As String
'CheckReelDC = False
'
'strSql = "select distinct date_code from customershippinguptbl where delivery = '" & tReel.T_DN_NUM & "' and batchnumber ='" & tReel.T_JOB_ID & "'"
'strDNDC = Get_OracleStr(strSql)
'
'If tReel.T_DATECODE <> strDNDC Then
'    MsgBox "该卷盘DC: " & tReel.T_DATECODE & "和 DN DC" & strDNDC & " 不一致 ", vbCritical, "警告"
'    Exit Function
'End If
'
'CheckReelDC = True
'End Function
'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckSamgJob_PN
' Description:       不可以跨JOB作业
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/8-15:51:31
'
' Parameters :       strReelID (String)
'                    strJobID (String)
'                    strLastJobID (String)
'--------------------------------------------------------------------------------
Private Function CheckSamgJob_PN(strDN As String, _
                                 strJobID As String, _
                                 lReelQty As Long, _
                                 strPN As String) As Boolean
CheckSamgJob_PN = False
Dim strSql           As String
Dim strLastJobID     As String
Dim lLastJobCurQty   As Long
Dim lLastJobTotalQty As Long
Dim strLastPN        As String
Dim lLastPNCurQty    As Long
Dim lLastPNTotalQty  As Long

With fpS(0)
    .Row = 1
    .Col = E_REEL_SCAN.E_REEL_JOBID
    strLastJobID = .text

End With

strLastPN = Get_OracleStr("select distinct marketingpn from customershippinguptbl where delivery = '" & strDN & "' and batchnumber = '" & strLastJobID & "'")
'判断上一个JOB是否满
If strJobID = strLastJobID Then
    strSql = "select sum(quantity) from customershippinguptbl where delivery = '" & strDN & "' and batchnumber = '" & strLastJobID & "'"
    lLastJobTotalQty = Get_OracleNo(strSql)
    strSql = "select sum(qty) from packing_detailed where dn_num = '" & strDN & "' and job_id = '" & strLastJobID & "'"
    lLastJobCurQty = Get_OracleNo(strSql)
    If (lLastJobCurQty + lReelQty) > lLastJobTotalQty Then
        MsgBox "JOBID: " & strJobID & "数量超出,挑料出错", vbCritical, "警告"
        Exit Function

    End If

Else
    strSql = "select sum(quantity) from customershippinguptbl where delivery = '" & strDN & "' and batchnumber = '" & strLastJobID & "'"
    lLastJobTotalQty = Get_OracleNo(strSql)
    strSql = "select sum(qty) from packing_detailed where dn_num = '" & strDN & "' and job_id = '" & strLastJobID & "'"
    lLastJobCurQty = Get_OracleNo(strSql)
    If lLastJobCurQty < lLastJobTotalQty Then
        Call PlaySound("上一个JOB没有扫完,请勿扫描其他JOB的卷盘")
        MsgBox "JOBID: " & strLastJobID & "数量未满,请继续扫描该JOB的卷盘", vbCritical, "警告"
        Exit Function

    End If

End If

'判断上一个机种是否满
If strPN = strLastPN Then
    strSql = "select sum(quantity) from customershippinguptbl where delivery = '" & strDN & "' and marketingpn = '" & strLastPN & "'"
    lLastPNTotalQty = Get_OracleNo(strSql)
    strSql = "select sum(qty) from packing_detailed where dn_num = '" & strDN & "' and customer_device = '" & strLastPN & "'"
    lLastPNCurQty = Get_OracleNo(strSql)
    If (lLastPNCurQty + lReelQty) > lLastPNTotalQty Then
        MsgBox "机种: " & strLastPN & "数量超出,挑料出错", vbCritical, "警告"
        Exit Function

    End If

Else
    strSql = "select sum(quantity) from customershippinguptbl where delivery = '" & strDN & "' and marketingpn = '" & strLastPN & "'"
    lLastPNTotalQty = Get_OracleNo(strSql)
    strSql = "select sum(qty) from packing_detailed where dn_num = '" & strDN & "' and customer_device = '" & strLastPN & "'"
    lLastPNCurQty = Get_OracleNo(strSql)
    If lLastPNCurQty < lLastPNTotalQty Then
        Call PlaySound("上一个机种没有扫完,请勿扫描其他机种的卷盘")
        MsgBox "机种: " & strLastPN & "数量未满,请继续扫描该机种的卷盘", vbCritical, "警告"
        Exit Function

    End If

End If

CheckSamgJob_PN = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetJobID
' Description:       获取JobID
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/8-15:10:22
'
' Parameters :       strReelID (String)
'--------------------------------------------------------------------------------
Private Function GetJobID(strReelID As String) As String
Dim strSql As String
Dim strRes As String

strSql = "select KEY_VALUE from erpdata..tblErpInStockDetailInfo a where SUBSTRING(a.KEY_VALUE,1,charindex('|',a.KEY_VALUE)-1) =  '" & strReelID & "' and a.KEY_NAME = 'CONTAINER_NAME' AND a.KEY_TYPE = 'T' and charindex('|',a.KEY_VALUE) > 0"
strRes = Get_SqlStr(strSql)
GetJobID = Mid(strRes, InStr(strRes, "|") + 1)

If GetJobID = "" Then
    GetJobID = Get_SqlStr("select customerlotid as jobid from erpdata..TblTSV_Tray_details where TRAYQBOXNUMBER = '" & strReelID & "'")
    
End If

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetJobID
' Description:       获取JobID
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/8-15:10:22
'
' Parameters :       strReelID (String)
'--------------------------------------------------------------------------------
Private Function GetReelQty(strReelID As String) As Long
Dim strSql As String

strSql = " select SUM(数量) from erpdata..tblPackMainInfSub where 箱号 = '" & strReelID & "' "
GetReelQty = Get_SqlserverNo(strSql)

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckEnough
' Description:       检查是否全部扫描完毕
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/8-16:27:21
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function CheckScanningComplate(strDN As String) As Boolean
Dim strSql    As String
Dim lCurQty   As Long
Dim lTotalQty As Long
Dim lMaxOP    As Long

CheckScanningComplate = False
strSql = "select sum(quantity) from customershippinguptbl where delivery = '" & strDN & "'"
lTotalQty = Get_OracleNo(strSql)
strSql = "select sum(qty) from packing_detailed where dn_num = '" & strDN & "'"
lCurQty = Get_OracleNo(strSql)
If lCurQty = lTotalQty Then
    strSql = "select max(OUTBOX_NUM) from packing_detailed where dn_num = '" & strDN & "'"
    txtMaxOP.text = Get_OracleNo(strSql)
    txtReelID.Enabled = False
    Toolbar1.Buttons("PRINT").Enabled = True
    CheckScanningComplate = True
    MsgBox "该DN所有卷盘已全部扫描完毕,请点击打印按钮,开始打印标签", vbInformation, "提示"
    Call UpdateERP_CARTON_NO(strDN)

End If

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetOut_IN_BoxNum
' Description:       获取外箱序号
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/9-10:38:59
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function GetOtherData(ByRef tReel As T_REELINFO)
Dim strSql         As String
Dim strLastMPN     As String
Dim strLastJobID   As String
Dim lLastOutBoxNum As Long
Dim lLastInboxNum  As Long
Dim lLastInboxCnt  As Long

strSql = "select nvl(max(OUTBOX_NUM),0) from PACKING_DETAILED where dn_num = '" & tReel.T_DN_NUM & "'"
lLastOutBoxNum = Get_OracleNo(strSql)
tReel.T_OUTBOX_NUM = lLastOutBoxNum
strSql = "select nvl(max(INBOX_NUM),0) from PACKING_DETAILED where dn_num = '" & tReel.T_DN_NUM & "' and OUTBOX_NUM = '" & tReel.T_OUTBOX_NUM & "' "
lLastInboxNum = Get_OracleStr(strSql)
tReel.T_INBOX_NUM = lLastInboxNum
strSql = "select CUSTOMER_DEVICE from packing_DETAILED where dn_num = '" & tReel.T_DN_NUM & "' order by seq desc"
strLastMPN = Get_OracleStr(strSql)
strSql = "select count(*) from packing_detailed where dn_num = '" & tReel.T_DN_NUM & "' and outbox_num = '" & tReel.T_OUTBOX_NUM & "' and inbox_num = '" & tReel.T_INBOX_NUM & "' "
lLastInboxCnt = Get_OracleNo(strSql)
strSql = "select job_id from packing_DETAILED where dn_num = '" & tReel.T_DN_NUM & "' order by seq desc"
strLastJobID = Get_OracleStr(strSql)
'Get OutboxNum InboxNum
If tReel.T_MPN <> strLastMPN Then
    tReel.T_OUTBOX_NUM = lLastOutBoxNum + 1
    tReel.T_INBOX_NUM = 1
Else
    If lLastInboxCnt = 9 Then
        tReel.T_INBOX_NUM = lLastInboxNum + 1
        If tReel.T_INBOX_NUM = 13 Then
            tReel.T_INBOX_NUM = 1
            tReel.T_OUTBOX_NUM = lLastOutBoxNum + 1

        End If

    End If

End If

tReel.T_KID = "K" & tReel.T_OUTBOX_NUM
'GetCID
strSql = "select CARTONID from packing_DETAILED where dn_num = '" & tReel.T_DN_NUM & "' order by seq desc"
tReel.T_CARTONID = Get_OracleStr(strSql)
If tReel.T_OUTBOX_NUM <> lLastOutBoxNum Then
    tReel.T_CARTONID = GetNewID(tReel, "-C")
Else
    If tReel.T_JOB_ID <> strLastJobID Then
        tReel.T_CARTONID = GetNewID(tReel, "-C")

    End If

End If

'GetBID
strSql = "select BOXID from packing_DETAILED where dn_num = '" & tReel.T_DN_NUM & "' order by seq desc"
tReel.T_BOXID = Get_OracleStr(strSql)
If tReel.T_OUTBOX_NUM <> lLastOutBoxNum Then
    tReel.T_BOXID = GetNewID(tReel, "-B")
ElseIf tReel.T_INBOX_NUM <> lLastInboxNum Then
    tReel.T_BOXID = GetNewID(tReel, "-B")
Else
    If tReel.T_JOB_ID <> strLastJobID Then
        tReel.T_BOXID = GetNewID(tReel, "-B")

    End If

End If

'GetQID
strSql = "select CARTON from packing_DETAILED where dn_num = '" & tReel.T_DN_NUM & "' order by seq desc"
tReel.T_CARTON = Get_OracleStr(strSql)
If tReel.T_OUTBOX_NUM <> lLastOutBoxNum Then
    tReel.T_CARTON = GetQID(tReel)

End If

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetNewID
' Description:       获取CID,BID
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/9-16:09:50
'
' Parameters :       ut (uReelInfo)
'                    strFlag (String)
'--------------------------------------------------------------------------------
Private Function GetNewID(tReel As T_REELINFO, strflag As String) As String
Dim strSql   As String
Dim strBase  As String
Dim strseq   As String
Dim strNewID As String

strBase = Left$(tReel.T_TRAYID, InStr(tReel.T_TRAYID, "-") - 1) & strflag
strSql = "select nvl(max(seq)+1, 1) from PKGIDSEQ_37 where val = '" & strBase & "' "
strseq = Get_OracleStr(strSql)
strNewID = strBase & Right$("0" & strseq, 2)
strSql = "insert into PKGIDSEQ_37(VAL,SEQ,CREATE_DATE,DN) values('" & strBase & "', '" & strseq & "', sysdate, '" & tReel.T_DN_NUM & "')"
AddSql (strSql)
GetNewID = strNewID

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetQID
' Description:       获取QID
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/9-16:10:44
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function GetQID(tReel As T_REELINFO)
Dim strSql As String
Dim strQID As String
Dim strBID As String

strSql = "select BOXID from PACKING_DETAILED where dn_num = '" & tReel.T_DN_NUM & "' and outbox_num = '" & tReel.T_OUTBOX_NUM & "' and inbox_num = 1"
strBID = Get_OracleStr(strSql)
strSql = "select trglabelseq.QTSeq_NotMesQbox('" & strBID & "')  from dual"
strQID = Get_OracleStr(strSql)
GetQID = strQID

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetCustPN
' Description:       获取客户机种
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/9-9:19:40
'
' Parameters :       strDN (String)
'                    strJobID (String)
'--------------------------------------------------------------------------------
Private Function GetCustPN(strDN As String, strJobID As String) As String
Dim strSql As String

strSql = "select distinct marketingpn as mpn from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "' and batchnumber = '" & strJobID & "'"
GetCustPN = Get_OracleStr(strSql)

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetCustReelID
' Description:       获取客户卷盘ID
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/28-16:47:08
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function GetCustReelID(strDN As String, strJobID As String, strReelID As String)
If txtShipTo.text = "HUAWEI" Then
    GetCustReelID = GetHWReelPSN(strDN, strJobID, strReelID)
ElseIf txtShipTo.text = "SSE2" Or txtShipTo.text = "SSSHORT" Then
    GetCustReelID = GetSSReelID(strJobID, strReelID)

End If

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetSSReelID
' Description:       获取出三星卷盘标签ID
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/28-16:56:50
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function GetSSReelID(strJobID As String, strReelID As String) As String
If Right$(strJobID, 1) = "M" Then
    GetSSReelID = strJobID & Right$(strReelID, 1)
Else
    GetSSReelID = strJobID & GetLableXHTW(strJobID)

End If

End Function

Private Function GetLableXHTW(strKey As String) As String
Dim strSql   As String
Dim rs       As New ADODB.Recordset
Dim strXH    As String
Dim intCount As Integer
Dim strLot1  As String

If strKey = "" Then Exit Function
intCount = 0
strLot1 = Replace(strKey, "M", "")
strSql = "SELECT dbo.F_GetPrintXH('" & strLot1 & "') 序号"
If rs.State = adStateOpen Then rs.Close
rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF Then
    strXH = Trim$("" & rs!序号)
    If strXH <> "" Then '如果有得到序号，就更新数据
        strSql = "Update erpdata..tblSysIncrement Set Para='" & strXH & "',ICount=ICount+1 Where Kind='" & strLot1 & "'"
        INIadoCon2.Execute strSql, intCount
        If intCount <= 0 Then   '表示不存在此LOT信息，就插入一笔
            strSql = "Insert Into erpdata..tblSysIncrement(Kind,Para,ICount) Values('" & strKey & "','" & strXH & "',1)"
            INIadoCon2.Execute strSql

        End If

    End If

End If

rs.Close
GetLableXHTW = strXH  '赋值回去

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetHWReelPSN
' Description:       获取出华为卷盘标签PSN
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/9-9:26:58
'
' Parameters :       strDN (String)
'                    strJobID (String)
'--------------------------------------------------------------------------------
Private Function GetHWReelPSN(strDN As String, _
                              strJobID As String, _
                              strReelID As String) As String
Dim strPSN  As String
Dim strSql  As String
Dim strCPN  As String
Dim strRand As String
Dim strMon  As String
Dim lBase   As Long, lCnt As Long

strSql = "select customerpartnumber from CUSTOMERSHIPPINGUPTBL where batchnumber = '" & strJobID & "' and delivery = '" & strDN & "'"
strCPN = UCase(Get_OracleStr(strSql))
strMon = strCPN & Right(Year(Now), 2) & Hex(Month(Now))
lBase = 166576   ' 004LXR  - 004WB8 =  4* 10
strSql = "select nvl(count(*) + 1, 1) from REEL_REC_37 where mon = '" & strMon & "'"
lCnt = lBase + Get_OracleNo(strSql)
strRand = Right("000000" & Get10To33(lCnt), 6)
If Len(strCPN) = 8 Then
    strPSN = "P" & strCPN & "S" & Right(Year(Now), 2) & Hex(Month(Now)) & strRand
Else
    strPSN = "P" & strCPN & "/" & "S" & Right(Year(Now), 2) & Hex(Month(Now)) & strRand

End If

If Left(strRand, 4) = "004W" Then
    MsgBox "PSN流水段吃紧, 请及时联系IT", vbInformation, "提示"

End If

GetHWReelPSN = strPSN
strSql = "insert into REEL_REC_37(REELID,MON,CREATE_DATE) values('" & strReelID & "','" & strMon & "', sysdate)"
AddSql (strSql)

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       get10To33
' Description:       10进制转33进制
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/9-9:26:09
'
' Parameters :       lData (Long)
'--------------------------------------------------------------------------------
Private Function Get10To33(lData As Long) As String
Dim strOut As String

strOut = ""
Do
    If (lData Mod 33) = 0 Then
        strOut = "0" & strOut
    Else
        strOut = get33Char(lData Mod 33) & strOut

    End If

    Get10To33 = strOut
    lData = lData \ 33
Loop Until (lData = 0)

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetSeq
' Description:       获取序列号
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/9-9:59:02
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function GetSeq(strDN As String)
Dim strSql As String

strSql = "select nvl(max(seq)+1, 1) from packing_detailed where dn_num = '" & strDN & "'  "
GetSeq = Get_OracleStr(strSql)

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetDateCode
' Description:       获取DATECODE
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/9-10:29:25
'
' Parameters :       strJob (String)
'--------------------------------------------------------------------------------
Private Function GetDateCode(strJob As String, strTrayID As String) As String
Dim strWaferID    As String
Dim strDateCode   As String
Dim strJobNew     As String
Dim strSql        As String
Dim strContent    As String
Dim str1          As String
Dim strBartenName As String

'str1 = "37_FIRST_FINISH_YYWW_MON"
str1 = "37_DATE_CODE"
strBartenName = "37TRAY.btw"
strSql = "select top 1 Content from erpdata..tblME_PrintInfo aa ," & "erpdata..tblErpInStockDetailInfo bb where bb.KEY_VALUE = '" & strTrayID & "' +  '|' +  '" & strJob & "' " & "and bb.keyid = aa.EVENT_ID and bb.KEY_NAME = 'CONTAINER_NAME'  and bb.KEY_TYPE = 'T' " & "and aa.BartenderName = '" & strBartenName & "' " & "order by ID desc"
strContent = Get_SqlStr(strSql)
If strContent = "" Then
    strSql = "select top 1 Content from erpdata..tblME_PrintInfo_BACK190603 aa ," & "erpdata..tblErpInStockDetailInfo bb where bb.KEY_VALUE = '" & strTrayID & "' +  '|' +  '" & strJob & "' " & "and bb.keyid = aa.EVENT_ID and bb.KEY_NAME = 'CONTAINER_NAME'  and bb.KEY_TYPE = 'T' " & "and aa.BartenderName = '" & strBartenName & "' " & "order by ID desc"
    strContent = Get_SqlStr(strSql)
    If strContent = "" Then
        strJobNew = Replace$(strJob, "M", "")
        strSql = "select distinct case when create_date >= to_date(to_char(create_date, 'yyyy') || '-12-31', 'yyyy-mm-dd') - mod(to_char(create_date, 'YYYY'), 7) - 5  then to_char(create_date, 'yyww') " & "else to_char(create_date + mod(mod(to_char(create_date, 'YYYY'), 7) + 5, 7),'yyww') end as PODATECODE " & "from customeroitbl_test a ,mappingdatatest b ,weight37 c where a.test_mtrl_desc = '" & strJobNew & "' and b.filename = to_char(a.id) and b.lotid = a.source_batch_id " & "and c.waferid = replace(b.substrateid,'+','') "
        GetDateCode = Get_OracleStr(strSql)
        Exit Function

    End If

End If

strDateCode = Mid$(strContent, InStr(strContent, str1) + Len(str1) + 3, 4)
GetDateCode = strDateCode

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckPackingDetail
' Description:       检查包装明细数据
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/30-9:26:46
'
' Parameters :       tReel (T_REELINFO)
'--------------------------------------------------------------------------------
Private Function CheckPackingDetail(tReel As T_REELINFO) As Boolean
Dim strDC  As String
Dim strSql As String

CheckPackingDetail = False
If tReel.T_DN_NUM = "" Then
    MsgBox "DN不可为空", vbInformation, "警告"
    Exit Function

End If

If tReel.T_REELID = "" Then
    MsgBox "客户卷盘ID不可为空", vbInformation, "警告"
    Exit Function

End If

If tReel.T_BOXID = "" Then
    MsgBox "BID不可为空", vbInformation, "警告"
    Exit Function

End If

If tReel.T_CARTONID = "" Then
    MsgBox "CID不可为空", vbInformation, "警告"
    Exit Function

End If

If tReel.T_CARTON = "" Then
    MsgBox "QID不可为空", vbInformation, "警告"
    Exit Function

End If

If tReel.T_JOB_ID = "" Then
    MsgBox "JOBID不可为空", vbInformation, "警告"
    Exit Function

End If

If tReel.T_KID = "" Then
    MsgBox "KID不可为空", vbInformation, "警告"
    Exit Function

End If

If tReel.T_MPN = "" Then
    MsgBox "MPN不可为空", vbInformation, "警告"
    Exit Function

End If

If Not tReel.T_QTY > 0 Then
    MsgBox "卷盘数量错误", vbInformation, "警告"
    Exit Function

End If

If tReel.T_DATECODE = "" Then
    MsgBox "DC不可为空", vbInformation, "警告"
    Exit Function

End If

strSql = "SELECT  right(datename(yy,t1.ERPCREATEDATE),2) + datename(ww,t1.ERPCREATEDATE) from [erpdata].[dbo].[tblTSVworkorder] t1 " & "inner join erpdata..tblStockNumSub t2 on t2.大工单 = t1.ORDERNAME " & "where t2.箱号 = '" & tReel.T_TRAYID & "'"
strDC = Get_SqlStr(strSql)
If strDC = "" Then
    strSql = "SELECT  distinct right(datename(yy,t1.ERPCREATEDATE),2) + datename(ww,t1.ERPCREATEDATE) from [erpdata].[dbo].[tblTSVworkorder] t1" & " inner join erpdata..tblStocksqfhsub t3 on t3.大工单 = t1.ORDERNAME " & " where  t3.箱号 ='" & tReel.T_TRAYID & "' "
    strDC = Get_SqlStr(strSql)

End If

If strDC <> tReel.T_DATECODE Then
    MsgBox "DateCode有错误,请联系IT", vbInformation, "提示"
    Exit Function

End If

CheckPackingDetail = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       SavePackingDetail
' Description:       保存包装明细数据
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/9-16:19:42
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub SavePackingDetail(tReel As T_REELINFO)
Dim strSql As String

strSql = "insert into PACKING_DETAILED(TRAYID,INBOX_NUM,OUTBOX_NUM,DN_NUM,JOB_ID,QTY,CUSTOMER_DEVICE,CREATE_DATE,CREATE_BY,PRINT_FLAG,FLAG,KID,SEQ,DATECODE,REELID,CARTON,CARTONID,BOXID) " & " values('" & tReel.T_TRAYID & "', '" & tReel.T_INBOX_NUM & "','" & tReel.T_OUTBOX_NUM & "', '" & tReel.T_DN_NUM & "','" & tReel.T_JOB_ID & "','" & tReel.T_QTY & "','" & tReel.T_MPN & "', '" & tReel.T_CREATE_DATE & "', '" & tReel.T_CREATE_BY & "' ,'" & tReel.T_PRINT_FLAG & "','" & tReel.T_FLAG & "','" & tReel.T_KID & "','" & tReel.T_SEQ & "', '" & tReel.T_DATECODE & "','" & tReel.T_REELID & "','" & tReel.T_CARTON & "','" & tReel.T_CARTONID & "','" & tReel.T_BOXID & "')"
AddSql (strSql)

End Sub

'-------------------------------------------------------------
'<<<<<<<<<<<<<<<<<<标签打印>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'-------------------------------------------------------------
Private Sub PrintHandler()
Dim lCurOutboxNum As Long
Dim lMaxOutboxNum As Long
Dim strSql        As String
Dim strDN         As String

strDN = Trim(txtDN.text)
lCurOutboxNum = CLng(Trim(txtCurOP.text))
strSql = "select max(OUTBOX_NUM) from PACKING_DETAILED where DN_NUM = '" & strDN & "'"
lMaxOutboxNum = Get_OracleNo(strSql)
If strDN = "" Then
    MsgBox "请输入DN", vbInformation, "提示"
    Exit Sub

End If

If lCurOutboxNum = 0 Then
    MsgBox "请输入第几箱", vbInformation, "提示"
    Exit Sub

End If

If lCurOutboxNum = 1 Then
    AddSql ("delete from TBL37QRVALUE where substr(key_name,1,8) = '" & strDN & "'")

End If

If lCurOutboxNum > lMaxOutboxNum Then
    MsgBox "标签已经全部打印完成", vbInformation, "提示"
    Exit Sub

End If

Call PrintLblByOutBoxNum(strDN, lCurOutboxNum)
MsgBox "第" & lCurOutboxNum & "箱标签已经全部打印完成", vbInformation, "提示"
lCurOutboxNum = lCurOutboxNum + 1
txtCurOP.text = lCurOutboxNum

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PrintLblByOutBoxNum
' Description:       打印标签接口:外箱序号
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-8:54:16
'
' Parameters :       strDN (String) lOutboxNum (Long)
'--------------------------------------------------------------------------------
Private Function PrintLblByOutBoxNum(strDN As String, lOutboxNum As Long)
Dim rsInboxNum As New ADODB.Recordset
Dim strSql     As String
Dim lInboxNum  As Long

strSql = "select distinct INBOX_NUM from PACKING_DETAILED where DN_NUM = '" & strDN & "' and OUTBOX_NUM = '" & lOutboxNum & "' order by INBOX_NUM"
Set rsInboxNum = Get_OracleRs(strSql)
If Not rsInboxNum.EOF Then

    Do While Not rsInboxNum.EOF
        lInboxNum = rsInboxNum!INBOX_NUM
        Call PrintLbl_IP(strDN, lOutboxNum, lInboxNum)
        rsInboxNum.MoveNext
    Loop

End If

Call PrintLbl_OP(strDN, lOutboxNum)
Set rsInboxNum = Nothing

End Function

'-------------------------------------------------------------
'<<<<<<<<<<<<<<<<<<内盒/卷盘打印>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'-------------------------------------------------------------
'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PrintLblByInBoxNum
' Description:       打印标签接口:外箱+内盒
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-9:13:50
'
' Parameters :       strDN (String)
'                    lOutboxNum (Long)
'                    lInboxNum (Long)
'--------------------------------------------------------------------------------
Private Function PrintLbl_IP(strDN As String, lOutboxNum As Long, lInboxNum As Long)
'1.打印内盒卷盘标签
Call Print37BoxLbl_OLD(strDN, lOutboxNum, lInboxNum, "") '37内盒B标签

Select Case Trim(txtShipTo.text)

    Case "HUAWEI"   '出华为
        Call PrintHWBoxLbl_OLD(strDN, lOutboxNum, lInboxNum, "") '华为内盒标签
        Call PrintHWReelLbl_OLD(strDN, lOutboxNum, lInboxNum, "") '华为卷盘标签

    Case "SSE2" '出三星E2
        Call PrintSSE2BoxLbl_OLD(strDN, lOutboxNum, lInboxNum, "") '三星E2内盒标签
        Call PrintSSE2ReelLbl_OLD(strDN, lOutboxNum, lInboxNum, "") '三星E2卷盘标签

    Case "SSSHORT"  '出三星SHORT
        Call PrintSSSHORTBoxLbl_OLD(strDN, lOutboxNum, lInboxNum, "") '三星SHORT内盒标签
        Call PrintSSSHORTReelLbl_OLD(strDN, lOutboxNum, lInboxNum, "") '三星SHORT卷盘标签

End Select

AddSql ("update PACKING_DETAILED set print_flag = 1 where dn_num = '" & strDN & "' and outbox_num = '" & lOutboxNum & "' and inbox_num = '" & lInboxNum & "'")

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       Print37BoxLbl_OLD
' Description:       打印37内盒B标签
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-9:23:02
'
' Parameters :       strDN (String)
'                    lOutboxNum (Integer)
'                    lInboxNum (Integer)
'--------------------------------------------------------------------------------
Private Sub Print37BoxLbl_OLD(strDN As String, _
                              lOutboxNum As Long, _
                              lInboxNum As Long, _
                              strBID As String)
Dim strSql      As String
Dim strTxt      As String
Dim strFlagTxt  As String
Dim strFileName As String
Dim rsJobID     As New ADODB.Recordset
Dim tSTBox      As STBox
Dim strQrCode   As String

If strBID <> "" Then
    strSql = "select JOB_ID,CUSTOMER_DEVICE,BOXID,DATECODE,SUM(QTY) as QTY from PACKING_DETAILED where DN_NUM = '" & strDN & "' and BOXID='" & strBID & "' group by JOB_ID,CUSTOMER_DEVICE,BOXID,DATECODE"
Else
    '标记
    strTxt = "BOX_" & lOutboxNum & "_" & lInboxNum
    strFileName = strDN & "-" & "FLAG_BOX_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
    Call CreateTxt(strFileName, strTxt, strFlagPath)
    Call Sleep(gSleepMicSec)
    strTxt = ""
    '正式
    strSql = "select JOB_ID,CUSTOMER_DEVICE,BOXID,DATECODE,SUM(QTY) as QTY from PACKING_DETAILED where DN_NUM = '" & strDN & "' and outbox_num = '" & lOutboxNum & "' and inbox_num =  '" & lInboxNum & "' group by JOB_ID,CUSTOMER_DEVICE,BOXID,DATECODE"

End If

Set rsJobID = Get_OracleRs(strSql)
If Not rsJobID.BOF Then
    rsJobID.MoveFirst

    Do While Not rsJobID.EOF
        tSTBox.JOB = Trim("" & rsJobID!JOB_ID)
        tSTBox.DEV = Trim("" & rsJobID!Customer_Device)
        tSTBox.lot = Trim("" & rsJobID!BOXID)
        tSTBox.DATECODE = Trim$("" & rsJobID!DATECODE)
        tSTBox.QTY = rsJobID!QTY
        tSTBox.FactoryFlow = Get_OracleStr("select distinct material from customershippinguptbl where marketingpn = '" & tSTBox.DEV & "' and delivery = '" & strDN & "'")
        strTxt = strTxt & tSTBox.DEV & "," & tSTBox.JOB & ",1T" & tSTBox.JOB & "," & tSTBox.DEV & "," & "1P" & tSTBox.DEV & "," & tSTBox.DATECODE & "," & tSTBox.DATECODE & "," & Mid(tSTBox.lot, 2) & "," & tSTBox.lot & "," & tSTBox.QTY & ",Q" & tSTBox.QTY & "," & tSTBox.DATECODE & "," & tSTBox.DATECODE & GetDevMark(tSTBox.DEV)
        strTxt = strTxt & "," & tSTBox.FactoryFlow & "," & "6P" & tSTBox.FactoryFlow & "," & "10D" & tSTBox.DATECODE & ","
        strQrCode = "[)>" & Chr(30) & "06" & Chr(29) & "1T" & tSTBox.JOB & Chr(29) & "1P" & tSTBox.DEV & Chr(29) & tSTBox.lot & Chr(29) & "Q" & tSTBox.QTY & Chr(29) & "6P" & tSTBox.FactoryFlow & Chr(29) & "10D" & tSTBox.DATECODE & Chr(30) & Chr(4)
        strTxt = strTxt & strQrCode & vbCrLf
        strQrCode = Replace(Replace(Replace(strQrCode, Chr(30), ""), Chr(29), ""), Chr(4), "")
        AddSql ("delete from TBL37QRVALUE where KEY_NAME = '" & strDN & "' || '_' ||  '" & tSTBox.lot & "'")
        AddSql ("insert into TBL37QRVALUE(KEY_NAME,KEY_VALUE,CREATE_DATE,CREATE_BY) values('" & strDN & "' || '_' ||  '" & tSTBox.lot & "','" & strQrCode & "',sysdate,'" & gUserName & "')  ")
        rsJobID.MoveNext
    Loop

End If

Set rsJobID = Nothing
strFileName = strDN & "-" & "BID" & "_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(strFileName, strTxt, str37BCIDPath)
Call Sleep(gSleepMicSec)

End Sub

'
'--------------------------------------------------------------------------------
'Project:            正式工程1
'Procedure:          PrintSSE2BoxLbl_OLD
'Description:        SSE2
' Created By:        Project Administrator
'Machine:            DESKTOP -MSUG5JD
' Date-Time  :       2019/8/28-17:51:21
'
' Parameters :       strDN (String)
'                    lOutboxNum (Integer)
'                    lInboxNum (Integer)
'--------------------------------------------------------------------------------
Private Sub PrintSSE2BoxLbl_OLD(strDN As String, _
                                lOutboxNum As Long, _
                                lInboxNum As Long, _
                                strBID As String)
Dim strSql          As String
Dim tCusBox         As CusBox
Dim strContent      As String
Dim strFileName     As String
Dim rs              As New ADODB.Recordset
Dim strFabSite      As String
Dim strAssemblySite As String
Dim strTestSite     As String

strFileName = strDN & "-" & "CUSBoxLbl" & "_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
If strBID <> "" Then
    strSql = "select sum(a.qty) as qty, a.CUSTOMER_DEVICE, b.customerpartnumber,b.FAB_SITE, b.ASSEMBLY_SITE,b.TEST_SITE from PACKING_DETAILED a , CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & strDN & "' and b.delivery = '" & strDN & "' and a.job_id = b.batchnumber and a.boxid = '" & strBID & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.FAB_SITE, b.ASSEMBLY_SITE,b.TEST_SITE"
Else
    strSql = "select sum(a.qty) as qty, a.CUSTOMER_DEVICE, b.customerpartnumber,b.FAB_SITE, b.ASSEMBLY_SITE,b.TEST_SITE from PACKING_DETAILED a , CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & strDN & "' and b.delivery = '" & strDN & "' and a.job_id = b.batchnumber and a.outbox_num = '" & lOutboxNum & "' and a.inbox_num = '" & lInboxNum & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.FAB_SITE, b.ASSEMBLY_SITE,b.TEST_SITE"

End If

Set rs = Get_OracleRs(strSql)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tCusBox.QTY = Trim("" & rs!QTY)
        tCusBox.DEV = Trim("" & rs!Customer_Device)
        tCusBox.PN = Trim("" & rs!CustomerPartnumber)
        strFabSite = Trim$("" & rs!FAB_SITE)
        strAssemblySite = Trim$("" & rs!ASSEMBLY_SITE)
        strTestSite = Trim("" & rs!TEST_SITE)
        If strFabSite <> "" Then
            strFabSite = "Fab:" & strFabSite

        End If

        If strAssemblySite <> "" Then
            strAssemblySite = "Assembly:" & strAssemblySite

        End If

        If strTestSite <> "" Then
            strTestSite = "Test:" & strTestSite

        End If

        strContent = strContent + tCusBox.PN + "DPTKE2" + Right$("000000" + tCusBox.QTY, 6) + ","
        strContent = strContent + tCusBox.PN + "," + "TVS DIODES" + "," + "PO TYPE,:E2" + "," + tCusBox.QTY + "," + tCusBox.DEV + "," + "DPTK" + "," + strFabSite + "," + strAssemblySite + "," + strTestSite
        rs.MoveNext
    Loop

End If

If tCusBox.DEV = "RCLAMP2581ZCTFT" Then
    Call CreateTxt(strFileName, strContent, strSSBoxPath2)
Else
    Call CreateTxt(strFileName, strContent, strSSBoxPath)

End If

Call Sleep(gSleepMicSec)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PrintSSE2ReelLbl_OLD
' Description:       三星E2卷盘标签
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/28-17:56:12
'
' Parameters :       strDN (String)
'                    lOutboxNum (Integer)
'                    lInboxNum (Integer)
'--------------------------------------------------------------------------------
Private Sub PrintSSE2ReelLbl_OLD(strDN As String, _
                                 lOutboxNum As Long, _
                                 lInboxNum As Long, _
                                 strRID As String)
Dim strSql          As String
Dim tCusReel        As CusReel
Dim strContent      As String
Dim strFileName     As String
Dim rs              As New ADODB.Recordset
Dim Rs2             As New ADODB.Recordset
Dim strFabSite      As String
Dim strAssemblySite As String
Dim strTestSite     As String
Dim strTxt          As String

'标记
If strRID = "" Then
    strTxt = "REEL_" & lOutboxNum & "_" & lInboxNum
    strFileName = strDN & "-" & "FLAG_REEL_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
    Call CreateTxt(strFileName, strTxt, strFlagPath)
    Call Sleep(gSleepMicSec)

End If

strFileName = strDN & "-" & "CUSREELLbl" & "_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
If strRID <> "" Then
    strSql = "select distinct trayid, reelid, qty, CUSTOMER_DEVICE, cpn,seq, FAB_SITE, ASSEMBLY_SITE,TEST_SITE from lpstbl where dn_num = '" & strDN & "' and trayid = '" & strRID & "'  "
Else
    strSql = "select distinct trayid, reelid, qty, CUSTOMER_DEVICE, cpn,seq, FAB_SITE, ASSEMBLY_SITE,TEST_SITE from lpstbl where dn_num = '" & strDN & "' and outbox_num = '" & lOutboxNum & "' and inbox_num = '" & lInboxNum & "'  order by seq "

End If

Set Rs2 = Get_OracleRs(strSql)
If Not Rs2.BOF Then
    Rs2.MoveFirst

    Do While Not Rs2.EOF
        tCusReel.TRAYID = Trim$("" & Rs2!TRAYID)
        tCusReel.lot = Trim$("" & Rs2!REELID)
        tCusReel.QTY = Trim("" & Rs2!QTY)
        tCusReel.DEV = Trim("" & Rs2!Customer_Device)
        tCusReel.PN = Trim("" & Rs2!CPN)
        strFabSite = Trim$("" & Rs2!FAB_SITE)
        strAssemblySite = Trim$("" & Rs2!ASSEMBLY_SITE)
        strTestSite = Trim("" & Rs2!TEST_SITE)
        If strFabSite <> "" Then
            strFabSite = "Fab:" & strFabSite

        End If

        If strAssemblySite <> "" Then
            strAssemblySite = "Assembly:" & strAssemblySite

        End If

        If strTestSite <> "" Then
            strTestSite = "Test:" & strTestSite

        End If

        strContent = strContent + tCusReel.PN + "DPTKE2" + tCusReel.lot + Right$("000000" + tCusReel.QTY, 6) + ","
        strContent = strContent + tCusReel.PN + "," + "TVS DIODES" + "," + "PO TYPE,:E2" + "," + tCusReel.lot + "," + tCusReel.QTY + ","
        strContent = strContent + tCusReel.DEV + "," + "DPTK" + "," + strFabSite + "," + strAssemblySite + "," + strTestSite + vbCrLf
        Rs2.MoveNext
    Loop

End If

If tCusReel.DEV = "RCLAMP2581ZCTFT" Then
    Call CreateTxt(strFileName, strContent, strSSReelPath2)
Else
    Call CreateTxt(strFileName, strContent, strSSReelPath)

End If

Call Sleep(gSleepMicSec)

End Sub

'--------------------------------------------------------------------------------
'Project:            正式工程1
'Procedure:          PrintSSSHORTBoxLbl_OLD
'Description:        SSSHORT
' Created By:        Project Administrator
'Machine:            DESKTOP -MSUG5JD
' Date-Time  :       2019/8/28-17:50:54
'
' Parameters :       strDN (String)
'                    lOutboxNum (Integer)
'                    lInboxNum (Integer)
'--------------------------------------------------------------------------------
Private Sub PrintSSSHORTBoxLbl_OLD(strDN As String, _
                                   lOutboxNum As Long, _
                                   lInboxNum As Long, _
                                   strBID As String)
Dim strSql      As String
Dim tCusBox     As CusBox
Dim strContent  As String
Dim strFileName As String
Dim rs          As New ADODB.Recordset

strFileName = strDN & "-" & "CUSBoxLbl2" & "_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
If strBID <> "" Then
    strSql = "select sum(a.qty) as qty, a.CUSTOMER_DEVICE, b.customerpartnumber from PACKING_DETAILED a , CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & strDN & "' and b.delivery = '" & strDN & "' and a.job_id = b.batchnumber and a.boxid = '" & strBID & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.FAB_SITE, b.ASSEMBLY_SITE,b.TEST_SITE"
Else
    strSql = "select sum(a.qty) as qty, a.CUSTOMER_DEVICE, b.customerpartnumber from PACKING_DETAILED a , CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & strDN & "' and b.delivery = '" & strDN & "' and a.job_id = b.batchnumber and a.outbox_num = '" & lOutboxNum & "' and a.inbox_num = '" & lInboxNum & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.FAB_SITE, b.ASSEMBLY_SITE,b.TEST_SITE"

End If

Set rs = Get_OracleRs(strSql)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tCusBox.QTY = Trim("" & rs!QTY)
        tCusBox.DEV = Trim("" & rs!Customer_Device)
        tCusBox.PN = Trim("" & rs!CustomerPartnumber)
        strContent = strContent + tCusBox.PN + "DPTK" + Right$("000000" + tCusBox.QTY, 6) + ","
        strContent = strContent + tCusBox.PN + "," + "TVS DIODES" + "," + tCusBox.QTY + "," + tCusBox.DEV + "," + "DPTK" + ","
        rs.MoveNext
    Loop

End If

If tCusBox.DEV = "RCLAMP2581ZCTFT" Then
    MsgBox "请联系IT确认标签模板", vbCritical, "警告"
    Exit Sub

End If

Call CreateTxt(strFileName, strContent, strSSBoxPath_Short)
Call Sleep(gSleepMicSec)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PrintSSSHORTReelLbl_OLD
' Description:       三星SHORT卷盘标签
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/28-17:55:50
'
' Parameters :       strDN (String)
'                    lOutboxNum (Integer)
'                    lInboxNum (Integer)
'--------------------------------------------------------------------------------
Private Sub PrintSSSHORTReelLbl_OLD(strDN As String, _
                                    lOutboxNum As Long, _
                                    lInboxNum As Long, _
                                    strRID As String)
Dim strSql      As String
Dim tCusReel    As CusReel
Dim strContent  As String
Dim strFileName As String
Dim rs          As New ADODB.Recordset
Dim strTxt      As String

'标记
If strRID = "" Then
    strTxt = "REEL_" & lOutboxNum & "_" & lInboxNum
    strFileName = strDN & "-" & "FLAG_REEL_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
    Call CreateTxt(strFileName, strTxt, strFlagPath)
    Call Sleep(gSleepMicSec)

End If

strFileName = strDN & "-" & "CUSREELLbl" & "_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
If strRID <> "" Then
    strSql = "select distinct trayid, reelid, qty, CUSTOMER_DEVICE, cpn,seq from lpstbl where dn_num = '" & strDN & "' and trayid = '" & strRID & "' "
Else
    strSql = "select distinct trayid, reelid, qty, CUSTOMER_DEVICE, cpn,seq from lpstbl where dn_num = '" & strDN & "' and outbox_num = '" & lOutboxNum & "' and inbox_num = '" & lInboxNum & "'  order by seq "

End If

Set rs = Get_OracleRs(strSql)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tCusReel.TRAYID = Trim$("" & rs!TRAYID)
        tCusReel.lot = Trim$("" & rs!REELID)
        tCusReel.QTY = Trim("" & rs!QTY)
        tCusReel.DEV = Trim("" & rs!Customer_Device)
        tCusReel.PN = Trim("" & rs!CPN)
        strContent = strContent + tCusReel.PN + "DPTK" + tCusReel.lot + Right$("000000" + tCusReel.QTY, 6) + ","
        strContent = strContent + tCusReel.PN + "," + "TVS DIODES" + "," + tCusReel.lot + "," + tCusReel.QTY + ","
        strContent = strContent + tCusReel.DEV + "," + "DPTK" + "," + vbCrLf
        rs.MoveNext
    Loop

End If

If tCusReel.DEV = "RCLAMP2581ZCTFT" Then
    MsgBox "请联系IT确认机种是否有问题", vbCritical, "警告"
    Exit Sub

End If

Call CreateTxt(strFileName, strContent, strSSReelPath_Short)
Call Sleep(gSleepMicSec)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PrintHWBoxLbl_OLD
' Description:       打印华为内盒标签
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-9:25:02
'
' Parameters :       strDN (String)
'                    lOutboxNum (Long)
'                    lInboxNum (Long)
'--------------------------------------------------------------------------------
Private Sub PrintHWBoxLbl_OLD(strDN As String, _
                              lOutboxNum As Long, _
                              lInboxNum As Long, _
                              strBID As String)
Dim strTxt      As String
Dim strBarcode  As String
Dim strQrCode   As String
Dim strFileName As String
Dim strSql      As String
Dim rsJobID     As New ADODB.Recordset
Dim tHWBox      As HWBox

'正式
If strBID <> "" Then
    strSql = "select job_id,mpn,cpn,datecode,sum(QTY) qty from LPSTBL where dn_num = '" & strDN & "' and boxid = '" & strBID & "' group by job_id,mpn,cpn,datecode"
Else
    strSql = "select job_id,mpn,cpn,datecode,sum(QTY) qty from LPSTBL where dn_num = '" & strDN & "' and outbox_num = '" & lOutboxNum & "' and inbox_num =  '" & lInboxNum & "' group by job_id,mpn,cpn,datecode"

End If

Set rsJobID = Get_OracleRs(strSql)
If Not rsJobID.BOF Then
    rsJobID.MoveFirst

    Do While Not rsJobID.EOF
        tHWBox.CPN = Trim$("" & rsJobID!CPN)
        tHWBox.MPN = Trim$("" & rsJobID!MPN)
        tHWBox.lot = Trim$("" & rsJobID!JOB_ID)
        tHWBox.PODATE = Trim$("" & rsJobID!DATECODE)
        tHWBox.QTY = rsJobID!QTY
        strBarcode = tHWBox.CPN & "," & "" & "," & "" & "," & tHWBox.MPN & "," & tHWBox.PODATE & "," & tHWBox.lot & "," & tHWBox.QTY & ","
        strQrCode = "[)>" & Chr(30) & "06" & Chr(29) & "F01001P" & Chr(29) & "18VLEHWT" & Chr(29) & "F02010I" & Chr(29) & "1P" & tHWBox.CPN & Chr(29) & "1V601024" & Chr(29) & "10D" & tHWBox.PODATE & Chr(29) & "1T" & tHWBox.lot & Chr(29) & "Q" & tHWBox.QTY & Chr(30) & Chr(4)
        strTxt = strTxt & strBarcode & strQrCode & vbCrLf
        rsJobID.MoveNext
    Loop

End If

strFileName = strDN & "-" & "HWBoxLbl" & "_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(strFileName, strTxt, strHWBoxPath)
Call Sleep(gSleepMicSec)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PrintHWReelLbl_OLD
' Description:       打印华为卷盘标签
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-9:26:23
'
' Parameters :       strDN (String)
'                    lOutboxNum (Long)
'                    lInboxNum (Long)
'--------------------------------------------------------------------------------
Private Sub PrintHWReelLbl_OLD(strDN As String, _
                               lOutboxNum As Long, _
                               lInboxNum As Long, _
                               strTrayID As String)
Dim strTxt      As String
Dim strBarcode  As String
Dim strQrCode   As String
Dim strFileName As String
Dim strSql      As String
Dim rsReel      As New ADODB.Recordset
Dim tHWBox      As HWBox

'正式
If strTrayID <> "" Then
    strSql = "select job_id,mpn,cpn, QTY,datecode,reelid,seq from LPSTBL where dn_num = '" & strDN & "' and trayid= '" & strTrayID & "' "
Else
    '标记
    strTxt = "REEL_" & lOutboxNum & "_" & lInboxNum
    strFileName = strDN & "-" & "FLAG_REEL_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
    Call CreateTxt(strFileName, strTxt, strFlagPath)
    Call Sleep(gSleepMicSec)
    strTxt = ""
    strSql = "select job_id,mpn,cpn, QTY,datecode,reelid,seq from LPSTBL where dn_num = '" & strDN & "' and outbox_num = '" & lOutboxNum & "' and inbox_num =  '" & lInboxNum & "' order by seq"

End If

Set rsReel = Get_OracleRs(strSql)
If Not rsReel.BOF Then
    rsReel.MoveFirst

    Do While Not rsReel.EOF
        tHWBox.CPN = Trim$("" & rsReel!CPN)
        tHWBox.MPN = Trim$("" & rsReel!MPN)
        tHWBox.lot = Trim("" & rsReel!JOB_ID)
        tHWBox.PODATE = Trim$("" & rsReel!DATECODE)
        tHWBox.PSN = Trim$("" & rsReel!REELID)
        tHWBox.QTY = rsReel!QTY
        strBarcode = tHWBox.CPN & "," & "" & "," & "" & "," & tHWBox.MPN & "," & tHWBox.PODATE & "," & tHWBox.lot & "," & tHWBox.QTY & "," & tHWBox.PSN & ","
        strQrCode = "[)>" & Chr(30) & "06" & Chr(29) & "F01001P" & Chr(29) & "52S" & tHWBox.PSN & Chr(29) & "18VLEHWT" & Chr(29) & "F02010I" & Chr(29) & "1P" & tHWBox.CPN & Chr(29) & "1V601024" & Chr(29) & "10D" & tHWBox.PODATE & Chr(29) & "1T" & tHWBox.lot & Chr(29) & "Q" & tHWBox.QTY & Chr(30) & Chr(4)
        strTxt = strTxt & strBarcode & strQrCode & vbCrLf
        rsReel.MoveNext
    Loop

End If

strFileName = strDN & "-" & "HWReelLbl" & "_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(strFileName, strTxt, strHWReelPath)
Call Sleep(gSleepMicSec)

End Sub

'-------------------------------------------------------------
'<<<<<<<<<<<<<<<<<<外箱打印>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'-------------------------------------------------------------
'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PrintLbl_OP
' Description:       打印外箱标签
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-10:40:32
'
' Parameters :       strDN (String)
'                    lOutboxNum (Long)
'--------------------------------------------------------------------------------
Private Function PrintLbl_OP(strDN As String, lOutboxNum As Long)
'1.打印外箱标签
Call Print37CartonLbl_OLD(strDN, lOutboxNum, "")   '37外箱C标签
Call PrintHTCartonLbl_OLD(strDN, lOutboxNum)    '华天Q箱号

Select Case Trim(txtShipTo.text)

    Case "HUAWEI"   '出华为
        Call Print37CartonStanderLbl_OLD(strDN, lOutboxNum, "") '37外箱标准大标签

    Case "SSE2" '出三星E2
        Call PrintSSE2CartonLbl_OLD(strDN, lOutboxNum, "") '三星E2外箱大标签

    Case "SSSHORT" '出三星SHORT
        Call PrintSSSHORTCartonLbl_OLD(strDN, lOutboxNum, "") '三星SHORT外箱大标签

    Case "ST"  '出37标准版
        Call Print37CartonStanderLbl_OLD(strDN, lOutboxNum, "") '37外箱标准大标签

End Select

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       Print37CartonLbl_OLD
' Description:       打印37外箱C标签
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-10:49:49
'
' Parameters :       strDN (String)
'                    lOutboxNum (Long)
'--------------------------------------------------------------------------------
Private Sub Print37CartonLbl_OLD(strDN As String, lOutboxNum As Long, strCID As String)
Dim strSql        As String
Dim tSTCarton     As STCarton
Dim strTxt        As String
Dim strFileName   As String
Dim sDatecode     As String
Dim sTestDateCode As String
Dim rsJobID       As New ADODB.Recordset
Dim sAdd          As String
Dim strQrCode     As String

If strCID <> "" Then
    strSql = "select JOB_ID,CUSTOMER_DEVICE,CARTONID,DATECODE,SUM(QTY) AS QTY from PACKING_DETAILED where dn_num = '" & strDN & "' and cartonid = '" & strCID & "' group by JOB_ID,CUSTOMER_DEVICE,CARTONID,DATECODE"
Else
    '标记
    strTxt = "CARTON_" & lOutboxNum
    strFileName = strDN & "-" & "FLAG_CARTON_" & lOutboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
    Call CreateTxt(strFileName, strTxt, strFlagPath)
    Call Sleep(gSleepMicSec)
    strTxt = ""
    '正式
    strSql = "select JOB_ID,CUSTOMER_DEVICE,CARTONID,DATECODE,SUM(QTY) AS QTY from PACKING_DETAILED where dn_num = '" & strDN & "' and outbox_num = '" & lOutboxNum & "' group by JOB_ID,CUSTOMER_DEVICE,CARTONID,DATECODE"

End If

Set rsJobID = Get_OracleRs(strSql)
If Not rsJobID.BOF Then
    rsJobID.MoveFirst

    Do While Not rsJobID.EOF
        tSTCarton.JOB = Trim("" & rsJobID!JOB_ID)
        tSTCarton.DEV = Trim$("" & rsJobID!Customer_Device)
        tSTCarton.lot = Trim("" & rsJobID!CARTONID)
        tSTCarton.DATECODE = Trim("" & rsJobID!DATECODE)
        tSTCarton.QTY = rsJobID!QTY
        tSTCarton.FactoryFlow = Get_OracleStr("select distinct material from customershippinguptbl where marketingpn = '" & tSTCarton.DEV & "' and delivery = '" & strDN & "'")
        strTxt = strTxt & tSTCarton.DEV & "," & tSTCarton.JOB & ",1T" & tSTCarton.JOB & "," & tSTCarton.DEV & "," & "1P" & tSTCarton.DEV & "," & tSTCarton.DATECODE & "," & tSTCarton.DATECODE & "," & Mid(tSTCarton.lot, 2) & "," & tSTCarton.lot & "," & tSTCarton.QTY & ",Q" & tSTCarton.QTY & "," & tSTCarton.testdateCode & "," & tSTCarton.testdateCode & GetDevMark(tSTCarton.DEV)
        strTxt = strTxt & "," & tSTCarton.FactoryFlow & "," & "6P" & tSTCarton.FactoryFlow & "," & "10D" & tSTCarton.DATECODE & ","
        strQrCode = "[)>" & Chr(30) & "06" & Chr(29) & "1T" & tSTCarton.JOB & Chr(29) & "1P" & tSTCarton.DEV & Chr(29) & tSTCarton.lot & Chr(29) & "Q" & tSTCarton.QTY & Chr(29) & "6P" & tSTCarton.FactoryFlow & Chr(29) & "10D" & tSTCarton.DATECODE & Chr(30) & Chr(4)
        strTxt = strTxt & strQrCode & vbCrLf
        strQrCode = Replace(Replace(Replace(strQrCode, Chr(30), ""), Chr(29), ""), Chr(4), "")
        AddSql ("delete from TBL37QRVALUE where KEY_NAME = '" & strDN & "' || '_' ||  '" & tSTCarton.lot & "'")
        AddSql ("insert into TBL37QRVALUE(KEY_NAME,KEY_VALUE,CREATE_DATE,CREATE_BY) values('" & strDN & "' || '_' ||  '" & tSTCarton.lot & "','" & strQrCode & "',sysdate,'" & gUserName & "')  ")
        rsJobID.MoveNext
    Loop

End If

strFileName = strDN & "-" & "CID" & "_" & lOutboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(strFileName, strTxt, str37BCIDPath)
Call Sleep(gSleepMicSec)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PrintHTCartonLbl_OLD
' Description:       打印华天Q标签
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-10:58:33
'
' Parameters :       strDN (String)
'                    lOutboxNum (Long)
'--------------------------------------------------------------------------------
Private Sub PrintHTCartonLbl_OLD(strDN As String, lOutboxNum As Long)
Dim strSql      As String
Dim strFileName As String
Dim strTxt      As String

strSql = "select distinct carton from PACKING_DETAILED where dn_num = '" & strDN & "' and outbox_num = '" & lOutboxNum & "'"
strTxt = Get_OracleStr(strSql)
strFileName = strDN & "-" & "QID_" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(strFileName, strTxt, strHTQCartonPath)
Call Sleep(gSleepMicSec)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PrintSSE2CartonLbl_OLD
' Description:       打印三星E2外箱大标签
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/29-9:47:44
'
' Parameters :       strDN (String)
'                    lOutboxNum (Long)
'--------------------------------------------------------------------------------
Private Sub PrintSSE2CartonLbl_OLD(strDN As String, lOutboxNum As Long, strQID As String)
Dim strSql      As String
Dim tCusCARTON  As CUSCARTON
Dim strFileName As String
Dim strContent  As String
Dim rs          As New ADODB.Recordset
Dim KID         As String
Dim sMaxOP      As String

sMaxOP = Get_OracleStr("select max(outbox_num) from PACKING_DETAILED where dn_num = '" & strDN & "'")
strFileName = strDN & "-" & "CUSCARTONLBL" + Format(Now(), "YYYYMMDDHHmmSS")
If strQID <> "" Then
    lOutboxNum = Get_OracleStr("select distinct outbox_num from packing_detailed where dn_num = '" & strDN & "' and carton = '" & strQID & "'")

End If

strSql = "select a.kid,a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & strDN & "'" & "and b.delivery = '" & strDN & "' and a.job_id = b.batchnumber and a.outbox_num = '" & lOutboxNum & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno, a.kid"
Set rs = Get_OracleRs(strSql)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tCusCARTON.dn = strDN
        tCusCARTON.PO = Left("" & rs!PO, 10)
        tCusCARTON.CPN = Trim$("" & rs!CustomerPartnumber)
        tCusCARTON.MPN = Trim$("" & rs!Customer_Device)
        tCusCARTON.QTY = "" & rs!QTY
        KID = rs!KID
        strContent = strContent & tCusCARTON.dn & ",I" & tCusCARTON.dn & "," & tCusCARTON.PO & ",K" & tCusCARTON.PO & ",E2," & tCusCARTON.CPN & ",P" & tCusCARTON.CPN & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & ","
        strContent = strContent & tCusCARTON.QTY & ",Q" & tCusCARTON.QTY & ","
        strContent = strContent & Get_OracleStr("select distinct trim(a.freightforwarder)|| ',CHINA,' || substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3)||','||trim(a.city) || ' ' || trim(a.state) || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ','|| 'Attn:;Tel:' || trim(a.phone) || ','  from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & strDN & "'")
        strContent = strContent & "N/A,PN/A,N/A,9DN/A," & lOutboxNum & "," & KID & "," & sMaxOP
        rs.MoveNext
    Loop

End If

Call CreateTxt(strFileName, strContent, strSSCartonPath)
Call Sleep(gSleepMicSec)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PrintSSSHORTCartonLbl_OLD
' Description:       打印三星SHORT外箱大标签
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/29-9:48:45
'
' Parameters :       strDN (String)
'                    lOutboxNum (Long)
'--------------------------------------------------------------------------------
Private Sub PrintSSSHORTCartonLbl_OLD(strDN As String, _
                                      lOutboxNum As Long, _
                                      strQID As String)
Dim strSql      As String
Dim tCusCARTON  As CUSCARTON
Dim strFileName As String
Dim strContent  As String
Dim rs          As New ADODB.Recordset
Dim KID         As String
Dim sMaxOP      As String

sMaxOP = Get_OracleStr("select max(outbox_num) from PACKING_DETAILED where dn_num = '" & strDN & "'")
strFileName = strDN & "-" & "CUSCARTONLBL" + Format(Now(), "YYYYMMDDHHmmSS")
If strQID <> "" Then
    lOutboxNum = Get_OracleStr("select distinct outbox_num from packing_detailed where dn_num = '" & strDN & "' and carton = '" & strQID & "'")

End If

strSql = "select a.kid,a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & strDN & "'" & "and b.delivery = '" & strDN & "' and a.job_id = b.batchnumber and a.outbox_num = '" & lOutboxNum & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno, a.kid"
Set rs = Get_OracleRs(strSql)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tCusCARTON.dn = strDN
        tCusCARTON.PO = Left("" & rs!PO, 10)
        tCusCARTON.CPN = Trim$("" & rs!CustomerPartnumber)
        tCusCARTON.MPN = Trim$("" & rs!Customer_Device)
        tCusCARTON.QTY = "" & rs!QTY
        KID = rs!KID
        strContent = strContent & tCusCARTON.dn & ",I" & tCusCARTON.dn & "," & tCusCARTON.PO & ",K" & tCusCARTON.PO & ",," & tCusCARTON.CPN & ",P" & tCusCARTON.CPN & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & ","
        strContent = strContent & tCusCARTON.QTY & ",Q" & tCusCARTON.QTY & ","
        strContent = strContent & Get_OracleStr("select distinct trim(a.freightforwarder)|| ',CHINA,' || substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3)||','||trim(a.city) || ' ' || trim(a.state) || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ','|| 'Attn:;Tel:' || trim(a.phone) || ','  from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & strDN & "'")
        strContent = strContent & "N/A,PN/A,N/A,9DN/A," & lOutboxNum & "," & KID & "," & sMaxOP
        rs.MoveNext
    Loop

End If

Call CreateTxt(strFileName, strContent, strSSCartonPath)
Call Sleep(gSleepMicSec)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       Print37CartonStanderLbl_OLD
' Description:       打印37标准大标签
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-10:59:05
'
' Parameters :       strDN (String)
'                    lOutboxNum (Long)
'--------------------------------------------------------------------------------
Private Sub Print37CartonStanderLbl_OLD(strDN As String, _
                                        lOutboxNum As Long, _
                                        strQID As String)
Dim strSql      As String
Dim tCusCARTON  As CUSCARTON
Dim strFileName As String
Dim strTxt      As String
Dim strKid      As String
Dim strMaxOP    As String
Dim stradd      As String
Dim rs          As New ADODB.Recordset
Dim strQrCode   As String
Dim strShipTo   As String

'出货分类
strSql = "select labelrequirement from customershippinguptbl where delivery = '" & strDN & "'"
strShipTo = UCase(Get_OracleStr(strSql))
If InStr(strShipTo, "HUAWEI") > 0 Then
    strShipTo = "HUAWEI"

End If

If InStr(strShipTo, "E2") > 0 Then
    strShipTo = "SSE2"

End If

If InStr(strShipTo, "SEMTECH") > 0 Then
    strShipTo = "ST"

End If

If InStr(strShipTo, "SHORT") > 0 Then
    strShipTo = "SSSHORT"

End If

strSql = "select max(OUTBOX_NUM) from PACKING_DETAILED where DN_NUM = '" & strDN & "'"
strMaxOP = Get_OracleStr(strSql)
If strQID <> "" Then
    lOutboxNum = Get_OracleStr("select distinct outbox_num from packing_detailed where dn_num = '" & strDN & "' and carton = '" & strQID & "'")

End If

strSql = "select a.CUSTOMER_DEVICE,a.kid, b.customerpartnumber,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & strDN & "' and b.delivery = '" & strDN & "' and a.job_id = b.batchnumber and a.outbox_num = '" & lOutboxNum & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno,a.kid"
Set rs = Get_OracleRs(strSql)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tCusCARTON.dn = strDN

        Select Case strShipTo

            Case "HUAWEI"   '出华为
                tCusCARTON.PO = UCase(IIf(IsNull(rs!PO), "N/A", rs!PO))

            Case "ST"
                tCusCARTON.PO = UCase(IIf(IsNull(rs!PO), "N/A", Left(rs!PO, 10)))

            Case Else
                tCusCARTON.PO = UCase(IIf(IsNull(rs!PO), "N/A", rs!PO))

        End Select

        tCusCARTON.CPN = UCase(IIf(IsNull(rs!CustomerPartnumber), "N/A", rs!CustomerPartnumber))
        tCusCARTON.MPN = UCase(IIf(IsNull(rs!Customer_Device), "N/A", rs!Customer_Device))
        tCusCARTON.KID = Trim("" & rs!KID)
        tCusCARTON.QTY = rs!QTY
        tCusCARTON.FactoryFlow = Get_OracleStr("select distinct material from customershippinguptbl where marketingpn = '" & tCusCARTON.MPN & "' and delivery = '" & strDN & "'")
        '判断
        If tCusCARTON.dn = "" Or tCusCARTON.PO = "" Or tCusCARTON.CPN = "" Or tCusCARTON.MPN = "" Or tCusCARTON.KID = "" Or tCusCARTON.FactoryFlow = "" Then
            MsgBox "标签特定字段不可为空,请联系IT处理", vbCritical, "警告"
            Exit Sub

        End If

        If tCusCARTON.QTY = 0 Then
            MsgBox "外箱数量不可为0", vbCritical, "警告"
            Exit Sub

        End If

        strTxt = strTxt & Get_OracleStr("select distinct substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3) || ','||trim(a.city) || ' ' || trim(a.state)  || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ',' || trim(a.contactname) || ',' || trim(a.phone) from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & strDN & "' ") & ","
        strTxt = strTxt & tCusCARTON.dn & ",I" & tCusCARTON.dn & "," & tCusCARTON.PO & ",K" & tCusCARTON.PO & "," & tCusCARTON.CPN & ",P" & tCusCARTON.CPN & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & "," & tCusCARTON.QTY & ",Q" & tCusCARTON.QTY & "," & Get_OracleStr("select distinct freightforwarder from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & strDN & "'") & "," & "" & "," & "" & "," & "" & "," & "COO:CHINA" & "," & "CHINA"
        stradd = "," & lOutboxNum & "," & tCusCARTON.KID
        
        If Check1.Value = 1 Then
             strTxt = strTxt & stradd & "," & strMaxOP & "," & "2S" & strDN & "," & "1P" & tCusCARTON.MPN & "," & tCusCARTON.FactoryFlow & ",6P" & tCusCARTON.FactoryFlow & ",3S" & tCusCARTON.dn & "-" & Right$("0" & Replace(tCusCARTON.KID, "K", ""), 2) & ","
        Else
             strTxt = strTxt & stradd & "," & strMaxOP & "," & "2S" & strDN & "," & "1P" & tCusCARTON.MPN & "," & tCusCARTON.FactoryFlow & ",6P" & tCusCARTON.FactoryFlow & ",3S" & tCusCARTON.KID & ","
        End If

        strQrCode = "[)>" & Chr(30) & "06" & Chr(29) & "2S" & strDN & Chr(29) & "K" & tCusCARTON.PO & Chr(29) & "P" & tCusCARTON.CPN & Chr(29) & "1P" & tCusCARTON.MPN & Chr(29) & "6P" & tCusCARTON.FactoryFlow & Chr(29) & "Q" & tCusCARTON.QTY & Chr(29) & "3S" & tCusCARTON.KID & Chr(29) & "4LCN" & Chr(30) & Chr(4)
        strTxt = strTxt & strQrCode
        strQrCode = Replace(Replace(Replace(strQrCode, Chr(30), ""), Chr(29), ""), Chr(4), "")
        AddSql ("delete from TBL37QRVALUE where KEY_NAME = '" & strDN & "' || '_' ||  '" & tCusCARTON.KID & "'")
        AddSql ("insert into TBL37QRVALUE(KEY_NAME,KEY_VALUE,CREATE_DATE,CREATE_BY) values('" & strDN & "' || '_' ||  '" & tCusCARTON.KID & "','" & strQrCode & "',sysdate,'" & gUserName & "')  ")
        rs.MoveNext
    Loop

End If

strFileName = strDN & "-" & "SemTechStanderCarton" + Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(strFileName, strTxt, str37CartonPath)
Call Sleep(gSleepMicSec)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PlaySound
' Description:       播放声音文件
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-9:32:48
'
' Parameters :       strSound (String)
'--------------------------------------------------------------------------------
Private Sub PlaySound(strSound As String)
player1.url = gMediaDir & strSound & ".wav"

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CreateTxt
' Description:       生成txt
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-9:32:58
'
' Parameters :       filename (String)
'                    msgTxt (String)
'                    dirtemp (String)
'--------------------------------------------------------------------------------
Private Sub CreateTxt(filename As String, msgTxt As String, dirtemp As String)
Dim fileNameTemp As String
Dim dirNameTemp  As String
Dim fileTemp     As String

dirNameTemp = dirtemp
fileNameTemp = Replace(filename, "'", "") & ".txt"
fileTemp = dirNameTemp & fileNameTemp
Open fileTemp For Output As #1
Print #1, msgTxt
Close #1

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       UpdateERP_CARTON_NO
' Description:       更新ERP箱号对照关系
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-12:03:00
'
' Parameters :       strDN (String)
'--------------------------------------------------------------------------------
Private Sub UpdateERP_CARTON_NO(strDN As String)
Dim strSql      As String
Dim rs          As ADODB.Recordset
Dim strCartonID As String, strCartonQty As String
Dim id          As String

On Error GoTo ERRON

INIadoCon.BeginTrans
strSql = "select CARTON, SUM(QTY) from PACKING_DETAILED where dn_num = '" & strDN & "' group by CARTON"
Set rs = Get_OracleRs(strSql)
If rs.EOF Then
    MsgBox "PACKING_DETAILED查询不到该DN", vbInformation, "提示"
    INIadoCon.RollbackTrans
    Exit Sub

End If

rs.MoveFirst

Do While Not rs.EOF
    strCartonID = Trim$("" & rs(0))
    strCartonQty = Trim$("" & rs(1))
    ' ---------------------------------------------------删除
    '0
    strSql = "delete from [erpdata].[dbo].[tblPackTreeInf] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strSql)
    strSql = "delete from [erpdata].[dbo].[tblPackMainInf] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strSql)
    strSql = "update [erpdata].[dbo].[tblPackTreeInf] set 上级序号 = '', Memo = '' where 箱号 in (select trayid from erpbase..PACKING_DETAILED where carton = '" & strCartonID & "')  "
    AddSql2 (strSql)
    strSql = "delete from [erpdata].[dbo].[tblStockNumTree] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strSql)
    strSql = "Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='',Memo='', dn='' where 箱号 in (select trayid from erpbase..PACKING_DETAILED where carton = '" & strCartonID & "') "
    AddSql2 (strSql)
    ' --------------------------------------------------更新
    '1 insert [erpdata].[dbo].[tblPackMainInf]
    strSql = "insert into [erpdata].[dbo].[tblPackMainInf](箱号,客户代码,数量,合格标记,装箱标记,产线标记) values('" & strCartonID & "','37'," & strCartonQty & ",'0','1','1')"
    If AddSql2(strSql) = 0 Then
        MsgBox "1 insert [erpdata].[dbo].[tblPackMainInf]:failed!!! ", vbCritical, "提示"
        Exit Sub

    End If

    '2 insert - update [erpdata].[dbo].[tblPackTreeInf]
    strSql = "insert into [erpdata].[dbo].[tblPackTreeInf](箱号,上级序号,基层标记 ,Memo) values('" & strCartonID & "',0,1,'37')"
    If AddSql2(strSql) = 0 Then
        MsgBox "2 insert [erpdata].[dbo].[tblPackTreeInf]:failed!!!", vbCritical, "提示"
        Exit Sub

    End If

    id = Get_SqlserverNo("select 序号 as ID from [erpdata].[dbo].[tblPackTreeInf] a where a.箱号='" & strCartonID & "' and Memo='37' ")
    strSql = "Update [erpdata].[dbo].[tblPackTreeInf] set 上级序号='" & id & "',Memo='37' " & " where 箱号 in ( select trayid from  OPENQUERY(ORACLEDB, 'SELECT * from packing_detailed where carton = ''" & strCartonID & "'' ')) "
    If AddSql2(strSql) = 0 Then
        MsgBox "2 update [erpdata].[dbo].[tblPackTreeInf]:failed!!!", vbCritical, "提示"
        Exit Sub

    End If

    '3 insert - update [erpdata].[dbo].[tblStockNumTree]
    strSql = "insert into [erpdata].[dbo].[tblStockNumTree](序号,箱号,上级序号,基层标记 ,尺寸,重量,Memo,DN) values( " & id & ",'" & strCartonID & "',0,1,'','','37','" & strDN & "')"
    If AddSql2(strSql) = 0 Then
        MsgBox "3 insert [erpdata].[dbo].[tblStockNumTree]", vbCritical, "提示"
        Exit Sub

    End If

    strSql = "Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='" & id & "',Memo='37', dn='" & strDN & "' " & " where 箱号 in ( select trayid from  OPENQUERY(ORACLEDB, 'SELECT * from packing_detailed where carton = ''" & strCartonID & "'' ')) "
    If AddSql2(strSql) = 0 Then
        MsgBox "3 update [erpdata].[dbo].[tblStockNumTree]", vbCritical, "提示"
        Exit Sub

    End If

    rs.MoveNext
Loop
INIadoCon.CommitTrans
'MsgBox "DN:" & strDN & "  :箱号已更新", vbInformation, "提示"
Exit Sub
ERRON:
INIadoCon.RollbackTrans
MsgBox "错误:" & Err.DESCRIPTION, vbCritical, "警告"

End Sub

'-------------------------------------------------------------
'<<<<<<<<<<<<<<<<<<补打标签>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'-------------------------------------------------------------
Private Sub PrintHandler2()
Dim strKey As String

strKey = UCase(Trim(txtScan2.text))
If Len(strKey) = 0 Then
    MsgBox "请输入需要补打的条码", vbInformation, "提示"
    Exit Sub

End If

Call printLblNew(strKey)
txtScan2.text = ""

End Sub

Private Sub Command1_Click()
Dim strSql As String

If txtUser2.text = txtUser.text Then
    MsgBox "员工不可输入组长的工号", vbCritical, "提示"
    Exit Sub

End If

strSql = "select * from tblOperatorData r where  r.状态标记=1  and r.用户号='10354'and r.密码='" & Replace(Trim(txtPassWd.text), "'", "") & "'"
If Get_SqlStr(strSql) = "" Then
    MsgBox "组长密码不正确", vbCritical, "提示"
    Exit Sub

End If

strSql = "select * from tblOperatorData r where  r.状态标记=1  and r.用户号='" & Trim(txtUser2.text) & "'and r.密码='" & Replace(Trim(txtPassWd2.text), "'", "") & "'"
If Get_SqlStr(strSql) = "" Then
    MsgBox "员工工号或者密码不正确", vbCritical, "提示"
    Exit Sub

End If

txtScan2.Visible = True

End Sub

Private Sub printLblNew(strKey As String)
Dim iQty       As Integer
Dim strDN      As String
Dim lOutboxNum As Long
Dim lInboxNum  As Long
Dim strShipTo  As String
Dim strSql     As String

If cbLblType.text = "" Then
    MsgBox "请选择补打的标签类型", vbInformation, "提示"
    Exit Sub

End If

Select Case cbLblType.text

    Case "37内盒-B标签"
        strDN = Get_OracleStr("select dn_num from packing_detailed where boxid = '" & strKey & "'")
        If strDN = "" Then
            MsgBox "查询不到该37内盒-BID,无法补打", vbInformation, "提示"
            Exit Sub

        End If

        Call Print37BoxLbl_OLD(strDN, lOutboxNum, lInboxNum, strKey)

    Case "37外箱-C标签"
        strDN = Get_OracleStr("select dn_num from packing_detailed where CARTONID = '" & strKey & "'")
        If strDN = "" Then
            MsgBox "查询不到该37外箱-C箱号,无法补打", vbInformation, "提示"
            Exit Sub

        End If

        Call Print37CartonLbl_OLD(strDN, lOutboxNum, strKey)

    Case "37外箱标准大标签"
        strDN = Get_OracleStr("select dn_num from packing_detailed where CARTON = '" & strKey & "'")
        If strDN = "" Then
            MsgBox "查询不到该37外箱Q箱号,无法补打", vbInformation, "提示"
            Exit Sub

        End If

        Call Print37CartonStanderLbl_OLD(strDN, lOutboxNum, strKey)

    Case "三星卷盘标签"
        strDN = Get_OracleStr("select dn_num from packing_detailed where trayid = '" & strKey & "'")
        If strDN = "" Then
            MsgBox "查询不到该37卷盘-RID,无法补打", vbInformation, "提示"
            Exit Sub

        End If

        strSql = "select labelrequirement from customershippinguptbl where delivery = '" & strDN & "'"
        strShipTo = UCase(Get_OracleStr(strSql))
        If InStr(strShipTo, "E2") > 0 Then
            Call PrintSSE2ReelLbl_OLD(strDN, lOutboxNum, lInboxNum, strKey)

        End If

        If InStr(strShipTo, "SHORT") > 0 Then
            Call PrintSSSHORTReelLbl_OLD(strDN, lOutboxNum, lInboxNum, strKey)

        End If

    Case "三星内盒标签"
        strDN = Get_OracleStr("select dn_num from packing_detailed where boxid = '" & strKey & "'")
        If strDN = "" Then
            MsgBox "查询不到该37内盒-BID,无法补打", vbInformation, "提示"
            Exit Sub

        End If

        strSql = "select labelrequirement from customershippinguptbl where delivery = '" & strDN & "'"
        strShipTo = UCase(Get_OracleStr(strSql))
        If InStr(strShipTo, "E2") > 0 Then
            Call PrintSSE2BoxLbl_OLD(strDN, lOutboxNum, lInboxNum, strKey)

        End If

        If InStr(strShipTo, "SHORT") > 0 Then
            Call PrintSSSHORTBoxLbl_OLD(strDN, lOutboxNum, lInboxNum, strKey)

        End If

    Case "三星外箱大标签"
        strDN = Get_OracleStr("select dn_num from packing_detailed where CARTON = '" & strKey & "'")
        If strDN = "" Then
            MsgBox "查询不到该37外箱Q箱号,无法补打", vbInformation, "提示"
            Exit Sub

        End If

        strSql = "select labelrequirement from customershippinguptbl where delivery = '" & strDN & "'"
        strShipTo = UCase(Get_OracleStr(strSql))
        If InStr(strShipTo, "E2") > 0 Then
            Call PrintSSE2CartonLbl_OLD(strDN, lOutboxNum, strKey)

        End If

        If InStr(strShipTo, "SHORT") > 0 Then
            Call PrintSSSHORTCartonLbl_OLD(strDN, lOutboxNum, strKey)

        End If

    Case "华为卷盘标签"
        strDN = Get_OracleStr("select dn_num from packing_detailed where trayid = '" & strKey & "'")
        If strDN = "" Then
            MsgBox "查询不到该37卷盘-RID,无法补打", vbInformation, "提示"
            Exit Sub

        End If

        Call PrintHWReelLbl_OLD(strDN, lOutboxNum, lInboxNum, strKey)

    Case "华为内盒标签"
        strDN = Get_OracleStr("select dn_num from packing_detailed where boxid = '" & strKey & "'")
        If strDN = "" Then
            MsgBox "查询不到该37内盒-BID,无法补打", vbInformation, "提示"
            Exit Sub

        End If

        Call PrintHWBoxLbl_OLD(strDN, lOutboxNum, lInboxNum, strKey)

    Case "华为外箱标准大标签"
        strDN = Get_OracleStr("select dn_num from packing_detailed where CARTON = '" & strKey & "'")
        If strDN = "" Then
            MsgBox "查询不到该37外箱Q箱号,无法补打", vbInformation, "提示"
            Exit Sub

        End If

        Call Print37CartonStanderLbl_OLD(strDN, lOutboxNum, strKey)

    Case "卷盘二维码标签转换"
        If txtDN2.text = "" Then
            MsgBox "DN不可为空", vbCritical, "警告"
            Exit Sub

        End If

        Call Print37QrReelLbl(strKey)

    Case Else
        MsgBox "暂未开发", vbInformation, "提示"
        Exit Sub

End Select

iQty = Get_OracleStr("select nvl(count(*) + 1, 1) from TBL_37_PRINT2_LIST where KEYNAME = '" & cbLblType.text & "' and KEYVALUE = '" & strKey & "'")
AddSql ("insert into TBL_37_PRINT2_LIST(KEYNAME,KEYVALUE,CREATE_DATE,CREATE_BY,CREATE_TIMES) values('" & cbLblType.text & "', '" & strKey & "', sysdate, '" & gUserName & "', '" & iQty & "')")
MsgBox "标签补打完成", vbInformation, "提示"

End Sub

Private Sub Print37QrReelLbl(strTrayID As String)
Dim strSql      As String
Dim strTxt      As String
Dim strFlagTxt  As String
Dim strFileName As String
Dim rsJobID     As New ADODB.Recordset
Dim tSTBox      As STBox
Dim strQrCode   As String
Dim strDN As String

strDN = Trim$(txtDN2.text)

tSTBox.JOB = GetJobID(strTrayID)
tSTBox.DEV = Get_SqlStr("  SELECT distinct  t2.mpn FROM erpdata..tblPackMainInfSub t1 " & _
"  inner join [erpdata].[dbo].tblTSVworkorder t2 on t1.大工单 = t2.ORDERNAME " & _
"  where t1.箱号 = '" & strTrayID & "' ")
tSTBox.lot = strTrayID
tSTBox.DATECODE = Get37TestDC(strDN, tSTBox.JOB)
tSTBox.QTY = GetReelQty(strTrayID)
tSTBox.FactoryFlow = Get_OracleStr("select distinct material from customershippinguptbl where marketingpn = '" & tSTBox.DEV & "' and delivery = '" & strDN & "'")
strTxt = strTxt & tSTBox.DEV & "," & tSTBox.JOB & ",1T" & tSTBox.JOB & "," & tSTBox.DEV & "," & "1P" & tSTBox.DEV & "," & tSTBox.DATECODE & "," & tSTBox.DATECODE & "," & Mid(tSTBox.lot, 2) & "," & tSTBox.lot & "," & tSTBox.QTY & ",Q" & tSTBox.QTY & "," & tSTBox.DATECODE & "," & tSTBox.DATECODE & GetDevMark(tSTBox.DEV)
strTxt = strTxt & "," & tSTBox.FactoryFlow & "," & "6P" & tSTBox.FactoryFlow & "," & "10D" & tSTBox.DATECODE & ","
strQrCode = "[)>" & Chr(30) & "06" & Chr(29) & "1T" & tSTBox.JOB & Chr(29) & "1P" & tSTBox.DEV & Chr(29) & tSTBox.lot & Chr(29) & "Q" & tSTBox.QTY & Chr(29) & "6P" & tSTBox.FactoryFlow & Chr(29) & "10D" & tSTBox.DATECODE & Chr(30) & Chr(4)
strTxt = strTxt & strQrCode & vbCrLf
strQrCode = Replace(Replace(Replace(strQrCode, Chr(30), ""), Chr(29), ""), Chr(4), "")
strFileName = "RID:" & strTrayID & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(strFileName, strTxt, str37BCIDPath)

'MsgBox "补打完成", vbInformation, "提示"

End Sub

Private Sub txtScan2_KeyPress(KeyAscii As Integer)
Dim strScan As String

strScan = UCase(Trim(txtScan2.text))
If KeyAscii <> vbKeyReturn Or Len(strScan) = 0 Then Exit Sub
Call printLblNew(strScan)
txtScan2.text = ""

End Sub
