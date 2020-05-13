VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLblPrint_37ToHW 
   BackColor       =   &H00C0C0C0&
   Caption         =   "标签打印系统_37出华为"
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
            Caption         =   "删除"
            Key             =   "DEL"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "导出"
            Key             =   "EXPORT"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
               Picture         =   "FrmLblPrint_37ToHW.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint_37ToHW.frx":213A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint_37ToHW.frx":4FC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint_37ToHW.frx":7776
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint_37ToHW.frx":98B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint_37ToHW.frx":C062
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint_37ToHW.frx":E814
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint_37ToHW.frx":11896
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint_37ToHW.frx":14048
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint_37ToHW.frx":14362
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint_37ToHW.frx":1503C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint_37ToHW.frx":180BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLblPrint_37ToHW.frx":1A870
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTTab0 
      Height          =   13455
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   20325
      _ExtentX        =   35851
      _ExtentY        =   23733
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483637
      ForeColor       =   255
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
      TabPicture(0)   =   "FrmLblPrint_37ToHW.frx":1B14A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraMnu"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraScanDetail"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "标签补打"
      TabPicture(1)   =   "FrmLblPrint_37ToHW.frx":1B166
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
      Tab(1).Control(9)=   "Command2"
      Tab(1).Control(10)=   "cbLblType"
      Tab(1).Control(11)=   "txtScan2"
      Tab(1).ControlCount=   12
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
         TabIndex        =   20
         Top             =   1560
         Width           =   19815
         Begin FPSpreadADO.fpSpread fpS 
            Height          =   9255
            Index           =   0
            Left            =   240
            TabIndex        =   29
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
            SpreadDesigner  =   "FrmLblPrint_37ToHW.frx":1B182
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
            TabIndex        =   26
            Top             =   3720
            Width           =   9855
         End
         Begin FPSpreadADO.fpSpread fpS 
            Height          =   3015
            Index           =   1
            Left            =   9600
            TabIndex        =   21
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
            SpreadDesigner  =   "FrmLblPrint_37ToHW.frx":1B5A4
            TextTip         =   2
            AppearanceStyle =   0
         End
         Begin FPSpreadADO.fpSpread fpS 
            Height          =   3015
            Index           =   2
            Left            =   14760
            TabIndex        =   22
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
            SpreadDesigner  =   "FrmLblPrint_37ToHW.frx":1BA16
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
            Top             =   360
            Width           =   885
         End
      End
      Begin VB.Frame fraMnu 
         Caption         =   "菜单"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   19815
         Begin VB.TextBox txtCurOP 
            BackColor       =   &H00E0E0E0&
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
            Left            =   7800
            TabIndex        =   33
            Text            =   "1"
            Top             =   675
            Width           =   975
         End
         Begin VB.TextBox txtMaxOP 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   675
            Width           =   1455
         End
         Begin VB.TextBox txtReelID 
            BackColor       =   &H00FFC0FF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   720
            TabIndex        =   28
            Top             =   675
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox txtDN 
            BackColor       =   &H00FFC0FF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   720
            TabIndex        =   16
            Top             =   330
            Width           =   2295
         End
         Begin VB.TextBox txtQty 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   330
            Width           =   1455
         End
         Begin VB.Label lblCurOP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "当前外箱序号"
            Height          =   195
            Left            =   6600
            TabIndex        =   32
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label lblMaxOp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "总箱数"
            Height          =   195
            Left            =   3960
            TabIndex        =   30
            Top             =   720
            Width           =   840
         End
         Begin VB.Label lblReelID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "卷盘"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   360
            TabIndex        =   27
            Top             =   720
            Visible         =   0   'False
            Width           =   360
         End
         Begin WMPLibCtl.WindowsMediaPlayer player1 
            Height          =   495
            Left            =   14160
            TabIndex        =   19
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
            Left            =   420
            TabIndex        =   18
            Top             =   375
            Width           =   210
         End
         Begin VB.Label lblQTY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "总数量(颗)"
            Height          =   195
            Left            =   3960
            TabIndex        =   17
            Top             =   375
            Width           =   840
         End
      End
      Begin VB.TextBox txtScan2 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   -73320
         TabIndex        =   9
         Top             =   780
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.ComboBox cbLblType 
         Height          =   315
         ItemData        =   "FrmLblPrint_37ToHW.frx":1BE88
         Left            =   -73320
         List            =   "FrmLblPrint_37ToHW.frx":1BE9B
         Style           =   2  'Dropdown List
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   780
         Width           =   2295
      End
      Begin VB.TextBox txtPassWd2 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -71640
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   4140
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "验证补打密码"
         Height          =   840
         Left            =   -68640
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3660
         Width           =   1575
      End
      Begin VB.TextBox txtPassWd 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -71640
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   3660
         Width           =   2895
      End
      Begin VB.TextBox txtUser 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -72960
         TabIndex        =   3
         Text            =   "10354"
         Top             =   3660
         Width           =   1215
      End
      Begin VB.TextBox txtUser2 
         Height          =   375
         Left            =   -72960
         TabIndex        =   2
         Top             =   4140
         Width           =   1215
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label266 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmLblPrint_37ToHW.frx":1BEE4
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
         TabIndex        =   11
         Top             =   4185
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmLblPrint_37ToHW.frx":1BEF8
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
         TabIndex        =   10
         Top             =   3720
         Width           =   1635
      End
   End
End
Attribute VB_Name = "FrmLblPrint_37ToHW"
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

Private gMediaDir As String

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

Private Type uReelInfo

    TYAYID As String
    INBOX_NUM As Long
    OUTBOX_NUM As Long
    DN_NUM As String
    JOB_ID As String
    QTY As Long
    Customer_Device As String
    CREATE_DATE As String
    CREATE_BY As String
    PRINT_FLAG As String
    FLAG As String
    carton As String
    REELID As String
    BOXID As String
    CARTONID As String
    KID As String
    SEQ As String
    DATECODE As String

End Type

Private Sub SSTTab0_Click(PreviousTab As Integer)
Select Case SSTTab0.Tab

    Case 0
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(1).Caption = "打印标签"
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(4).Enabled = True
        Toolbar1.Buttons(4).Caption = "导出打印记录"

    Case 1
        Toolbar1.Buttons(1).Enabled = True
        Toolbar1.Buttons(1).Caption = "补打标签"
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(4).Enabled = True
        Toolbar1.Buttons(4).Caption = "导出补打记录"

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
                Call Print2Handler

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

strDN = Trim$(txtDN.Text)
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
str37BCIDPath = "\\10.160.1.84\public\BarCode\37\37内箱\"        ' 37B,C,R小标签
str37CartonPath = "\\10.160.1.84\public\BarCode\37\37外箱\"      ' 37自家外箱大标签
' 出三星标签
strSSBoxPath = "\\10.160.1.84\public\BarCode\37\37BoxNH\"      ' 三星内盒小标签
strSSReelPath = "\\10.160.1.84\public\BarCode\37\37BoxJP\"     ' 三星卷盘小标签
strSSCartonPath = "\\10.160.1.84\public\BarCode\37\37BoxOut\"  ' 三星外箱大标签
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
strSSReelPath = "C:\test\"    ' 三星卷盘小标签
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
With Fps(0)
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
    .TypeVAlign = TypeVAlignCenter
    .TypeHAlign = TypeVAlignCenter
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
With Fps(1)
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
    .TypeVAlign = TypeVAlignCenter
    .TypeHAlign = TypeVAlignCenter
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
With Fps(2)
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
    .TypeVAlign = TypeVAlignCenter
    .TypeHAlign = TypeVAlignCenter
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
strDN = Right$(Trim(txtDN.Text), consDNLen)
If Len(strDN) <> consDNLen Then
    MsgBox "请扫描正确的DN", vbInformation, "DN扫描"
    txtDN.Text = ""
    Exit Sub

End If

If CheckDN(strDN) = False Then
    txtDN.Text = ""
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
Dim strSql As String

txtDN.Text = strDN
strSql = "select sum(quantity) from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "'"
txtQty.Text = Get_OracleStr(strSql)

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

With Fps(0)
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

With Fps(1)
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

With Fps(2)
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
txtStatus.Text = vbCrLf & Get_OracleNo(strSql)

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
Dim strDN     As String
Dim strReelID As String
Dim strJobID  As String
Dim lReelQty  As Long

If KeyAscii <> vbKeyReturn Then Exit Sub
If (Len(Trim(txtReelID.Text)) <> consReelIDLen) And (Len(Trim(txtReelID.Text)) <> consReelIDLen + 1) Then
    MsgBox "请扫描正确的卷盘号", vbInformation, "卷盘扫描"
    txtReelID.Text = ""
    Exit Sub

End If

strDN = Trim$(txtDN.Text)
strReelID = UCase(Trim(txtReelID.Text))
strJobID = GetJobID(strReelID)
lReelQty = GetReelQty(strReelID)
If strJobID = "" Or CheckReelID(strDN, strJobID, strReelID, lReelQty) = False Then
    txtStatus.BackColor = vbRed
    txtReelID.Text = ""
    Exit Sub
Else
    txtStatus.BackColor = vbWhite

End If

Call AddReelInfo(strDN, strJobID, strReelID, lReelQty)
Call ShowScanInfo(strDN)
Call CheckScanningComplate(strDN)
txtReelID.Text = ""

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
Private Function CheckReelID(strDN As String, _
                             strJobID As String, _
                             strReelID As String, _
                             lReelQty As Long) As Boolean
Dim strSql As String

CheckReelID = False
strSql = "select * from packing_detailed where trayid = '" & strReelID & "' and dn_num <> '" & strDN & "'"
If Get_OracleCnt(strSql) > 0 Then
    MsgBox "该卷盘: " & strReelID & " 有扫描历史,请确认是否有误", vbCritical, "警告"
    Exit Function

End If

strSql = "select * from packing_detailed where dn_num = '" & strDN & "' and trayid = '" & strReelID & "'"
If Get_OracleCnt(strSql) > 0 Then
    Call PlaySound("该卷盘已经扫描过, 请勿重复扫描")
    Exit Function

End If

strSql = "select * from customershippinguptbl where delivery =  '" & strDN & "' and batchnumber = '" & strJobID & "'"
If Get_OracleCnt(strSql) = 0 Then
    MsgBox "该卷盘: " & strReelID & " 的JobID: " & strJobID & " 不属于本次DN: " & strDN, vbCritical, "警告"
    Exit Function

End If

If CheckSamgJob(strDN, strJobID, lReelQty) = False Then
    Exit Function

End If

Call PlaySound("卷盘号正确")
CheckReelID = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckSamgJob
' Description:       不可以跨JOB作业
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/8-15:51:31
'
' Parameters :       strReelID (String)
'                    strJobID (String)
'                    strLastJobID (String)
'--------------------------------------------------------------------------------
Private Function CheckSamgJob(strDN As String, _
                              strJobID As String, _
                              lReelQty As Long) As Boolean
CheckSamgJob = False
Dim strSql           As String
Dim strLastJobID     As String
Dim lLastJobCurQty   As Long
Dim lLastJobTotalQty As Long

With Fps(0)
    .Row = 1
    .Col = E_REEL_SCAN.E_REEL_JOBID
    strLastJobID = .Text

End With

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

CheckSamgJob = True

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

strSql = "select KEY_VALUE from erpdata..tblErpInStockDetailInfo a where CHARINDEX('" & strReelID & "',a.KEY_VALUE) > 0 and a.KEY_NAME = 'CONTAINER_NAME' AND a.KEY_TYPE = 'T'"
strRes = Get_SqlStr(strSql)
GetJobID = Mid(strRes, InStr(strRes, "|") + 1)

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

strSql = "select SUM(入库数) from erpdata..tblPackToHouseSub where 箱号 = '" & strReelID & "'"
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

strSql = "select sum(quantity) from customershippinguptbl where delivery = '" & strDN & "'"
lTotalQty = Get_OracleNo(strSql)
strSql = "select sum(qty) from packing_detailed where dn_num = '" & strDN & "'"
lCurQty = Get_OracleNo(strSql)
If lCurQty = lTotalQty Then
    strSql = "select max(OUTBOX_NUM) from packing_detailed where dn_num = '" & strDN & "'"
    txtMaxOP.Text = Get_OracleNo(strSql)
    txtReelID.Enabled = False
    Toolbar1.Buttons("PRINT").Enabled = True
    MsgBox "该DN所有卷盘已全部扫描完毕,请点击打印按钮,开始打印标签", vbInformation, "提示"
    Call UpdateERP_CARTON_NO(strDN)

End If

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       AddReelInfo
' Description:       卷盘信息存入PACKING_DETAILED
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/9-9:09:20
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub AddReelInfo(strDN As String, _
                        strJobID As String, _
                        strReelID As String, _
                        lReelQty As Long)
Dim ut As uReelInfo

ut.DN_NUM = strDN
ut.JOB_ID = strJobID
ut.TYAYID = strReelID
ut.QTY = lReelQty
ut.Customer_Device = GetCustPN(strDN, strJobID)
ut.REELID = GetPSN(strDN, strJobID, strReelID)
ut.SEQ = GetSeq(strDN)
ut.CREATE_DATE = Now
ut.CREATE_BY = gUserName
ut.PRINT_FLAG = "0"
ut.FLAG = "0"
ut.DATECODE = Get37TestDC(strDN, strJobID)
Call GetOtherData(ut)
Call SaveReelInfo(ut)

End Sub

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
Private Function GetOtherData(ByRef ut As uReelInfo)
Dim strSql         As String
Dim strLastMPN     As String
Dim strLastJobID   As String
Dim lLastOutBoxNum As Long
Dim lLastInboxNum  As Long
Dim lLastInboxCnt  As Long

strSql = "select nvl(max(OUTBOX_NUM),0) from PACKING_DETAILED where dn_num = '" & ut.DN_NUM & "'"
lLastOutBoxNum = Get_OracleNo(strSql)
ut.OUTBOX_NUM = lLastOutBoxNum
strSql = "select nvl(max(INBOX_NUM),0) from PACKING_DETAILED where dn_num = '" & ut.DN_NUM & "' and OUTBOX_NUM = '" & ut.OUTBOX_NUM & "' "
lLastInboxNum = Get_OracleStr(strSql)
ut.INBOX_NUM = lLastInboxNum
strSql = "select CUSTOMER_DEVICE from packing_DETAILED where dn_num = '" & ut.DN_NUM & "' order by seq desc"
strLastMPN = Get_OracleStr(strSql)
strSql = "select count(*) from packing_detailed where dn_num = '" & ut.DN_NUM & "' and outbox_num = '" & ut.OUTBOX_NUM & "' and inbox_num = '" & ut.INBOX_NUM & "' "
lLastInboxCnt = Get_OracleNo(strSql)
strSql = "select job_id from packing_DETAILED where dn_num = '" & ut.DN_NUM & "' order by seq desc"
strLastJobID = Get_OracleStr(strSql)
'Get OutboxNum InboxNum
If ut.Customer_Device <> strLastMPN Then
    ut.OUTBOX_NUM = lLastOutBoxNum + 1
    ut.INBOX_NUM = 1
Else
    If lLastInboxCnt = 9 Then
        ut.INBOX_NUM = lLastInboxNum + 1
        If ut.INBOX_NUM = 13 Then
            ut.INBOX_NUM = 1
            ut.OUTBOX_NUM = lLastOutBoxNum + 1

        End If

    End If

End If

ut.KID = "K" & ut.OUTBOX_NUM
'GetCID
strSql = "select CARTONID from packing_DETAILED where dn_num = '" & ut.DN_NUM & "' order by seq desc"
ut.CARTONID = Get_OracleStr(strSql)
If ut.OUTBOX_NUM <> lLastOutBoxNum Then
    ut.CARTONID = GetNewID(ut, "-C")
Else
    If ut.JOB_ID <> strLastJobID Then
        ut.CARTONID = GetNewID(ut, "-C")

    End If

End If

'GetBID
strSql = "select BOXID from packing_DETAILED where dn_num = '" & ut.DN_NUM & "' order by seq desc"
ut.BOXID = Get_OracleStr(strSql)
If ut.OUTBOX_NUM <> lLastOutBoxNum Then
    ut.BOXID = GetNewID(ut, "-B")
ElseIf ut.INBOX_NUM <> lLastInboxNum Then
    ut.BOXID = GetNewID(ut, "-B")
Else
    If ut.JOB_ID <> strLastJobID Then
        ut.BOXID = GetNewID(ut, "-B")

    End If

End If

'GetQID
strSql = "select CARTON from packing_DETAILED where dn_num = '" & ut.DN_NUM & "' order by seq desc"
ut.carton = Get_OracleStr(strSql)
If ut.OUTBOX_NUM <> lLastOutBoxNum Then
    ut.carton = GetQID(ut)

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
Private Function GetNewID(ut As uReelInfo, strflag As String) As String
Dim strSql   As String
Dim strBase  As String
Dim strseq   As String
Dim strNewID As String

'strBase = Left$(ut.TYAYID, 9) & strflag

strBase = Left$(ut.TYAYID, InStr(ut.TYAYID, "-") - 1) & strflag

strSql = "select nvl(max(seq)+1, 1) from PKGIDSEQ_37 where val = '" & strBase & "' "
strseq = Get_OracleStr(strSql)
strNewID = strBase & Right$("0" & strseq, 2)
strSql = "insert into PKGIDSEQ_37(VAL,SEQ,CREATE_DATE,DN) values('" & strBase & "', '" & strseq & "', sysdate, '" & ut.DN_NUM & "')"
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
Private Function GetQID(ut As uReelInfo)
Dim strSql As String
Dim strQID As String
Dim strBID As String

strSql = "select BOXID from PACKING_DETAILED where dn_num = '" & ut.DN_NUM & "' and outbox_num = '" & ut.OUTBOX_NUM & "' and inbox_num = 1"
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
' Procedure  :       GetPSN
' Description:       获取PSN
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/9-9:26:58
'
' Parameters :       strDN (String)
'                    strJobID (String)
'--------------------------------------------------------------------------------
Private Function GetPSN(strDN As String, _
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

GetPSN = strPSN
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
Dim strWaferID  As String
Dim strDateCode As String
Dim strJobNew   As String
Dim strSql        As String
Dim strContent    As String
Dim str1          As String
Dim strBartenName As String

str1 = "37_FIRST_FINISH_YYWW_MON"
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
' Procedure  :       SaveReelInfo
' Description:       保存卷盘数据
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/9-16:19:42
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub SaveReelInfo(ut As uReelInfo)
Dim strSql As String

strSql = "insert into PACKING_DETAILED(TRAYID,INBOX_NUM,OUTBOX_NUM,DN_NUM,JOB_ID,QTY,CUSTOMER_DEVICE,CREATE_DATE,CREATE_BY,PRINT_FLAG,FLAG,KID,SEQ,DATECODE,REELID,CARTON,CARTONID,BOXID) " & " values('" & ut.TYAYID & "', '" & ut.INBOX_NUM & "','" & ut.OUTBOX_NUM & "', '" & ut.DN_NUM & "','" & ut.JOB_ID & "','" & ut.QTY & "','" & ut.Customer_Device & "', sysdate, '" & gUserName & "' ,'0','0','" & ut.KID & "','" & ut.SEQ & "', '" & ut.DATECODE & "','" & ut.REELID & "','" & ut.carton & "','" & ut.CARTONID & "','" & ut.BOXID & "')"
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

strDN = Trim(txtDN.Text)
lCurOutboxNum = CLng(Trim(txtCurOP.Text))
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

If lCurOutboxNum > lMaxOutboxNum Then
    MsgBox "标签已经全部打印完成", vbInformation, "提示"
    Exit Sub

End If

Call PrintLblByOutBoxNum(strDN, lCurOutboxNum)
MsgBox "第" & lCurOutboxNum & "箱标签已经全部打印完成", vbInformation, "提示"
lCurOutboxNum = lCurOutboxNum + 1
txtCurOP.Text = lCurOutboxNum

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
Call Print37BoxLbl_OLD(strDN, lOutboxNum, lInboxNum) '37内盒B标签
Call PrintHWBoxLbl_OLD(strDN, lOutboxNum, lInboxNum) '华为内盒标签
Call PrintHWReelLbl_OLD(strDN, lOutboxNum, lInboxNum) '华为卷盘标签
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
Private Sub Print37BoxLbl_OLD(strDN As String, lOutboxNum As Long, lInboxNum As Long)
Dim strSql      As String
Dim strTxt      As String
Dim strFlagTxt  As String
Dim StrFileName As String
Dim rsJobID     As New ADODB.Recordset
Dim tSTBox      As STBox

'标记
strTxt = "BOX_" & lOutboxNum & "_" & lInboxNum
StrFileName = strDN & "-" & "FLAG_BOX_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(StrFileName, strTxt, strFlagPath)
Call Sleep(gSleepMicSec)
strTxt = ""
'正式
strSql = "select JOB_ID,CUSTOMER_DEVICE,BOXID,DATECODE,SUM(QTY) as QTY from PACKING_DETAILED where DN_NUM = '" & strDN & "' and outbox_num = '" & lOutboxNum & "' and inbox_num =  '" & lInboxNum & "' group by JOB_ID,CUSTOMER_DEVICE,BOXID,DATECODE"
Set rsJobID = Get_OracleRs(strSql)
If Not rsJobID.BOF Then
    rsJobID.MoveFirst

    Do While Not rsJobID.EOF
        tSTBox.JOB = Trim("" & rsJobID!JOB_ID)
        tSTBox.DEV = Trim("" & rsJobID!Customer_Device)
        tSTBox.lot = Trim("" & rsJobID!BOXID)
        tSTBox.DATECODE = Trim$("" & rsJobID!DATECODE)
        tSTBox.QTY = rsJobID!QTY
        strTxt = strTxt & tSTBox.DEV & "," & tSTBox.JOB & ",1T" & tSTBox.JOB & "," & tSTBox.DEV & "," & "1P" & tSTBox.DEV & "," & tSTBox.DATECODE & "," & tSTBox.DATECODE & "," & Mid(tSTBox.lot, 2) & "," & tSTBox.lot & "," & tSTBox.QTY & ",Q" & tSTBox.QTY & ",," & GetDevMark(tSTBox.DEV) & vbCrLf
        rsJobID.MoveNext
    Loop

End If

Set rsJobID = Nothing
StrFileName = strDN & "-" & "BID" & "_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(StrFileName, strTxt, str37BCIDPath)
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
Private Sub PrintHWBoxLbl_OLD(strDN As String, lOutboxNum As Long, lInboxNum As Long)
Dim strTxt      As String
Dim strBarcode  As String
Dim strQrCode   As String
Dim StrFileName As String
Dim strSql      As String
Dim rsJobID     As New ADODB.Recordset
Dim tHWBox      As HWBox

'正式
strSql = "select job_id,mpn,cpn,datecode,sum(QTY) qty from LPSTBL where dn_num = '" & strDN & "' and outbox_num = '" & lOutboxNum & "' and inbox_num =  '" & lInboxNum & "' group by job_id,mpn,cpn,datecode"
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

StrFileName = strDN & "-" & "HWBoxLbl" & "_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(StrFileName, strTxt, strHWBoxPath)
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
Private Sub PrintHWReelLbl_OLD(strDN As String, lOutboxNum As Long, lInboxNum As Long)
Dim strTxt      As String
Dim strBarcode  As String
Dim strQrCode   As String
Dim StrFileName As String
Dim strSql      As String
Dim rsReel      As New ADODB.Recordset
Dim tHWBox      As HWBox

'标记
strTxt = "REEL_" & lOutboxNum & "_" & lInboxNum
StrFileName = strDN & "-" & "FLAG_REEL_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(StrFileName, strTxt, strFlagPath)
Call Sleep(gSleepMicSec)
strTxt = ""
'正式
strSql = "select job_id,mpn,cpn, QTY,datecode,reelid,seq from LPSTBL where dn_num = '" & strDN & "' and outbox_num = '" & lOutboxNum & "' and inbox_num =  '" & lInboxNum & "' order by seq"
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

StrFileName = strDN & "-" & "HWReelLbl" & "_" & lOutboxNum & "_" & lInboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(StrFileName, strTxt, strHWReelPath)
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
Call Print37CartonLbl_OLD(strDN, lOutboxNum)    '37外箱C标签
Call PrintHTCartonLbl_OLD(strDN, lOutboxNum)        '华天Q箱号
Call Print37CartonStanderLbl_OLD(strDN, lOutboxNum) '37外箱标准大标签

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
Private Sub Print37CartonLbl_OLD(strDN As String, lOutboxNum As Long)
Dim strSql        As String
Dim tSTCarton     As STCarton
Dim strTxt        As String
Dim StrFileName   As String
Dim sDatecode     As String
Dim sTestDateCode As String
Dim rsJobID       As New ADODB.Recordset
Dim sAdd          As String

'标记
strTxt = "CARTON_" & lOutboxNum
StrFileName = strDN & "-" & "FLAG_CARTON_" & lOutboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(StrFileName, strTxt, strFlagPath)
Call Sleep(gSleepMicSec)
strTxt = ""
'正式
strSql = "select JOB_ID,CUSTOMER_DEVICE,CARTONID,DATECODE,SUM(QTY) AS QTY from PACKING_DETAILED where dn_num = '" & strDN & "' and outbox_num = '" & lOutboxNum & "' group by JOB_ID,CUSTOMER_DEVICE,CARTONID,DATECODE"
Set rsJobID = Get_OracleRs(strSql)
If Not rsJobID.BOF Then
    rsJobID.MoveFirst

    Do While Not rsJobID.EOF
        tSTCarton.JOB = Trim("" & rsJobID!JOB_ID)
        tSTCarton.DEV = Trim$("" & rsJobID!Customer_Device)
        tSTCarton.lot = Trim("" & rsJobID!CARTONID)
        tSTCarton.DATECODE = Trim("" & rsJobID!DATECODE)
        tSTCarton.QTY = rsJobID!QTY
        strTxt = strTxt & tSTCarton.DEV & "," & tSTCarton.JOB & ",1T" & tSTCarton.JOB & "," & tSTCarton.DEV & "," & "1P" & tSTCarton.DEV & "," & tSTCarton.DATECODE & "," & tSTCarton.DATECODE & "," & Mid(tSTCarton.lot, 2) & "," & tSTCarton.lot & "," & tSTCarton.QTY & ",Q" & tSTCarton.QTY & "," & tSTCarton.testdateCode & "," & tSTCarton.testdateCode & GetDevMark(tSTCarton.DEV) & vbCrLf
        rsJobID.MoveNext
    Loop

End If

StrFileName = strDN & "-" & "CID" & "_" & lOutboxNum & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(StrFileName, strTxt, str37BCIDPath)
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
Dim StrFileName As String
Dim strTxt      As String

strSql = "select distinct carton from PACKING_DETAILED where dn_num = '" & strDN & "' and outbox_num = '" & lOutboxNum & "'"
strTxt = Get_OracleStr(strSql)
StrFileName = strDN & "-" & "QID_" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(StrFileName, strTxt, strHTQCartonPath)
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
Private Sub Print37CartonStanderLbl_OLD(strDN As String, lOutboxNum As Long)
Dim strSql      As String
Dim tCusCARTON  As CUSCARTON
Dim StrFileName As String
Dim strTxt      As String
Dim strKid      As String
Dim strMaxOP    As String
Dim strAdd      As String
Dim rs          As New ADODB.Recordset

strSql = "select max(OUTBOX_NUM) from PACKING_DETAILED where DN_NUM = '" & strDN & "'"
strMaxOP = Get_OracleStr(strSql)
strSql = "select a.CUSTOMER_DEVICE,a.kid, b.customerpartnumber,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & strDN & "' and b.delivery = '" & strDN & "' and a.job_id = b.batchnumber and a.outbox_num = '" & lOutboxNum & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno,a.kid"
Set rs = Get_OracleRs(strSql)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tCusCARTON.dn = strDN
       ' tCusCARTON.PO = UCase(IIf(IsNull(rs!PO), "N/A", Left(rs!PO, 10)))
        
        tCusCARTON.PO = UCase(IIf(IsNull(rs!PO), "N/A", rs!PO))
        tCusCARTON.CPN = UCase(IIf(IsNull(rs!CustomerPartnumber), "N/A", rs!CustomerPartnumber))
        tCusCARTON.MPN = UCase(IIf(IsNull(rs!Customer_Device), "N/A", rs!Customer_Device))
        tCusCARTON.KID = Trim("" & rs!KID)
        tCusCARTON.QTY = rs!QTY
        strTxt = strTxt & Get_OracleStr("select distinct substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3) || ','||trim(a.city) || ' ' || trim(a.state)  || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ',' || trim(a.contactname) || ',' || trim(a.phone) from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & strDN & "' ") & ","
        strTxt = strTxt & tCusCARTON.dn & ",I" & tCusCARTON.dn & "," & tCusCARTON.PO & ",K" & tCusCARTON.PO & "," & tCusCARTON.CPN & ",P" & tCusCARTON.CPN & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & "," & tCusCARTON.QTY & ",Q" & tCusCARTON.QTY & "," & Get_OracleStr("select distinct freightforwarder from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & strDN & "'") & "," & "" & "," & "" & "," & "" & "," & "COO:CHINA" & "," & "CHINA"
        strAdd = "," & lOutboxNum & "," & tCusCARTON.KID
        strTxt = strTxt & strAdd & "," & strMaxOP
        rs.MoveNext
    Loop

End If

StrFileName = strDN & "-" & "SemTechStanderCarton" + Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(StrFileName, strTxt, str37CartonPath)
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
Dim ID          As String

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

    ID = Get_SqlserverNo("select 序号 as ID from [erpdata].[dbo].[tblPackTreeInf] a where a.箱号='" & strCartonID & "' and Memo='37' ")
    strSql = "Update [erpdata].[dbo].[tblPackTreeInf] set 上级序号='" & ID & "',Memo='37' " & " where 箱号 in ( select trayid from  OPENQUERY(ORACLEDB, 'SELECT * from packing_detailed where carton = ''" & strCartonID & "'' ')) "
    If AddSql2(strSql) = 0 Then
        MsgBox "2 update [erpdata].[dbo].[tblPackTreeInf]:failed!!!", vbCritical, "提示"
        Exit Sub

    End If

    '3 insert - update [erpdata].[dbo].[tblStockNumTree]
    strSql = "insert into [erpdata].[dbo].[tblStockNumTree](序号,箱号,上级序号,基层标记 ,尺寸,重量,Memo,DN) values( " & ID & ",'" & strCartonID & "',0,1,'','','37','" & strDN & "')"
    If AddSql2(strSql) = 0 Then
        MsgBox "3 insert [erpdata].[dbo].[tblStockNumTree]", vbCritical, "提示"
        Exit Sub

    End If

    strSql = "Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='" & ID & "',Memo='37', dn='" & strDN & "' " & " where 箱号 in ( select trayid from  OPENQUERY(ORACLEDB, 'SELECT * from packing_detailed where carton = ''" & strCartonID & "'' ')) "
    If AddSql2(strSql) = 0 Then
        MsgBox "3 update [erpdata].[dbo].[tblStockNumTree]", vbCritical, "提示"
        Exit Sub

    End If

    rs.MoveNext
Loop
INIadoCon.CommitTrans
MsgBox "DN:" & strDN & "  :箱号已更新", vbInformation, "提示"
Exit Sub
ERRON:
INIadoCon.RollbackTrans
MsgBox "错误:" & Err.DESCRIPTION, vbCritical, "警告"

End Sub

'-------------------------------------------------------------
'<<<<<<<<<<<<<<<<<<补打标签>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'-------------------------------------------------------------
Private Sub Print2Handler()
Dim strKey As String

strKey = UCase(Trim(txtScan2.Text))
If Len(strKey) = 0 Then
    MsgBox "请输入需要补打的条码", vbInformation, "提示"
    Exit Sub

End If

Call printLblNew(strKey)
txtScan2.Text = ""

End Sub

Private Sub Command1_Click()
Dim strSql As String

If txtUser2.Text = txtUser.Text Then
    MsgBox "员工不可输入组长的工号", vbCritical, "提示"
    Exit Sub

End If

strSql = "select * from tblOperatorData r where  r.状态标记=1  and r.用户号='10354'and r.密码='" & Replace(Trim(txtPassWd.Text), "'", "") & "'"
If Get_SqlStr(strSql) = "" Then
    MsgBox "组长密码不正确", vbCritical, "提示"
    Exit Sub

End If

strSql = "select * from tblOperatorData r where  r.状态标记=1  and r.用户号='" & Trim(txtUser2.Text) & "'and r.密码='" & Replace(Trim(txtPassWd2.Text), "'", "") & "'"
If Get_SqlStr(strSql) = "" Then
    MsgBox "员工工号或者密码不正确", vbCritical, "提示"
    Exit Sub

End If

txtScan2.Visible = True

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

Private Sub txtScan2_KeyPress(KeyAscii As Integer)
Dim strScan As String

strScan = UCase(Trim(txtScan2.Text))
If KeyAscii <> vbKeyReturn Or Len(strScan) = 0 Then Exit Sub
Call printLblNew(strScan)
txtScan2.Text = ""

End Sub
