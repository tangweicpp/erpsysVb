VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_37LblPrint_ONELOT 
   BackColor       =   &H00C0C0C0&
   Caption         =   "SEMTECH-标签打印"
   ClientHeight    =   12870
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
   ScaleHeight     =   12870
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTTab0 
      Height          =   12015
      Left            =   -120
      TabIndex        =   0
      Top             =   840
      Width           =   23565
      _ExtentX        =   41566
      _ExtentY        =   21193
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
      TabPicture(0)   =   "Frm_37ToHW_LblPrint_ONELOT.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "m1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDN"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblPO"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblQTY"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblS"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblREELS"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblBOXS"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblCARTONS"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblJOBID"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "printRate"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ProgressBar1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "UpDown1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame3"
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
      Tab(0).Control(20)=   "chk"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtPrintInterval"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtReelQty"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtBoxQty"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtCartonQty"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtJobID"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtTimer"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "标签扫描补打"
      TabPicture(1)   =   "Frm_37ToHW_LblPrint_ONELOT.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPassWd2"
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(2)=   "txtPassWd"
      Tab(1).Control(3)=   "txtUser"
      Tab(1).Control(4)=   "txtUser2"
      Tab(1).Control(5)=   "cbLblType"
      Tab(1).Control(6)=   "txtScan2"
      Tab(1).Control(7)=   "Label266"
      Tab(1).Control(8)=   "Label1"
      Tab(1).Control(9)=   "lblType"
      Tab(1).Control(10)=   "lblBarcodeScan2"
      Tab(1).ControlCount=   11
      Begin VB.TextBox txtPassWd2 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -71760
         PasswordChar    =   "*"
         TabIndex        =   43
         Top             =   3360
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "验证补打密码"
         Height          =   840
         Left            =   -68760
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtPassWd 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -71760
         PasswordChar    =   "*"
         TabIndex        =   41
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox txtUser 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -73080
         TabIndex        =   40
         Text            =   "10354"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtUser2 
         Height          =   375
         Left            =   -73080
         TabIndex        =   39
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ComboBox cbLblType 
         Height          =   315
         ItemData        =   "Frm_37ToHW_LblPrint_ONELOT.frx":0038
         Left            =   -73560
         List            =   "Frm_37ToHW_LblPrint_ONELOT.frx":004B
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtScan2 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   -73560
         TabIndex        =   33
         Top             =   720
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.TextBox txtTimer 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1860
         Left            =   6360
         TabIndex        =   21
         Text            =   "60"
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtJobID 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtCartonQty 
         BackColor       =   &H00C0C0C0&
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1620
         Width           =   975
      End
      Begin VB.TextBox txtBoxQty 
         BackColor       =   &H00C0C0C0&
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtReelQty 
         BackColor       =   &H00C0C0C0&
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox txtPrintInterval 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "6"
         Top             =   60
         Width           =   300
      End
      Begin VB.CheckBox chk 
         Caption         =   "测试用"
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtTmpQty 
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0"
         Top             =   1620
         Width           =   735
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1620
         Width           =   735
      End
      Begin VB.TextBox txtPO 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtDN 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   1020
         Width           =   1455
      End
      Begin VB.TextBox txtScan 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   525
         Width           =   4575
      End
      Begin VB.Frame Frame3 
         Caption         =   "扫描状态"
         ForeColor       =   &H00FF0000&
         Height          =   9495
         Left            =   13800
         TabIndex        =   8
         Top             =   2220
         Width           =   9375
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
            Height          =   8295
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   480
            Width           =   9135
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "扫描信息"
         ForeColor       =   &H00FF0000&
         Height          =   9495
         Left            =   240
         TabIndex        =   2
         Top             =   2220
         Width           =   13455
         Begin FPSpreadADO.fpSpread fps 
            Height          =   8895
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   8535
            _Version        =   524288
            _ExtentX        =   15055
            _ExtentY        =   15690
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
            MaxRows         =   0
            SpreadDesigner  =   "Frm_37ToHW_LblPrint_ONELOT.frx":0094
            AppearanceStyle =   0
         End
         Begin FPSpreadADO.fpSpread fps 
            Height          =   2535
            Index           =   1
            Left            =   8760
            TabIndex        =   3
            Top             =   360
            Width           =   4575
            _Version        =   524288
            _ExtentX        =   8070
            _ExtentY        =   4471
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
            SpreadDesigner  =   "Frm_37ToHW_LblPrint_ONELOT.frx":04E0
            TextTip         =   2
            AppearanceStyle =   0
         End
         Begin FPSpreadADO.fpSpread fps 
            Height          =   5895
            Index           =   2
            Left            =   8760
            TabIndex        =   4
            Top             =   3360
            Width           =   4575
            _Version        =   524288
            _ExtentX        =   8070
            _ExtentY        =   10398
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
            SpreadDesigner  =   "Frm_37ToHW_LblPrint_ONELOT.frx":097C
            TextTip         =   2
            AppearanceStyle =   0
         End
         Begin VB.Label lblReelList 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "卷盘已扫描:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2520
            TabIndex        =   7
            Top             =   120
            Width           =   1125
         End
         Begin VB.Label lblMP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "机种已扫描:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   10320
            TabIndex        =   6
            Top             =   120
            Width           =   1125
         End
         Begin VB.Label lblJOBList 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "JOB已扫描:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   10320
            TabIndex        =   5
            Top             =   3120
            Width           =   1065
         End
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   10140
         TabIndex        =   22
         Top             =   60
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txtReelQty"
         BuddyDispid     =   196620
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   14880
         TabIndex        =   37
         Top             =   1320
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   873
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label266 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm_37ToHW_LblPrint_ONELOT.frx":0E18
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
         Left            =   -74760
         TabIndex        =   45
         Top             =   3405
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm_37ToHW_LblPrint_ONELOT.frx":0E2C
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
         Left            =   -74760
         TabIndex        =   44
         Top             =   2940
         Width           =   1635
      End
      Begin VB.Label printRate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "打印进度:"
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
         Left            =   13800
         TabIndex        =   38
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "补打类型:"
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
         Left            =   -74640
         TabIndex        =   36
         Top             =   1110
         Width           =   1035
      End
      Begin VB.Label lblBarcodeScan2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫码框:"
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
         Left            =   -74400
         TabIndex        =   35
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblJOBID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JOBID:"
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
         Left            =   780
         TabIndex        =   32
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblCARTONS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARTONS:"
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
         Left            =   4080
         TabIndex        =   31
         Top             =   1620
         Width           =   1035
      End
      Begin VB.Label lblBOXS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BOXS:"
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
         Left            =   4440
         TabIndex        =   30
         Top             =   1335
         Width           =   675
      End
      Begin VB.Label lblREELS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REELS:"
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
         Left            =   4380
         TabIndex        =   29
         Top             =   1020
         Width           =   615
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
         Left            =   8520
         TabIndex        =   28
         Top             =   135
         Width           =   1320
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
         Left            =   960
         TabIndex        =   27
         Top             =   1620
         Width           =   435
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
         Left            =   1035
         TabIndex        =   26
         Top             =   1320
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
         Left            =   1020
         TabIndex        =   25
         Top             =   1020
         Width           =   390
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫码框:"
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
         Left            =   600
         TabIndex        =   24
         Top             =   540
         Width           =   795
      End
      Begin WMPLibCtl.WindowsMediaPlayer m1 
         Height          =   495
         Left            =   15720
         TabIndex        =   23
         Top             =   480
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   1482
      ButtonWidth     =   2090
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "打印标签"
            Key             =   "PRINT"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "删除打印记录"
            Key             =   "DEL"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "导出打印记录"
            Key             =   "EXPORT"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出窗体"
            Key             =   "EXIT"
            ImageIndex      =   12
         EndProperty
      EndProperty
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   6840
         Top             =   0
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   7800
         Top             =   0
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
               Picture         =   "Frm_37ToHW_LblPrint_ONELOT.frx":0E40
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint_ONELOT.frx":2F7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint_ONELOT.frx":5E04
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint_ONELOT.frx":85B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint_ONELOT.frx":A6F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint_ONELOT.frx":CEA2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint_ONELOT.frx":F654
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint_ONELOT.frx":126D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint_ONELOT.frx":14E88
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint_ONELOT.frx":151A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint_ONELOT.frx":15E7C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint_ONELOT.frx":18EFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_37ToHW_LblPrint_ONELOT.frx":1B6B0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   7320
         Top             =   0
      End
   End
End
Attribute VB_Name = "Frm_37LblPrint_ONELOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilllMicroSeconds As Long)
Dim lMicroSec  As Long
Dim bLocked_IP As Boolean
Dim bLocked_OP As Boolean

Private Const MAXINBOXQTY = 12
Private Const MAXINREELQTY = 9
Dim strLastReelID    As String
Dim strLastJobID     As String
Dim strLastMPN       As String

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

Private Sub Form_Activate()
SSTTab0.Tab = 0
txtScan.SetFocus

End Sub

Private Sub Form_Load()
Call InitCtrls
Call InitData

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       initCtrls
' Description:       初始化控件状态
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-9:39:09
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitCtrls()

If gUserName = "07885" Then
    Toolbar1.Buttons(3).Enabled = True
    txtDN.Locked = False
    chk.Value = 1
ElseIf gUserName = "10354" Then
    Toolbar1.Buttons(3).Enabled = True
    txtDN.Locked = False

End If

Call InitFps
Call initMedia

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
    .DAutoCellTypes = True
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
    .ColWidth(E_REEL_SCAN.E_REEL_SCANTIME) = 12
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
    .TypeVAlign = TypeVAlignCenter
    .TypeHAlign = TypeVAlignCenter
    .SelForeColor = &HFF8080
    .SetText E_MPN_SCAN.E_MPN_ID, 0, "M.P.N"
    .SetText E_MPN_SCAN.E_MPN_TOTAL_QTY, 0, "总数量"
    .SetText E_MPN_SCAN.E_MPN_CUR_QTY, 0, "已扫描数量"
    .ColWidth(E_MPN_SCAN.E_MPN_ID) = 14
    .ColWidth(E_MPN_SCAN.E_MPN_TOTAL_QTY) = 8
    .ColWidth(E_MPN_SCAN.E_MPN_CUR_QTY) = 8
    .ReDraw = True

End With

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
    .TypeVAlign = TypeVAlignCenter
    .TypeHAlign = TypeVAlignCenter
    .SelForeColor = &HFF8080
    .SetText E_JOB_SCAN.E_JOB_ID, 0, "JOB"
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
' Procedure  :       initMedia
' Description:       初始化声音
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-9:58:27
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub initMedia()
Dim strDocCfgDir    As String
Dim strLocalDocDir  As String
Dim strRemoteDocDir As String
Dim strTemp         As String
Dim strArr()        As String

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

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       initData
' Description:       初始化变量
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-9:38:31
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitData()
lMicroSec = CLng(Trim(txtPrintInterval.Text)) * 500
bLocked_IP = False
bLocked_OP = False

If chk.Value = 1 Then
    Call setTestPrintPath
Else
    Call setPrintPath

End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       txtScan_KeyPress
' Description:       扫描入口
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-9:40:05
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub txtScan_KeyPress(KeyAscii As Integer)
Dim strScan As String

strScan = UCase(Trim(txtScan.Text))

If KeyAscii <> vbKeyReturn Or Len(strScan) = 0 Then Exit Sub
If Len(Trim(txtDN.Text)) = 0 Then
    Call scanDN(strScan)
Else
    Call scanReel(strScan)

End If

txtScan.Text = ""

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       scanDN
' Description:       扫描DN
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-10:13:09
'
' Parameters :       strScan (String)
'--------------------------------------------------------------------------------
Private Sub scanDN(strScan As String)
Dim strDN As String

'1.检查DN是否正确
If Left$(strScan, 1) <> "I" Or Len(strScan) <> 9 Then
    m1.url = getMediaUrl("DN未获取,请扫描DN条码")
    txtStatus.BackColor = vbRed
    Exit Sub

End If

strDN = Mid(strScan, 2)

If chkDNToHW_ONELOT(strDN) = False Then
    txtStatus.BackColor = vbRed
    Exit Sub

End If

'2.获取DN信息
Call getDNInfo(strDN)
'3.刷新扫描状态
Call updateScanInfo(strDN)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       getDNInfo
' Description:       获取DN信息
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-10:16:31
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub getDNInfo(strDN As String)
Dim strsql As String

txtDN.Text = strDN

strsql = "select distinct purchasingdocno from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "'"
txtPO.Text = Get_OracleStr(strsql)

strsql = "select sum(quantity) from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "'"
txtQty.Text = Get_OracleStr(strsql)

strsql = "select sum(quantity) / 15000 from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "'"
txtReelQty.Text = Get_OracleStr(strsql)

strsql = "select sum(trunc(((quantity / 15000) -1) / 9) +1)  from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "' "
txtBoxQty.Text = Get_OracleStr(strsql)

strsql = "select sum(AA.qty) from(select  a.marketingpn,trunc(((sum(a.quantity) / 15000) - 1) / 108 ) + 1 as qty from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & strDN & "' group by a.marketingpn) AA"
txtCartonQty.Text = Get_OracleStr(strsql)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       updateScanInfo
' Description:       更新扫描状态
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-10:45:34
'
' Parameters :       strDN (String)
'--------------------------------------------------------------------------------
Private Sub updateScanInfo(strDN As String)
Dim strsql           As String
Dim i                As Integer
Dim rs               As New ADODB.Recordset
Dim lScanQty         As Long
Dim lScanReels       As Long
Dim lLastJobCurQty   As Long
Dim lLastJobTotalQty As Long
Dim bLastJobEnough   As Boolean
Dim lLastMPNCurQty   As Long
Dim lLastMPNTotalQty As Long
Dim bLastMPNEnough   As Boolean
Dim iLastInBoxNO     As Integer
Dim iLastInBoxReels  As Integer

bLastJobEnough = False
bLastMPNEnough = False
'1.获取最后一个卷盘的REELID,JOBID,MPN
lScanReels = getScanReels(strDN)

If lScanReels <> 0 Then
    strLastReelID = getLastScanReelID(strDN)
    strLastJobID = getLastScanJOBID(strDN)
    strLastMPN = getLastScanMPN(strDN)
    lScanQty = getScanQty(strDN)
    iLastInBoxNO = getLastInBoxNO(strDN, strLastReelID)
    iLastInBoxReels = getLastInBoxReels(strDN, strLastReelID)
Else
    strLastReelID = ""
    strLastJobID = ""
    strLastMPN = ""
    lScanQty = 0
    iLastInBoxNO = 0
    iLastInBoxReels = 0

End If

'2.更新已扫描数量
txtTmpQty.Text = lScanQty
'3.更新已扫描卷盘数量
txtStatus.BackColor = vbWhite
txtStatus.Text = vbCrLf & lScanReels
'4.更新卷盘已扫描fps
strsql = "select outbox_num 外箱,inbox_num 内箱,trayid 卷盘ID,reelid PSN,job_id JOBID,seq 第几卷,create_date 扫描时间 from PACKING_DETAILED where dn_num = '" & strDN & "' order by seq desc"
Set rs = Get_OracleRs(strsql)

With fpS(0)
    .MaxRows = 0

    If rs.RecordCount > 0 Then
        Set .DataSource = rs

    End If

End With

If strLastJobID <> "" Then

    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_REEL_SCAN.E_REEL_JOBID

            If .Text = strLastJobID Then
                .BackColor = vbGreen
            Else
                .BackColor = vbWhite

            End If

        Next

    End With

End If

'5.更新机种已扫描fps
strsql = " select AA.marketingpn,AA.realqtys, BB.thisqtys from (select marketingpn, sum(quantity) as realqtys from CUSTOMERSHIPPINGUPTBL  where delivery = '" & strDN & "' group by marketingpn) AA left join (select customer_device, sum(qty) as thisqtys from PACKING_DETAILED where dn_num = '" & strDN & "' group by customer_device) BB on AA.marketingpn = BB.customer_device "
Set rs = Get_OracleRs(strsql)

With fpS(1)
    .MaxRows = 0

    If rs.RecordCount > 0 Then
        Set .DataSource = rs

    End If

End With

If strLastMPN <> "" Then

    With fpS(1)

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_MPN_SCAN.E_MPN_ID

            If .Text = strLastMPN Then
                .BackColor = vbGreen
                .Col = E_MPN_SCAN.E_MPN_CUR_QTY
                lLastMPNCurQty = Val(Trim$(.Text))
                .Col = E_MPN_SCAN.E_MPN_TOTAL_QTY
                lLastMPNTotalQty = Val(Trim$(.Text))

                If lLastMPNCurQty = lLastMPNTotalQty Then
                    bLastMPNEnough = True
                ElseIf lLastMPNCurQty < lLastMPNTotalQty Then
                    bLastMPNEnough = False
                Else
                    MsgBox "最后一个机种:" & strLastMPN & "的数量大于实际总数量,请确认是否有误", vbCritical, "警告"
                    txtStatus.BackColor = vbRed
                    txtScan.Enabled = False
                    Exit Sub

                End If

            Else
                .BackColor = vbWhite

            End If

        Next

    End With

End If

'6.更新JOBID已扫描fps
strsql = " select AA.batchnumber,AA.realqtys, BB.thisqtys from (select batchnumber, sum(quantity) as realqtys from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "' group by batchnumber) AA left join (select job_id, sum(qty) as thisqtys from PACKING_DETAILED where dn_num = '" & strDN & "' group by job_id) BB on AA.batchnumber = BB.job_id "
Set rs = Get_OracleRs(strsql)

With fpS(2)
    .MaxRows = 0

    If rs.RecordCount > 0 Then
        Set .DataSource = rs

    End If

End With

If strLastJobID <> "" Then

    With fpS(2)

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_JOB_SCAN.E_JOB_ID

            If .Text = strLastJobID Then
                .BackColor = vbGreen
                .Col = E_JOB_SCAN.E_JOB_CUR_QTY
                lLastJobCurQty = Val(Trim$(.Text))
                .Col = E_JOB_SCAN.E_JOB_TOTAL_QTY
                lLastJobTotalQty = Val(Trim$(.Text))

                If lLastJobCurQty = lLastJobTotalQty Then
                    bLastJobEnough = True
                ElseIf lLastJobCurQty < lLastJobTotalQty Then
                    bLastJobEnough = False
                Else
                    MsgBox "最后一个JOBID:" & strLastJobID & "的数量大于实际总数量,请确认是否有误", vbCritical, "警告"
                    txtStatus.BackColor = vbRed
                    txtScan.Enabled = False
                    Exit Sub

                End If

            Else
                .BackColor = vbWhite

            End If

        Next

    End With

End If

'打印内盒卷盘标签
If bLastJobEnough = True Then
   Call updatePackingBYJOBID(strDN, strLastJobID)
   Call printIPCase
End If

'打印外箱标签
If bLastMPNEnough = True Then
   Call updatePackingBYJOBID(strDN, strLastJobID)
   Call printOPCase
   Exit Sub
End If

If iLastInBoxNO = MAXINBOXQTY And bLastJobEnough = True Then
   Call updatePackingBYJOBID(strDN, strLastJobID)
   Call printOPCase
   Exit Sub
End If

If iLastInBoxNO = MAXINBOXQTY And iLastInBoxReels = MAXINREELQTY Then
    Call updatePackingBYJOBID(strDN, strLastJobID)
    Call printOPCase
    Exit Sub
End If

End Sub

Private Sub printIPCase()
m1.url = getMediaUrl("该JOB已经扫描完毕,请点击打印按钮")
MsgBox "该JOB已经扫描完毕,请点击打印按钮", vbInformation, "提示"
bLocked_IP = True
Toolbar1.Buttons(1).Enabled = True

End Sub

Private Sub printOPCase()
m1.url = getMediaUrl("该外箱已经扫描完毕,请点击打印按钮")
MsgBox "该外箱已经扫描完毕,请点击打印按钮", vbInformation, "提示"
bLocked_OP = True
Toolbar1.Buttons(1).Enabled = True

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       scanReel
' Description:       扫描卷盘
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-13:17:11
'
' Parameters :       strReelID (String)
'--------------------------------------------------------------------------------
Private Sub scanReel(strReelID As String)
Dim strsql    As String
Dim strDN     As String
Dim rs        As New ADODB.Recordset
Dim strMPN    As String
Dim strJobID  As String
Dim lREEL_Qty As Long
Dim lMaxQty   As Long
Dim i         As Integer

'If bLocked_IP = True Then
'    m1.url = getMediaUrl("当前Job已扫描完毕, 请勿扫描其他卷盘,请点击打印按钮")
'    Exit Sub
'
'End If
'
'If bLocked_OP = True Then
'    m1.url = getMediaUrl("当前外箱已扫描完毕, 请勿扫描其他卷盘,请点击打印按钮")
'    Exit Sub
'
'End If
strDN = Trim(txtDN.Text)
lMaxQty = CLng(txtQty.Text)

If Left$(strReelID, 1) <> "S" Or (Len(strReelID) <> 13 And Len(strReelID) <> 14) Then
    m1.url = getMediaUrl("卷盘号未获取,请扫描卷盘号")
    txtStatus.BackColor = vbRed
    Exit Sub

End If

'检查是否重复扫描
With fpS(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_REEL_SCAN.E_REEL_RID

        If strReelID = Trim$("" & .Text) Then
            m1.url = getMediaUrl("该卷盘已经扫描过, 请勿重复扫描")
            txtStatus.BackColor = vbRed
            Exit Sub

        End If

    Next

End With

'获取JOBID和卷盘数量
strJobID = getReel_JOBID_ONELOT(strReelID)

'检查JOBID
If checkJobID_ONELOT(strDN, strJobID) Then
    m1.url = getMediaUrl("卷盘号不正确")
    txtStatus.BackColor = vbRed
    Exit Sub
End If

txtJobID.Text = strJobID

'获取卷盘数量
lREEL_Qty = getReelQty_ONELOT(strReelID)

'获取卷盘机种
strMPN = getReelMPN_ONELOT(strDN, strJobID)

' Check
If chkReelID_ONELOT(strReelID, strDN, strJobID, strMPN) = False Then
    txtStatus.BackColor = vbRed
    Exit Sub

End If

Call insertReelID_ONELOT(strReelID, strDN, strJobID, strMPN, lREEL_Qty)
Call updateScanInfo(strDN)

End Sub

Private Sub chk_Click()

If chk.Value = 1 Then
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

Private Sub Timer1_Timer()

'm1.url = getMediaUrl("扫描已超时,请退出重新扫描")
'txtScan.Visible = False
End Sub

Private Sub Timer2_Timer()

If txtTimer.Text <> "0" Then
    txtTimer.Text = txtTimer.Text - 1

End If

End Sub

Private Sub txtPrintInterval_Change()
lMicroSec = CLng(txtPrintInterval.Text) * 500

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case SSTTab0.Tab

    Case 0

        Select Case Button.Key

            Case "PRINT"
                printDN

            Case "DEL"
                delDN

            Case "EXPORT"
                exptDN

            Case "EXIT"
                Unload Me

        End Select

    Case 1

        Select Case Button.Key

            Case "PRINT"
                printLbl2

            Case "EXPORT"
                exptDN2

            Case "EXIT"
                Unload Me

        End Select

    Case Else

End Select

End Sub

Private Sub exptDN2()
ExporToExcel ("select KEYNAME 补打类型,keyvalue 补打值,CREATE_DATE 补打时间,CREATE_BY 补打人员工号,CREATE_TIMES 第几次补打 from TBL_37_PRINT2_LIST order by CREATE_date desc")

End Sub

Private Sub exptDN()
Dim strDN  As String
Dim strsql As String

strDN = Trim$(txtDN.Text)

If Len(strDN) = 0 Then
    MsgBox "请输入要导出的DN", vbInformation, "提示"
    Exit Sub

End If

strsql = "select dn_num dn, OUTBOX_NUM 外箱,INBOX_NUM 内箱,trayid, reelid PSN,job_id job, customer_device 客户机种, QTY 数量, KID, DATECODE, CREATE_BY 打印人员, CREATE_DATE 打印时间,'' as 备注 from packing_detailed where dn_num = '" & strDN & "' order by seq  "
ExporToExcel (strsql)

End Sub

Private Sub delDN()
DialogDNDel.Show 1

End Sub

Private Sub printDN()
Dim strsql    As String
Dim rs        As New ADODB.Recordset
Dim i         As Integer
Dim lTotalQty As Long
Dim lCurQty   As Long
Dim strDN     As String
Dim strJob    As String

strDN = txtDN.Text
strJob = strLastJobID
Toolbar1.Buttons(1).Enabled = False
ProgressBar1.Value = 0

If bLocked_IP = True Then
    Call printIPkgLbl(strDN, strJob)
    bLocked_IP = False

End If

If bLocked_OP = True Then
    Call printOPkgLbl(strDN, strJob)
    bLocked_OP = False

End If

Toolbar1.Buttons(1).Enabled = True
ProgressBar1.Value = 100
'判断已打印数量
strsql = "select sum(quantity) from customershippinguptbl where delivery = '" & strDN & "'  "
lTotalQty = Get_OracleNo(strsql)
strsql = "select sum(qty) from packing_detailed where dn_num = '" & strDN & "' and print_flag = '1'"
lCurQty = Get_OracleNo(strsql)

If lCurQty = lTotalQty And lTotalQty <> 0 Then
    m1.url = getMediaUrl("该D N所有标签已全部打印完毕")
    MsgBox "该D N所有标签已全部打印完毕", vbInformation, "提示"
    Toolbar1.Buttons(1).Enabled = False
    txtScan.Visible = False
    Exit Sub

End If

End Sub

' 打印内箱小标签
Private Sub printIPkgLbl(strDN As String, strJob As String)
Dim strsql As String
Dim i      As Integer
Dim J      As Integer
Dim rs     As New ADODB.Recordset
Dim Rs2    As New ADODB.Recordset

Call PrintJobFlag(strDN, strJob)
strsql = "select distinct OUTBOX_NUM from PACKING_DETAILED where dn_NUM = '" & strDN & "' and JOB_ID = '" & strJob & "' order by OUTBOX_NUM "
Set rs = Get_OracleRs(strsql)

If Not rs.EOF Then

    Do While Not rs.EOF
        i = rs!OUTBOX_NUM
        strsql = "select distinct inbox_num from PACKING_DETAILED where dn_NUM = '" & strDN & "' and outbox_num = '" & i & "' and JOB_ID = '" & strJob & "' order by inbox_num "
        Set Rs2 = Get_OracleRs(strsql)

        If Not Rs2.EOF Then

            Do While Not Rs2.EOF
                J = Rs2!INBOX_NUM
                ' 内盒 开始
                Call PrintINNERBOXFlag(strDN, i, J)
                Sleep (lMicroSec)
                ' 37 内盒B标签
                Call PrintSTBoxLbl(strDN, i, J)
                Sleep (lMicroSec)
                ' 华为 内盒标签
                Call PrintHWBoxLbl(strDN, i, J)
                Sleep (lMicroSec)
                ' 卷盘 开始
                Call PrintREELFlag(strDN, i, J)
                Sleep (lMicroSec)
                ' 华为 卷盘PSN标签
                Call PrintHWReelLbl(strDN, i, J)
                Sleep (lMicroSec)
                AddSql ("update PACKING_DETAILED set print_flag = 1 where dn_num = '" & strDN & "' and outbox_num = '" & i & "' and inbox_num = '" & J & "'")
                
                Rs2.MoveNext
            Loop

        End If

        rs.MoveNext
    Loop

End If

MsgBox "JOBID:" & strJob & "已经打印完毕", vbInformation, "友情提示:"
End Sub

' 打印外箱标签
Private Sub printOPkgLbl(strDN As String, strJob As String)
Dim strsql As String
Dim i      As Integer
Dim iOpMax As Integer
Dim rs     As New ADODB.Recordset

iOpMax = Val(txtCartonQty.Text)
strsql = "select distinct OUTBOX_NUM from PACKING_DETAILED where dn_NUM = '" & strDN & "' and JOB_ID = '" & strJob & "' order by OUTBOX_NUM "
Set rs = Get_OracleRs(strsql)

If Not rs.EOF Then

    Do While Not rs.EOF
        i = rs!OUTBOX_NUM
        Call PrintOuterCFlag(strDN, i)
        Sleep (lMicroSec)
        ' 37 外箱C标签
        Call PrintSTCartonLbl(strDN, i)
        Sleep (lMicroSec)
        Call PrintHTCartonLbl(strDN, i)
        Sleep (lMicroSec)
        Call PrintSTCartonStanderLbl_ONELOT(strDN, i, iOpMax)
        MsgBox "外箱:" & i & "已经打印完毕", vbInformation, "友情提示:"
        rs.MoveNext
    Loop

End If

End Sub

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

Private Sub txtScan2_KeyPress(KeyAscii As Integer)
Dim strScan As String

strScan = UCase(Trim(txtScan2.Text))

If KeyAscii <> vbKeyReturn Or Len(strScan) = 0 Then Exit Sub
Call printLblNew(strScan)
txtScan2.Text = ""

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
