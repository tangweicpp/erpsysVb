VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form MyLabelSystemNew 
   BackColor       =   &H00C0C0C0&
   Caption         =   "LPS[标签打印系统]"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13380
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
   ScaleHeight     =   11055
   ScaleWidth      =   13380
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab DDD 
      Height          =   12255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   21975
      _ExtentX        =   38761
      _ExtentY        =   21616
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12632256
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
      TabPicture(0)   =   "MyLabelSystemNew.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDnCode"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblScanningFrame"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPurchaseOrder"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblQuantity"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblCusMaterial"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblReelList"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblStatus"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblPnInfo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblJobInfo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblShipTo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblSTBoxPath"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblCusBoxPath"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblCusReelPath"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblCusCartonPath"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblHTCartonPath"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblFlagPath"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblHWReelPath"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "media"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblMediaPath"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblSTCarton_PATH"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Line1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblDN222"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "fpsJob"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "fpsPN"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtSF"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtDN"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtPO"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtQty"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtPN"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtReelID"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtStatus"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "D"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtSTBoxPath"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtCusBoxPath"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtCusReelPath"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtCusCartonPath"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtHTCartonPath"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtFlagPath"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtShip"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtHWReelPath"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtMediaPath"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtSEMTCartonPAth"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cmdDNClear"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtCleanDNINfo"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cmdUpdateBoxNo"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "cmdExport"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).ControlCount=   46
      TabCaption(1)   =   "标签补打"
      TabPicture(1)   =   "MyLabelSystemNew.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblType"
      Tab(1).Control(1)=   "lbl"
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(3)=   "Label266"
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(5)=   "cbType"
      Tab(1).Control(6)=   "tReferTo"
      Tab(1).Control(7)=   "cmdPrint"
      Tab(1).Control(8)=   "chk"
      Tab(1).Control(9)=   "Command2"
      Tab(1).Control(10)=   "txtUser2"
      Tab(1).Control(11)=   "txtUser"
      Tab(1).Control(12)=   "txtPassWd"
      Tab(1).Control(13)=   "Command1"
      Tab(1).Control(14)=   "txtPassWd2"
      Tab(1).Control(15)=   "txtDN2"
      Tab(1).ControlCount=   16
      Begin VB.TextBox txtDN2 
         Height          =   285
         Left            =   -72720
         TabIndex        =   70
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtPassWd2 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -71160
         PasswordChar    =   "*"
         TabIndex        =   66
         Top             =   4200
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "验证补打密码"
         Height          =   840
         Left            =   -68160
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox txtPassWd 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -71160
         PasswordChar    =   "*"
         TabIndex        =   64
         Top             =   3720
         Width           =   2895
      End
      Begin VB.TextBox txtUser 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -72480
         TabIndex        =   63
         Text            =   "10354"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtUser2 
         Height          =   375
         Left            =   -72480
         TabIndex        =   62
         Top             =   4200
         Width           =   1215
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
         Height          =   465
         Left            =   -67680
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "导出DN信息"
         Height          =   360
         Left            =   6000
         TabIndex        =   60
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdateBoxNo 
         Caption         =   "箱号不存在"
         Height          =   360
         Left            =   8040
         TabIndex        =   59
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtCleanDNINfo 
         Height          =   285
         Left            =   5880
         TabIndex        =   56
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdDNClear 
         Caption         =   "清除DN信息"
         Height          =   360
         Left            =   6000
         TabIndex        =   55
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtSEMTCartonPAth 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         TabIndex        =   53
         Text            =   "\\10.160.1.84\public\BarCode\37\37外箱\"
         Top             =   8040
         Width           =   3375
      End
      Begin VB.CheckBox chk 
         Caption         =   "手动输入"
         Height          =   375
         Left            =   -69120
         TabIndex        =   52
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FF80FF&
         Caption         =   "手动编辑补打"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -67800
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox tReferTo 
         Height          =   285
         Left            =   -72720
         TabIndex        =   50
         Top             =   1905
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ComboBox cbType 
         Height          =   315
         ItemData        =   "MyLabelSystemNew.frx":0038
         Left            =   -72720
         List            =   "MyLabelSystemNew.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtMediaPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         TabIndex        =   45
         Text            =   "C:\media_source\"
         Top             =   7200
         Width           =   3375
      End
      Begin VB.TextBox txtHWReelPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         TabIndex        =   43
         Text            =   "\\10.160.1.84\public\BarCode\37\37HW\HW卷盘\"
         Top             =   6360
         Width           =   3735
      End
      Begin VB.TextBox txtShip 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1478
         Width           =   2055
      End
      Begin VB.TextBox txtFlagPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         TabIndex        =   40
         Text            =   "\\10.160.1.84\public\BarCode\37\37Flag\"
         Top             =   5580
         Width           =   3375
      End
      Begin VB.TextBox txtHTCartonPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         TabIndex        =   38
         Text            =   "\\10.160.1.84\public\BarCode\37\37Box\"
         Top             =   4740
         Width           =   3375
      End
      Begin VB.TextBox txtCusCartonPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         TabIndex        =   36
         Text            =   "\\10.160.1.84\public\BarCode\37\37BoxOut\"
         Top             =   3900
         Width           =   3375
      End
      Begin VB.TextBox txtCusReelPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         TabIndex        =   30
         Text            =   "\\10.160.1.84\public\BarCode\37\37BoxJP\"
         Top             =   3156
         Width           =   3375
      End
      Begin VB.TextBox txtCusBoxPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         TabIndex        =   29
         Text            =   "\\10.160.1.84\public\BarCode\37\37BoxNH\"
         Top             =   2388
         Width           =   3375
      End
      Begin VB.TextBox txtSTBoxPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         TabIndex        =   28
         Text            =   "\\10.160.1.84\public\BarCode\37\37内箱\"
         Top             =   1620
         Width           =   3375
      End
      Begin VB.Frame D 
         Height          =   1335
         Left            =   360
         TabIndex        =   19
         Top             =   10800
         Width           =   19575
         Begin VB.CommandButton cmdTest 
            Caption         =   "Test"
            Height          =   480
            Left            =   17640
            TabIndex        =   58
            Top             =   480
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtPrintInterval 
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
            Left            =   360
            TabIndex        =   26
            Text            =   "2"
            Top             =   518
            Width           =   300
         End
         Begin VB.CommandButton cmdCombine 
            BackColor       =   &H0080FFFF&
            Caption         =   "合箱"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   480
            Width           =   2055
         End
         Begin VB.CommandButton cmdPrintReel_Box 
            BackColor       =   &H008080FF&
            Caption         =   "打印卷盘/内盒"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   6210
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   480
            Width           =   2055
         End
         Begin VB.CommandButton cmdPrintCarton 
            BackColor       =   &H0080FF80&
            Caption         =   "打印外箱标签"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   9060
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   480
            Width           =   2055
         End
         Begin VB.CommandButton cmdReset 
            BackColor       =   &H00FF8080&
            Caption         =   "重置窗体"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   11910
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   480
            Width           =   2055
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF80FF&
            Caption         =   "退出窗体"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   14760
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   480
            Width           =   2055
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   405
            Left            =   661
            TabIndex        =   25
            Top             =   518
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   714
            _Version        =   393216
            Value           =   2
            BuddyControl    =   "txtPrintInterval"
            BuddyDispid     =   196636
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
         Begin VB.Label lblPrintInternvl 
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
            Left            =   1080
            TabIndex        =   27
            Top             =   600
            Width           =   1320
         End
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
         Height          =   5775
         Left            =   7440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   4980
         Width           =   9855
      End
      Begin VB.TextBox txtReelID 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   5535
         Left            =   2280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   5220
         Width           =   2175
      End
      Begin VB.TextBox txtPN 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   2280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3660
         Width           =   3735
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3060
         Width           =   2055
      End
      Begin VB.TextBox txtPO 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtDN 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   1980
         Width           =   2055
      End
      Begin VB.TextBox txtSF 
         BackColor       =   &H0080FFFF&
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   1020
         Width           =   3495
      End
      Begin FPSpreadADO.fpSpread fpsPN 
         Height          =   3375
         Left            =   7440
         TabIndex        =   11
         Top             =   1500
         Width           =   4935
         _Version        =   524288
         _ExtentX        =   8705
         _ExtentY        =   5953
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
         SpreadDesigner  =   "MyLabelSystemNew.frx":003C
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread fpsJob 
         Height          =   3375
         Left            =   12360
         TabIndex        =   12
         Top             =   1500
         Width           =   4935
         _Version        =   524288
         _ExtentX        =   8705
         _ExtentY        =   5953
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
         SpreadDesigner  =   "MyLabelSystemNew.frx":04D6
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "补打DN:"
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
         Left            =   -73560
         TabIndex        =   69
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label Label266 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"MyLabelSystemNew.frx":0970
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
         Left            =   -74160
         TabIndex        =   68
         Top             =   4245
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"MyLabelSystemNew.frx":0984
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
         Left            =   -74160
         TabIndex        =   67
         Top             =   3780
         Width           =   1635
      End
      Begin VB.Label lblDN222 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN号码:"
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
         Left            =   4920
         TabIndex        =   57
         Top             =   600
         Width           =   810
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   6720
         X2              =   6720
         Y1              =   1440
         Y2              =   840
      End
      Begin VB.Label lblSTCarton_PATH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STCarton_PATH:"
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
         Left            =   18000
         TabIndex        =   54
         Top             =   7800
         Width           =   1590
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参考值:"
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
         Left            =   -73560
         TabIndex        =   49
         Top             =   1920
         Width           =   795
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
         Left            =   -73800
         TabIndex        =   47
         Top             =   1230
         Width           =   1035
      End
      Begin VB.Label lblMediaPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MediaPath:"
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
         Left            =   18000
         TabIndex        =   46
         Top             =   6960
         Width           =   1095
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   18000
         TabIndex        =   44
         Top             =   8520
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
      Begin VB.Label lblHWReelPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HWReelPath:"
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
         Left            =   18000
         TabIndex        =   42
         Top             =   6120
         Width           =   1275
      End
      Begin VB.Label lblFlagPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FlagPath:"
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
         Left            =   18000
         TabIndex        =   39
         Top             =   5340
         Width           =   900
      End
      Begin VB.Label lblHTCartonPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HTCartonPath:"
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
         Left            =   18000
         TabIndex        =   37
         Top             =   4500
         Width           =   1425
      End
      Begin VB.Label lblCusCartonPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CusCartonPath:"
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
         Left            =   18000
         TabIndex        =   35
         Top             =   3660
         Width           =   1530
      End
      Begin VB.Label lblCusReelPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CusReelPath:"
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
         Left            =   18000
         TabIndex        =   34
         Top             =   2940
         Width           =   1290
      End
      Begin VB.Label lblCusBoxPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CusBoxPath:"
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
         Left            =   18000
         TabIndex        =   33
         Top             =   2100
         Width           =   1215
      End
      Begin VB.Label lblSTBoxPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STBox_Carton_Path:"
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
         Left            =   18000
         TabIndex        =   32
         Top             =   1380
         Width           =   1995
      End
      Begin VB.Label lblShipTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To:"
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
         Left            =   1425
         TabIndex        =   31
         Top             =   1500
         Width           =   765
      End
      Begin VB.Label lblJobInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Info:"
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
         Left            =   14580
         TabIndex        =   18
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label lblPnInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MPN Info:"
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
         Left            =   9345
         TabIndex        =   17
         Top             =   1140
         Width           =   930
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   6600
         TabIndex        =   16
         Top             =   8287
         Width           =   720
      End
      Begin VB.Label lblReelList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reel List:"
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
         Left            =   1290
         TabIndex        =   15
         Top             =   8287
         Width           =   900
      End
      Begin VB.Label lblCusMaterial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus& Material P/N:"
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
         Left            =   495
         TabIndex        =   9
         Top             =   4147
         Width           =   1695
      End
      Begin VB.Label lblQuantity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
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
         Left            =   1275
         TabIndex        =   7
         Top             =   3060
         Width           =   915
      End
      Begin VB.Label lblPurchaseOrder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Order:"
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
         Left            =   600
         TabIndex        =   5
         Top             =   2542
         Width           =   1590
      End
      Begin VB.Label lblScanningFrame 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scanning Frame:"
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
         Left            =   600
         TabIndex        =   4
         Top             =   1020
         Width           =   1590
      End
      Begin VB.Label lblDnCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dn Code:"
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
         Left            =   1320
         TabIndex        =   2
         Top             =   1980
         Width           =   870
      End
   End
End
Attribute VB_Name = "MyLabelSystemNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim glReelListCnt As Long
Dim gsLastPN      As String
Dim gsLastJob     As String
Dim bAdmin        As Boolean
Dim bCheckDC      As Boolean


Private Sub chk_Click()
If chk.Value = 1 Then
    cmdPrint.Visible = True
Else
    cmdPrint.Visible = False

End If

End Sub

Private Sub cmdDNClear_Click()
Dim strDN As String

If txtCleanDNINfo.Text = "" Then
    MsgBox "请填入DN号", vbInformation, "提示"
    Exit Sub

End If

strDN = Trim$(txtCleanDNINfo.Text)
If MsgBox("你确认要删除吗?", vbOKCancel, "提示") = vbCancel Then
    Exit Sub

End If

Call DelToErp(strDN)
AddSql ("insert into packing_detailed_bak select * from packing_detailed where dn_num = '" & strDN & "' ")
AddSql ("delete from packing_detailed where dn_num = '" & strDN & "' ")
AddSql2 ("delete from erpbase..packing_detailed where dn_num = '" & strDN & "' ")
MsgBox "备份删除完成", vbInformation, "提示"

End Sub

Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub cmdExit1_Click()
Unload Me

End Sub

Private Sub cmdExport_Click()
Dim strDN  As String
Dim strSql As String

strDN = Trim$(txtCleanDNINfo.Text)
If Len(strDN) = 0 Then
    MsgBox "请输入要导出的DN", vbInformation, "提示"
    Exit Sub

End If

strSql = "select dn_num dn, trayid, reelid,job_id job, customer_device 客户机种, QTY 数量, KID, DATECODE, CREATE_BY 打印人员, CREATE_DATE 打印时间,'' as 备注 from packing_detailed where dn_num = '" & strDN & "' order by seq desc"
ExporToExcel (strSql)

End Sub

Private Sub cmdReset_Click()
Unload Me
MyLabelSystem.Show

End Sub

Private Sub cmdTest_Click()
txtSTBoxPath.Text = "C:\Users\ks015918\Desktop\test\"
txtCusBoxPath.Text = "C:\Users\ks015918\Desktop\test\"
txtCusReelPath.Text = "C:\Users\ks015918\Desktop\test\"
txtCusCartonPath.Text = "C:\Users\ks015918\Desktop\test\"
txtHTCartonPath.Text = "C:\Users\ks015918\Desktop\test\"
txtFlagPath.Text = "C:\Users\ks015918\Desktop\test\"
txtHWReelPath.Text = "C:\Users\ks015918\Desktop\test\"
txtMediaPath.Text = "\\10.160.1.84\public\media_source\"
cmdCombine.Enabled = True
cmdPrintReel_Box.Enabled = True
cmdPrintCarton.Enabled = True

End Sub

Private Sub cmdUpdateBoxNo_Click()
Dim strDN As String

If txtCleanDNINfo.Text = "" Then
    MsgBox "请输入DN", vbInformation, "提示"
    Exit Sub

End If

strDN = Trim$(txtCleanDNINfo.Text)
TransToErp (strDN)

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

tReferTo.Visible = True
MsgBox "密码输入正确,可以补打", vbInformation, "提示"

End Sub

Private Sub Command2_Click()
ExporToExcel ("select * from TBL_37_PRINT2_LIST order by KEYNAME, KEYVALUE, CREATE_TIMES desc")

End Sub

Private Sub Form_Activate()
txtSF.SetFocus

End Sub

Private Sub Form_Load()
Call InitData
Call InitFps
Call InitStatus
bAdmin = False

End Sub

Private Sub InitShip()
Dim sType As String
Dim sOra  As String
Dim sDN   As String

sDN = Trim(txtDN.Text)
If sDN = "" Then
    Exit Sub

End If

sOra = "select UPPER(labelrequirement) as type from CUSTOMERSHIPPINGUPTBL where delivery = '" & sDN & "'"
sType = Get_OracleStr(sOra)
If InStr(sType, "E2") Then
    txtShip.Text = "SSE2"

End If

If InStr(sType, "HUAWEI") Then
    txtShip.Text = "HW"

End If

If InStr(sType, "SEMTECH") Then
    txtShip.Text = "ST"

End If

If InStr(sType, "SHORT") Then
    txtShip.Text = "SSSHORT"

End If

End Sub

Private Sub InitData()
' 标签打印
If gUserName = "07885" Then
    cmdTest.Visible = True

End If

glReelListCnt = 0
gsLastPN = ""
gsLastJob = ""
If gUserName = "07885" Or gUserName = "10354" Then
    cmdPrint.Enabled = True
    cmdDNClear.Visible = True

End If

' 标签补打
cbType.AddItem ("按内盒补打")
cbType.AddItem ("三星卷盘标签")
cbType.AddItem ("Semtech内盒标签")
cbType.AddItem ("三星内盒标签")
cbType.AddItem ("Semtech外箱分标签")
cbType.AddItem ("外箱大标签补打")
cbType.ListIndex = 0

End Sub

Private Sub DelTmpTbl()

End Sub

Private Sub InitFps()

With fpsPN
    .ReDraw = False
    .MaxCols = 3
    .MaxRows = 0
    .FontBold = True
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .TypeVAlign = TypeVAlignCenter
    .SelForeColor = &HFF8080
    .SetText 1, 0, "MPN"
    .SetText 2, 0, "目标数量"
    .SetText 3, 0, "累计数量"
    .ColWidth(1) = 14
    .ColWidth(2) = 10
    .ColWidth(3) = 10
    .RowHeight(0) = 20
    .RowHeight(-1) = 15
    .ReDraw = True

End With

With fpsJob
    .ReDraw = False
    .MaxCols = 3
    .MaxRows = 0
    .FontBold = True
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

Private Sub InitStatus()
txtStatus.Text = vbCrLf & glReelListCnt

End Sub

Private Sub tReferTo_KeyPress(KeyAscii As Integer)
Dim iQty   As Integer
Dim sRefer As String
Dim sType  As String

If KeyAscii <> vbKeyReturn Then
    Exit Sub

End If

If tReferTo.Text = "" Then
    MsgBox "请扫描参考值", vbInformation, "提示:"
    Exit Sub

End If

sRefer = UCase(Trim$(tReferTo.Text))
sType = cbType.Text

Select Case sType

    Case "按内盒补打"
        Call PrintA_Inner(sRefer)

    Case "三星卷盘标签"
        Call PrintA_SSREEL(sRefer)

    Case "Semtech内盒标签"
        Call PrintA_STBOX(sRefer)

    Case "三星内盒标签"
        Call PrintA_SSBOX(sRefer)

    Case "Semtech外箱分标签"
        Call PrintA_STCARTONSUB(sRefer)

    Case "外箱大标签补打"
        Call PrintAOut

End Select

iQty = Get_OracleStr("select nvl(count(*) + 1, 1) from TBL_37_PRINT2_LIST where KEYNAME = '" & sType & "' and KEYVALUE = '" & sRefer & "'")
AddSql ("insert into TBL_37_PRINT2_LIST(KEYNAME,KEYVALUE,CREATE_DATE,CREATE_BY,CREATE_TIMES) values('" & sType & "', '" & sRefer & "', sysdate, '" & gUserName & "', '" & iQty & "')")
tReferTo.Text = ""

End Sub

Private Sub txtSF_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Then
    Exit Sub

End If

If gUserName = "07885" And bAdmin = False Then
    If MsgBox("是否进行特权打印", vbYesNo, "提示") = vbYes Then
        bAdmin = True

    End If

End If

If bAdmin = True Then
    Call ForDN2
Else
    If txtDN.Text = "" Then
        Call ForDN
    Else
        Call ForReelID

    End If

End If

txtSF.Text = ""
txtSF.SetFocus

End Sub

Rem: Scanning DN
Private Sub ForDN()
Dim sScan    As String
Dim sKeyWord As String
Dim sDN      As String

sScan = UCase(Trim(txtSF.Text))
sKeyWord = Left$(sScan, 1)
If sKeyWord <> "I" Then
    Play ("noDN")
    Exit Sub

End If

sDN = Mid$(sScan, 2)
If CheckDN(sDN) = False Then
    Play ("wrongDN")
    Exit Sub

End If

Dim sOra As String

sOra = "DELETE FROM ST_TR_SEQ where dn = '" & sDN & "' "
AddSql (sOra)
txtDN.Text = sDN
Play ("DN_OK")
Call InitShip
Call ShowDNInfo(sDN)
Call ShowFps(sDN)

End Sub

Private Sub ForDN2()
Dim sScan    As String
Dim sKeyWord As String
Dim sDN      As String

sScan = UCase(Trim(txtSF.Text))
sKeyWord = Left$(sScan, 1)
If sKeyWord <> "I" Then
    Play ("noDN")
    Exit Sub

End If

sDN = Mid$(sScan, 2)
Dim sOra As String

txtDN.Text = sDN
Play ("DN_OK")
Call InitShip
cmdCombine.Enabled = True
cmdPrintReel_Box.Enabled = True
cmdPrintCarton.Enabled = True

End Sub

Public Function CheckDN(sDN As String) As Boolean
CheckDN = False
If IsExist(sDN) = False Then
    'MsgBox "DN: " & sDN & " 不存在", vbInformation, "友情提示!!!"
    Exit Function

End If

If IsPrint(sDN) = True Then
    'MsgBox "DN: " & sDN & " 已经打印,不可再次打印", vbInformatio, "友情提示!!!"
    Exit Function

End If

If IsDNCreated(sDN) = True Then
    Exit Function

End If

CheckDN = True

End Function

Private Function IsExist(sDN As String) As Boolean
Dim sOra As String

sOra = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & sDN & "'"
IsExist = IsOraRecord(sOra)

End Function

Private Function IsDNCreated(sDN As String) As Boolean
Dim sOra As String

sOra = "select * from packing_detailed where dn_num = '" & sDN & "'"
If Get_OracleStr(sOra) <> "" Then
    MsgBox "该DN有历史扫描记录, 请确认是否要删除旧数据,否则不可再次扫描该DN", vbInformation, "提示"
    IsDNCreated = True
    Exit Function
Else
    IsDNCreated = False

End If

End Function

Private Function IsPrint(sDN As String) As Boolean
Dim sOra As String

sOra = "select * from LPS_PRINTHISTORY where dn = '" & sDN & "'"
IsPrint = IsOraRecord(sOra)

End Function

Private Sub ShowDNInfo(sDN As String)
Dim sOra As String
Dim rs   As New ADODB.Recordset

sOra = "select trim(nvl(purchasingdocno, 'NULL')) as po,sum(nvl(quantity, 0)) as qty from CUSTOMERSHIPPINGUPTBL where delivery = '" & sDN & "' group by purchasingdocno"
Set rs = Get_OracleRs(sOra)
txtDN.Text = "" & sDN
txtPO.Text = "" & rs!PO
txtQty.Text = "" & rs!QTY
txtpn.Text = ""
sOra = "select distinct trim(nvl(marketingpn, 'NULL')) as mpn, trim(nvl(customerpartnumber, 'NULL')) as cpn from CUSTOMERSHIPPINGUPTBL a where delivery  = '" & sDN & "'"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        txtpn.Text = txtpn.Text & rs!CPN & "_" & rs!MPN & vbCrLf
        rs.MoveNext
    Loop

End If

rs.Close

End Sub

Private Sub ShowFps(sDN As String)
Dim sOra As String
Dim rs   As New ADODB.Recordset

sOra = " select AA.marketingpn,AA.realqtys, BB.thisqtys from " & "  (select marketingpn, sum(quantity) as realqtys " & "       from CUSTOMERSHIPPINGUPTBL " & "       where delivery = '" & sDN & "' " & "       group by marketingpn) AA  " & "  left join (select dev, sum(qty) as thisqtys " & "       from ST_TR_SEQ " & "       where dn = '" & sDN & "' " & "       group by dev) BB " & "       on AA.marketingpn = BB.dev "
Set rs = Get_OracleRs(sOra)

With fpsPN
    .MaxRows = 0
    If rs.RecordCount > 0 Then
        Set .DataSource = rs

    End If

End With

sOra = " select AA.batchnumber,AA.realqtys, BB.thisqtys from " & "  (select batchnumber, sum(quantity) as realqtys " & "       from CUSTOMERSHIPPINGUPTBL " & "       where delivery = '" & sDN & "' " & "       group by batchnumber) AA  " & "  left join (select job, sum(qty) as thisqtys " & "       from ST_TR_SEQ " & "       where dn = '" & sDN & "' " & "       group by job) BB " & "       on AA.batchnumber = BB.job "
Set rs = Get_OracleRs(sOra)

With fpsJob
    .MaxRows = 0
    If rs.RecordCount > 0 Then
        Set .DataSource = rs

    End If

End With

rs.Close

End Sub

Private Sub ForReelID()
Dim sScan    As String
Dim sKeyWord As String
Dim sReelID  As String
Dim sDN      As String
Dim strDC    As String

sScan = UCase(Trim$(txtSF.Text))
sKeyWord = Left$(sScan, 1)
If sKeyWord <> "S" Then
    Play ("noReel")
    Exit Sub

End If

sReelID = sScan
sDN = txtDN.Text
If CheckReelID(sReelID) = False Then
    Exit Sub

End If

If InsertReelIDToTmpTbl(sReelID) = False Then
    Exit Sub

End If

Play ("Reelright")

Call ShowFps(sDN)
Call CheckCurrentQty(sReelID)
Call UpdateCurrentStatus

End Sub

Private Function CheckReelID(sReelID As String) As Boolean
CheckReelID = False
If IsInErp(sReelID) = False Then
    MsgBox "该卷盘ID: " & sReelID & "不存在于ERP仓库", vbInformation, "友情提示!!!"
    Exit Function

End If

If IsErrSite(sReelID) = True Then
    MsgBox "该卷盘ID: " & sReelID & "在6000,6001仓库, 不可合箱", vbInformation, "友情提示!!！"
    Play ("wrongReel")
    Exit Function

End If

If IsInBox(sReelID) = True Then
    MsgBox "该卷盘ID: " & sReelID & "已经合过箱", vbInformation, "友情提示!!!"
    Play ("wrongReel")
    Exit Function

End If

If IsRightRId(sReelID) = False Then
    MsgBox "该卷盘Job和DN不匹配", vbInformation, "友情提示!!!"
    Play ("wrongReel")
    Exit Function

End If

If Get_OracleCnt("select * from ST_TR_SEQ where reelid = '" & sReelID & "'") Then
    MsgBox "请勿重复扫描", vbInformation, "警告"
    Exit Function

End If

CheckReelID = True

End Function

Private Function IsInErp(sReelID As String) As Boolean
Dim sSql As String

sSql = "SELECT * FROM  [erpdata].[dbo].[tblstocknumsub] WHERE  箱号= '" & sReelID & "'  "
IsInErp = IsSqlRecord(sSql)

End Function

Private Function IsErrSite(sReelID As String) As Boolean
Dim sSql As String

sSql = "SELECT * FROM  [erpdata].[dbo].[tblstocknumsub] WHERE  箱号= '" & sReelID & "' and 库房编号 in (36,37)"
IsErrSite = IsSqlRecord(sSql)

End Function

Private Function IsInBox(sReelID As String) As Boolean
Dim sOra As String

sOra = "SELECT * FROM PACKING_DETAILED where TRAYID = '" & sReelID & "'"
IsInBox = IsOraRecord(sOra)

End Function

Private Function IsRightRId(sReelID As String) As Boolean
Dim sSql As String
Dim sJob As String
Dim sOra As String

If Len(sReelID) = 13 Or Len(sReelID) = 14 Then
    sSql = "SELECT y.TEST_MTRL_DESC + case  SUBSTRING(KEY_VALUE,CHARINDEX('M-',KEY_VALUE,1),1) when 'M' then 'M' else '' end " & " FROM erpdata..tblErpInStockDetailInfo c ,erpdata..tblErpInStockRelation b  " & " ,ERPBASE..tblmappingData x,tblCustomerOI y  " & " where c.KEY_NAME ='CONTAINER_NAME'  " & " and charindex('" & sReelID & "', KEY_VALUE) > 0 " & " and b.BOX_ID = c.BOX_ID  " & " and SUBSTRING(c.KEY_VALUE,2,8) = SUBSTRING(REPLACE(B.SFC_ID,'SFCBO:1020,',''),1,8)  " & " and SUBSTRING( replace(b.WAFER_ID,b.SFC_ID+',',''),1,CHARINDEX('::',replace(b.WAFER_ID,b.SFC_ID+',',''),1)-1) = x.SUBSTRATEID  " & " and convert(varchar(100),y.id) = x.FILENAME"
Else
    sSql = "select isnull(CUSTOMERLOTID, 'null') as jobid from [erpdata].[dbo].[TblTSV_Tray_details] where TRAYQBOXNUMBER = '" & sReelID & "'"

End If

sJob = Trim(Get_SqlStr(sSql))
sOra = "select * from CUSTOMERSHIPPINGUPTBL where batchnumber = '" & sJob & "' and delivery = '" & txtDN.Text & "'"
IsRightRId = IsOraRecord(sOra)

End Function

Private Function InsertReelIDToTmpTbl(sReelID As String) As Boolean
Dim sOra         As String
Dim sSql         As String
Dim sSql2        As String
Dim lLastDevQty  As Long
Dim lTotalDevQty As Long
Dim lLastJobQty  As Long
Dim lTotalJobQty As Long
Dim tOra         As ST_TR_SEQ
Dim rs           As New ADODB.Recordset

InsertReelIDToTmpTbl = False
sSql2 = " SELECT RTRIM(SUBSTRING(a.KEY_VALUE,CHARINDEX('|',a.KEY_VALUE)+1,20)) AS JOB ,sum(b.数量) FROM erpdata..tblErpInStockDetailInfo a ,erpdata..tblStockNumSub b  " & " WHERE a.KEY_VALUE LIKE '%|%'  AND a.KEY_NAME = 'CONTAINER_NAME' AND a.KEY_TYPE = 'T' AND SUBSTRING(a.KEY_VALUE,1,CHARINDEX('|',a.KEY_VALUE)-1) = b.箱号 and b.箱号 = '" & sReelID & "' GROUP BY RTRIM(SUBSTRING(a.KEY_VALUE,CHARINDEX('|',a.KEY_VALUE)+1,20)) "
Set rs = Get_SqlserveRs(sSql2)
If Len(sReelID) = 13 Or Len(sReelID) = 14 Then
    If rs.RecordCount > 0 Then
        tOra.sJob = Trim$(rs(0).Value)
        tOra.lQty = Trim$(rs(1).Value)

    End If

    tOra.sDN = Trim(txtDN.Text)
    tOra.sReelID = sReelID
    '    tOra.sLot = Mid$(sReelID, 2, Len(sReelID) - 5)
    tOra.sLot = Mid(sReelID, 2, InStr(sReelID, "-") - 2)
    tOra.lSeq = glReelListCnt + 1
Else
    sSql = "select customerlotid as jobid, htlotid as lotid, qty from [erpdata].[dbo].[TblTSV_Tray_details] where TRAYQBOXNUMBER = '" & sReelID & "'"
    Set rs = Get_SqlserveRs(sSql)
    If rs.RecordCount = 0 Then
        '   MsgBox "系统TblTSV_Tray_details无此卷盘数据", vbInformation, "友情提示!!!"
        rs.Close
        Exit Function

    End If

    If IsNull(rs!JOBID) Or IsNull(rs!LOTID) Or IsNull(rs!QTY) Then
        '   MsgBox "系统TblTSV_Tray_details,无job信息", vbInformation, "友情提示!!!"
        rs.Close
        Exit Function

    End If

    tOra.sDN = txtDN.Text
    tOra.sReelID = sReelID
    tOra.sJob = Trim("" & rs!JOBID)
    tOra.sLot = Trim$("" & rs!LOTID)
    tOra.lQty = rs!QTY
    tOra.lSeq = glReelListCnt + 1

End If

sOra = "select marketingpn as mpn from CUSTOMERSHIPPINGUPTBL where delivery = '" & txtDN.Text & "' and batchnumber = '" & tOra.sJob & "'"
Set rs = Get_OracleRs(sOra)
If rs.RecordCount = 0 Then
    MsgBox "DN无机种信息", vbInformation, "友情提示!!!"
    rs.Close
    Exit Function

End If

tOra.sDev = Trim(rs!MPN)
If gsLastPN <> "" Then
    sOra = "select sum(qty) as qty from ST_TR_SEQ where dev = '" & gsLastPN & "'  and dn = '" & txtDN.Text & "' "
    lLastDevQty = Get_OracleNo(sOra)
    sOra = "select sum(quantity) as qty from CUSTOMERSHIPPINGUPTBL where delivery = '" & tOra.sDN & "' and marketingpn = '" & gsLastPN & "'"
    lTotalDevQty = Get_OracleNo(sOra)
    If lLastDevQty < lTotalDevQty Then
        If tOra.sDev <> gsLastPN Then
            Play ("oneDev")
            MsgBox "上一机种尚未扫完, 请按机种顺序扫描", vbInformation, "友情提示!!!"
            Exit Function

        End If

    End If

End If

If gsLastJob <> "" Then
    sOra = "select sum(qty) as qty from ST_TR_SEQ where job = '" & gsLastJob & "'  and dn = '" & txtDN.Text & "' "
    lLastJobQty = Get_OracleNo(sOra)
    sOra = "select sum(quantity) as qty from CUSTOMERSHIPPINGUPTBL where delivery = '" & tOra.sDN & "' and batchnumber = '" & gsLastJob & "'"
    lTotalJobQty = Get_OracleNo(sOra)
    If lLastJobQty < lTotalJobQty Then
        If tOra.sJob <> gsLastJob Then
            Play ("oneJob")
            MsgBox "上一JOB未扫完, 请按Job顺序扫描", vbInformation, "友情提示!!!"
            Exit Function

        End If

    End If

End If

gsLastJob = tOra.sJob
gsLastPN = tOra.sDev
sOra = "insert into ST_TR_SEQ values('" & tOra.sDN & "', '" & tOra.sJob & "', '" & tOra.sDev & "', '" & tOra.lQty & "', sysdate, '" & tOra.sLot & "','" & tOra.sReelID & "', '" & tOra.lSeq & "' )"
AddSql (sOra)
' 备份
'sOra = "insert into ST_TR_SEQ_BACK values('" & tOra.sDN & "', '" & tOra.sJob & "', '" & tOra.sDev & "', '" & tOra.lQty & "', sysdate, '" & tOra.sLot & "','" & tOra.sReelID & "', '" & tOra.lSeq & "' )"
'AddSql (sOra)
InsertReelIDToTmpTbl = True

End Function

Private Sub CheckCurrentQty(sReelID As String)
Dim lTotalQty   As Long
Dim lCurrentQty As Long
Dim i           As Integer
Dim bFinish     As Boolean

bFinish = True

With fpsJob

    For i = 1 To .MaxRows
        .Row = i
        .Col = 2
        lTotalQty = Val(.Text)
        .Col = 3
        lCurrentQty = Val(.Text)
        If lCurrentQty > lTotalQty Then
            Play ("wrongCnt")
            MsgBox "已经超出所需数量, 挑料出错", vbInformation, "友情提示!!!"
            AddSql ("delete from ST_TR_SEQ where REELID = '" & sReelID & "'")
            Call ShowFps(txtDN.Text)
            Exit Sub

        End If

        If lCurrentQty < lTotalQty Then
            bFinish = False

        End If

    Next

End With

txtReelID.Text = sReelID & vbCrLf & txtReelID.Text
If bFinish = True Then
    txtSF.Locked = True
    Play ("ReelFinish")
    cmdCombine.Enabled = True
    MsgBox "已经全部扫描完毕, 请不要再扫描", vbInformation, "友情提示!!!"

End If

End Sub

Private Sub UpdateCurrentStatus()
txtStatus.BackColor = vbWhite
glReelListCnt = UBound(Split(txtReelID, vbCrLf))
Call InitStatus

End Sub

Rem: CombineData
Private Sub cmdCombine_Click()
Dim sOra As String
Dim rs   As New ADODB.Recordset

cmdCombine.Enabled = False
If txtDN.Text = "" Then
    MsgBox "DN不可为空", vbInformation, "友情提示!!!"
    Exit Sub

End If

sOra = "select * from ST_TR_SEQ where dn = '" & txtDN.Text & "' order by seq"
Set rs = Get_OracleRs(sOra)
If rs.RecordCount = 0 Then
    MsgBox "临时表没有卷盘的数据", vbInformation, "友情提示!!!"
    rs.Close
    Exit Sub

End If

rs.MoveFirst

Do While Not rs.EOF
    Call InsertTmpToHistory(rs)
    rs.MoveNext
Loop
Call MakeBoxID(txtDN.Text)
MsgBox "合箱完成", vbInformation, "友情提示!!!"
cmdPrintReel_Box.Enabled = True
rs.Close

End Sub

Private Sub InsertTmpToHistory(rs As ADODB.Recordset)
Dim sOra  As String
Dim tData As tSTData
Dim lCnt  As Long
Dim sSeq  As String

sOra = "select count(outbox_num) from PACKING_DETAILED where dn_num = '" & rs.Fields("DN") & "' and customer_device = '" & rs.Fields("DEV") & "' " & " and outbox_num in (select nvl(max(outbox_num), '1') from PACKING_DETAILED where dn_num = '" & rs.Fields("DN") & "' and customer_device = '" & rs.Fields("DEV") & "')"
lCnt = Get_OracleNo(sOra)
If lCnt <= 107 Then
    sOra = "select * from PACKING_DETAILED where dn_num = '" & rs.Fields("DN") & "' and customer_device = '" & rs.Fields("DEV") & "'"
    If Get_OracleCnt(sOra) > 0 Then
        sOra = "select nvl(max(outbox_num), '1') from PACKING_DETAILED where dn_num = '" & rs.Fields("DN") & "' and customer_device = '" & rs.Fields("DEV") & "' "
        tData.OUTBOX_NUM = Get_OracleNo(sOra)
    Else
        sOra = " select nvl(max(outbox_num), '0') +1 from PACKING_DETAILED where dn_num = '" & rs.Fields("DN") & "'"
        tData.OUTBOX_NUM = Get_OracleNo(sOra)

    End If

Else
    sOra = "select (nvl(max(outbox_num), '1') + 1) from PACKING_DETAILED where dn_num = '" & rs.Fields("DN") & "' and customer_device = '" & rs.Fields("DEV") & "' "
    tData.OUTBOX_NUM = Get_OracleStr(sOra)

End If

sOra = "select count(inbox_num) from PACKING_DETAILED where dn_num = '" & rs.Fields("DN") & "' and customer_device = '" & rs.Fields("DEV") & "'  and outbox_num = '" & tData.OUTBOX_NUM & "' " & "and inbox_num in (select nvl(max(inbox_num), '1') from PACKING_DETAILED where dn_num = '" & rs.Fields("DN") & "' and customer_device = '" & rs.Fields("DEV") & "'  and outbox_num = '" & tData.OUTBOX_NUM & "')"
lCnt = Get_OracleNo(sOra)
If lCnt <= 8 Then
    sOra = "select nvl(max(inbox_num), '1') from PACKING_DETAILED where dn_num = '" & rs.Fields("DN") & "' and customer_device = '" & rs.Fields("DEV") & "'  and outbox_num = '" & tData.OUTBOX_NUM & "'"
    tData.INBOX_NUM = Get_OracleStr(sOra)
Else
    sOra = "select nvl(max(inbox_num), '1') + 1 from PACKING_DETAILED where dn_num = '" & rs.Fields("DN") & "' and customer_device = '" & rs.Fields("DEV") & "'  and outbox_num = '" & tData.OUTBOX_NUM & "'"
    tData.INBOX_NUM = Get_OracleStr(sOra)

End If

tData.TRAYID = "" & rs.Fields("REELID")
tData.CREATE_BY = gUserName
tData.DN_NUM = "" & rs.Fields("DN")
tData.JOB_ID = "" & rs.Fields("JOB")
tData.Customer_Device = "" & rs.Fields("DEV")
tData.QTY = rs.Fields("QTY")
tData.REEL_ID = Replace(tData.JOB_ID, "M", "") & GetLableXHTW(tData.JOB_ID)
tData.SEQ = rs.Fields("SEQ")
If Right(rs.Fields("LOTID"), 1) = "M" Then
    tData.REEL_ID = tData.JOB_ID & Right$(tData.TRAYID, 1)

End If

' 求datecode
sDatecode = Get37TestDC(Trim(tData.DN_NUM), Trim(tData.JOB_ID))

sOra = "insert into PACKING_DETAILED values('" & tData.TRAYID & "','" & tData.INBOX_NUM & "','" & tData.OUTBOX_NUM & "','" & tData.DN_NUM & "','" & tData.JOB_ID & "','" & tData.QTY & "','" & tData.Customer_Device & "',sysdate,'" & tData.CREATE_BY & "','0','0','', '" & tData.REEL_ID & "','','','', '" & tData.SEQ & "', '" & sDatecode & "') "
AddSql (sOra)

End Sub

Private Sub MakeBoxID(sDN As String)
Dim rs          As New ADODB.Recordset
Dim sOra        As String
Dim sAppend     As String
Dim iOp         As Integer
Dim iIp         As Integer
Dim iOpMax      As Integer
Dim iIpMax      As Integer
Dim sLot        As String
Dim iSeq        As Integer
Dim sLotBox     As String
Dim sLastLot    As String
Dim sFirstBoxId As String
Dim sQboxNo     As String
Dim sCartonID   As String
Dim lQty        As Long
Dim iBoxIDSeq   As String
Dim stqtpj      As String

iOp = 1
iIp = 1
sOra = "select max(outbox_num) from PACKING_DETAILED where dn_num = '" & sDN & "'  "
iOpMax = Get_OracleNo(sOra)

For iOp = 1 To iOpMax
    sOra = "select max(inbox_num) from PACKING_DETAILED where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "'"
    iIpMax = Get_OracleNo(sOra)
    sOra = "update PACKING_DETAILED set kid = 'K'||'" & iOp & "' where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "'"
    Exec_Ora (sOra)
    Rem: Make C ID
    sAppend = "select a.job_id, b.lotid, sum(a.Qty) from PACKING_DETAILED a, ST_TR_SEQ b where a.dn_num = '" & txtDN.Text & "'   and b.dn = '" & txtDN.Text & "'  and a.outbox_num = '" & iOp & "' and b.job = a.job_id group by a.job_id, b.lotid"
    Set rs = Get_OracleRs(sAppend)
    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            sLot = "S" & rs!LOTID & "-C"
            sOra = " select nvl(max(seq)+1, 1) from PKGIDSEQ_37 where val = '" & sLot & "' "
            stqtpj = Right("0" & Get_OracleStr(sOra), 2)
            sLotBox = sLot & stqtpj
            sOra = "update PACKING_DETAILED set CARTONID = '" & sLotBox & "' where dn_num = '" & txtDN.Text & "' and outbox_num = '" & iOp & "' and job_id = '" & rs.Fields("job_id") & "'"
            AddSql (sOra)
            strSql = "insert into PKGIDSEQ_37(VAL,SEQ,CREATE_DATE,DN) values('" & sLot & "', '" & stqtpj & "', sysdate, '" & sDN & "')"
            AddSql (strSql)
            rs.MoveNext
        Loop

    End If

    rs.Close
    Rem: Make B ID

    For iIp = 1 To iIpMax
        sAppend = "select a.job_id, b.lotid, sum(a.Qty) from PACKING_DETAILED a, ST_TR_SEQ b where a.dn_num = '" & txtDN.Text & "' and b.dn = '" & txtDN.Text & "'  and a.outbox_num = '" & iOp & "' and a.inbox_num = '" & iIp & "' and b.job = a.job_id group by a.job_id, b.lotid"
        Set rs = Get_OracleRs(sAppend)
        If Not rs.BOF Then
            rs.MoveFirst

            Do While Not rs.EOF
                sLot = "S" & rs!LOTID & "-B"
                sOra = " select nvl(max(seq)+1, 1) from PKGIDSEQ_37 where val = '" & sLot & "' "
                stqtpj = Right("0" & Get_OracleStr(sOra), 2)
                sLotBox = sLot & stqtpj
                ' 更新进数据表
                sOra = "update PACKING_DETAILED set BOXID = '" & sLotBox & "' where dn_num = '" & txtDN.Text & "' and outbox_num = '" & iOp & "' and inbox_num = '" & iIp & "' and job_id = '" & rs.Fields("job_id") & "'"
                AddSql (sOra)
                strSql = "insert into PKGIDSEQ_37(VAL,SEQ,CREATE_DATE,DN) values('" & sLot & "', '" & stqtpj & "', sysdate, '" & sDN & "')"
                AddSql (strSql)
                rs.MoveNext
            Loop

        End If

        rs.Close
    Next
    Rem: Make Q ID
    sFirstBoxId = Get_OracleStr("select boxid from PACKING_DETAILED where dn_num = '" & txtDN.Text & "' and outbox_num='" & iOp & "' and inbox_num = '1'")
    If sFirstBoxId <> "" Then
        sQboxNo = Get_OracleStr("select  trglabelseq.QTSeq_NotMesQbox('" & sFirstBoxId & "')  from dual")
        sOra = "update PACKING_DETAILED set CARTON = '" & sQboxNo & "' where dn_num = '" & txtDN.Text & "' and outbox_num = '" & iOp & "' "
        AddSql (sOra)

    End If

Next
' 插入SQl SERVER
AddSql2 ("delete from  ERPBASE..PACKING_DETAILED where dn_num = '" & sDN & "'")
AddSql2 ("insert into ERPBASE..PACKING_DETAILED select * from OPENQUERY(ORACLEDB, 'SELECT * from PACKING_DETAILED where dn_num = ''" & sDN & "'' ' ) ")
' 更新箱号关系
Call TransToErp(sDN)

End Sub

Rem: Print Inner start
Private Sub cmdPrintReel_Box_Click()
Dim sOra   As String
Dim sDN    As String
Dim iOpMax As Integer
Dim iIpMax As Integer
Dim iOp    As Integer
Dim iIp    As Integer
Dim iSec   As Integer
Dim sType  As String
Dim rs     As New ADODB.Recordset

cmdPrintReel_Box.Enabled = False
sType = Trim(txtShip.Text)
cmdPrintReel_Box.Enabled = False
iSec = txtPrintInterval.Text * 1000
sDN = txtDN.Text
If sDN = "" Then
    MsgBox "DN不可为空", vbInformation, "友情提示!!!"
    Exit Sub

End If

If Get_OracleStr("select * from PACKING_DETAILED where dn_num = '" & sDN & "' ") = "" Then
    MsgBox "系统无此标签数据", vbInformation, "提示:"
    Exit Sub

End If

' modify: 按照单个外箱打印
sOra = "select  min(outbox_num) from PACKING_DETAILED where dn_num = '" & sDN & "' and print_flag = '0' order by outbox_num"
iOp = Get_OracleNo(sOra)
If iOp = 0 Then
    MsgBox "内盒卷盘标签已全部打印完毕", vbInformation, "友情提示:"
    cmdPrintReel_Box.Enabled = False
    cmdPrintCarton.Enabled = True
    Exit Sub

End If

sOra = "select max(inbox_num) from PACKING_DETAILED where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "'"
iIpMax = Get_OracleNo(sOra)

For iIp = 1 To iIpMax
    Call PrintINNERBOXFlag(sDN, iOp, iIp)
    Sleep (iSec)
    Call PrintSTBoxLbl(sDN, iOp, iIp)
    Sleep (iSec)
    ' 三星
    If sType = "SSE2" Then
        Call PrintCusBoxLbl(sDN, iOp, iIp)
        Sleep (iSec)
        Call PrintREELFlag(sDN, iOp, iIp)
        Sleep (iSec)
        Call PrintCusReelLbl(sDN, iOp, iIp)
        Sleep (iSec)
    ElseIf sType = "SSSHORT" Then
        Call PrintCusBoxLbl2(sDN, iOp, iIp)
        Sleep (iSec)
        Call PrintREELFlag(sDN, iOp, iIp)
        Sleep (iSec)
        Call PrintCusReelLbl2(sDN, iOp, iIp)
        Sleep (iSec)
    ElseIf sType = "HW" Then
        ' 华为
        Call PrintHWBoxLbl(sDN, iOp, iIp)
        Sleep (iSec)
        Call PrintREELFlag(sDN, iOp, iIp)
        Sleep (iSec)
        Call PrintHWReelLbl(sDN, iOp, iIp)
        Sleep (iSec)

    End If

Next
Call PrintOuterCFlag(sDN, iOp)
Call PrintSTCartonLbl(sDN, iOp)
Call PrintHTCartonLbl(sDN, iOp)
Call UpdatePrintStatus(sDN, iOp)
MsgBox "第" & iOp & "个外箱已经打印完成", vbInformation, "友情提示:"
cmdPrintReel_Box.Enabled = True

End Sub

Rem: Step 1 :ST BoxLable
Private Sub PrintSTBoxLbl(sDN As String, iOp As Integer, iIp As Integer)
Dim sOra          As String
Dim sDatecode     As String
Dim sTestDateCode As String
Dim tSTBox        As STBox
Dim sContent      As String
Dim sFileName     As String
Dim sPath         As String
Dim sAdd          As String
Dim rs            As New ADODB.Recordset

sPath = Trim(txtSTBoxPath.Text)
sFileName = sDN & "-" & "STBoxLbl" & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select job_id,CUSTOMER_DEVICE,boxid, sum(QTY) as qty from PACKING_DETAILED where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' and inbox_num =  '" & iIp & "'  group by job_id,CUSTOMER_DEVICE,boxid"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tSTBox.JOB = Trim("" & rs!JOB_ID)
        tSTBox.DEV = Trim("" & rs!Customer_Device)
        tSTBox.lot = Trim("" & rs!BOXID)
        tSTBox.QTY = Trim("" & rs!QTY)
        sDatecode = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTBox.JOB & "'")
        tSTBox.DATECODE = sDatecode
        sTestDateCode = Get_SqlStr("select distinct HTDATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTBox.JOB & "'")
        tSTBox.testdateCode = sTestDateCode
        If tSTBox.DATECODE = "" Or tSTBox.testdateCode = "" Then
            tSTBox.DATECODE = Get_OracleStr("select datecode from PACKING_DETAILED where dn_num = '" & sDN & "' and JOB_ID = '" & tSTBox.JOB & "'")
            tSTBox.testdateCode = Get_OracleStr("select datecode from PACKING_DETAILED where dn_num = '" & sDN & "' and JOB_ID = '" & tSTBox.JOB & "'")

        End If

        sContent = sContent + tSTBox.DEV + "," + tSTBox.JOB + ",1T" + tSTBox.JOB + "," + tSTBox.DEV + "," + "1P" + tSTBox.DEV + "," + tSTBox.DATECODE + "," + tSTBox.DATECODE + "," + Mid(tSTBox.lot, 2) + "," + tSTBox.lot + ","
        sContent = sContent + tSTBox.QTY + ",Q" + tSTBox.QTY + "," + tSTBox.testdateCode + "," + tSTBox.testdateCode
        sAdd = ""
        sAdd = GetDevMark(tSTBox.DEV)
        sContent = sContent + sAdd + vbCrLf
        rs.MoveNext
    Loop

End If

Call CreateTxt(sFileName, sContent, sPath)

End Sub

Rem: Step 2 :Samgsung Box Lable
Private Sub PrintCusBoxLbl(sDN As String, iOp As Integer, iIp As Integer)
Dim sOra       As String
Dim tCusBox    As CusBox
Dim sContent   As String
Dim sFileName  As String
Dim rs         As New ADODB.Recordset
Dim sPath      As String
Dim strFabSite As String, strAssemblySite As String, strTestSite As String

sPath = Trim(txtCusBoxPath.Text)
sFileName = sDN & "-" & "CUSBoxLbl" & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select sum(a.qty) as qty, a.CUSTOMER_DEVICE, b.customerpartnumber,b.FAB_SITE, b.ASSEMBLY_SITE,b.TEST_SITE from PACKING_DETAILED a , CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & sDN & "' and b.delivery = '" & sDN & "' and a.job_id = b.batchnumber and a.outbox_num = '" & iOp & "' and a.inbox_num = '" & iIp & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.FAB_SITE, b.ASSEMBLY_SITE,b.TEST_SITE"
Set rs = Get_OracleRs(sOra)
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

        sContent = sContent + tCusBox.PN + "DPTKE2" + Right$("000000" + tCusBox.QTY, 6) + ","
        sContent = sContent + tCusBox.PN + "," + "TVS DIODES" + "," + "PO TYPE,:E2" + "," + tCusBox.QTY + "," + tCusBox.DEV + "," + "DPTK" + "," + strFabSite + "," + strAssemblySite + "," + strTestSite
        rs.MoveNext
    Loop

End If

If tCusBox.DEV = "RCLAMP2581ZCTFT" Then
    sPath = "\\10.160.1.84\public\BarCode\37\37BoxNH-新\"

End If

Call CreateTxt(sFileName, sContent, sPath)

End Sub

Private Sub PrintCusBoxLbl2(sDN As String, iOp As Integer, iIp As Integer)
Dim sOra      As String
Dim tCusBox   As CusBox
Dim sContent  As String
Dim sFileName As String
Dim rs        As New ADODB.Recordset
Dim sPath     As String

sPath = "\\10.160.1.84\public\BarCode\37\37NH2\"
sFileName = sDN & "-" & "CUSBoxLbl2" & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select sum(a.qty) as qty, a.CUSTOMER_DEVICE, b.customerpartnumber from PACKING_DETAILED a , CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & sDN & "' and b.delivery = '" & sDN & "' and a.job_id = b.batchnumber and a.outbox_num = '" & iOp & "' and a.inbox_num = '" & iIp & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.FAB_SITE, b.ASSEMBLY_SITE,b.TEST_SITE"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tCusBox.QTY = Trim("" & rs!QTY)
        tCusBox.DEV = Trim("" & rs!Customer_Device)
        tCusBox.PN = Trim("" & rs!CustomerPartnumber)
        sContent = sContent + tCusBox.PN + "DPTK" + Right$("000000" + tCusBox.QTY, 6) + ","
        sContent = sContent + tCusBox.PN + "," + "TVS DIODES" + "," + tCusBox.QTY + "," + tCusBox.DEV + "," + "DPTK" + ","
        rs.MoveNext
    Loop

End If

If tCusBox.DEV = "RCLAMP2581ZCTFT" Then
    MsgBox "请联系IT确认标签模板", vbCritical, "警告"
    Exit Sub

End If

Call CreateTxt(sFileName, sContent, sPath)

End Sub

Private Sub PrintHWBoxLbl(sDN As String, iOp As Integer, iIp As Integer)
Dim tHWBox    As HWBox
Dim sContent  As String
Dim sFileName As String
Dim sPath     As String
Dim sOra      As String
Dim rs        As New ADODB.Recordset

sPath = Trim(txtHWReelPath.Text)
sFileName = sDN & "-" & "HWBoxLbl" & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select job_id,mpn,cpn, sum(QTY) as qty from LPSTBL where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' and inbox_num =  '" & iIp & "' group by job_id,mpn,cpn"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tHWBox.CPN = Trim("" & rs!CPN)
        tHWBox.MPN = Trim("" & rs!MPN)
        tHWBox.lot = Trim("" & rs!JOB_ID)
        strFAB_SITE = Trim$("" & rs!FAB_SITE)
        strASSEMBLY_SITE = Trim$("" & rs!ASSEMBLY_SITE)
        strTEST_SITE = Trim("" & rs!TEST_SITE)
        tHWBox.QTY = "" & rs!QTY
        tHWBox.PODATE = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tHWBox.lot & "'")
        sContent = sContent + tHWBox.CPN + ",P" + tHWBox.CPN + "," + tHWBox.MPN + ",1P" + tHWBox.MPN + "," + tHWBox.QTY + ",Q" + tHWBox.QTY + ","
        sContent = sContent + tHWBox.PODATE + ",10D" + tHWBox.PODATE + "," + "CN" + "," + "VCN" + "," + "601024" + "," + "4L601024" + "," + tHWBox.lot + ",1T" + tHWBox.lot + ","
        sContent = sContent + tHWBox.CPN + ";" + tHWBox.MPN + ";" + tHWBox.QTY + ";" + tHWBox.PODATE + ";" + "CN" + ";" + "601024" + ";" + tHWBox.lot + vbCrLf
        rs.MoveNext
    Loop

End If

Call CreateTxt(sFileName, sContent, sPath)

End Sub

Rem: Step 3 :Cus Reel Lable
Private Sub PrintCusReelLbl(sDN As String, iOp As Integer, iIp As Integer)
Dim sOra            As String
Dim tCusReel        As CusReel
Dim sContent        As String
Dim sFileName       As String
Dim rs              As New ADODB.Recordset
Dim Rs2             As New ADODB.Recordset
Dim sPath           As String
Dim strFabSite      As String
Dim strAssemblySite As String
Dim strTestSite     As String

sPath = Trim(txtCusReelPath.Text)
sFileName = sDN & "-" & "CUSREELLbl" & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select distinct trayid, reelid, qty, CUSTOMER_DEVICE, cpn,seq, FAB_SITE, ASSEMBLY_SITE,TEST_SITE from lpstbl where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' and inbox_num = '" & iIp & "'  order by seq "
Set Rs2 = Get_OracleRs(sOra)
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

        sContent = sContent + tCusReel.PN + "DPTKE2" + tCusReel.lot + Right$("000000" + tCusReel.QTY, 6) + ","
        sContent = sContent + tCusReel.PN + "," + "TVS DIODES" + "," + "PO TYPE,:E2" + "," + tCusReel.lot + "," + tCusReel.QTY + ","
        sContent = sContent + tCusReel.DEV + "," + "DPTK" + "," + strFabSite + "," + strAssemblySite + "," + strTestSite + vbCrLf
        Rs2.MoveNext
    Loop

End If

If tCusReel.DEV = "RCLAMP2581ZCTFT" Then
    sPath = "\\10.160.1.84\public\BarCode\37\37BoxJP-新\"

End If

Call CreateTxt(sFileName, sContent, sPath)

End Sub

Rem: Step 3 :Cus Reel Lable
Private Sub PrintCusReelLbl2(sDN As String, iOp As Integer, iIp As Integer)
Dim sOra      As String
Dim tCusReel  As CusReel
Dim sContent  As String
Dim sFileName As String
Dim rs        As New ADODB.Recordset
Dim Rs2       As New ADODB.Recordset
Dim sPath     As String

sPath = "\\10.160.1.84\public\BarCode\37\37JP2\"
sFileName = sDN & "-" & "CUSREELLbl" & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select distinct trayid, reelid, qty, CUSTOMER_DEVICE, cpn,seq from lpstbl where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' and inbox_num = '" & iIp & "'  order by seq "
Set Rs2 = Get_OracleRs(sOra)
If Not Rs2.BOF Then
    Rs2.MoveFirst

    Do While Not Rs2.EOF
        tCusReel.TRAYID = Trim$("" & Rs2!TRAYID)
        tCusReel.lot = Trim$("" & Rs2!REELID)
        tCusReel.QTY = Trim("" & Rs2!QTY)
        tCusReel.DEV = Trim("" & Rs2!Customer_Device)
        tCusReel.PN = Trim("" & Rs2!CPN)
        sContent = sContent + tCusReel.PN + "DPTK" + tCusReel.lot + Right$("000000" + tCusReel.QTY, 6) + ","
        sContent = sContent + tCusReel.PN + "," + "TVS DIODES" + "," + tCusReel.lot + "," + tCusReel.QTY + ","
        sContent = sContent + tCusReel.DEV + "," + "DPTK" + "," + vbCrLf
        Rs2.MoveNext
    Loop

End If

If tCusReel.DEV = "RCLAMP2581ZCTFT" Then
    MsgBox "请联系IT确认机种是否有问题", vbCritical, "警告"
    Exit Sub

End If

Call CreateTxt(sFileName, sContent, sPath)

End Sub

Private Sub PrintHWReelLbl(sDN As String, iOp As Integer, iIp As Integer)
Dim tHWBox    As HWBox
Dim sContent  As String
Dim sFileName As String
Dim sPath     As String
Dim sOra      As String
Dim rs        As New ADODB.Recordset

sPath = Trim(txtHWReelPath.Text)
sFileName = sDN & "-" & "HWReelLbl" & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select job_id,mpn,cpn, QTY from LPSTBL where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' and inbox_num =  '" & iIp & "' "
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tHWBox.CPN = Trim("" & rs!CPN)
        tHWBox.MPN = Trim("" & rs!MPN)
        tHWBox.lot = Trim("" & rs!JOB_ID)
        tHWBox.QTY = rs!QTY
        tHWBox.PODATE = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tHWBox.lot & "'")
        sContent = sContent + tHWBox.CPN + ",P" + tHWBox.CPN + "," + tHWBox.MPN + ",1P" + tHWBox.MPN + "," + tHWBox.QTY + ",Q" + tHWBox.QTY + ","
        sContent = sContent + tHWBox.PODATE + ",10D" + tHWBox.PODATE + "," + "CN" + "," + "VCN" + "," + "601024" + "," + "4L601024" + "," + tHWBox.lot + ",1T" + tHWBox.lot + ","
        sContent = sContent + tHWBox.CPN + ";" + tHWBox.MPN + ";" + tHWBox.QTY + ";" + tHWBox.PODATE + ";" + "CN" + ";" + "601024" + ";" + tHWBox.lot + vbCrLf
        rs.MoveNext
    Loop

End If

Call CreateTxt(sFileName, sContent, sPath)

End Sub

Private Sub PrintFlag(sDN As String, iOp As Integer, iIp As Integer)
Dim sContent  As String
Dim sPath     As String
Dim sFileName As String

sPath = Trim(txtFlagPath.Text)
sContent = "BOX_REEL" & iOp & iIp
sFileName = sDN & "-" & "FLAGLBL" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(sFileName, sContent, sPath)

End Sub

Private Sub PrintINNERBOXFlag(sDN As String, iOp As Integer, iIp As Integer)
Dim sContent  As String
Dim sPath     As String
Dim sFileName As String

sPath = Trim(txtFlagPath.Text)
sContent = "BOX_" & iOp & "_" & iIp
sFileName = sDN & "-" & "FLAG_BOX_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(sFileName, sContent, sPath)

End Sub

Private Sub PrintREELFlag(sDN As String, iOp As Integer, iIp As Integer)
Dim sContent  As String
Dim sPath     As String
Dim sFileName As String

sPath = Trim(txtFlagPath.Text)
sContent = "REEL_" & iOp & "_" & iIp
sFileName = sDN & "-" & "FLAG_REEL_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(sFileName, sContent, sPath)

End Sub

Private Sub PrintOuterCFlag(sDN As String, iOp As Integer)
Dim sContent  As String
Dim sPath     As String
Dim sFileName As String

sPath = Trim(txtFlagPath.Text)
sContent = "CARTON_" & iOp
sFileName = sDN & "-" & "FLAG_CARTON_" & iOp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
Call CreateTxt(sFileName, sContent, sPath)

End Sub

Rem: Step 4 :ST Carton Lable
Private Sub PrintSTCartonLbl(sDN As String, iOp As Integer)
Dim sOra          As String
Dim tSTCarton     As STCarton
Dim sContent      As String
Dim sFileName     As String
Dim sDatecode     As String
Dim sTestDateCode As String
Dim rs            As New ADODB.Recordset
Dim sPath         As String
Dim sAdd          As String

sPath = Trim(txtSTBoxPath.Text)
sFileName = sDN & "-" & "STCARTONLBL" & "_" & iOp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = " select job_id,CUSTOMER_DEVICE,cartonid, sum(qty) as qty from PACKING_DETAILED where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' group by job_id,CUSTOMER_DEVICE,cartonid"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tSTCarton.JOB = Trim("" & rs!JOB_ID)
        tSTCarton.DEV = Trim$("" & rs!Customer_Device)
        tSTCarton.lot = Trim("" & rs!CARTONID)
        tSTCarton.QTY = Trim("" & rs!QTY)
        sDatecode = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTCarton.JOB & "'")
        tSTCarton.DATECODE = sDatecode
        sTestDateCode = Get_SqlStr("select distinct HTDATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTCarton.JOB & "'")
        tSTCarton.testdateCode = sTestDateCode
        If tSTCarton.DATECODE = "" Or tSTCarton.testdateCode = "" Then
            tSTCarton.DATECODE = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTCarton.JOB & "'")
            tSTCarton.testdateCode = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTCarton.JOB & "'")

        End If

        sContent = sContent + tSTCarton.DEV + "," + tSTCarton.JOB + ",1T" + tSTCarton.JOB + "," + tSTCarton.DEV + "," + "1P" + tSTCarton.DEV + "," + tSTCarton.DATECODE + "," + tSTCarton.DATECODE + "," + Mid(tSTCarton.lot, 2) + "," + tSTCarton.lot + ","
        sContent = sContent + tSTCarton.QTY + ",Q" + tSTCarton.QTY + "," + tSTCarton.testdateCode + "," + tSTCarton.testdateCode
        sAdd = GetDevMark(tSTCarton.DEV)
        sContent = sContent + sAdd + vbCrLf
        rs.MoveNext
    Loop

End If

Call CreateTxt(sFileName, sContent, sPath)

End Sub

Rem: Print OutPkg Start
Private Sub cmdPrintCarton_Click()
Dim sDN    As String
Dim sOra   As String
Dim iOp    As Integer
Dim iOpMax As Integer
Dim iSec   As Integer
Dim sType  As String

cmdPrintCarton.Enabled = False
iSec = (txtPrintInterval.Text) * 1000
sType = Trim(txtShip.Text)
sDN = txtDN.Text
If sDN = "" Then
    MsgBox "DN不可为空", vbInformation
    Exit Sub

End If

If Get_OracleStr("select * from PACKING_DETAILED where dn_num = '" & sDN & "' ") = "" Then
    MsgBox "系统无此标签数据", vbInformation, "提示:"
    Exit Sub

End If

sOra = "select max(outbox_num) from PACKING_DETAILED where dn_num = '" & sDN & "'  "
iOpMax = Get_OracleNo(sOra)

For iOp = 1 To iOpMax
    Sleep (iSec)
    If sType = "SSE2" Then
        Call PrintCusCartonLbl(sDN, iOp)
        Sleep (iSec)
    ElseIf sType = "SSSHORT" Then
        Call PrintCusCartonLbl2(sDN, iOp)
        Sleep (iSec)
    Else
        Call PrintSTCartonStanderLbl(sDN, iOp)
        Sleep (iSec)

    End If

Next
MsgBox "外箱大标签打印完成", vbInformation, "友情提示:"

End Sub

Rem: Step 5 :Customer Carton Lable
Private Sub PrintCusCartonSubLbl(sDN As String, iOp As Integer)
Dim sOra       As String
Dim tCusCARTON As CUSCARTON
Dim sFileName  As String
Dim sContent   As String
Dim rs         As New ADODB.Recordset
Dim sPath      As String
Dim KID        As String

sPath = Trim(txtCusCartonPath.Text)
sFileName = sDN & "-" & "CUSCARTONSUBLBL" + Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select a.kid,a.CUSTOMER_DEVICE, b.customerpartnumber, a.job_id,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & txtDN.Text & "'" & "and b.delivery = '" & txtDN.Text & "' and a.job_id = b.batchnumber and a.outbox_num = '" & iOp & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber, a.job_id, b.purchasingdocno,a.kid"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tCusCARTON.dn = txtDN.Text
        tCusCARTON.PO = Trim$("" & rs!PO)
        tCusCARTON.CPN = Trim$("" & rs!CustomerPartnumber)
        tCusCARTON.MPN = Trim$("" & rs!Customer_Device)
        tCusCARTON.JOB = Trim$("" & rs!JOB_ID)
        tCusCARTON.QTY = Trim$("" & rs!QTY)
        KID = "" & rs!KID
        tCusCARTON.DATECODE = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tCusCARTON.JOB & "'")
        If tCusCARTON.DATECODE = "" Then
            tCusCARTON.DATECODE = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tCusCARTON.JOB & "'")

        End If

        sContent = sContent & tCusCARTON.dn & ",I" & tCusCARTON.dn & "," & tCusCARTON.PO & ",K" & tCusCARTON.PO & ",E2," & tCusCARTON.CPN & ",P" & tCusCARTON.CPN & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & ","
        sContent = sContent & tCusCARTON.QTY & ",Q" & tCusCARTON.QTY & ","
        sContent = sContent & Get_OracleStr("select distinct trim(a.freightforwarder)|| ',CHINA,' || substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3)||','||trim(a.city) || ' ' || trim(a.state) || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ',' || trim(a.phone) || ','  from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & txtDN.Text & "'")
        sContent = sContent & tCusCARTON.JOB & ",P" & tCusCARTON.JOB & "," & tCusCARTON.DATECODE & ",9D" & tCusCARTON.DATECODE & "," & iOp & "," & KID & vbCrLf
        rs.MoveNext
    Loop

End If

Call CreateTxt(sFileName, sContent, sPath)

End Sub

Rem: Step 6 :Customer Carton Lable
Private Sub PrintCusCartonLbl(sDN As String, iOp As Integer)
Dim sOra       As String
Dim tCusCARTON As CUSCARTON
Dim sFileName  As String
Dim sContent   As String
Dim rs         As New ADODB.Recordset
Dim sPath      As String
Dim KID        As String
Dim sMaxOP     As String

sMaxOP = Get_OracleStr("select max(outbox_num) from PACKING_DETAILED where dn_num = '" & sDN & "'")
sPath = Trim(txtCusCartonPath.Text)
sFileName = sDN & "-" & "CUSCARTONLBL" + Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select a.kid,a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & sDN & "'" & "and b.delivery = '" & sDN & "' and a.job_id = b.batchnumber and a.outbox_num = '" & iOp & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno, a.kid"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tCusCARTON.dn = sDN
        tCusCARTON.PO = Left("" & rs!PO, 10)
        tCusCARTON.CPN = "" & rs!CustomerPartnumber
        tCusCARTON.MPN = "" & rs!Customer_Device
        tCusCARTON.QTY = "" & rs!QTY
        KID = rs!KID
        sContent = sContent & tCusCARTON.dn & ",I" & tCusCARTON.dn & "," & tCusCARTON.PO & ",K" & tCusCARTON.PO & ",E2," & tCusCARTON.CPN & ",P" & tCusCARTON.CPN & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & ","
        sContent = sContent & tCusCARTON.QTY & ",Q" & tCusCARTON.QTY & ","
        sContent = sContent & Get_OracleStr("select distinct trim(a.freightforwarder)|| ',CHINA,' || substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3)||','||trim(a.city) || ' ' || trim(a.state) || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ','|| 'Attn:;Tel:' || trim(a.phone) || ','  from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & sDN & "'")
        sContent = sContent & "N/A,PN/A,N/A,9DN/A," & iOp & "," & KID & "," & sMaxOP
        rs.MoveNext
    Loop

End If

Call CreateTxt(sFileName, sContent, sPath)

End Sub

Private Sub PrintCusCartonLbl2(sDN As String, iOp As Integer)
Dim sOra       As String
Dim tCusCARTON As CUSCARTON
Dim sFileName  As String
Dim sContent   As String
Dim rs         As New ADODB.Recordset
Dim sPath      As String
Dim KID        As String
Dim sMaxOP     As String

sMaxOP = Get_OracleStr("select max(outbox_num) from PACKING_DETAILED where dn_num = '" & sDN & "'")
sPath = Trim(txtCusCartonPath.Text)
sFileName = sDN & "-" & "CUSCARTONLBL" + Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select a.kid,a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & sDN & "'" & "and b.delivery = '" & sDN & "' and a.job_id = b.batchnumber and a.outbox_num = '" & iOp & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno, a.kid"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tCusCARTON.dn = sDN
        tCusCARTON.PO = Left("" & rs!PO, 10)
        tCusCARTON.CPN = "" & rs!CustomerPartnumber
        tCusCARTON.MPN = "" & rs!Customer_Device
        tCusCARTON.QTY = "" & rs!QTY
        KID = rs!KID
        sContent = sContent & tCusCARTON.dn & ",I" & tCusCARTON.dn & "," & tCusCARTON.PO & ",K" & tCusCARTON.PO & ",," & tCusCARTON.CPN & ",P" & tCusCARTON.CPN & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & ","
        sContent = sContent & tCusCARTON.QTY & ",Q" & tCusCARTON.QTY & ","
        sContent = sContent & Get_OracleStr("select distinct trim(a.freightforwarder)|| ',CHINA,' || substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3)||','||trim(a.city) || ' ' || trim(a.state) || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ','|| 'Attn:;Tel:' || trim(a.phone) || ','  from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & sDN & "'")
        sContent = sContent & "N/A,PN/A,N/A,9DN/A," & iOp & "," & KID & "," & sMaxOP
        rs.MoveNext
    Loop

End If

Call CreateTxt(sFileName, sContent, sPath)

End Sub

Private Sub PrintSTCartonStanderLbl(sDN As String, iOp As Integer)
Dim sOra       As String
Dim tCusCARTON As CUSCARTON
Dim sFileName  As String
Dim sContent   As String
Dim rs         As New ADODB.Recordset
Dim sPath      As String
Dim sAdd       As String
Dim sKid       As String
Dim sMaxOP     As String

sMaxOP = Get_OracleStr("select max(outbox_num) from PACKING_DETAILED where dn_num = '" & sDN & "'")
sPath = Trim(txtSEMTCartonPAth.Text)
sFileName = sDN & "-" & "SemTechStanderCarton" + Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select a.CUSTOMER_DEVICE,a.kid, b.customerpartnumber,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & sDN & "'" & "and b.delivery = '" & sDN & "' and a.job_id = b.batchnumber and a.outbox_num = '" & iOp & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno,a.kid"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tCusCARTON.dn = sDN
        tCusCARTON.PO = IIf(IsNull(rs!PO), "N/A", Left(rs!PO, 10))
        tCusCARTON.CPN = IIf(IsNull(rs!CustomerPartnumber), "N/A", rs!CustomerPartnumber)
        tCusCARTON.MPN = IIf(IsNull(rs!Customer_Device), "N/A", rs!Customer_Device)
        tCusCARTON.QTY = rs!QTY
        sKid = rs!KID
        sContent = sContent & Get_OracleStr("select distinct substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3) || ','||trim(a.city) || ' ' || trim(a.state)  || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ',' || trim(a.contactname) || ',' || trim(a.phone) from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & sDN & "' ") & ","
        sContent = sContent & tCusCARTON.dn & ",I" & tCusCARTON.dn & "," & Left(tCusCARTON.PO, 10) & ",K" & Left(tCusCARTON.PO, 10) & "," & Left$(tCusCARTON.CPN, 11) & ",P" & Left$(tCusCARTON.CPN, 11) & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & "," & tCusCARTON.QTY & ",Q" & tCusCARTON.QTY & "," & Get_OracleStr("select distinct freightforwarder from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & sDN & "'") & "," & "" & "," & "" & "," & "" & "," & "COO:CHINA" & "," & "CHINA"
        sAdd = "," & iOp & "," & sKid & "," & sMaxOP
        sContent = sContent & sAdd
        rs.MoveNext
    Loop

End If

Call CreateTxt(sFileName, sContent, sPath)

End Sub

Rem: Step 7 :HT Carton Lable
Private Sub PrintHTCartonLbl(sDN As String, iOp As Integer)
Dim sCartonNo As String
Dim sOra      As String
Dim sFileName As String
Dim sContent  As String
Dim sPath     As String

sPath = Trim(txtHTCartonPath.Text)
sFileName = sDN & "-" & "STCARTONLBL" + Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select distinct carton from PACKING_DETAILED where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "'"
sCartonNo = Get_OracleStr(sOra)
sFileName = "HTCARTONLBL" & Format(Now(), "YYYYMMDDHHmmSS")
sContent = sCartonNo
Call CreateTxt(sFileName, sContent, sPath)

End Sub

Rem: 打印标签
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
Sleep (1000)

End Sub

Rem: 播放音频提醒
Private Sub Play(sFileName As String)
Dim sPath   As String
Dim sSuffix As String

sPath = txtMediaPath.Text
sSuffix = ".wav"
media.url = sPath & sFileName & sSuffix
Sleep (200)

End Sub

Rem: 补打
Private Sub cmdPrint_Click()
Dim sRefer As String
Dim sType  As String
Dim iQty   As Integer

If tReferTo.Text = "" Then
    MsgBox "请输入参考值", vbInformation, "提示:"
    Exit Sub

End If

sRefer = UCase(Trim$(tReferTo.Text))
sType = cbType.Text

Select Case sType

    Case "按内盒补打"
        Call PrintA_Inner(sRefer)

    Case "三星卷盘标签"
        Call PrintA_SSREEL(sRefer)

    Case "Semtech内盒标签"
        Call PrintA_STBOX(sRefer)

    Case "三星内盒标签"
        Call PrintA_SSBOX(sRefer)

    Case "Semtech外箱分标签"
        Call PrintA_STCARTONSUB(sRefer)

    Case "外箱大标签补打"
        Call PrintAOut

End Select

iQty = Get_OracleStr("select nvl(count(*) + 1, 1) from TBL_37_PRINT2_LIST where KEYNAME = '" & sType & "' and KEYVALUE = '" & sRefer & "'")
AddSql ("insert into TBL_37_PRINT2_LIST(KEYNAME,KEYVALUE,CREATE_DATE,CREATE_BY,CREATE_TIMES) values('" & sType & "', '" & sRefer & "', sysdate, '" & gUserName & "', '" & iQty & "')")

End Sub

'按内盒补打
Private Sub PrintA_Inner(sRefer As String)
Dim sOra  As String
Dim rs    As New ADODB.Recordset
Dim iOp   As Integer
Dim iIp   As Integer
Dim sType As String
Dim sDN   As String

sDN = txtDN2.Text
sType = Trim(txtShip.Text)
If sDN = "" Then
    MsgBox "请先输入DN号", vbInformation, "提示:"
    Exit Sub

End If

sOra = "select distinct outbox_num, inbox_num  from lpstbl where boxid = '" & sRefer & "'"
If Get_OracleStr(sOra) = "" Then
    MsgBox "系统查不到此内盒的数据", vbInformation, "提示:"
    Exit Sub

End If

Set rs = Get_OracleRs(sOra)
iOp = rs!OUTBOX_NUM
iIp = rs!INBOX_NUM
sOra = "select distinct print_flag from lpstbl where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' and inbox_num = '" & iIp & "'"
If Get_OracleStr(sOra) <> "1" Then
    MsgBox "该内盒未打印, 不可补打", vbInformation, "提示:"
    Exit Sub

End If

Call PrintINNERBOXFlag(sDN, iOp, iIp)
Sleep (iSec)
Call PrintSTBoxLbl(sDN, iOp, iIp)
Sleep (iSec)
' 三星
If sType = "SSE2" Then
    Call PrintCusBoxLbl(sDN, iOp, iIp)
    Sleep (iSec)
    Call PrintREELFlag(sDN, iOp, iIp)
    Sleep (iSec)
    Call PrintCusReelLbl(sDN, iOp, iIp)
    Sleep (iSec)
ElseIf sType = "SSSHORT" Then
    Call PrintCusBoxLbl2(sDN, iOp, iIp)
    Sleep (iSec)
    Call PrintREELFlag(sDN, iOp, iIp)
    Sleep (iSec)
    Call PrintCusReelLbl2(sDN, iOp, iIp)
    Sleep (iSec)
ElseIf sType = "HW" Then
    ' 华为
    Call PrintHWBoxLbl(sDN, iOp, iIp)
    Sleep (iSec)
    Call PrintHWReelLbl(sDN, iOp, iIp)
    Sleep (iSec)

End If

Call PrintOuterCFlag(sDN, iOp)
Call PrintSTCartonLbl(sDN, iOp)
Call UpdatePrintAStatus(sDN, iOp, iIp)
MsgBox "补打成功", vbInformation, "提示:"

End Sub

'补打三星卷盘标签
Private Sub PrintA_SSREEL(sRefer As String)
Dim sOra            As String
Dim tCusReel        As CusReel
Dim sContent        As String
Dim sFileName       As String
Dim rs              As New ADODB.Recordset
Dim sPath           As String
Dim strSSType       As String
Dim strFabSite      As String
Dim strAssemblySite As String
Dim strTestSite     As String
Dim strSql          As String
Dim strDN           As String

sPath = Trim(txtCusReelPath.Text)
sFileName = "CUSREELLbl" + Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
If InStr(sRefer, "DPTKE2") > 0 Then
    sRefer = Mid(sRefer, InStr(sRefer, "DPTKE2") + 6, 10)

End If

sOra = "select * from lpstbl where reelid = '" & sRefer & "'"
If Get_OracleStr(sOra) = "" Then
    MsgBox "系统查不到此卷盘的数据", vbInformation, "提示:"
    Exit Sub

End If

sOra = "select dn_num,reelid, qty, mpn, cpn,FAB_SITE,ASSEMBLY_SITE,TEST_SITE from lpstbl where reelid = '" & sRefer & "'"
Set rs = Get_OracleRs(sOra)
strDN = Trim("" & rs!DN_NUM)
tCusReel.lot = UCase(Trim("" & rs!REELID))
tCusReel.QTY = UCase(Trim("" & rs!QTY))
tCusReel.DEV = UCase(Trim("" & rs!MPN))
tCusReel.PN = UCase(Trim("" & rs!CPN))
strFabSite = Trim$("" & rs!FAB_SITE)
strAssemblySite = Trim$("" & rs!ASSEMBLY_SITE)
strTestSite = Trim$("" & rs!TEST_SITE)
If strFabSite <> "" Then
    strFabSite = "Fab:" & strFabSite

End If

If strAssemblySite <> "" Then
    strAssemblySite = "Assembly:" & strAssemblySite

End If

If strTestSite <> "" Then
    strTestSite = "Test:" & strTestSite

End If

strSql = "select distinct Labelrequirement from customershippinguptbl where delivery = '" & strDN & "'"
If InStr(UCase(Get_OracleStr(strSql)), "E2") Then
    strSSType = "E2"
ElseIf InStr(UCase(Get_OracleStr(strSql)), "SHORT") Then
    strSSType = "SHORT"
Else
    MsgBox "三星标签类型不正确,请确定问题", vbCritical, "警告"
    Exit Sub

End If

If strSSType = "E2" Then
    sContent = sContent + tCusReel.PN + "DPTKE2" + tCusReel.lot + Right$("000000" + tCusReel.QTY, 6) + ","
    sContent = sContent + tCusReel.PN + "," + "TVS DIODES" + "," + "PO TYPE,:E2" + "," + tCusReel.lot + "," + tCusReel.QTY + ","
    sContent = sContent + tCusReel.DEV + "," + "DPTK" & "," & strFabSite & "," & strAssemblySite & "," & strTestSite
    If tCusReel.DEV = "RCLAMP2581ZCTFT" Then
        sPath = "\\10.160.1.14\BarCode\37\37BOXJP-新\"

    End If

    Call CreateTxt(sFileName, sContent, sPath)
    MsgBox "补打成功", vbInformation, "提示:"
    tReferTo.Text = ""
Else
    sContent = sContent + tCusReel.PN + "DPTK" + tCusReel.lot + Right$("000000" + tCusReel.QTY, 6) + ","
    sContent = sContent + tCusReel.PN + "," + "TVS DIODES" + "," + tCusReel.lot + "," + tCusReel.QTY + ","
    sContent = sContent + tCusReel.DEV + "," + "DPTK"
    sPath = "\\10.160.1.14\BarCode\37\37JP2\"
    sFileName = sDN & "-" & "CUSREELLbl" & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
    Call CreateTxt(sFileName, sContent, sPath)
    MsgBox "补打成功", vbInformation, "提示:"
    tReferTo.Text = ""

End If

End Sub

'补打SEMTECH内盒标签
Private Sub PrintA_STBOX(sRefer As String)
Dim sOra          As String
Dim sDatecode     As String
Dim sTestDateCode As String
Dim tSTBox        As STBox
Dim sContent      As String
Dim sFileName     As String
Dim sPath         As String
Dim rs            As New ADODB.Recordset
Dim sAdd          As String

sAdd = ""
sPath = Trim(txtSTBoxPath.Text)
sFileName = "STBoxLbl" + Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select * from lpstbl where boxid = '" & sRefer & "'"
If Get_OracleStr(sOra) = "" Then
    MsgBox "系统查不到此内盒的数据", vbInformation, "提示:"
    Exit Sub

End If

sOra = "select job_id,  mpn, cpn, boxid ,sum(qty) as qty from lpstbl where boxid = '" & sRefer & "' group by  job_id,  mpn, cpn, boxid"
Set rs = Get_OracleRs(sOra)
tSTBox.JOB = Trim("" & rs!JOB_ID)
tSTBox.DEV = Trim("" & rs!MPN)
tSTBox.lot = Trim("" & rs!BOXID)
tSTBox.QTY = Trim("" & rs!QTY)
sDatecode = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTBox.JOB & "'")
tSTBox.DATECODE = sDatecode
sTestDateCode = Get_SqlStr("select distinct HTDATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTBox.JOB & "'")
tSTBox.testdateCode = sTestDateCode
If tSTBox.DATECODE = "" Or tSTBox.testdateCode = "" Then
    tSTBox.DATECODE = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTBox.JOB & "'")
    tSTBox.testdateCode = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTBox.JOB & "'")

End If

sContent = sContent + tSTBox.DEV + "," + tSTBox.JOB + ",1T" + tSTBox.JOB + "," + tSTBox.DEV + "," + "1P" + tSTBox.DEV + "," + tSTBox.DATECODE + "," + tSTBox.DATECODE + "," + Mid(tSTBox.lot, 2) + "," + tSTBox.lot + ","
sContent = sContent + tSTBox.QTY + ",Q" + tSTBox.QTY + "," + tSTBox.testdateCode + "," + tSTBox.testdateCode
sAdd = GetDevMark(tSTBox.DEV)
sContent = sContent + sAdd
Call CreateTxt(sFileName, sContent, sPath)
MsgBox "补打成功", vbInformation, "提示:"
tReferTo.Text = ""

End Sub

' 补打三星内盒标签
Private Sub PrintA_SSBOX(sRefer As String)
Dim sOra            As String
Dim tCusBox         As CusBox
Dim sContent        As String
Dim sFileName       As String
Dim rs              As New ADODB.Recordset
Dim sPath           As String
Dim iIp             As Integer
Dim iOp             As Integer
Dim sDN             As String
Dim strSSType       As String
Dim strFabSite      As String
Dim strAssemblySite As String
Dim strTestSite     As String
Dim strSql          As String

sPath = Trim(txtCusBoxPath.Text)
sFileName = "CUSBoxLbl" + Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select distinct dn_num, inbox_num, outbox_num, FAB_SITE,ASSEMBLY_SITE,TEST_SITE  from lpstbl where boxid ='" & sRefer & "'"
If Get_OracleStr(sOra) = "" Then
    MsgBox "系统查不到此内盒的数据", vbInformation, "提示:"
    Exit Sub

End If

Set rs = Get_OracleRs(sOra)
iIp = rs!INBOX_NUM
iOp = rs!OUTBOX_NUM
sDN = rs!DN_NUM
strFabSite = Trim$("" & rs!FAB_SITE)
strAssemblySite = Trim$("" & rs!ASSEMBLY_SITE)
strTestSite = Trim$("" & rs!TEST_SITE)
If strFabSite <> "" Then
    strFabSite = "Fab:" & strFabSite

End If

If strAssemblySite <> "" Then
    strAssemblySite = "Assembly:" & strAssemblySite

End If

If strTestSite <> "" Then
    strTestSite = "Test:" & strTestSite

End If

sOra = "select cpn, mpn, sum(qty) as qty from lpstbl where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' and inbox_num = '" & iIp & "' group by cpn, mpn"
Set rs = Get_OracleRs(sOra)
tCusBox.QTY = Trim$("" & rs!QTY)
tCusBox.DEV = Trim("" & rs!MPN)
tCusBox.PN = Trim("" & rs!CPN)
strSql = "select distinct Labelrequirement from customershippinguptbl where delivery = '" & sDN & "'"
If InStr(UCase(Get_OracleStr(strSql)), "E2") Then
    strSSType = "E2"
ElseIf InStr(UCase(Get_OracleStr(strSql)), "SHORT") Then
    strSSType = "SHORT"
Else
    MsgBox "三星标签类型不正确,请确定问题", vbCritical, "警告"
    Exit Sub

End If

If strSSType = "E2" Then
    sContent = sContent + tCusBox.PN + "DPTKE2" + Right$("000000" + tCusBox.QTY, 6) + ","
    sContent = sContent + tCusBox.PN + "," + "TVS DIODES" + "," + "PO TYPE,:E2" + "," + tCusBox.QTY + "," + tCusBox.DEV + "," + "DPTK" & "," & strFabSite & "," & strAssemblySite & "," & strTestSite
    If tCusBox.DEV = "RCLAMP2581ZCTFT" Then
        sPath = "\\10.160.1.14\BarCode\37\37BoxNH-新\"

    End If

    Call CreateTxt(sFileName, sContent, sPath)
    MsgBox "补打成功", vbInformation, "提示:"
    tReferTo.Text = ""
Else
    sPath = "\\10.160.1.14\BarCode\37\37NH2\"
    sFileName = sDN & "-" & "CUSBoxLbl2" & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
    sContent = sContent + tCusBox.PN + "DPTK" + Right$("000000" + tCusBox.QTY, 6) + ","
    sContent = sContent + tCusBox.PN + "," + "TVS DIODES" + "," + tCusBox.QTY + "," + tCusBox.DEV + "," + "DPTK"
    If tCusBox.DEV = "RCLAMP2581ZCTFT" Then
        MsgBox "请联系IT确认标签模板", vbCritical, "警告"
        Exit Sub

    End If

    Call CreateTxt(sFileName, sContent, sPath)
    MsgBox "补打成功", vbInformation, "提示:"
    tReferTo.Text = ""

End If

End Sub

' 补打SEMTECH外箱小标签
Private Sub PrintA_STCARTONSUB(sRefer As String)
Dim sOra          As String
Dim tSTCarton     As STCarton
Dim sContent      As String
Dim sFileName     As String
Dim sDatecode     As String
Dim sTestDateCode As String
Dim rs            As New ADODB.Recordset
Dim sPath         As String
Dim sAdd          As String

sAdd = ""
sPath = Trim(txtSTBoxPath.Text)
sFileName = "STCARTONLBL" + Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select * from lpstbl where cartonid ='" & sRefer & "'"
If Get_OracleStr(sOra) = "" Then
    MsgBox "系统查不到此外箱的数据", vbInformation, "提示:"
    Exit Sub

End If

sOra = " select job_id,CUSTOMER_DEVICE,cartonid, sum(qty) as qty from PACKING_DETAILED where cartonid = '" & sRefer & "' group by job_id,CUSTOMER_DEVICE,cartonid"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tSTCarton.JOB = Trim("" & rs!JOB_ID)
        tSTCarton.DEV = Trim("" & rs!Customer_Device)
        tSTCarton.lot = Trim("" & rs!CARTONID)
        tSTCarton.QTY = Trim("" & rs!QTY)
        sDatecode = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTCarton.JOB & "'")
        tSTCarton.DATECODE = sDatecode
        sTestDateCode = Get_SqlStr("select distinct HTDATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTCarton.JOB & "'")
        tSTCarton.testdateCode = sTestDateCode
        If tSTCarton.DATECODE = "" Or tSTCarton.testdateCode = "" Then
            tSTCarton.DATECODE = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTCarton.JOB & "'")
            tSTCarton.testdateCode = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTCarton.JOB & "'")

        End If

        sFileName = "SEMT_CARTON_" + tSTCarton.lot
        sContent = sContent + tSTCarton.DEV + "," + tSTCarton.JOB + ",1T" + tSTCarton.JOB + "," + tSTCarton.DEV + "," + "1P" + tSTCarton.DEV + "," + tSTCarton.DATECODE + "," + tSTCarton.DATECODE + "," + Mid(tSTCarton.lot, 2) + "," + tSTCarton.lot + ","
        sContent = sContent + tSTCarton.QTY + ",Q" + tSTCarton.QTY + "," + tSTCarton.testdateCode + "," + tSTCarton.testdateCode
        sAdd = GetDevMark(tSTCarton.DEV)
        sContent = sContent + sAdd
        rs.MoveNext
    Loop

End If

Call CreateTxt(sFileName, sContent, sPath)
MsgBox "补打成功", vbInformation, "提示:"
tReferTo.Text = ""

End Sub

' 补打外箱大标签
Private Sub PrintAOut()
Dim sDN    As String
Dim sOra   As String
Dim iOp    As Integer
Dim iOpMax As Integer
Dim iSec   As Integer
Dim sType  As String

cmdPrintCarton.Enabled = False
iSec = (txtPrintInterval.Text) * 1000
sType = Trim(txtShip.Text)
sDN = txtDN2.Text
If sDN = "" Then
    MsgBox "DN不可为空", vbInformation
    Exit Sub

End If

sOra = "select UPPER(labelrequirement) as type from CUSTOMERSHIPPINGUPTBL where delivery = '" & sDN & "'"
sType = Get_OracleStr(sOra)
If InStr(sType, "E2") > 0 Then
    sType = "SSE2"

End If

If InStr(sType, "SHORT") > 0 Then
    sType = "SSSHORT"

End If

If InStr(sType, "HUAWEI") > 0 Then
    sType = "HW"

End If

If InStr(sType, "SEMTECH") > 0 Then
    sType = "ST"

End If

If Get_OracleStr("select * from PACKING_DETAILED where dn_num = '" & sDN & "' ") = "" Then
    MsgBox "系统无此标签数据", vbInformation, "提示:"
    Exit Sub

End If

sOra = "select outbox_num from PACKING_DETAILED where dn_num = '" & sDN & "' and  CARTON = '" & Trim(tReferTo.Text) & "'"
If Get_OracleCnt(sOra) > 0 Then
    iOp = Get_OracleNo(sOra)
    If sType = "SSE2" Then
        Call PrintCusCartonLbl(sDN, iOp)
        Sleep (iSec)
    ElseIf sType = "SSSHORT" Then
        Call PrintCusCartonLbl2(sDN, iOp)
        Sleep (iSec)
    Else
        Call PrintSTCartonStanderLbl(sDN, iOp)
        Sleep (iSec)

    End If

End If

MsgBox "外箱大标签补打完成", vbInformation, "友情提示:"

End Sub

' 更新打印状态
Private Sub UpdatePrintStatus(sDN As String, iOp As Integer)
Dim sOra As String

sOra = "update PACKING_DETAILED set print_flag = '1' where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "'"
Exec_Ora (sOra)

End Sub

' 更新补打状态
Private Sub UpdatePrintAStatus(sDN As String, iOp As Integer, iIp As Integer)
Dim sOra As String

sOra = "update PACKING_DETAILED set print_flag = '2' where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' and inbox_num = '" & iIp & "'"
Exec_Ora (sOra)

End Sub

'Private Function Get37TestDC(strJob As String, strTrayID As String) As String
'Dim strWaferID  As String
'Dim strDateCode As String
'Dim strSql      As String
'Dim strJobNew   As String
'Dim strContent  As String
'Dim str1 As String
'Dim strBartenName As String
'Dim strTestDateCode As String
'
'strJobNew = Replace$(strJob, "M", "")
'strSql = "select distinct case when create_date >= to_date(to_char(create_date, 'yyyy') || '-12-31', 'yyyy-mm-dd') - mod(to_char(create_date, 'YYYY'), 7) - 5  then to_char(create_date, 'yyww') " & "else to_char(create_date + mod(mod(to_char(create_date, 'YYYY'), 7) + 5, 7),'yyww') end as PODATECODE " & "from customeroitbl_test a ,mappingdatatest b ,weight37 c where a.test_mtrl_desc = '" & strJobNew & "' and b.filename = to_char(a.id) and b.lotid = a.source_batch_id " & "and c.waferid = replace(b.substrateid,'+','') "
'strDateCode = Get_OracleStr(strSql)
'
'strSql = "select TT.TEST_DATECODE from (  " & _
'" select t4.TEST_MTRL_DESC as JOBID,min(DATE_CODE_CONVERT.DC_CONVERT(to_char(t2.ERPCREATEDATE, 'YYYY-MM-DD'),1)) as TEST_DATECODE " & _
'" from ib_waferlist t1 " & _
'" inner join ib_workorder t2 on t1.ORDERNAME = t2.ORDERNAME " & _
'" inner join mappingdatatest t3 on t3.SUBSTRATEID = t1.WAFERID and t3.LOTID = t1.WAFERLOT " & _
'" inner join  customeroitbl_test t4 on t4.SOURCE_BATCH_ID = t3.LOTID  and to_char(t4.ID) = t3.FILENAME " & _
'" where t4.TEST_MTRL_DESC in ('" & strJobNew & "')  group by t4.TEST_MTRL_DESC  ) TT where TT.TEST_DATECODE >=1929 "
'
'strTestDateCode = Get_OracleStr(strSql)
'
'If strTestDateCode <> "" And strTestDateCode <> strDateCode Then
'
'    Get37TestDC = strTestDateCode
'Else
'    Get37TestDC = strDateCode
'End If
'
'End Function



'获取标签序号
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
