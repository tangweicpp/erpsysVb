VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form MyLabelSystem 
   BackColor       =   &H00C0C0C0&
   Caption         =   "LPS[标签打印系统]"
   ClientHeight    =   12660
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21480
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
   ScaleHeight     =   12660
   ScaleWidth      =   21480
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
      TabPicture(0)   =   "MyLabelSystem.frx":0000
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
      Tab(0).Control(20)=   "fpsJob"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "fpsPN"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtSF"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtDN"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtPO"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtQty"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtPN"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtReelID"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtStatus"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "D"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtSTBoxPath"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtCusBoxPath"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtCusReelPath"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtCusCartonPath"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtHTCartonPath"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtFlagPath"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtShip"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtHWReelPath"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtMediaPath"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtSEMTCartonPAth"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cmdDNClear"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).ControlCount=   41
      TabCaption(1)   =   "标签补打"
      TabPicture(1)   =   "MyLabelSystem.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chk"
      Tab(1).Control(1)=   "cmdPrint"
      Tab(1).Control(2)=   "tReferTo"
      Tab(1).Control(3)=   "cbType"
      Tab(1).Control(4)=   "lbl"
      Tab(1).Control(5)=   "lblType"
      Tab(1).ControlCount=   6
      Begin VB.CommandButton cmdDNClear 
         Caption         =   "清除DN信息"
         Height          =   360
         Left            =   6120
         TabIndex        =   56
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtSEMTCartonPAth 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "\\10.160.1.14\BarCode\37\37外箱\"
         Top             =   8040
         Width           =   3375
      End
      Begin VB.CheckBox chk 
         Caption         =   "手动输入"
         Height          =   375
         Left            =   -68760
         TabIndex        =   53
         Top             =   2160
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
         Height          =   825
         Left            =   -66360
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2040
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox tReferTo 
         Height          =   285
         Left            =   -72720
         TabIndex        =   51
         Top             =   2138
         Width           =   2895
      End
      Begin VB.ComboBox cbType 
         Height          =   315
         Left            =   -72720
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtMediaPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         TabIndex        =   46
         Text            =   "C:\media_source\"
         Top             =   7200
         Width           =   3375
      End
      Begin VB.TextBox txtHWReelPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "\\10.160.1.14\BarCode\37\37HW\HW卷盘\"
         Top             =   6360
         Width           =   3375
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
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "\\10.160.1.14\BarCode\37\37Flag\"
         Top             =   5580
         Width           =   3375
      End
      Begin VB.TextBox txtHTCartonPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "\\10.160.1.14\BarCode\37\37BOX\"
         Top             =   4740
         Width           =   3375
      End
      Begin VB.TextBox txtCusCartonPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "\\10.160.1.14\BarCode\37\37BoxOut\"
         Top             =   3900
         Width           =   3375
      End
      Begin VB.TextBox txtCusReelPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "\\10.160.1.14\BarCode\37\37BoxJP\"
         Top             =   3156
         Width           =   3375
      End
      Begin VB.TextBox txtCusBoxPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "\\10.160.1.14\BarCode\37\37BoxNH\"
         Top             =   2388
         Width           =   3375
      End
      Begin VB.TextBox txtSTBoxPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18000
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "\\10.160.1.14\BarCode\37\37内箱\"
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
            TabIndex        =   44
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
            BuddyDispid     =   196626
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
         Locked          =   -1  'True
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
         SpreadDesigner  =   "MyLabelSystem.frx":0038
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
         SpreadDesigner  =   "MyLabelSystem.frx":04D2
         TextTip         =   2
         AppearanceStyle =   0
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
         TabIndex        =   55
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
         Left            =   -73800
         TabIndex        =   50
         Top             =   2160
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
         Left            =   -73920
         TabIndex        =   48
         Top             =   1237
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
         TabIndex        =   47
         Top             =   6960
         Width           =   1095
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   18000
         TabIndex        =   45
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
Attribute VB_Name = "MyLabelSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim glReelListCnt As Long

Dim gsLastPN      As String

Dim gsLastJob     As String

Private Sub chk_Click()

    If chk.Value = 1 Then
        cmdPrint.Visible = True
    Else
        cmdPrint.Visible = False
    End If

End Sub

Private Sub cmdDNClear_Click()

    Dim QboxNumber As String

    Dim rs As New ADODB.Recordset

    If txtdn.Text = "" Then
        MsgBox "请扫入DN号", vbInformation, "提示"
        Exit Sub

    End If
    
    If MsgBox("你确认要删除吗?", vbOKCancel, "提示") = vbCancel Then
        Exit Sub
    End If
    
    Set rs.ActiveConnection = OraConnect
    rs.Source = "select distinct carton from packing_detailed where dn_num = '" & Trim(txtdn.Text) & "'  and carton is not null"

    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        
        rs.MoveFirst

        For I = 1 To rs.RecordCount
            QboxNumber = Trim(rs(0))
             
            FrmDelERPQbox.Text1 = QboxNumber
            
            Call FrmDelERPQbox.Command1_Click
            
            rs.MoveNext
        Next I

    Else
        'MsgBox "查询不到该DN信息, 请确认", vbCritical, "警告"

       ' Exit Sub

    End If

    rs.Close
    Set rs = Nothing
    
    AddSql ("insert into packing_detailed_bak select * from packing_detailed where dn_num = '" & txtdn.Text & "'")
    AddSql ("delete from packing_detailed where dn_num = '" & txtdn.Text & "'")

    MsgBox "备份删除完成", vbInformation, "提示"

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExit1_Click()
    Unload Me
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
    txtMediaPath.Text = "\\10.160.1.31\software\media_source\"

    cmdCombine.Enabled = True
    cmdPrintReel_Box.Enabled = True
    cmdPrintCarton.Enabled = True
End Sub






Private Sub Form_Activate()
    txtSF.SetFocus
End Sub

Private Sub Form_Load()
    Call initData
    Call InitFps
    Call InitStatus
   ' Call DelTmpTbl
End Sub

Private Sub InitShip()

    Dim sType As String

    Dim sOra  As String

    Dim sDN   As String

    sDN = Trim(txtdn.Text)

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

End Sub

Private Sub initData()

    ' 标签打印
    If gUserName = "07885" Then
        cmdTest.Visible = True
    End If

    glReelListCnt = 0
    gsLastPN = ""
    gsLastJob = ""

    If gUserName <> "07885" And gUserName <> "10354" Then
        cmdPrint.Enabled = False
        tReferTo.Visible = False
        cmdDNClear.Visible = False
        
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

    If KeyAscii <> vbKeyReturn Then

        Exit Sub

    End If

    Dim sRefer As String

    Dim sType  As String

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

    tReferTo.Text = ""
End Sub

Rem: For Scanning
    Private Sub txtSF_KeyPress(KeyAscii As Integer)

        If KeyAscii <> vbKeyReturn Then

            Exit Sub

        End If

        If txtdn.Text = "" Then
            Call ForDN
        Else
            Call ForReelID
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
    
    txtdn.Text = sDN
    Play ("DN_OK")

    Call InitShip
    Call ShowDNInfo(sDN)
    Call ShowFps(sDN)
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

    CheckDN = True
End Function

Private Function IsExist(sDN As String) As Boolean

    Dim sOra As String

    sOra = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & sDN & "'"
    IsExist = IsOraRecord(sOra)
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

    txtdn.Text = sDN
    txtpo.Text = rs!PO
    txtqty.Text = rs!qty
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

    Dim rs As New ADODB.Recordset

    sOra = " select AA.marketingpn,AA.realqtys, BB.thisqtys from " & _
  "  (select marketingpn, sum(quantity) as realqtys " & _
  "       from CUSTOMERSHIPPINGUPTBL " & _
  "       where delivery = '" & sDN & "' " & _
  "       group by marketingpn) AA  " & _
  "  left join (select dev, sum(qty) as thisqtys " & _
  "       from ST_TR_SEQ " & _
  "       where dn = '" & sDN & "' " & _
  "       group by dev) BB " & _
  "       on AA.marketingpn = BB.dev "
    Set rs = Get_OracleRs(sOra)

    With fpsPN
        .MaxRows = 0

        If rs.RecordCount > 0 Then
            Set .DataSource = rs
        End If

    End With

     sOra = " select AA.batchnumber,AA.realqtys, BB.thisqtys from " & _
  "  (select batchnumber, sum(quantity) as realqtys " & _
  "       from CUSTOMERSHIPPINGUPTBL " & _
  "       where delivery = '" & sDN & "' " & _
  "       group by batchnumber) AA  " & _
  "  left join (select job, sum(qty) as thisqtys " & _
  "       from ST_TR_SEQ " & _
  "       where dn = '" & sDN & "' " & _
  "       group by job) BB " & _
  "       on AA.batchnumber = BB.job "

    Set rs = Get_OracleRs(sOra)

    With fpsJob
        .MaxRows = 0

        If rs.RecordCount > 0 Then
            Set .DataSource = rs
        End If

    End With

    rs.Close
End Sub

Rem: Scanning ReelID
Private Sub ForReelID()

    Dim sScan    As String

    Dim sKeyWord As String

    Dim sReelID  As String

    Dim sDN      As String

    sScan = UCase(Trim$(txtSF.Text))
    sKeyWord = Left$(sScan, 1)

    If sKeyWord <> "S" Then
        'MsgBox "没有扫描到卷盘ID", vbInformation, "友情提示!!!"
        Play ("noReel")

        Exit Sub

    End If

    sReelID = sScan
    sDN = txtdn.Text

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

    If InStr(txtReelID, sReelID) Then
        txtStatus.BackColor = vbRed
        Play ("noRep")

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

    sOra = "select * from CUSTOMERSHIPPINGUPTBL where batchnumber = '" & sJob & "' and delivery = '" & txtdn.Text & "'"
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

    sSql2 = "SELECT y.TEST_MTRL_DESC + case  SUBSTRING(KEY_VALUE,CHARINDEX('M-',KEY_VALUE,1),1) when 'M' then 'M' else '' end " & " FROM erpdata..tblErpInStockDetailInfo c ,erpdata..tblErpInStockRelation b  " & " ,ERPBASE..tblmappingData x,tblCustomerOI y  " & " where c.KEY_NAME ='CONTAINER_NAME'  " & " and charindex('" & sReelID & "', c.KEY_VALUE) > 0 " & " and b.BOX_ID = c.BOX_ID  " & " and SUBSTRING(c.KEY_VALUE,2,8) = SUBSTRING(REPLACE(B.SFC_ID,'SFCBO:1020,',''),1,8)  " & " and SUBSTRING( replace(b.WAFER_ID,b.SFC_ID+',',''),1,CHARINDEX('::',replace(b.WAFER_ID,b.SFC_ID+',',''),1)-1) = x.SUBSTRATEID  " & " and convert(varchar(100),y.id) = x.FILENAME"

    If Len(sReelID) = 13 Or Len(sReelID) = 14 Then
        tOra.sDN = txtdn.Text
        tOra.sReelID = sReelID
        tOra.sJob = Trim(Get_SqlStr(sSql2))
        tOra.sLot = Mid$(sReelID, 2, Len(sReelID) - 5)
        tOra.lQty = Get_SqlserverNo("SELECT z.KEY_VALUE " & " FROM erpdata..tblErpInStockDetailInfo c ,erpdata..tblErpInStockRelation b " & "  ,ERPBASE..tblmappingData x,tblCustomerOI y,erpdata..tblErpInStockDetailInfo z " & " where c.KEY_NAME ='CONTAINER_NAME' " & " and charindex('" & sReelID & "', c.KEY_VALUE) > 0  " & " and b.BOX_ID = c.BOX_ID " & " and SUBSTRING(c.KEY_VALUE,2,8) = SUBSTRING(REPLACE(B.SFC_ID,'SFCBO:1020,',''),1,8) " & " and SUBSTRING( replace(b.WAFER_ID,b.SFC_ID+',',''),1,CHARINDEX('::',replace(b.WAFER_ID,b.SFC_ID+',',''),1)-1) = x.SUBSTRATEID " & " and convert(varchar(100),y.id) = x.FILENAME " & " and z.BOX_ID = c.BOX_ID " & " and z.KEY_TYPE = 'T' " & " and z.KEY_NAME = 'GOOD_DIE'")

        tOra.lSeq = glReelListCnt + 1
    Else
        sSql = "select customerlotid as jobid, htlotid as lotid, qty from [erpdata].[dbo].[TblTSV_Tray_details] where TRAYQBOXNUMBER = '" & sReelID & "'"
        Set rs = Get_SqlserveRs(sSql)

        If rs.RecordCount = 0 Then
            '   MsgBox "系统TblTSV_Tray_details无此卷盘数据", vbInformation, "友情提示!!!"
    
            rs.Close

            Exit Function

        End If

        If IsNull(rs!JobID) Or IsNull(rs!LOTID) Or IsNull(rs!qty) Then
            '   MsgBox "系统TblTSV_Tray_details,无job信息", vbInformation, "友情提示!!!"
    
            rs.Close

            Exit Function

        End If
    
        tOra.sDN = txtdn.Text
        tOra.sReelID = sReelID
        tOra.sJob = Trim(rs!JobID)
        tOra.sLot = Trim$(rs!LOTID)
        tOra.lQty = rs!qty
        tOra.lSeq = glReelListCnt + 1
    End If

    sOra = "select marketingpn as mpn from CUSTOMERSHIPPINGUPTBL where delivery = '" & txtdn.Text & "' and batchnumber = '" & tOra.sJob & "'"
    Set rs = Get_OracleRs(sOra)

    If rs.RecordCount = 0 Then
        MsgBox "DN无机种信息", vbInformation, "友情提示!!!"
    
        rs.Close

        Exit Function

    End If

    tOra.sDev = Trim(rs!MPN)

    If gsLastPN <> "" Then
        sOra = "select sum(qty) as qty from ST_TR_SEQ where dev = '" & gsLastPN & "'  and dn = '" & txtdn.Text & "'"
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
        sOra = "select sum(qty) as qty from ST_TR_SEQ where job = '" & gsLastJob & "' and dn = '" & txtdn.Text & "' "
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
    Exec_Ora (sOra)

    ' 备份
    sOra = "insert into ST_TR_SEQ_BACK values('" & tOra.sDN & "', '" & tOra.sJob & "', '" & tOra.sDev & "', '" & tOra.lQty & "', sysdate, '" & tOra.sLot & "','" & tOra.sReelID & "', '" & tOra.lSeq & "' )"
    Exec_Ora (sOra)

    InsertReelIDToTmpTbl = True
End Function

Private Sub CheckCurrentQty(sReelID As String)

    Dim lTotalQty   As Long

    Dim lCurrentQty As Long

    Dim I           As Integer

    Dim bFinish     As Boolean

    bFinish = True

    With fpsJob

        For I = 1 To .MaxRows
            .Row = I
        
            .Col = 2
            lTotalQty = Val(.Text)
        
            .Col = 3
            lCurrentQty = Val(.Text)
        
            If lCurrentQty > lTotalQty Then
                Play ("wrongCnt")
                MsgBox "已经超出所需数量, 挑料出错", vbInformation, "友情提示!!!"
            
                Exec_Ora ("delete from ST_TR_SEQ where REELID = '" & sReelID & "'")
                Call ShowFps(txtdn.Text)
            
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

    If txtdn.Text = "" Then
        MsgBox "DN不可为空", vbInformation, "友情提示!!!"

        Exit Sub

    End If

    sOra = "select * from ST_TR_SEQ where dn = '" & txtdn.Text & "' order by seq"

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

    Call MakeBoxID(txtdn.Text)

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

    tData.TRAYID = rs.Fields("REELID")
    tData.CREATE_BY = gUserName
    tData.DN_NUM = rs.Fields("DN")
    tData.JOB_ID = rs.Fields("JOB")
    tData.CUSTOMER_DEVICE = rs.Fields("DEV")

    tData.qty = rs.Fields("QTY")
    tData.REEL_ID = Replace(tData.JOB_ID, "M", "") & GetLableXHTW(tData.JOB_ID)
    tData.seq = rs.Fields("SEQ")

    If Right(rs.Fields("LOTID"), 1) = "M" Then
        tData.REEL_ID = tData.JOB_ID & Right$(tData.TRAYID, 1)
    End If

    ' 求datecode
    Dim sWaferID  As String

    Dim sDatecode As String

    sWaferID = Get_SqlStr("select waferid from [erpdata].[dbo].[v_job_m] j where charindex( '" & tData.TRAYID & "',j.KEY_VALUE ) > 0 ")
    sDatecode = Get_OracleStr("select to_char(aa.create_date+1,'YYWW') from weight37 aa where aa.waferid = '" & sWaferID & "'")

    sOra = "insert into PACKING_DETAILED values('" & tData.TRAYID & "','" & tData.INBOX_NUM & "','" & tData.OUTBOX_NUM & "','" & tData.DN_NUM & "','" & tData.JOB_ID & "','" & tData.qty & "','" & tData.CUSTOMER_DEVICE & "',sysdate,'" & tData.CREATE_BY & "','0','0','', '" & tData.REEL_ID & "','','','', '" & tData.seq & "', '" & sDatecode & "') "
    Exec_Ora (sOra)

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

    ' 删除M
    sOra = "select max(outbox_num) from PACKING_DETAILED where dn_num = '" & sDN & "'  "
    iOpMax = Get_OracleNo(sOra)

    For iOp = 1 To iOpMax

        sOra = "select max(inbox_num) from PACKING_DETAILED where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "'"
        iIpMax = Get_OracleNo(sOra)
       
        sOra = "update PACKING_DETAILED set kid = 'K'||'" & iOp & "' where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "'"
        Exec_Ora (sOra)
       
        Rem: Make C ID
        sAppend = "select a.job_id, b.lotid, sum(a.Qty) from PACKING_DETAILED a, ST_TR_SEQ b where a.dn_num = '" & txtdn.Text & "'  and b.dn = '" & txtdn.Text & "'  and a.outbox_num = '" & iOp & "' and b.job = a.job_id group by a.job_id, b.lotid"
        Set rs = Get_OracleRs(sAppend)

        If Not rs.BOF Then
        
            rs.MoveFirst

            Do While Not rs.EOF
                sLot = "S" & rs.Fields("lotid") & "-C"
                
                sOra = "select max(CARTONID) from PACKING_DETAILED where CARTONID like '" & sLot & "%' "
                sLastLot = Get_OracleStr(sOra)

                If sLastLot = "" Then
                    sLotBox = sLot & "01"
                Else
                    iSeq = Val(Mid$(sLastLot, InStr(sLastLot, "-C") + 2))
                    
                    sLotBox = sLot & Right("0" & (iSeq + 1), 2)
                    
                End If

                sOra = "update PACKING_DETAILED set CARTONID = '" & sLotBox & "' where dn_num = '" & txtdn.Text & "' and outbox_num = '" & iOp & "' and job_id = '" & rs.Fields("job_id") & "'"
                Exec_Ora (sOra)
                rs.MoveNext
            Loop

        End If
        
        rs.Close

        Rem: Make B ID

        For iIp = 1 To iIpMax
        
            sAppend = "select a.job_id, b.lotid, sum(a.Qty) from PACKING_DETAILED a, ST_TR_SEQ b where a.dn_num = '" & txtdn.Text & "' and b.dn = '" & txtdn.Text & "' and a.outbox_num = '" & iOp & "' and a.inbox_num = '" & iIp & "' and b.job = a.job_id group by a.job_id, b.lotid"
        
            Set rs = Get_OracleRs(sAppend)
        
            If Not rs.BOF Then
        
                rs.MoveFirst

                Do While Not rs.EOF
                    sLot = "S" & rs.Fields("lotid")
                
                    sOra = "select  '-B'|| substr('00'||(nvl(max(a.seqtxt),0)+1),-2)  from TSV_QBOXTBL_37SEQ a where a.firtname = '" & rs.Fields("lotid") & "' group by a.firtname"
                    iBoxIDSeq = Get_OracleStr(sOra)

                    If iBoxIDSeq = "" Then
                        sLotBox = sLot & "-B01"
                    Else
                        sLotBox = sLot & iBoxIDSeq
                    End If
               
                    ' 更新进数据表
                    sOra = "update PACKING_DETAILED set BOXID = '" & sLotBox & "' where dn_num = '" & txtdn.Text & "' and outbox_num = '" & iOp & "' and inbox_num = '" & iIp & "' and job_id = '" & rs.Fields("job_id") & "'"
                    Exec_Ora (sOra)
                
                    ' 更新进序列表
                    stqtpj = Right$(sLotBox, 2)
                    sOra = "insert into TSV_QBOXTBL_37SEQ(typename,createdate,seqtxt,containername,Firtname) values('INQbox',sysdate,'" & stqtpj & "','','" & rs.Fields("lotid") & "')"
                    Exec_Ora (sOra)
                
                    rs.MoveNext
                Loop
                
            End If
        
            rs.Close
        Next
    
        Rem: Make Q ID
        sFirstBoxId = Get_OracleStr("select boxid from PACKING_DETAILED where dn_num = '" & txtdn.Text & "' and outbox_num='" & iOp & "' and inbox_num = '1'")

        If sFirstBoxId <> "" Then
            sQboxNo = Get_OracleStr("select  trglabelseq.QTSeq_NotMesQbox('" & sFirstBoxId & "')  from dual")
            sOra = "update PACKING_DETAILED set CARTON = '" & sQboxNo & "' where dn_num = '" & txtdn.Text & "' and outbox_num = '" & iOp & "' "
            Exec_Ora (sOra)
        
            lQty = Get_OracleNo("select sum(qty) from PACKING_DETAILED where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "'")
            Call TransToErp(sQboxNo, lQty)
        End If

    Next

End Sub

Rem: Print Inner start
Private Sub cmdPrintReel_Box_Click()

    Dim sOra      As String

    Dim sDN       As String

    Dim iOpMax    As Integer

    Dim iIpMax    As Integer

    Dim iOp       As Integer

    Dim iIp       As Integer

    Dim iSec As Integer

    Dim sType     As String

    Dim rs        As New ADODB.Recordset

    cmdPrintReel_Box.Enabled = False

    sType = Trim(txtShip.Text)
    cmdPrintReel_Box.Enabled = False
    iSec = txtPrintInterval.Text * 1000

    sDN = txtdn.Text

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
            tSTBox.Job = Trim(rs!JOB_ID)
            tSTBox.DEV = Trim(rs!CUSTOMER_DEVICE)
            tSTBox.lot = Trim(rs!BoxID)
            tSTBox.qty = Trim(rs!qty)
    
            sDatecode = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTBox.Job & "'")
            tSTBox.DATECODE = sDatecode
        
            sTestDateCode = Get_SqlStr("select distinct HTDATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTBox.Job & "'")
            tSTBox.testdateCode = sTestDateCode
        
            If tSTBox.DATECODE = "" Or tSTBox.testdateCode = "" Then
            
                tSTBox.DATECODE = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTBox.Job & "'")
                tSTBox.testdateCode = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTBox.Job & "'")
            End If
        
            sContent = sContent + tSTBox.DEV + "," + tSTBox.Job + ",1T" + tSTBox.Job + "," + tSTBox.DEV + "," + "1P" + tSTBox.DEV + "," + tSTBox.DATECODE + "," + tSTBox.DATECODE + "," + Mid(tSTBox.lot, 2) + "," + tSTBox.lot + ","
            sContent = sContent + tSTBox.qty + ",Q" + tSTBox.qty + "," + tSTBox.testdateCode + "," + tSTBox.testdateCode
        
            sAdd = ""
        
            sAdd = GetDevMark(tSTBox.DEV)
        
            sContent = sContent + sAdd + vbCrLf
        
            rs.MoveNext
        Loop

    End If

    Call CreateTxt(sFileName, sContent, sPath)
End Sub

Rem: Step 2 :Cus Box Lable
Private Sub PrintCusBoxLbl(sDN As String, iOp As Integer, iIp As Integer)

    Dim sOra      As String

    Dim tCusBox   As CusBox

    Dim sContent  As String

    Dim sFileName As String

    Dim rs        As New ADODB.Recordset

    Dim sPath     As String

    sPath = Trim(txtCusBoxPath.Text)
    sFileName = sDN & "-" & "CUSBoxLbl" & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
    sContent = ""

    sOra = "select  sum(a.qty) as qty, a.CUSTOMER_DEVICE, b.customerpartnumber from PACKING_DETAILED a , CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & sDN & "' and b.delivery = '" & sDN & "' and a.job_id = b.batchnumber and a.outbox_num = '" & iOp & "' and a.inbox_num = '" & iIp & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber"
    Set rs = Get_OracleRs(sOra)

    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            tCusBox.qty = Trim(rs!qty)
            tCusBox.DEV = Trim(rs!CUSTOMER_DEVICE)
            tCusBox.PN = Trim(rs!customerpartnumber)
        
            sContent = sContent + tCusBox.PN + "DPTKE2" + Right$("000000" + tCusBox.qty, 6) + ","
            sContent = sContent + tCusBox.PN + "," + "TVS DIODES" + "," + "PO TYPE,:E2" + "," + tCusBox.qty + "," + tCusBox.DEV + "," + "DPTK"
    
            rs.MoveNext
        Loop
                
    End If
    
    If tCusBox.DEV = "RCLAMP2581ZCTFT" Then
        sPath = "\\10.160.1.14\BarCode\37\37BoxNH-新\"
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
            tHWBox.CPN = Trim(rs!CPN)
            tHWBox.MPN = Trim(rs!MPN)
            tHWBox.lot = Trim(rs!JOB_ID)
            tHWBox.qty = rs!qty
            tHWBox.PODATE = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tHWBox.lot & "'")
        
            sContent = sContent + tHWBox.CPN + ",P" + tHWBox.CPN + "," + tHWBox.MPN + ",1P" + tHWBox.MPN + "," + tHWBox.qty + ",Q" + tHWBox.qty + ","
            sContent = sContent + tHWBox.PODATE + ",10D" + tHWBox.PODATE + "," + "CN" + "," + "VCN" + "," + "601024" + "," + "4L601024" + "," + tHWBox.lot + ",1T" + tHWBox.lot + ","
            sContent = sContent + tHWBox.CPN + ";" + tHWBox.MPN + ";" + tHWBox.qty + ";" + tHWBox.PODATE + ";" + "CN" + ";" + "601024" + ";" + tHWBox.lot + vbCrLf
        
            rs.MoveNext
        Loop
                
    End If

    Call CreateTxt(sFileName, sContent, sPath)
End Sub

Rem: Step 3 :Cus Reel Lable
Private Sub PrintCusReelLbl(sDN As String, iOp As Integer, iIp As Integer)

    Dim sOra      As String

    Dim tCusReel  As CusReel

    Dim sContent  As String

    Dim sFileName As String

    Dim rs        As New ADODB.Recordset

    Dim rs2       As New ADODB.Recordset

    Dim sPath     As String

    sPath = Trim(txtCusReelPath.Text)
    sFileName = sDN & "-" & "CUSREELLbl" & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
    sContent = ""

    sOra = "select trayid, reelid, qty, CUSTOMER_DEVICE, cpn from lpstbl where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' and inbox_num = '" & iIp & "'  order by seq "
    Set rs2 = Get_OracleRs(sOra)

    If Not rs2.BOF Then
        rs2.MoveFirst

        Do While Not rs2.EOF
            tCusReel.TRAYID = Trim$(rs2!TRAYID)
            tCusReel.lot = Trim$(rs2!ReelID)
            tCusReel.qty = Trim(rs2!qty)
            tCusReel.DEV = Trim(rs2!CUSTOMER_DEVICE)
            tCusReel.PN = Trim(rs2!CPN)

            sContent = sContent + tCusReel.PN + "DPTKE2" + tCusReel.lot + Right$("000000" + tCusReel.qty, 6) + ","
            sContent = sContent + tCusReel.PN + "," + "TVS DIODES" + "," + "PO TYPE,:E2" + "," + tCusReel.lot + "," + tCusReel.qty + ","
            sContent = sContent + tCusReel.DEV + "," + "DPTK" + vbCrLf

            rs2.MoveNext
        Loop

    End If

    If tCusReel.DEV = "RCLAMP2581ZCTFT" Then
        sPath = "\\10.160.1.14\BarCode\37\37BOXJP-新\"
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
            tHWBox.CPN = Trim(rs!CPN)
            tHWBox.MPN = Trim(rs!MPN)
            tHWBox.lot = Trim(rs!JOB_ID)
            tHWBox.qty = rs!qty
            tHWBox.PODATE = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tHWBox.lot & "'")
        
            sContent = sContent + tHWBox.CPN + ",P" + tHWBox.CPN + "," + tHWBox.MPN + ",1P" + tHWBox.MPN + "," + tHWBox.qty + ",Q" + tHWBox.qty + ","
            sContent = sContent + tHWBox.PODATE + ",10D" + tHWBox.PODATE + "," + "CN" + "," + "VCN" + "," + "601024" + "," + "4L601024" + "," + tHWBox.lot + ",1T" + tHWBox.lot + ","
            sContent = sContent + tHWBox.CPN + ";" + tHWBox.MPN + ";" + tHWBox.qty + ";" + tHWBox.PODATE + ";" + "CN" + ";" + "601024" + ";" + tHWBox.lot + vbCrLf
        
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
            tSTCarton.Job = Trim(rs!JOB_ID)
            tSTCarton.DEV = Trim$(rs!CUSTOMER_DEVICE)
            tSTCarton.lot = Trim(rs!CartonID)
            tSTCarton.qty = Trim(rs!qty)
        
            sDatecode = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTCarton.Job & "'")
            tSTCarton.DATECODE = sDatecode
        
            sTestDateCode = Get_SqlStr("select distinct HTDATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTCarton.Job & "'")
            tSTCarton.testdateCode = sTestDateCode

            If tSTCarton.DATECODE = "" Or tSTCarton.testdateCode = "" Then
            
                tSTCarton.DATECODE = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTCarton.Job & "'")
                tSTCarton.testdateCode = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTCarton.Job & "'")
            End If

            sContent = sContent + tSTCarton.DEV + "," + tSTCarton.Job + ",1T" + tSTCarton.Job + "," + tSTCarton.DEV + "," + "1P" + tSTCarton.DEV + "," + tSTCarton.DATECODE + "," + tSTCarton.DATECODE + "," + Mid(tSTCarton.lot, 2) + "," + tSTCarton.lot + ","
            sContent = sContent + tSTCarton.qty + ",Q" + tSTCarton.qty + "," + tSTCarton.testdateCode + "," + tSTCarton.testdateCode
        
            sAdd = GetDevMark(tSTCarton.DEV)
            sContent = sContent + sAdd + vbCrLf
        
            rs.MoveNext
        Loop

    End If

    Call CreateTxt(sFileName, sContent, sPath)
End Sub

Rem: Print OutPkg Start
Private Sub cmdPrintCarton_Click()

    Dim sDN       As String

    Dim sOra      As String

    Dim iOp       As Integer

    Dim iOpMax    As Integer

    Dim iSec As Integer

    Dim sType     As String

    cmdPrintCarton.Enabled = False
    iSec = (txtPrintInterval.Text) * 1000
    sType = Trim(txtShip.Text)

    sDN = txtdn.Text

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
        
            Call PrintCusCartonSubLbl(sDN, iOp)
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

    sOra = "select a.kid,a.CUSTOMER_DEVICE, b.customerpartnumber, a.job_id,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & txtdn.Text & "'" & "and b.delivery = '" & txtdn.Text & "' and a.job_id = b.batchnumber and a.outbox_num = '" & iOp & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber, a.job_id, b.purchasingdocno,a.kid"
    
    Set rs = Get_OracleRs(sOra)

    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            tCusCARTON.DN = txtdn.Text
            tCusCARTON.PO = rs!PO
            tCusCARTON.CPN = rs!customerpartnumber
            tCusCARTON.MPN = rs!CUSTOMER_DEVICE
            tCusCARTON.Job = rs!JOB_ID
            tCusCARTON.qty = rs!qty
            KID = rs!KID
        
            tCusCARTON.DATECODE = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tCusCARTON.Job & "'")
        
            If tCusCARTON.DATECODE = "" Then
                tCusCARTON.DATECODE = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tCusCARTON.Job & "'")
            End If
        
            sContent = sContent & tCusCARTON.DN & ",I" & tCusCARTON.DN & "," & tCusCARTON.PO & ",K" & tCusCARTON.PO & ",E2," & tCusCARTON.CPN & ",P" & tCusCARTON.CPN & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & ","
            sContent = sContent & tCusCARTON.qty & ",Q" & tCusCARTON.qty & ","
            sContent = sContent & Get_OracleStr("select distinct trim(a.freightforwarder)|| ',CHINA,' || substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3)||','||trim(a.city) || ' ' || trim(a.state) || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ',' || trim(a.phone) || ','  from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & txtdn.Text & "'")
            sContent = sContent & tCusCARTON.Job & ",P" & tCusCARTON.Job & "," & tCusCARTON.DATECODE & ",9D" & tCusCARTON.DATECODE & "," & iOp & "," & KID & vbCrLf
        
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

    sPath = Trim(txtCusCartonPath.Text)
    sFileName = sDN & "-" & "CUSCARTONLBL" + Format(Now(), "YYYYMMDDHHmmSS")
    sContent = ""

    sOra = "select a.kid,a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & txtdn.Text & "'" & "and b.delivery = '" & txtdn.Text & "' and a.job_id = b.batchnumber and a.outbox_num = '" & iOp & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno, a.kid"
    
    Set rs = Get_OracleRs(sOra)

    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            tCusCARTON.DN = txtdn.Text
            tCusCARTON.PO = rs!PO
            tCusCARTON.CPN = rs!customerpartnumber
            tCusCARTON.MPN = rs!CUSTOMER_DEVICE
            tCusCARTON.qty = rs!qty
            KID = rs!KID
            sContent = sContent & tCusCARTON.DN & ",I" & tCusCARTON.DN & "," & tCusCARTON.PO & ",K" & tCusCARTON.PO & ",E2," & tCusCARTON.CPN & ",P" & tCusCARTON.CPN & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & ","
            sContent = sContent & tCusCARTON.qty & ",Q" & tCusCARTON.qty & ","
            sContent = sContent & Get_OracleStr("select distinct trim(a.freightforwarder)|| ',CHINA,' || substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3)||','||trim(a.city) || ' ' || trim(a.state) || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ','|| 'Attn:;Tel:' || trim(a.phone) || ','  from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & txtdn.Text & "'")
            sContent = sContent & "N/A,PN/A,N/A,9DN/A," & iOp & "," & KID
        
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

    sPath = Trim(txtSEMTCartonPAth.Text)
    sFileName = sDN & "-" & "SemTechStanderCarton" + Format(Now(), "YYYYMMDDHHmmSS")
    sContent = ""

    sOra = "select a.CUSTOMER_DEVICE,a.kid, b.customerpartnumber,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & txtdn.Text & "'" & "and b.delivery = '" & txtdn.Text & "' and a.job_id = b.batchnumber and a.outbox_num = '" & iOp & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno,a.kid"
    
    Set rs = Get_OracleRs(sOra)

    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            tCusCARTON.DN = txtdn.Text
            tCusCARTON.PO = IIf(IsNull(rs!PO), "N/A", rs!PO)
            tCusCARTON.CPN = IIf(IsNull(rs!customerpartnumber), "N/A", rs!customerpartnumber)
            tCusCARTON.MPN = IIf(IsNull(rs!CUSTOMER_DEVICE), "N/A", rs!CUSTOMER_DEVICE)
            tCusCARTON.qty = rs!qty
            sKid = rs!KID
        
            sContent = sContent & Get_OracleStr("select distinct substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3) || ','||trim(a.city) || ' ' || trim(a.state)  || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ',' || trim(a.contactname) || ',' || trim(a.phone) from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & txtdn.Text & "' ") & ","
            sContent = sContent & tCusCARTON.DN & ",I" & tCusCARTON.DN & "," & Left(tCusCARTON.PO, 10) & ",K" & Left(tCusCARTON.PO, 10) & "," & Left$(tCusCARTON.CPN, 11) & ",P" & Left$(tCusCARTON.CPN, 11) & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & "," & tCusCARTON.qty & ",Q" & tCusCARTON.qty & "," & Get_OracleStr("select distinct freightforwarder from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & txtdn.Text & "'") & "," & "" & "," & "" & "," & "" & "," & "COO:CHINA" & "," & "CHINA"

            sAdd = "," & iOp & "," & sKid

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

End Sub

'按内盒补打
Private Sub PrintA_Inner(sRefer As String)

    Dim sOra  As String

    Dim rs    As New ADODB.Recordset

    Dim iOp   As Integer

    Dim iIp   As Integer

    Dim sType As String

    Dim sDN   As String

    sDN = txtdn.Text
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

    Dim sOra      As String

    Dim tCusReel  As CusReel

    Dim sContent  As String

    Dim sFileName As String

    Dim rs        As New ADODB.Recordset

    Dim sPath     As String

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

    sOra = "select reelid, qty, mpn, cpn from lpstbl where reelid = '" & sRefer & "'"
    Set rs = Get_OracleRs(sOra)

    tCusReel.lot = UCase(Trim(rs!ReelID))
    tCusReel.qty = UCase(Trim(rs!qty))
    tCusReel.DEV = UCase(Trim(rs!MPN))
    tCusReel.PN = UCase(Trim(rs!CPN))
    
    sContent = sContent + tCusReel.PN + "DPTKE2" + tCusReel.lot + Right$("000000" + tCusReel.qty, 6) + ","
    sContent = sContent + tCusReel.PN + "," + "TVS DIODES" + "," + "PO TYPE,:E2" + "," + tCusReel.lot + "," + tCusReel.qty + ","
    sContent = sContent + tCusReel.DEV + "," + "DPTK"

    If tCusReel.DEV = "RCLAMP2581ZCTFT" Then
        sPath = "\\10.160.1.14\BarCode\37\37BOXJP-新\"
    End If

    Call CreateTxt(sFileName, sContent, sPath)

    MsgBox "补打成功", vbInformation, "提示:"
    tReferTo.Text = ""
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

    tSTBox.Job = Trim(rs!JOB_ID)
    tSTBox.DEV = Trim(rs!MPN)
    tSTBox.lot = Trim(rs!BoxID)
    tSTBox.qty = Trim(rs!qty)
        
    sDatecode = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTBox.Job & "'")
    tSTBox.DATECODE = sDatecode
        
    sTestDateCode = Get_SqlStr("select distinct HTDATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTBox.Job & "'")
    tSTBox.testdateCode = sTestDateCode

    If tSTBox.DATECODE = "" Or tSTBox.testdateCode = "" Then
        tSTBox.DATECODE = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTBox.Job & "'")
        tSTBox.testdateCode = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTBox.Job & "'")
    End If

    '
    'If Left$(tSTBox.DEV, 2) = "RC" Then
    '    sAdd = "," & Left$(tSTBox.DEV, 6) & ",{R}," & Mid$(tSTBox.DEV, 7) & "," & "RailClamp{R}"
    'Else
    '    If Left$(tSTBox.DEV, 2) <> "PA" Then
    '        sAdd = "," & Left$(tSTBox.DEV, 6) & ",{TM}," & Mid$(tSTBox.DEV, 7) & "," & "MicroClamp{TM}"
    '    End If
    'End If
    '
    sContent = sContent + tSTBox.DEV + "," + tSTBox.Job + ",1T" + tSTBox.Job + "," + tSTBox.DEV + "," + "1P" + tSTBox.DEV + "," + tSTBox.DATECODE + "," + tSTBox.DATECODE + "," + Mid(tSTBox.lot, 2) + "," + tSTBox.lot + ","
    sContent = sContent + tSTBox.qty + ",Q" + tSTBox.qty + "," + tSTBox.testdateCode + "," + tSTBox.testdateCode

    sAdd = GetDevMark(tSTBox.DEV)

    sContent = sContent + sAdd

    Call CreateTxt(sFileName, sContent, sPath)

    MsgBox "补打成功", vbInformation, "提示:"
    tReferTo.Text = ""
End Sub

' 补打三星内盒标签
Private Sub PrintA_SSBOX(sRefer As String)

    Dim sOra      As String

    Dim tCusBox   As CusBox

    Dim sContent  As String

    Dim sFileName As String

    Dim rs        As New ADODB.Recordset

    Dim sPath     As String

    Dim iIp       As Integer

    Dim iOp       As Integer

    Dim sDN       As String

    sPath = Trim(txtCusBoxPath.Text)
    sFileName = "CUSBoxLbl" + Format(Now(), "YYYYMMDDHHmmSS")
    sContent = ""

    sOra = "select distinct dn_num, inbox_num, outbox_num from lpstbl where boxid ='" & sRefer & "'"

    If Get_OracleStr(sOra) = "" Then
        MsgBox "系统查不到此内盒的数据", vbInformation, "提示:"

        Exit Sub

    End If

    Set rs = Get_OracleRs(sOra)
    iIp = rs!INBOX_NUM
    iOp = rs!OUTBOX_NUM
    sDN = rs!DN_NUM

    sOra = "select cpn, mpn, sum(qty) as qty from lpstbl where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' and inbox_num = '" & iIp & "' group by cpn, mpn"
    Set rs = Get_OracleRs(sOra)

    tCusBox.qty = Trim$(rs!qty)
    tCusBox.DEV = Trim(rs!MPN)
    tCusBox.PN = Trim(rs!CPN)

    sContent = sContent + tCusBox.PN + "DPTKE2" + Right$("000000" + tCusBox.qty, 6) + ","
    sContent = sContent + tCusBox.PN + "," + "TVS DIODES" + "," + "PO TYPE,:E2" + "," + tCusBox.qty + "," + tCusBox.DEV + "," + "DPTK"

    If tCusBox.DEV = "RCLAMP2581ZCTFT" Then
        sPath = "\\10.160.1.14\BarCode\37\37BoxNH-新\"
    End If


    Call CreateTxt(sFileName, sContent, sPath)

    MsgBox "补打成功", vbInformation, "提示:"
    tReferTo.Text = ""
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
            tSTCarton.Job = Trim(rs!JOB_ID)
            tSTCarton.DEV = Trim$(rs!CUSTOMER_DEVICE)
            tSTCarton.lot = Trim(rs!CartonID)
            tSTCarton.qty = Trim(rs!qty)
        
            sDatecode = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTCarton.Job & "'")
            tSTCarton.DATECODE = sDatecode
        
            sTestDateCode = Get_SqlStr("select distinct HTDATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTCarton.Job & "'")
            tSTCarton.testdateCode = sTestDateCode
        
            If tSTCarton.DATECODE = "" Or tSTCarton.testdateCode = "" Then
                tSTCarton.DATECODE = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTCarton.Job & "'")
                tSTCarton.testdateCode = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTCarton.Job & "'")
            End If
  
            sFileName = "SEMT_CARTON_" + tSTCarton.lot
    
            sContent = sContent + tSTCarton.DEV + "," + tSTCarton.Job + ",1T" + tSTCarton.Job + "," + tSTCarton.DEV + "," + "1P" + tSTCarton.DEV + "," + tSTCarton.DATECODE + "," + tSTCarton.DATECODE + "," + Mid(tSTCarton.lot, 2) + "," + tSTCarton.lot + ","
            sContent = sContent + tSTCarton.qty + ",Q" + tSTCarton.qty + "," + tSTCarton.testdateCode + "," + tSTCarton.testdateCode
        
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

    Dim sDN       As String

    Dim sOra      As String

    Dim iOp       As Integer

    Dim iOpMax    As Integer

    Dim iSec As Integer

    Dim sType     As String

    cmdPrintCarton.Enabled = False
    iSec = (txtPrintInterval.Text) * 1000
    sType = Trim(txtShip.Text)

    sDN = txtdn.Text

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
        Call PrintHTCartonLbl(sDN, iOp)
        Sleep (iSec)
    
        If sType = "SSE2" Then
            Call PrintCusCartonLbl(sDN, iOp)
            Sleep (iSec)
        
            Call PrintCusCartonSubLbl(sDN, iOp)
            Sleep (iSec)
        Else
            Call PrintSTCartonStanderLbl(sDN, iOp)
            Sleep (iSec)
        End If

    Next

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

' 记录到erp
Private Sub TransToErp(sCartonID As String, lQty As Long)

    Dim sqlTemp As String

    Dim idTemp  As String

    Dim boxdn   As String
    
      ' 先删掉相同QBoxNo
    AddSql2 ("DELETE FROM [erpdata].[dbo].[tblPackTreeInf] WHERE 箱号 = '" & sCartonID & "' ")
    AddSql2 ("DELETE FROM [erpdata].[dbo].[tblPackMainInf] WHERE 箱号 = '" & sCartonID & "' ")
    AddSql2 ("DELETE FROM [erpdata].[dbo].[tblStockNumTree] WHERE 箱号 = '" & sCartonID & "' ")
    
    ' 再更新信息
    
    sqlTemp = "insert into [erpdata].[dbo].[tblPackMainInf](箱号,客户代码,数量,合格标记,装箱标记,产线标记) values('" & sCartonID & "','37'," & lQty & ",'0','1','1') "

    AddSql2 (sqlTemp)

    '插入Sqlserver   tblPackTreeInf
    sqlTemp = "insert into [erpdata].[dbo].[tblPackTreeInf](箱号,上级序号,基层标记 ,Memo) values('" & sCartonID & "',0,1,'37')"
    AddSql2 (sqlTemp)

    '再更新小箱的上级序号

    '把序号先查出来，再整体更新

    idTemp = Get37BigQboxIDV1(sCartonID)
    boxdn = UCase(Trim(txtdn.Text))

    sqlTemp = "insert into [erpdata].[dbo].[tblStockNumTree](序号,箱号,上级序号,基层标记 ,尺寸,重量,Memo,DN) values( " & idTemp & ",'" & sCartonID & "',0,1,'','','37','" & boxdn & "')"
    AddSql2 (sqlTemp)

    sqlTemp = " Update [erpdata].[dbo].[tblPackTreeInf] set 上级序号='" & idTemp & "',Memo='37' " & " where 箱号 in ( select trayid from (select * from OPENQUERY(ORACLEDB, 'SELECT * from lpstbl' )) X where X.carton = '" & sCartonID & "') "

    AddSql2 (sqlTemp)

    sqlTemp = " Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='" & idTemp & "',Memo='37', dn='" & boxdn & "' " & " where 箱号 in (select trayid from (select * from OPENQUERY(ORACLEDB, 'SELECT * from lpstbl' )) X where X.carton = '" & sCartonID & "') "

    AddSql2 (sqlTemp)

End Sub

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

Private Function GetDevMark(sDev As String) As String

    Dim PartNOPre     As String

    Dim ProductFamily As String

    Dim sCode         As String
            
    sCode = Left$(sDev, 2)
        
    Select Case sCode
        
        Case "RC"
            PartNOPre = "RCLAMP,{R}," & Mid$(sDev, 7)
            ProductFamily = "RailClamp{R}"
        
        Case "UC"
            PartNOPre = "UCLAMP,{R}," & Mid$(sDev, 7)
            ProductFamily = "MicroClamp{TM}"

        Case "EC"
            PartNOPre = "ECLAMP,{TM}," & Mid$(sDev, 7)
            ProductFamily = "EMIClamp{TM}"

        Case "TC"
            PartNOPre = "TCLAMP,{TM}," & Mid$(sDev, 7)
            ProductFamily = "TransClamp{TM}"

        Case "HC"
            PartNOPre = "HCLAMP,{TM}," & Mid$(sDev, 7)
            ProductFamily = ""

        Case "PC"
            PartNOPre = "PCLAMP,{TM}," & Mid$(sDev, 7)
            ProductFamily = ""

        Case "HS"
            PartNOPre = "HS"
            ProductFamily = "HotSwitch{TM}"

    End Select
    
    GetDevMark = "," & PartNOPre & "," & ProductFamily

End Function

