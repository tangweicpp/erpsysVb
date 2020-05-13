VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmUpLoadOI 
   Caption         =   "上传客户资料"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
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
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Aptina及CN 导入"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label25"
      Tab(0).Control(1)=   "lblLOT"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(4)=   "Combo1"
      Tab(0).Control(5)=   "cmdddd"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "WO 导入(GC)"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label24"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label26"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label27"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label28"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lbl111"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblM"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "CmbCustomer"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame5"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chkGC_GC50253"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "ComCbBand"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtMsg"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "chkMsgAppend"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "手工创建OI"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(2)=   "Label3"
      Tab(2).Control(3)=   "Label4"
      Tab(2).Control(4)=   "Label5"
      Tab(2).Control(5)=   "Label6"
      Tab(2).Control(6)=   "Label7"
      Tab(2).Control(7)=   "Label8"
      Tab(2).Control(8)=   "Label9"
      Tab(2).Control(9)=   "Label10"
      Tab(2).Control(10)=   "Label11"
      Tab(2).Control(11)=   "Label12"
      Tab(2).Control(12)=   "Label13"
      Tab(2).Control(13)=   "Label14"
      Tab(2).Control(14)=   "Label15"
      Tab(2).Control(15)=   "Label16"
      Tab(2).Control(16)=   "Label17"
      Tab(2).Control(17)=   "Label18"
      Tab(2).Control(18)=   "Label19"
      Tab(2).Control(19)=   "Label20"
      Tab(2).Control(20)=   "TxtCustomer"
      Tab(2).Control(21)=   "TxtPO"
      Tab(2).Control(22)=   "TxtPOItem"
      Tab(2).Control(23)=   "TxtLotId"
      Tab(2).Control(24)=   "TxtMpn"
      Tab(2).Control(25)=   "TxtMpnDesc"
      Tab(2).Control(26)=   "TxtWaferQty"
      Tab(2).Control(27)=   "TxtDieQty"
      Tab(2).Control(28)=   "TxtDesign"
      Tab(2).Control(29)=   "TxtCountryFab"
      Tab(2).Control(30)=   "TxtImageRev"
      Tab(2).Control(31)=   "TxtFFacility"
      Tab(2).Control(32)=   "TxtMarkId"
      Tab(2).Control(33)=   "TxtLotPriority"
      Tab(2).Control(34)=   "TxtFilmApld"
      Tab(2).Control(35)=   "TxtShip260"
      Tab(2).Control(36)=   "TxtShipLevel"
      Tab(2).Control(37)=   "TxtMicMaterial"
      Tab(2).Control(38)=   "TxtShipSite"
      Tab(2).Control(39)=   "TxtLotStatus"
      Tab(2).Control(40)=   "CmdSaveOI"
      Tab(2).Control(41)=   "CmdClearOI"
      Tab(2).ControlCount=   42
      TabCaption(3)   =   "手工创建Mapping"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label21"
      Tab(3).Control(1)=   "Label22"
      Tab(3).Control(2)=   "Label23"
      Tab(3).Control(3)=   "Frame4"
      Tab(3).Control(4)=   "TxtCustomerName"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "手工创建新客户WO"
      TabPicture(4)   =   "Form1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label29"
      Tab(4).Control(1)=   "Label30"
      Tab(4).Control(2)=   "Label31"
      Tab(4).Control(3)=   "Label32"
      Tab(4).Control(4)=   "Label33"
      Tab(4).Control(5)=   "Label34"
      Tab(4).Control(6)=   "Label35"
      Tab(4).Control(7)=   "Label36"
      Tab(4).Control(8)=   "Label37"
      Tab(4).Control(9)=   "Label38"
      Tab(4).Control(10)=   "Label39"
      Tab(4).Control(11)=   "Label40"
      Tab(4).Control(12)=   "Label41"
      Tab(4).Control(13)=   "Label42"
      Tab(4).Control(14)=   "Label43"
      Tab(4).Control(15)=   "TxtCustomerID"
      Tab(4).Control(16)=   "TxtItem"
      Tab(4).Control(17)=   "TxtPONO"
      Tab(4).Control(18)=   "TxtSupplier"
      Tab(4).Control(19)=   "TxtFabDevice"
      Tab(4).Control(20)=   "TxtCustomerDevice"
      Tab(4).Control(21)=   "TxtWaferVersion"
      Tab(4).Control(22)=   "TxtMarkingLotID"
      Tab(4).Control(23)=   "DTPicker1"
      Tab(4).Control(24)=   "TxtLotId2"
      Tab(4).Control(25)=   "TxtWaferID2"
      Tab(4).Control(26)=   "TxtTotailDie"
      Tab(4).Control(27)=   "TxtGoodDieQty"
      Tab(4).Control(28)=   "TxtWONO2"
      Tab(4).Control(29)=   "TxtShipTo"
      Tab(4).Control(30)=   "CmdSave"
      Tab(4).Control(31)=   "ComdClear"
      Tab(4).Control(32)=   "Command18"
      Tab(4).ControlCount=   33
      Begin VB.CheckBox chkMsgAppend 
         Caption         =   "是否需要"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2040
         TabIndex        =   127
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtMsg 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   1770
         Left            =   4320
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   125
         Top             =   480
         Width           =   4575
      End
      Begin VB.CommandButton cmdddd 
         Caption         =   "导出"
         Height          =   360
         Left            =   -72360
         TabIndex        =   124
         Top             =   7440
         Width           =   990
      End
      Begin VB.ComboBox ComCbBand 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":008C
         Left            =   1680
         List            =   "Form1.frx":0096
         Style           =   2  'Dropdown List
         TabIndex        =   122
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkGC_GC50253 
         Caption         =   "GC_GC5025-3- 2212"
         Height          =   375
         Left            =   360
         TabIndex        =   120
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CommandButton Command18 
         Caption         =   "退出"
         Height          =   600
         Left            =   -64920
         TabIndex        =   117
         Top             =   4320
         Width           =   1575
      End
      Begin VB.CommandButton ComdClear 
         Caption         =   "清空"
         Height          =   600
         Left            =   -67800
         TabIndex        =   116
         Top             =   4320
         Width           =   1575
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "保存"
         Height          =   600
         Left            =   -71040
         TabIndex        =   115
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox TxtShipTo 
         Height          =   375
         Left            =   -73200
         TabIndex        =   114
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox TxtWONO2 
         Height          =   375
         Left            =   -65280
         TabIndex        =   113
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox TxtGoodDieQty 
         Height          =   375
         Left            =   -69480
         TabIndex        =   111
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox TxtTotailDie 
         Height          =   375
         Left            =   -73200
         TabIndex        =   109
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox TxtWaferID2 
         Height          =   375
         Left            =   -61680
         TabIndex        =   107
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox TxtLotId2 
         Height          =   375
         Left            =   -65280
         TabIndex        =   105
         Top             =   2160
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -69480
         TabIndex        =   103
         Top             =   2160
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   219348993
         CurrentDate     =   42173
      End
      Begin VB.TextBox TxtMarkingLotID 
         Height          =   375
         Left            =   -73200
         TabIndex        =   101
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox TxtWaferVersion 
         Height          =   375
         Left            =   -61680
         TabIndex        =   99
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox TxtCustomerDevice 
         Height          =   375
         Left            =   -65280
         TabIndex        =   97
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox TxtFabDevice 
         Height          =   375
         Left            =   -69480
         TabIndex        =   95
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox TxtSupplier 
         Height          =   375
         Left            =   -61680
         TabIndex        =   92
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox TxtPONO 
         Height          =   375
         Left            =   -65280
         TabIndex        =   90
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox TxtItem 
         Height          =   375
         Left            =   -69480
         TabIndex        =   88
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox TxtCustomerID 
         Height          =   375
         Left            =   -73200
         TabIndex        =   86
         Top             =   840
         Width           =   2175
      End
      Begin VB.Frame Frame5 
         Caption         =   "Mapping上传"
         Height          =   2895
         Left            =   240
         TabIndex        =   76
         Top             =   5040
         Width           =   8655
         Begin VB.ComboBox CmbGCType 
            Height          =   315
            ItemData        =   "Form1.frx":00A8
            Left            =   1320
            List            =   "Form1.frx":00AA
            Style           =   2  'Dropdown List
            TabIndex        =   118
            Top             =   1680
            Width           =   2055
         End
         Begin VB.CommandButton Command15 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   3840
            TabIndex        =   80
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CommandButton Command14 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   1320
            TabIndex        =   79
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CommandButton Command13 
            Caption         =   ".."
            Height          =   495
            Left            =   6120
            TabIndex        =   78
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox TxtSI 
            Enabled         =   0   'False
            Height          =   975
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   77
            Top             =   480
            Width           =   5295
         End
         Begin MSComDlg.CommonDialog ComSI 
            Left            =   6360
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            MaxFileSize     =   10000
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GC类型："
            Height          =   195
            Left            =   600
            TabIndex        =   119
            Top             =   1680
            Width           =   750
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的map："
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   840
            TabIndex        =   81
            Top             =   240
            Width           =   1560
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":00AC
         Left            =   -73320
         List            =   "Form1.frx":00AE
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   780
         Width           =   1695
      End
      Begin VB.ComboBox CmbCustomer 
         Height          =   315
         ItemData        =   "Form1.frx":00B0
         Left            =   1680
         List            =   "Form1.frx":00B2
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox TxtCustomerName 
         Height          =   375
         Left            =   -73440
         TabIndex        =   70
         Top             =   780
         Width           =   2415
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mapping_XML"
         Height          =   2295
         Left            =   -74040
         TabIndex        =   62
         Top             =   1380
         Width           =   9015
         Begin VB.TextBox Text4 
            Enabled         =   0   'False
            Height          =   495
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   66
            Top             =   840
            Width           =   4935
         End
         Begin VB.CommandButton Command12 
            Caption         =   ".."
            Height          =   495
            Left            =   6120
            TabIndex        =   65
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton Command11 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   1200
            TabIndex        =   64
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton Command10 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   4080
            TabIndex        =   63
            Top             =   1560
            Width           =   1335
         End
         Begin MSComDlg.CommonDialog CommonDialog3 
            Left            =   3000
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的xls："
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   840
            TabIndex        =   67
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.CommandButton CmdClearOI 
         Caption         =   "清空"
         Height          =   480
         Left            =   -68160
         TabIndex        =   61
         Top             =   5580
         Width           =   1335
      End
      Begin VB.CommandButton CmdSaveOI 
         Caption         =   "保存"
         Height          =   480
         Left            =   -71040
         TabIndex        =   60
         Top             =   5580
         Width           =   1335
      End
      Begin VB.TextBox TxtLotStatus 
         Height          =   375
         Left            =   -68640
         TabIndex        =   59
         Top             =   4620
         Width           =   2415
      End
      Begin VB.TextBox TxtShipSite 
         Height          =   375
         Left            =   -73080
         TabIndex        =   57
         Top             =   4620
         Width           =   2415
      End
      Begin VB.TextBox TxtMicMaterial 
         Height          =   375
         Left            =   -64440
         TabIndex        =   55
         Top             =   4000
         Width           =   2415
      End
      Begin VB.TextBox TxtShipLevel 
         Height          =   375
         Left            =   -68640
         TabIndex        =   53
         Top             =   4000
         Width           =   2415
      End
      Begin VB.TextBox TxtShip260 
         Height          =   375
         Left            =   -73080
         TabIndex        =   51
         Top             =   4000
         Width           =   2415
      End
      Begin VB.TextBox TxtFilmApld 
         Height          =   375
         Left            =   -64440
         TabIndex        =   49
         Top             =   3380
         Width           =   2415
      End
      Begin VB.TextBox TxtLotPriority 
         Height          =   375
         Left            =   -68640
         TabIndex        =   47
         Top             =   3380
         Width           =   2415
      End
      Begin VB.TextBox TxtMarkId 
         Height          =   375
         Left            =   -73080
         TabIndex        =   45
         Top             =   3380
         Width           =   2415
      End
      Begin VB.TextBox TxtFFacility 
         Height          =   375
         Left            =   -64440
         TabIndex        =   43
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox TxtImageRev 
         Height          =   375
         Left            =   -68640
         TabIndex        =   41
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox TxtCountryFab 
         Height          =   375
         Left            =   -73080
         TabIndex        =   39
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox TxtDesign 
         Height          =   375
         Left            =   -64440
         TabIndex        =   37
         Top             =   2140
         Width           =   2415
      End
      Begin VB.TextBox TxtDieQty 
         Height          =   375
         Left            =   -68640
         TabIndex        =   35
         Top             =   2140
         Width           =   2415
      End
      Begin VB.TextBox TxtWaferQty 
         Height          =   375
         Left            =   -73080
         TabIndex        =   33
         Top             =   2140
         Width           =   2415
      End
      Begin VB.TextBox TxtMpnDesc 
         Height          =   375
         Left            =   -64440
         TabIndex        =   31
         Top             =   1520
         Width           =   2415
      End
      Begin VB.TextBox TxtMpn 
         Height          =   375
         Left            =   -68640
         TabIndex        =   29
         Top             =   1500
         Width           =   2415
      End
      Begin VB.TextBox TxtLotId 
         Height          =   375
         Left            =   -73080
         TabIndex        =   27
         Top             =   1520
         Width           =   2415
      End
      Begin VB.TextBox TxtPOItem 
         Height          =   375
         Left            =   -64440
         TabIndex        =   25
         Top             =   900
         Width           =   2415
      End
      Begin VB.TextBox TxtPO 
         Height          =   375
         Left            =   -68640
         TabIndex        =   23
         Top             =   900
         Width           =   2415
      End
      Begin VB.TextBox TxtCustomer 
         Height          =   375
         Left            =   -73080
         TabIndex        =   21
         Top             =   900
         Width           =   2415
      End
      Begin VB.Frame Frame3 
         Caption         =   "WO上传"
         Height          =   2535
         Left            =   240
         TabIndex        =   13
         Top             =   2400
         Width           =   8655
         Begin VB.CommandButton Command9 
            Caption         =   "导出明细表"
            Height          =   480
            Left            =   5400
            TabIndex        =   19
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CommandButton Command8 
            Caption         =   "导出主表"
            Height          =   480
            Left            =   3600
            TabIndex        =   17
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CommandButton Command7 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   1200
            TabIndex        =   16
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   ".."
            Height          =   495
            Left            =   6120
            TabIndex        =   15
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   495
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   840
            Width           =   4935
         End
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   6240
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的CSV："
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   840
            TabIndex        =   18
            Top             =   480
            Width           =   1545
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mapping_XML"
         Height          =   2415
         Left            =   -74040
         TabIndex        =   7
         Top             =   1320
         Width           =   7095
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   975
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   480
            Width           =   5295
         End
         Begin VB.CommandButton cmd 
            Caption         =   ".."
            Height          =   495
            Index           =   0
            Left            =   6120
            TabIndex        =   10
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   1200
            TabIndex        =   9
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   4080
            TabIndex        =   8
            Top             =   1680
            Width           =   1335
         End
         Begin MSComDlg.CommonDialog Com 
            Left            =   6360
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            MaxFileSize     =   10000
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的XML："
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   840
            TabIndex        =   12
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "OI_CSV"
         Height          =   2535
         Left            =   -74040
         TabIndex        =   1
         Top             =   4080
         Width           =   7095
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   495
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   840
            Width           =   4935
         End
         Begin VB.CommandButton Command2 
            Caption         =   ".."
            Height          =   495
            Left            =   6120
            TabIndex        =   4
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton Command3 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   1200
            TabIndex        =   3
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command5 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   4080
            TabIndex        =   2
            Top             =   1680
            Width           =   1335
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   3000
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的CSV："
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   840
            TabIndex        =   6
            Top             =   480
            Width           =   1545
         End
      End
      Begin VB.Label lblM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "邮件正文补充(M)                                                                                                "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   240
         TabIndex        =   126
         Top             =   2160
         Width           =   13230
      End
      Begin VB.Label lblLOT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "导出LOT对照表"
         Height          =   195
         Left            =   -73680
         TabIndex        =   123
         Top             =   7440
         Width           =   1185
      End
      Begin VB.Label lbl111 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "保税/非保税:"
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
         Left            =   360
         TabIndex        =   121
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WO NO："
         Height          =   195
         Left            =   -66240
         TabIndex        =   112
         Top             =   2880
         Width           =   795
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Good Die Qty："
         Height          =   195
         Left            =   -70680
         TabIndex        =   110
         Top             =   3000
         Width           =   1275
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Die Qty："
         Height          =   195
         Left            =   -74520
         TabIndex        =   108
         Top             =   3000
         Width           =   1275
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wafer ID："
         Height          =   195
         Left            =   -62640
         TabIndex        =   106
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot ID："
         Height          =   195
         Left            =   -66240
         TabIndex        =   104
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date："
         Height          =   195
         Left            =   -70320
         TabIndex        =   102
         Top             =   2280
         Width           =   675
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marking Lot ID："
         Height          =   195
         Left            =   -74640
         TabIndex        =   100
         Top             =   2280
         Width           =   1275
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "wafer Version："
         Height          =   195
         Left            =   -63000
         TabIndex        =   98
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Device："
         Height          =   195
         Left            =   -66960
         TabIndex        =   96
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FAB Device："
         Height          =   195
         Left            =   -70560
         TabIndex        =   94
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To："
         Height          =   195
         Left            =   -74160
         TabIndex        =   93
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier："
         Height          =   195
         Left            =   -62640
         TabIndex        =   91
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO NO："
         Height          =   195
         Left            =   -66240
         TabIndex        =   89
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item："
         Height          =   195
         Left            =   -70320
         TabIndex        =   87
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码："
         Height          =   195
         Left            =   -74280
         TabIndex        =   85
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GC客户WLT、COG，MG客户上传时，先上传WO，后再上传Mapping。"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1320
         TabIndex        =   84
         Top             =   7980
         Width           =   5370
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   83
         Top             =   480
         Width           =   90
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请先选择客户代码，然后再上传WO或Mapping"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   82
         Top             =   480
         Width           =   3570
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户"
         Height          =   195
         Left            =   -73800
         TabIndex        =   75
         Top             =   780
         Width           =   360
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户"
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
         Left            =   360
         TabIndex        =   73
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户："
         Height          =   195
         Left            =   -74040
         TabIndex        =   71
         Top             =   900
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excel模板格式为：WaferId LotId ProductId 良品数 不良数"
         Height          =   195
         Left            =   -73440
         TabIndex        =   69
         Top             =   4500
         Width           =   4395
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注："
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -73920
         TabIndex        =   68
         Top             =   4140
         Width           =   540
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot_Status："
         Height          =   195
         Left            =   -69720
         TabIndex        =   58
         Top             =   4740
         Width           =   960
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ship_Site："
         Height          =   195
         Left            =   -74040
         TabIndex        =   56
         Top             =   4740
         Width           =   840
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Micron_Material："
         Height          =   195
         Left            =   -65760
         TabIndex        =   54
         Top             =   4140
         Width           =   1305
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping_Mst_Level："
         Height          =   195
         Left            =   -70320
         TabIndex        =   52
         Top             =   4140
         Width           =   1590
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping_Mst_260："
         Height          =   195
         Left            =   -74640
         TabIndex        =   50
         Top             =   4140
         Width           =   1485
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Protective_Film_Apld："
         Height          =   195
         Left            =   -66120
         TabIndex        =   48
         Top             =   3420
         Width           =   1680
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot_Priority："
         Height          =   195
         Left            =   -69720
         TabIndex        =   46
         Top             =   3420
         Width           =   1005
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encoded_Mark_Id："
         Height          =   195
         Left            =   -74640
         TabIndex        =   44
         Top             =   3420
         Width           =   1470
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fabrication_Facility："
         Height          =   195
         Left            =   -66000
         TabIndex        =   42
         Top             =   2820
         Width           =   1560
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imager_Customer_Rev："
         Height          =   195
         Left            =   -70560
         TabIndex        =   40
         Top             =   2820
         Width           =   1845
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country_Of_Fab："
         Height          =   195
         Left            =   -74520
         TabIndex        =   38
         Top             =   2820
         Width           =   1395
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Design_Id："
         Height          =   195
         Left            =   -65400
         TabIndex        =   36
         Top             =   2220
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Die_Qty："
         Height          =   195
         Left            =   -69480
         TabIndex        =   34
         Top             =   2340
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current_Wafer_Qty："
         Height          =   195
         Left            =   -74760
         TabIndex        =   32
         Top             =   2220
         Width           =   1635
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mpn_Desc："
         Height          =   195
         Left            =   -65400
         TabIndex        =   30
         Top             =   1620
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mpn："
         Height          =   195
         Left            =   -69240
         TabIndex        =   28
         Top             =   1620
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source_Batch_Id："
         Height          =   195
         Left            =   -74520
         TabIndex        =   26
         Top             =   1620
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Po_Item："
         Height          =   195
         Left            =   -65280
         TabIndex        =   24
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Po_Num："
         Height          =   195
         Left            =   -69480
         TabIndex        =   22
         Top             =   1020
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户："
         Height          =   195
         Left            =   -73680
         TabIndex        =   20
         Top             =   1020
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmUpLoadOI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Dim mapTemp        As MapRecord

Dim gcHeaderTemp   As GCHeader

Dim eqISHeaderTemp As EQISHeader

Dim gcDetailTemp   As GCDetail

Dim gTax As String

'Dim SumCount As Integer
Dim ErrorInf       As String

Dim updateRS       As New ADODB.Recordset

Dim oiRS           As New ADODB.Recordset

Dim strFileName As String

Dim gUpID As String


Private Sub cmd_Click(Index As Integer)
 On Error Resume Next

    Dim FName

    '帅选文件
    Com.Filter = "XML文件(*.xml)|*.xml"
    Com.ShowOpen
    '得到文件名
    FName = Com.filename

    If FName <> "" Then
        Text1.text = Replace(FName, Chr(0), ",")

    End If
End Sub

Private Sub cmdddd_Click()
ExporToExcel ("select * from AA_MARKCODE_STD order by id  ")
MsgBox "导出完毕", vbInformation, "提示"

End Sub

Private Sub CmdClearOI_Click()
    ClearData

End Sub

Private Sub ClearData()
    TxtCustomer.text = ""
    TxtPO.text = ""
    TxtPOItem.text = ""
    TxtLotId.text = ""
    TxtMpn.text = ""

    TxtMpnDesc.text = ""
    TxtWaferQty.text = ""
    TxtDieQty.text = ""
    TxtDesign.text = ""
    TxtCountryFab.text = ""

    TxtImageRev.text = ""
    TxtFFacility.text = ""
    TxtMarkId.text = ""
    TxtLotPriority.text = ""
    TxtFilmApld.text = ""

    TxtShip260.text = ""
    TxtShipLevel.text = ""
    TxtMicMaterial.text = ""
    TxtShipSite.text = ""
    TxtLotStatus.text = ""

    TxtCustomer.SetFocus

End Sub

Private Sub CmdSaveOI_Click()

    Dim oiRecordTemp As OIRecord

    If TxtWaferQty.text = "" Then
        MsgBox "片数不可以为空！"
        Exit Sub

    End If

    If TxtDieQty.text = "" Then
        MsgBox "片数不可以为空！"
        Exit Sub

    End If

    oiRecordTemp.id = GetMaxID()
    oiRecordTemp.PoNum = Trim(TxtPO.text)
    oiRecordTemp.PoItem = Trim(TxtPOItem.text)
    oiRecordTemp.LOTID = Trim(TxtLotId.text)
    oiRecordTemp.MPN = Trim(TxtMpn.text)
    oiRecordTemp.MPNDec = Trim(TxtMpnDesc.text)

    oiRecordTemp.WaferQTY = CInt(Trim(TxtWaferQty.text))
    oiRecordTemp.DIEQTY = CInt(Trim(TxtDieQty.text))
    oiRecordTemp.DESIGNID = Trim(TxtDesign.text)
    oiRecordTemp.CountryFab = Trim(TxtCountryFab.text)
    oiRecordTemp.ImageRev = Trim(TxtImageRev.text)

    oiRecordTemp.FFacility = Trim(TxtFFacility.text)
    oiRecordTemp.MarkId = Trim(TxtMarkId.text)
    oiRecordTemp.LotPriority = Trim(TxtLotPriority.text)
    oiRecordTemp.FilmApld = Trim(TxtFilmApld.text)
    oiRecordTemp.Ship260 = Trim(TxtShip260.text)

    oiRecordTemp.ShipLevel = Trim(TxtShipLevel.text)
    oiRecordTemp.MicMaterial = Trim(TxtMicMaterial.text)
    oiRecordTemp.ShipSite = Trim(TxtShipSite.text)
    oiRecordTemp.LotStatus = Trim(TxtLotStatus.text)
    oiRecordTemp.customerName = Trim(TxtCustomer.text)

    oiRecordTemp.FLAG = "Y"
    oiRecordTemp.CreateBy = "Auto"

    Call AddOIRecord(oiRecordTemp)

    ClearData

End Sub

Private Sub Qtech_OrderMapping()

    SumCount = 0
    ErrorInf = ""

    If Text1.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub
    
    End If
    
    Dim filename As String

    filename = Text1.text

    Dim dirtemp() As String

    Dim i         As Integer
    
    If InStr(1, filename, ",") > 0 Then
        dirtemp = Split(filename, ",")
        
        For i = 1 To UBound(dirtemp)
            UpMxlForQtech (dirtemp(0) + "\" + dirtemp(i))
        Next
        
    Else
        
        UpMxlForQtech (filename)

    End If
    
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
    
    If ErrorInf <> "" Then
        MsgBox "上传失败的有:" + ErrorInf + "数据库中已存在！"

    End If

End Sub

Private Sub Command1_Click()
    SumCount = 0

    '2013-01-21 jiayun add  Qtech 自购 Mapping
    If Combo1.text = "自购" Then

        Qtech_OrderMapping

    Else

        SumCount = 0
        ErrorInf = ""

        If Text1.text = "" Then
            MsgBox "先选择待上传的文件"
            Exit Sub
    
        End If
    
        Dim filename As String

        filename = Text1.text

        Dim dirtemp() As String

        Dim i         As Integer
    
        If InStr(1, filename, ",") > 0 Then
            dirtemp = Split(filename, ",")
        
            For i = 1 To UBound(dirtemp)
                UpMxl (dirtemp(0) + "\" + dirtemp(i))
            Next
        
        Else
        
            UpMxl (filename)

        End If
    
        If SumCount > 0 Then
            MsgBox "已成功上传" & SumCount & "笔！"

        End If
    
        If ErrorInf <> "" Then
            MsgBox "上传失败的有:" + ErrorInf + "数据库中已存在！"

        End If

    End If

End Sub

Private Sub UpMxl(dirtemp As String)

    '--定义XML

    Dim XMLDoc           As DOMDocument

    Dim xn               As IXMLDOMNode

    Dim xn01             As IXMLDOMNode

    Dim xn02             As IXMLDOMNode

    Dim xn03             As IXMLDOMNode

    Dim FLAG             As Integer

    Dim JudgeFlag        As Boolean

    Dim goodDieQty       As Integer

    Dim badDieQty        As Integer

    Dim customerNameTemp As String

    customerNameTemp = ""

    customerNameTemp = Combo1.text

    If customerNameTemp = "" Then
        customerNameTemp = "AA"

    End If

    Dim fileNameTemp As String

    fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)

    Set XMLDoc = New DOMDocument
    XMLDoc.Load dirtemp

    Set xn = XMLDoc.documentElement
    'SumCount = 0

    If Not xn Is Nothing Then
        '循环 Map

        For Each xn01 In xn.childNodes

            JudgeFlag = False
            goodDieQty = 0
            badDieQty = 0

            mapTemp.SUBSTRATEID = xn01.Attributes(1).nodeValue
        
            ' 判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
            If (JudgeFlagStauts(mapTemp.SUBSTRATEID)) Then
                '          MsgBox "这笔：" & mapTemp.SubstrateId & "已存在，无需上传!"
                ErrorInf = ErrorInf + "," + mapTemp.SUBSTRATEID

                GoTo NextRecord

            End If

            mapTemp.substratetype = xn01.Attributes(2).nodeValue

            '循环 Device
            If xn01.nodeName = "Map" Then

                For Each xn02 In xn01.childNodes

                    mapTemp.LOTID = xn02.Attributes(1).nodeValue
                    mapTemp.LOTID = Replace$(mapTemp.LOTID, ".", "")
                    mapTemp.PRODUCTID = xn02.Attributes(6).nodeValue
                    mapTemp.CreateDate = xn02.Attributes(8).nodeValue
                    mapTemp.MicronLotId = xn02.Attributes(14).nodeValue
                    mapTemp.MicronLotId = Replace$(mapTemp.MicronLotId, ".", "")

                    '循环 ReferenceDevice
                    If xn02.nodeName = "Device" Then
                        FLAG = 0

                        For Each xn03 In xn02.childNodes

                            '定义这一行的，三个临时变量
                            Dim field1      As String

                            Dim field2      As String

                            Dim field3      As String

                            Dim field1Value As String

                            Dim field2Value As String

                            Dim field3Value As String
                        
                            If xn03.nodeName = "Bin" Then
                                '2012-10-25 这行只有三个关键点 BinCode ,BinCount ,BinQuality
                                field1 = xn03.Attributes(0).nodeName
                                field1Value = xn03.Attributes(0).nodeValue
                            
                                field2 = xn03.Attributes(1).nodeName
                                field2Value = xn03.Attributes(1).nodeValue
                            
                                field3 = xn03.Attributes(2).nodeName
                                field3Value = xn03.Attributes(2).nodeValue
                            
                                If (field1 = "BinCode" And field1Value = "1") Or (field2 = "BinCode" And field2Value = "1") Or (field3 = "BinCode" And field3Value = "1") Then
                            
                                    '说明为良品数
                                    If field1 = "BinCount" Then
                                        goodDieQty = goodDieQty + CInt(field1Value)
                                
                                    ElseIf field2 = "BinCount" Then
                                        goodDieQty = goodDieQty + CInt(field2Value)
                                
                                    ElseIf field3 = "BinCount" Then
                                        goodDieQty = goodDieQty + CInt(field3Value)
                                
                                    End If

                                End If
                            
                                If (field1 = "BinCode" And (field1Value = "3" Or field1Value = "4" Or field1Value = "5")) Or (field2 = "BinCode" And (field2Value = "3" Or field2Value = "4" Or field2Value = "5")) Or (field3 = "BinCode" And (field3Value = "3" Or field3Value = "4" Or field3Value = "5")) Then

                                    '说明为不良品数
                                    If field1 = "BinCount" Then
                                        badDieQty = badDieQty + CInt(field1Value)
                                
                                    ElseIf field2 = "BinCount" Then
                                        badDieQty = badDieQty + CInt(field2Value)
                                
                                    ElseIf field3 = "BinCount" Then
                                        badDieQty = badDieQty + CInt(field3Value)
                                
                                    End If

                                End If
                            
                            ElseIf xn03.nodeName = "Data" Then

                                Exit For
                              
                            End If
                                  
                        Next   '  xn03 end
                    
                    End If   'Device end
                    
                    mapTemp.PASSBINCOUNT = goodDieQty
                    mapTemp.FailBinCount = badDieQty
                            
                Next

                '上传到DB
                mapTemp.filename = fileNameTemp
        
                '2014-04-22 jiayun  针对Y开头的，替换lotid 为文件名的
        
                '        If UCase(Mid(fileNameTemp, 1, 2)) = "YP" Then
                '
                '        mapTemp.lotid = Replace(Replace(fileNameTemp, ".xml", ""), ".XML", "")
                '
                '        End If
  
                '2016-02-20 jiayun 添加 AA lotid 取消._ -
        
                mapTemp.LOTID = Replace(Replace(Replace(Replace(Replace(fileNameTemp, ".xml", ""), ".XML", ""), ".", ""), "-", ""), "_", "")
        
                Call AddMap(mapTemp, customerNameTemp)
      
            End If

NextRecord:
        Next

    End If

End Sub

Private Sub UpMxlForQtech(dirtemp As String)
    'Qtech 自购Mapping 处理

    '--定义XML

    Dim XMLDoc       As DOMDocument

    Dim xn           As IXMLDOMNode

    Dim xn01         As IXMLDOMNode

    Dim xn02         As IXMLDOMNode

    Dim xn03         As IXMLDOMNode

    Dim FLAG         As Integer

    Dim JudgeFlag    As Boolean

    Dim goodDieQty   As Integer

    Dim badDieQty    As Integer

    Dim fileNameTemp As String

    fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)

    Set XMLDoc = New DOMDocument
    XMLDoc.Load dirtemp

    Set xn = XMLDoc.documentElement
    'SumCount = 0

    If Not xn Is Nothing Then
        '循环 Map

        For Each xn01 In xn.childNodes

            JudgeFlag = False
            goodDieQty = 0
            badDieQty = 0

            mapTemp.SUBSTRATEID = xn01.Attributes(1).nodeValue
        
            '判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
            If (JudgeFlagStauts(mapTemp.SUBSTRATEID)) Then
                '          MsgBox "这笔：" & mapTemp.SubstrateId & "已存在，无需上传!"
                ErrorInf = ErrorInf + "," + mapTemp.SUBSTRATEID

                GoTo NextRecord

            End If

            mapTemp.substratetype = xn01.Attributes(2).nodeValue

            '循环 Device
            If xn01.nodeName = "Map" Then

                For Each xn02 In xn01.childNodes

                    mapTemp.LOTID = xn02.Attributes(1).nodeValue
                    mapTemp.LOTID = Replace$(mapTemp.LOTID, ".", "")
                    mapTemp.PRODUCTID = xn02.Attributes(6).nodeValue
                    mapTemp.CreateDate = xn02.Attributes(8).nodeValue
                    mapTemp.MicronLotId = xn02.Attributes(14).nodeValue
                    mapTemp.MicronLotId = Replace$(mapTemp.MicronLotId, ".", "")

                    '循环 ReferenceDevice
                    If xn02.nodeName = "Device" Then
                        FLAG = 0

                        For Each xn03 In xn02.childNodes

                            '定义这一行的，三个临时变量
                            Dim field1      As String

                            Dim field2      As String

                            Dim field3      As String

                            Dim field1Value As String

                            Dim field2Value As String

                            Dim field3Value As String
                        
                            If xn03.nodeName = "Bin" Then
                                '2012-10-25 这行只有三个关键点 BinCode ,BinCount ,BinQuality
                                field1 = xn03.Attributes(0).nodeName
                                field1Value = xn03.Attributes(0).nodeValue
                            
                                field2 = xn03.Attributes(1).nodeName
                                field2Value = xn03.Attributes(1).nodeValue
                            
                                field3 = xn03.Attributes(2).nodeName
                                field3Value = xn03.Attributes(2).nodeValue
                            
                                If (field1 = "BinCode" And field1Value = "G") Or (field2 = "BinCode" And field2Value = "G") Or (field3 = "BinCode" And field3Value = "G") Then
                            
                                    '说明为良品数
                                    If field1 = "BinCount" Then
                                        goodDieQty = goodDieQty + CInt(field1Value)
                                
                                    ElseIf field2 = "BinCount" Then
                                        goodDieQty = goodDieQty + CInt(field2Value)
                                
                                    ElseIf field3 = "BinCount" Then
                                        goodDieQty = goodDieQty + CInt(field3Value)
                                
                                    End If

                                End If
                            
                                If (field1 = "BinCode" And (field1Value = "X")) Or (field2 = "BinCode" And (field2Value = "X")) Or (field3 = "BinCode" And (field3Value = "X")) Then

                                    '说明为不良品数
                                    If field1 = "BinCount" Then
                                        badDieQty = badDieQty + CInt(field1Value)
                                
                                    ElseIf field2 = "BinCount" Then
                                        badDieQty = badDieQty + CInt(field2Value)
                                
                                    ElseIf field3 = "BinCount" Then
                                        badDieQty = badDieQty + CInt(field3Value)
                                
                                    End If

                                End If
                            
                            ElseIf xn03.nodeName = "Data" Then

                                Exit For
                              
                            End If
                                  
                        Next   '  xn03 end
                    
                    End If   'Device end
                    
                    mapTemp.PASSBINCOUNT = goodDieQty
                    mapTemp.FailBinCount = badDieQty
                            
                Next

                '上传到DB
                mapTemp.filename = fileNameTemp
                Call AddMap(mapTemp, "QT")
                SumCount = SumCount + 1

            End If

NextRecord:
        Next

    End If

End Sub

Private Sub Command10_Click()

    If TxtCustomerName.text = "" Then
        MsgBox "请先输入客户代码！"
        Exit Sub
    
    Else
 
        ExporToExcel ("select SUBSTRATEID, SUBSTRATETYPE, LOTID, PRODUCTID, CREATEDATE,MICRONLOTID, PASSBINCOUNT, FAILBINCOUNT, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE ,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE from mappingDataTest where customershortname ='" & Trim(TxtCustomerName.text) & "' and qtech_created_date>sysdate-30  order by qtech_created_date desc , lotid, substrateid")

    End If

End Sub

Private Sub Command11_Click()

    Dim mapTemp As MapRecord

    If TxtCustomerName.text = "" Then
        MsgBox "请先输入客户代码！"
        Exit Sub

    End If

    If Text4.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    Dim dirName  As String

    Dim filename As String

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text4.text)    '打开文件

    Set xlSheet = xlBook.Worksheets("sheet1")        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 5 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0
    BCResultFlag = False

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
    
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
           
            If j = 1 Then
                mapTemp.SUBSTRATEID = Trim(tempVal) 'WaferId
            
                '判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
                If (JudgeFlagStauts(mapTemp.SUBSTRATEID)) Then
                    MsgBox "这笔：" & mapTemp.SUBSTRATEID & "已存在，无需上传!"
                    '              ErrorInf = ErrorInf + "," + mapTemp.SubstrateId
              
                    GoTo NextRecord2
    
                End If
            
            End If
        
            If j = 2 Then
                mapTemp.LOTID = Trim(tempVal) 'LotId

            End If
        
            If j = 3 Then
                mapTemp.PRODUCTID = Trim(tempVal) 'ProductId

            End If
        
            If j = 4 Then
                mapTemp.PASSBINCOUNT = Trim(tempVal) 'PassBinCount

            End If
        
            If j = 5 Then
                mapTemp.FailBinCount = Trim(tempVal) 'FailBinCount

            End If
        
        Next j
    
        mapTemp.CreateDate = ""
        mapTemp.MicronLotId = ""
        mapTemp.filename = ""
    
        Call AddMap2(mapTemp, Trim$(TxtCustomerName.text))
        SumCount = SumCount + 1

NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！", , "友情提醒"
    
    Else

        If BCResultFlag = True Then
            MsgBox "上传失败，请确认资料格式！", , "友情提醒"
            Exit Sub

        End If
    
    End If

End Sub

Private Sub Command12_Click()

    '打开选择文件
    On Error Resume Next

    Dim FName

    '帅选文件
    CommonDialog1.Filter = "EXCEL文件(*.xls)|*.xls"
    CommonDialog1.ShowOpen
    '得到文件名
    FName = CommonDialog1.filename

    If FName <> "" Then
        Text4.text = FName

    End If

End Sub

Private Sub Command13_Click()

    On Error Resume Next

    'si map

    Dim FName

    '帅选文件
    ComSI.Filter = "map文件(*.map)|*.map|txt文件(*.txt)|*.txt|XML文件(*.XML)|*.XML|CSV文件(*.csv)|*.csv|smic文件(*.smic)|*.smic"

    ComSI.ShowOpen
    '得到文件名
    FName = ComSI.filename

    If FName <> "" Then
        'TxtSI.Text = Replace(FName, " ", "")
        TxtSI.text = FName

    End If

End Sub

Private Sub Command14_Click()
    'si map

    If CmbCustomer.text = "" Then
        MsgBox "请先选择客户！"
        Exit Sub

    End If

    If CmbCustomer.text = "GC" Then

        If CmbGCType.text = "" Then
            MsgBox "GC客户上传Mapping时，请先选择是GC哪一型产品，再点上传按钮！"
            Exit Sub

        End If

    End If

    SumCount = 0
    ErrorInf = ""

    If TxtSI.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub
    
    End If
    
    Dim filename As String

    filename = TxtSI.text

    Dim dirtemp() As String

    Dim i         As Integer
    
    If InStr(1, filename, " ") > 0 Then
        dirtemp = Split(filename, " ")
        
        For i = 1 To UBound(dirtemp)

            If CmbCustomer.text = "GT" Or CmbCustomer.text = "SI" Then
             
                UpMap (dirtemp(0) + "\" + dirtemp(i))
            
            ElseIf CmbCustomer.text = "HD" Then
                'HD客户
                UpMapHD (dirtemp(0) + "\" + dirtemp(i))
                 
            ElseIf CmbCustomer.text = "GC" And CmbGCType.text = "WLT" Then
                'GC WLT客户
                UpMapGCWlt (dirtemp(0) + "\" + dirtemp(i))
                 
            ElseIf CmbCustomer.text = "GC" And CmbGCType.text = "COG" Then
                'GC客户 COG
                UpMapGCCOG (dirtemp(0) + "\" + dirtemp(i))
                 
            ElseIf CmbCustomer.text = "MG" Then
                'MG客户
                UpMapMG (dirtemp(0) + "\" + dirtemp(i))
                  
            ElseIf CmbCustomer.text = "56" Then
             
                UpMap56 (dirtemp(0) + "\" + dirtemp(i))
            ElseIf CmbCustomer.text = "95" Then
             
                UpMap95 (dirtemp(0) + "\" + dirtemp(i))
                
            ElseIf CmbCustomer.text = "TW058" Then
             
                UpMapTW058 (dirtemp(0) + "\" + dirtemp(i))
            
            End If
            
        Next
        
    Else

        If CmbCustomer.text = "GT" Or CmbCustomer.text = "SI" Then
        
            UpMap (filename)
        
        ElseIf CmbCustomer.text = "HD" Then
            'HD客户
            UpMapHD (filename)
        ElseIf CmbCustomer.text = "GC" And CmbGCType.text = "WLT" Then
            'GC客户   2015-03-20 jiayun add
            UpMapGCWlt (filename)
         
        ElseIf CmbCustomer.text = "GC" And CmbGCType.text = "COG" Then
            'GC客户 COG
            UpMapGCCOG (filename)
         
        ElseIf CmbCustomer.text = "MG" Then
            UpMapMG (filename)
         
        ElseIf CmbCustomer.text = "56" Then
        
            UpMap56 (filename)
        ElseIf CmbCustomer.text = "95" Then
             
            UpMap95 (filename)
        
        ElseIf CmbCustomer.text = "TW058" Then
             
            UpMapTW058 (filename)

        End If
        
    End If
    
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
    
    If ErrorInf <> "" Then
        MsgBox "上传失败的有:" + ErrorInf + "数据库中已存在！"

    End If

End Sub

Private Sub UpMap(dirtemp As String)

    Dim FLAG             As Integer

    Dim JudgeFlag        As Boolean

    Dim customerNameTemp As String

    Dim waferIDSeq       As String

    Dim allDieQty        As Integer

    Dim goodDieQty       As Integer

    Dim badDieQty        As Integer

    Dim fileNameTemp     As String

    fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
    mapTemp.filename = fileNameTemp
    customerNameTemp = "GT"
 
    'SI Mapping

    Dim TextLine As String

    Open dirtemp For Input As #1

    ' 打开文件。
    Do While Not EOF(1)
        ' 循环至文件尾。
        Line Input #1, TextLine

        '判断这行，是否要取资料，是则处理；否则下一行
        If InStr(TextLine, "LOT_NO") > 0 Then
            'lotid
            mapTemp.LOTID = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
            waferIDSeq = Trim(Mid(TextLine, InStrRev(TextLine, ":") + 1, Len(TextLine) - InStrRev(TextLine, ":")))
            mapTemp.SUBSTRATEID = mapTemp.LOTID & waferIDSeq
     
        End If
    
        If InStr(TextLine, "GOOD_DIE") > 0 Then
            'qty
            mapTemp.PASSBINCOUNT = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
     
            allDieQty = Trim(Mid(TextLine, InStrRev(TextLine, ":") + 1, Len(TextLine) - InStrRev(TextLine, ":")))
     
            mapTemp.FailBinCount = allDieQty - mapTemp.PASSBINCOUNT
    
        End If

        If InStr(TextLine, "[FLAT") > 0 Then
            GoTo ContinueFlag
    
        End If

    Loop

ContinueFlag:

    Close #1    ' 关闭文件。

    ' 判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
       
    If (JudgeFlagStauts(mapTemp.SUBSTRATEID)) Then
        MsgBox "这笔：" & mapTemp.SUBSTRATEID & "已存在，无需上传!"
       
    Else
       
        Call AddMap(mapTemp, customerNameTemp)

    End If

End Sub

' 56 old  jiayun backup 2016-04-25
'Private Sub UpMap56(dirtemp As String)
'Dim Flag As Integer
'Dim JudgeFlag As Boolean
'Dim customerNameTemp As String
'Dim productaNameTenp As String
'
'Dim waferIDSeq As String
'Dim allDieQty As Long
'Dim goodDieQty As Long
'Dim badDieQty As Long
'
'Dim fileNameTemp As String
'fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
'mapTemp.FileName = fileNameTemp
'customerNameTemp = "56"
'
''56 Mapping
'
'Dim TextLine As String
'Open dirtemp For Input As #1
'' 打开文件。
'Do While Not EOF(1)
'' 循环至文件尾。
'Line Input #1, TextLine
'
'    '判断这行，是否要取资料，是则处理；否则下一行
'    If InStr(TextLine, "Product Name") > 0 Then
'
'    mapTemp.SubstrateType = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
'
'    End If
'
'
'     If InStr(TextLine, "Lot id") > 0 Then
'    mapTemp.lotid = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
'    End If
'
'
'     If InStr(TextLine, "Wafer ID") > 0 Then
'     waferIDSeq = Right("0" & Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20)), 2)
'     mapTemp.SubstrateId = mapTemp.lotid & waferIDSeq
'    End If
'
'     If InStr(TextLine, "Gross Dice") > 0 Then
'    'qty
'     allDieQty = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
'
'     End If
'
'     If InStr(TextLine, "Good Dice") > 0 Then
'    'qty
'     mapTemp.PassBinCount = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
'
'     mapTemp.FailBinCount = CLng(allDieQty) - mapTemp.PassBinCount
'
'    End If
'
'
'    If InStr(TextLine, "Yield") > 0 Then
'      GoTo ContinueFlag
'
'    End If
'
'
'
'Loop
'
'
'ContinueFlag:
'
'
'Close #1    ' 关闭文件。
'
'       ' 判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
'
'       If (JudgeFlagStauts(mapTemp.SubstrateId)) Then
'            MsgBox "这笔：" & mapTemp.SubstrateId & "已存在，无需上传!"
'
'       Else
'
'            Call AddMap(mapTemp, customerNameTemp)
'
'       End If
'
'End Sub

Private Sub UpMap56(dirtemp As String)

    Dim FLAG             As Integer

    Dim JudgeFlag        As Boolean

    Dim customerNameTemp As String

    Dim productaNameTenp As String

    Dim waferIDSeq       As String

    Dim allDieQty        As Long

    Dim goodDieQty       As Long

    Dim badDieQty        As Long

    Dim kk1              As String

    Dim kk2              As String

    Dim fileNameTemp     As String

    fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
    mapTemp.filename = fileNameTemp
    customerNameTemp = "56"
 
    '56 Mapping

    Dim TextLine As String

    Dim temp1    As String

    Dim temp2    As String

    Open dirtemp For Input As #1

    ' 打开文件。
    Do While Not EOF(1)
        ' 循环至文件尾。
        Line Input #1, TextLine

        '判断这行，是否要取资料，是则处理；否则下一行
        If InStr(TextLine, "WAFER_BATCH_ID") > 0 Then
    
            kk1 = Trim(TextLine)
            temp1 = Mid(kk1, 17, Len(kk1) - 34 + 1)
   
            kk2 = Mid(temp1, InStr(temp1, "-") + 1, Len(temp1) - InStr(temp1, "-") + 1)
   
            mapTemp.LOTID = Left(kk2, InStr(kk2, "-") - 1)
   
            'D7W183-CP2-1
            temp2 = Right(kk2, Len(kk2) - InStr(kk2, "-CP2-") - 4)
   
            waferIDSeq = Right(CStr("0" & temp2), 2)
   
            mapTemp.SUBSTRATEID = mapTemp.LOTID & waferIDSeq

        End If
    
        If InStr(TextLine, "BIN_COUNT_PASS") > 0 Then
    
            kk1 = Trim(TextLine)
   
            temp1 = Mid(kk1, 17, Len(kk1) - 34 + 1)
   
            kk2 = Mid(temp1, InStr(temp1, "-") + 1, Len(temp1) - InStr(temp1, "-"))
   
            goodDieQty = CLng(kk2)
   
            mapTemp.PASSBINCOUNT = goodDieQty
   
            mapTemp.FailBinCount = 98184 - goodDieQty

        End If

        If InStr(TextLine, "</HEADER>") > 0 Then
            GoTo ContinueFlag
    
        End If

    Loop

ContinueFlag:

    Close #1    ' 关闭文件。

    ' 判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
       
    If (JudgeFlagStauts(mapTemp.SUBSTRATEID)) Then
        MsgBox "这笔：" & mapTemp.SUBSTRATEID & "已存在，无需上传!"
       
    Else
       
        Call AddMap56(mapTemp, customerNameTemp, CInt(waferIDSeq))

    End If

End Sub

Private Sub UpMap95(dirtemp As String)

    Dim FLAG             As Integer

    Dim JudgeFlag        As Boolean

    Dim customerNameTemp As String

    Dim productaNameTenp As String

    Dim waferIDSeq       As String

    Dim allDieQty        As Long

    Dim goodDieQty       As Long

    Dim badDieQty        As Long

    Dim kk1              As String

    Dim kk2              As String

    Dim fileNameTemp     As String

    fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
    mapTemp.filename = fileNameTemp
    customerNameTemp = "95"

    Dim TextLine As String

    Dim temp1    As String

    Dim temp2    As String

    Open dirtemp For Input As #1

    ' 打开文件。
    Do While Not EOF(1)
        ' 循环至文件尾。
        Line Input #1, TextLine

        '取LOT号
        If InStr(TextLine, "Lot ID:") > 0 Then

            mapTemp.LOTID = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
   
        End If
    
        '取WAFERid
        If InStr(TextLine, "smic") > 0 Then
            waferIDSeq = Trim(Mid(TextLine, InStr(TextLine, "-") + 1, 2))
            mapTemp.SUBSTRATEID = mapTemp.LOTID & waferIDSeq
   
            mapTemp.PASSBINCOUNT = Val(Right(Trim(TextLine), 3))
            mapTemp.FailBinCount = 705 - Val(Right(Trim(TextLine), 3))

        End If
    
        Call AddMap95(mapTemp, customerNameTemp)
    
        If InStr(TextLine, "Dice:") > 0 Then
            GoTo ContinueFlag
    
        End If

    Loop

ContinueFlag:

    Close #1

    MsgBox "已上传成功请查询确认对错 ！"

End Sub

Private Sub UpMapTW058(dirtemp As String)

    Dim FLAG             As Integer

    Dim JudgeFlag        As Boolean

    Dim customerNameTemp As String

    Dim productaNameTenp As String

    Dim waferIDSeq       As String

    Dim pj               As String

    Dim pj1              As String

    Dim allDieQty        As Long

    Dim goodDieQty       As Long

    Dim badDieQty        As Long

    Dim fileNameTemp     As String

    fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
    mapTemp.filename = fileNameTemp
    customerNameTemp = "TW058"
 
    '56 Mapping

    Dim TextLine As String

    Open dirtemp For Input As #1

    ' 打开文件。
    Do While Not EOF(1)
        ' 循环至文件尾。
        Line Input #1, TextLine

        '判断这行，是否要取资料，是则处理；否则下一行
        If InStr(TextLine, "Product Name") > 0 Then
    
            mapTemp.substratetype = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
    
        End If
    
        If InStr(TextLine, "Lot No") > 0 Then
            mapTemp.LOTID = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))

        End If
    
        If InStr(TextLine, "Slot No") > 0 Then
            waferIDSeq = Right("0" & Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20)), 2)
            mapTemp.SUBSTRATEID = mapTemp.LOTID & waferIDSeq

        End If
    
        If InStr(TextLine, "Total") > 0 Then
            'qty
     
            allDieQty = Trim(Mid(TextLine, InStr(TextLine, "=") + 1, InStr(TextLine, "e") - InStr(TextLine, "=") - 1))
            mapTemp.FailBinCount = allDieQty - mapTemp.PASSBINCOUNT
     
        End If
    
        If InStr(TextLine, "Bin  1") > 0 Then
            'qty
      
            mapTemp.PASSBINCOUNT = Trim(Mid(TextLine, InStr(TextLine, "=") + 1, InStr(TextLine, "e") - InStr(TextLine, "=") - 1))

        End If

        If InStr(TextLine, "Yield") > 0 Then
            GoTo ContinueFlag
    
        End If

    Loop

ContinueFlag:

    Close #1    ' 关闭文件。

    Call AddMap95(mapTemp, customerNameTemp)

End Sub

'2015-04-20 jiayun add MG

Private Sub UpMapMG(dirtemp As String)

    Dim FLAG             As Integer

    Dim JudgeFlag        As Boolean

    Dim customerNameTemp As String

    Dim waferIDSeq       As String

    Dim allDieQty        As Integer

    Dim goodDieQty       As Integer

    Dim badDieQty        As Integer

    Dim fileNameTemp     As String

    fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
    mapTemp.filename = fileNameTemp
    customerNameTemp = "MG"
 
    'SI Mapping

    Dim TextLine As String

    Open dirtemp For Input As #1

    ' 打开文件。
    Do While Not EOF(1)
        ' 循环至文件尾。
        Line Input #1, TextLine

        '判断这行，是否要取资料，是则处理；否则下一行
        If InStr(TextLine, "LOT_NO") > 0 Then
            'lotid
            mapTemp.LOTID = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
            waferIDSeq = Trim(Mid(TextLine, InStrRev(TextLine, ":") + 1, 3))
            mapTemp.SUBSTRATEID = mapTemp.LOTID & waferIDSeq
     
        End If
    
        If InStr(TextLine, "GOOD_DIE") > 0 Then
            'qty
            mapTemp.PASSBINCOUNT = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
     
            allDieQty = Trim(Mid(TextLine, InStrRev(TextLine, ":") + 1, Len(TextLine) - InStrRev(TextLine, ":")))
     
            mapTemp.FailBinCount = allDieQty - mapTemp.PASSBINCOUNT
    
        End If

        If InStr(TextLine, "TEST_TIME") > 0 Then
            GoTo ContinueFlag
    
        End If

    Loop

ContinueFlag:

    Close #1    ' 关闭文件。

    ' 判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
       
    '       If (JudgeFlagStauts(mapTemp.SubstrateId)) Then
    '            MsgBox "这笔：" & mapTemp.SubstrateId & "已存在，无需上传!"
    '
    '       Else
       
    'Call AddMap(mapTemp, customerNameTemp)
            
    Call updateMGMap(mapTemp.SUBSTRATEID, mapTemp.PASSBINCOUNT, mapTemp.FailBinCount)

    '       End If

End Sub

Private Sub UpMapHD(dirtemp As String)

    Dim FLAG             As Integer

    Dim JudgeFlag        As Boolean

    Dim customerNameTemp As String

    Dim waferIdTemp      As String

    Dim waferIDSeq       As String

    Dim allDieQty        As Integer

    Dim goodDieQty       As Integer

    Dim badDieQty        As Integer

    Dim fileNameTemp     As String

    fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
    mapTemp.filename = fileNameTemp
    customerNameTemp = "HD"

    Dim lotlot   As String
 
    'SI Mapping

    Dim TextLine As String

    Open dirtemp For Input As #1

    ' 打开文件。
    Do While Not EOF(1)
        ' 循环至文件尾。
        Line Input #1, TextLine

        '判断这行，是否要取资料，是则处理；否则下一行
        'LotID
        'If InStr(TextLine, "Lot No") > 0 Then
        If InStr(TextLine, "Lot No.") > 0 Then
            'lotid
            mapTemp.LOTID = Mid(Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20)), 1, 6)
            '     Len(Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))))
            '     waferIDSeq = Trim(Mid(TextLine, InStrRev(TextLine, ":") + 1, Len(TextLine) - InStrRev(TextLine, ":")))
            '     mapTemp.SubstrateId = mapTemp.lotID & waferIDSeq
    
            '     mapTemp.lotid = Mid(mapTemp.lotid, 1, InStr(mapTemp.lotid, ".") - 1)
     
        End If
    
        'WaferID
        'If InStr(TextLine, "Wafer ID") > 0 Then
        If InStr(TextLine, "Wafer ID") > 0 Then
            'lotid
            ' mapTemp.lotID = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     
            'D02939-1
            waferIdTemp = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
            waferIdTemp = Mid(waferIdTemp, InStr(waferIdTemp, "-") + 1, 2)
     
            waferIDSeq = Right("0" & waferIdTemp, 2)
            mapTemp.SUBSTRATEID = mapTemp.LOTID & waferIDSeq
     
        End If
    
        If InStr(TextLine, "Total Tested") > 0 Then '获得总数

            mapTemp.TotalQty = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
     
        End If
    
        'If InStr(TextLine, "Total Pass") > 0 Then
        If InStr(TextLine, "Total Pass") > 0 Then
            'qty
            mapTemp.PASSBINCOUNT = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
     
            '     allDieQty = Trim(Mid(TextLine, InStrRev(TextLine, ":") + 1, Len(TextLine) - InStrRev(TextLine, ":")))
            '
            '     mapTemp.FailBinCount = allDieQty - mapTemp.PassBinCount
    
        End If
    
        'If InStr(TextLine, "Total Fail") > 0 Then
        If InStr(TextLine, "Yield") > 0 Then  '改变算法现在的模板里面没有NGDIE需要用总数-良品得到NG数量
            'qty
            mapTemp.FailBinCount = mapTemp.TotalQty - mapTemp.PASSBINCOUNT
     
        End If

        If InStr(TextLine, "Yield") > 0 Then
            GoTo ContinueFlag
    
        End If

    Loop

ContinueFlag:

    Close #1    ' 关闭文件。

    ' 判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
       
    If (JudgeFlagStautsMapping2(mapTemp.SUBSTRATEID)) Then
        MsgBox "这笔：" & mapTemp.SUBSTRATEID & "已存在，无需上传!"
       
    Else
       
        Call AddTSVMap(mapTemp, customerNameTemp)

    End If

End Sub

Private Sub UpMapGCWlt(dirtemp As String)

    Dim customerNameTemp As String

    Dim waferidGCTemp    As String

    Dim gcGoodDieQty     As Long

    Dim fileNameTemp     As String

    fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
    mapTemp.filename = fileNameTemp
    customerNameTemp = "GC"
 
    'SI Mapping

    Dim TextLine As String

    Open dirtemp For Input As #1

    ' 打开文件。
    Do While Not EOF(1)
        ' 循环至文件尾。
        Line Input #1, TextLine

        '判断这行，是否要取资料，是则处理；否则下一行
        'LotID
        If InStr(TextLine, "Lot:") > 0 Then

            waferidGCTemp = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     
        End If
  
        If InStr(TextLine, "BIN_1") > 0 Then
    
            gcGoodDieQty = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
      
            GoTo ContinueFlag
    
        End If

    Loop

ContinueFlag:

    Close #1    ' 关闭文件。
       
    Call updateGCWltMap(waferidGCTemp, gcGoodDieQty)

End Sub

Private Sub UpMapGCCOG(dirtemp As String)

    Dim TPriceFlag           As Boolean

    Dim source_batch_id_Temp As String

    Dim dirName              As String

    Dim filename             As String

    TPriceFlag = False

    '获取文件名
    If InStrRev(Trim(dirtemp), "\") > 0 Then
        strFileName = Mid(Trim(dirtemp), InStrRev(Trim(dirtemp), "\") + 1)
        dirName = Mid$(Trim(dirtemp), 1, InStrRev(Trim(dirtemp), "\"))

    End If

    Dim con As New ADODB.Connection

    Dim rs  As New ADODB.Recordset

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(dirtemp)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    Dim i             As Integer

    Dim j             As Integer

    Dim id            As Long

    Dim TEMP          As String

    Dim temp2         As String

    Dim tempVal       As String

    Dim WV_inspect    As String

    Dim Comp_codeTemp As String

    Dim waferIdTemp   As String

    Dim gdQtyTemp     As Long

    Dim ngQtyTemp     As Long

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count

        strChar = Chr(96 + 1)
        tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
 
        waferIdTemp = Replace(Mid(Trim(tempVal), 5, 11), "-", "")
 
        strChar = Chr(96 + 5)
        tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
        gdQtyTemp = CLng(tempVal)
  
        strChar = Chr(96 + 6)
        tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
        ngQtyTemp = CLng(tempVal)
       
        Call updateGCCOGMap(waferIdTemp, gdQtyTemp, ngQtyTemp)
    
    Next i

End Sub

Private Sub Command15_Click()
    ExporToExcel ("select SUBSTRATEID, SUBSTRATETYPE, LOTID, PRODUCTID, CREATEDATE,MICRONLOTID, PASSBINCOUNT, FAILBINCOUNT, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE ,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE from mappingDataTest where customershortname  in ('SI','GT')  and qtech_created_date>sysdate-30  order by qtech_created_date desc , lotid, substrateid")

End Sub

Private Sub Command16_Click()

    Dim tt As String

    Set xml = New DOMDocument
    Call xml.Load("C:\FJ801Z-1.XML")    'index.xml为描述图书信息的XML文档

    Dim root As IXMLDOMElement

    Set root = xml.documentElement

    Dim node As IXMLDOMNode

    For Each node In root.childNodes

        tt = tt & node.text
    Next

End Sub

Private Sub Command17_Click()

End Sub

Private Sub Command2_Click()

    On Error Resume Next

    Dim FName

    '帅选文件
    CommonDialog1.Filter = "CSV文件(*.csv)|*.csv"
    CommonDialog1.ShowOpen
    '得到文件名
    FName = CommonDialog1.filename

    If FName <> "" Then
        Text2.text = FName

    End If

End Sub

Private Sub Command3_Click()

    Dim source_batch_id_Temp As String

    '上传OI的CSV
    '处理文件名
    If Text2.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    Dim dirName  As String

    Dim filename As String

    '获取文件名
    If InStrRev(Trim(Text2.text), "\") > 0 Then
        strFileName = Mid(Trim(Text2.text), InStrRev(Trim(Text2.text), "\") + 1)
        dirName = Mid$(Trim(Text2.text), 1, InStrRev(Trim(Text2.text), "\"))

    End If

    Dim con As New ADODB.Connection

    Dim rs  As New ADODB.Recordset

    'con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
    'Rs.open "Select * From " & strfilename, con, adOpenStatic, adLockReadOnly, adCmdText

    '2012-07-03 jiayunzhang 修改读CSV的方式

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text2.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同
    '2012-10-08 jiayunzhang 市场部要求新增一列 comp_code

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 73 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i             As Integer

    Dim j             As Integer

    Dim id            As Long

    Dim TEMP          As String

    Dim temp2         As String

    Dim tempVal       As String

    Dim WV_inspect    As String

    Dim Comp_codeTemp As String

    Dim SumCount      As Integer

    SumCount = 0
    'Rs.MoveFirst
    'For i = 0 To Rs.RecordCount - 1

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count

        TEMP = ""
        source_batch_id_Temp = ""
        '    For j = 0 To Rs.fields.Count - 1

        '2012-07-03 因客户OI添加字段，数据库新增在最后一列，所以程序要特殊处理。 把列数，xlSheet.Range("A1").CurrentRegion.Columns.Count 改为 71
        For j = 1 To 71
      
            If j > 26 Then
                strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
            Else
                strChar = Chr(96 + j)

            End If
      
            '        strChar = Chr(96 + j)
        
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
    
            If j = 46 Or j = 60 Then
                TEMP = TEMP & "," & newStrDate("" & tempVal)
        
            Else

                If j = 61 Then
                    tempVal = Format(xlSheet.Range(strChar & i).Value, "HH:mm:SS")
                    TEMP = TEMP & "," & newStr("" & tempVal)
                Else
            
                    TEMP = TEMP & "," & newStr("" & tempVal)

                End If
        
            End If
        
            If j = 3 Then
                source_batch_id_Temp = tempVal

            End If
    
        Next j
    
        j = 72
        strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
    
        WV_inspect = newStr("" & xlSheet.Range(strChar & i).Value)
    
        j = 73
        strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
    
        Comp_codeTemp = newStr("" & xlSheet.Range(strChar & i).Value)
    
        '取目前DB最大的ID号
        id = GetMaxID()
        TEMP = id & TEMP
        temp2 = TEMP & ",'Y','" & gUserName & "',GETDATE(),'','','AA',0," & WV_inspect & "," & Comp_codeTemp
        TEMP = TEMP & ",'Y','" & gUserName & "',sysdate,'','','AA',0,1," & WV_inspect & "," & Comp_codeTemp
    
        '    Debug.Print temp

        '             '判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
        If (JudgeFlagStautsOI(source_batch_id_Temp)) Then
            MsgBox "这笔：" & source_batch_id_Temp & "已存在，无需上传!"
            GoTo NextRecord2

        End If
    
        Call AddOI(TEMP, temp2)
        SumCount = SumCount + 1
    
        '上传到DB
    
NextRecord2:
        '    Rs.MoveNext

    Next i

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If

End Sub

Private Function newStrDate(TEMP As String)

    Dim mmTemp  As String

    Dim ddTemp  As String

    Dim newTemp As String

    '2012-09-14 jiayunzhang Modify 时间格式不需转化。
    If TEMP <> "" Then

        '    mmTemp = Mid$(temp, 6, InStr(6, temp, "-") - 6)
        '    ddTemp = Right$(temp, Len(temp) - InStr(6, temp, "-"))
    
        '    If Val(mmTemp) >= 1 And Val(mmTemp) <= 12 And Val(ddTemp) >= 1 And Val(ddTemp) <= 12 Then
        '        '此时需要转换
        '
        '        newTemp = Left$(temp, 4) & "-" & ddTemp & "-" & mmTemp
        '        newStrDate = "'" & newTemp & "'"
        '
        '    Else
        newStrDate = "'" & TEMP & "'"
        '    End If

    Else
        newStrDate = "''"

    End If

End Function

Private Function newStr(TEMP As String)

    If TEMP <> "" Then
        newStr = "'" & TEMP & "'"
    Else
        newStr = "''"

    End If

End Function

Private Sub Command4_Click()
    ExporToExcel ("select SUBSTRATEID, SUBSTRATETYPE, LOTID, PRODUCTID, CREATEDATE,MICRONLOTID, PASSBINCOUNT, FAILBINCOUNT, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE ,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE from mappingDataTest where customershortname ='AA' and qtech_created_date>sysdate-90  order by qtech_created_date desc , lotid, substrateid")

End Sub

Private Sub Command5_Click()

    '    ExporToExcel (" select ID,PO_NUM,PO_ITEM,SOURCE_BATCH_ID,SOURCE_MTRL_NUM,MTRL_NUM,MTRL_DESC,TEST_MTRL_NUM,TEST_MTRL_DESC, MPN, " & _
    '                 " MPN_DESC, SOURCE_MTRL_SLOC, MTRL_NUM_MTRLGRP,PROBE_SHIP_PART_TYPE, OFFSHORE_ASM_COMPANY,OFFSHORE_TEST_COMPANY, CURRENT_WAFER_QTY, DIE_QTY, DESIGN_ID,COUNTRY_OF_FAB," & _
    '                 " FAB_CONV_ID,FAB_EXCR_ID,RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,WAFER_SIZE,IMAGER_CUSTOMER_REV, CHROMATICITY, MICRO_LENS_SHIFT, TEMPERATURE_SPEC," & _
    '                 " PRB_CONTAINMENT_TYPE, FABRICATION_FACILITY, PRB_EXCR_ID, BATCH_COMMENT_PROBE, ASSY_PROCESS_ID, DARK_BOND_PAD_ASSY, ASSY_SERIAL_TYPE, STICKY_BACKS_TO_SAVE, OPTICAL_QUALITY, ENCODED_MARK_ID, " & _
    '                 " PLANNED_LASER_SCRIBE, PACKAGE_LID_TYPE, PACKAGE_TYPE, PB_FREE_PACKAGE, TARGET_WAF_THICKNESS, RELIABILITY_SAMPLING, LOT_PRIORITY, WAFER_BOX_TYPE, TEST_SITE,ASSEMBLY_FACILITY, " & _
    '                 " BATCH_COMMENT_ASSY, TST_PROCESS_ID,ELEC_SPECIAL_TEST, BOX_TYPE, PROTECTIVE_FILM_APLD, SHIPPING_MST_260,SHIPPING_MST_LEVEL, T_PRICE, SHIP_COMMENT, BATCH_COMMENT_TEST, " & _
    '                 " CREATED_DATE, CREATED_TIME, UNIT_PRICE,REF_PO, REF_PO_ITEM, COUNTRY_OF_ASSEMBLY, MICRON_MATERIAL,DATE_CODE, SHIP_SITE, SPECIAL_PROCESS_LOT, " & _
    '                 " LOT_STATUS, CUSTOM_PART_NO, FLAG,QTECH_CREATED_BY,QTECH_CREATED_DATE, QTECH_LASTUPDATE_BY, QTECH_LASTUPDATE_DATE from CustomerOItbl_test  where (customershortname = 'AA' or customershortname is null)  and (source_batch_id like '6%' or source_batch_id like '7%')  order by id ")
    '
    
    '2012-05-15 jiayunzhang Modify
    
    ExporToExcel (" select ID,PO_NUM,PO_ITEM,SOURCE_BATCH_ID,SOURCE_MTRL_NUM,MTRL_NUM,MTRL_DESC,TEST_MTRL_NUM,TEST_MTRL_DESC, MPN, " & _
       " MPN_DESC, SOURCE_MTRL_SLOC, MTRL_NUM_MTRLGRP,PROBE_SHIP_PART_TYPE, OFFSHORE_ASM_COMPANY,OFFSHORE_TEST_COMPANY, CURRENT_WAFER_QTY, DIE_QTY, DESIGN_ID,COUNTRY_OF_FAB," & _
       " FAB_CONV_ID,FAB_EXCR_ID,RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,WAFER_SIZE,IMAGER_CUSTOMER_REV, CHROMATICITY, MICRO_LENS_SHIFT, TEMPERATURE_SPEC," & _
       " PRB_CONTAINMENT_TYPE, FABRICATION_FACILITY, PRB_EXCR_ID, BATCH_COMMENT_PROBE, ASSY_PROCESS_ID, DARK_BOND_PAD_ASSY, ASSY_SERIAL_TYPE, STICKY_BACKS_TO_SAVE, OPTICAL_QUALITY, ENCODED_MARK_ID, " & _
       " PLANNED_LASER_SCRIBE, PACKAGE_LID_TYPE, PACKAGE_TYPE, PB_FREE_PACKAGE, TARGET_WAF_THICKNESS, RELIABILITY_SAMPLING, LOT_PRIORITY, WAFER_BOX_TYPE, TEST_SITE,ASSEMBLY_FACILITY, " & _
       " BATCH_COMMENT_ASSY, TST_PROCESS_ID,ELEC_SPECIAL_TEST, BOX_TYPE, PROTECTIVE_FILM_APLD, SHIPPING_MST_260,SHIPPING_MST_LEVEL, T_PRICE, SHIP_COMMENT, BATCH_COMMENT_TEST, " & _
       " CREATED_DATE, CREATED_TIME, UNIT_PRICE,REF_PO, REF_PO_ITEM, COUNTRY_OF_ASSEMBLY, MICRON_MATERIAL,DATE_CODE, SHIP_SITE, SPECIAL_PROCESS_LOT, " & _
       " LOT_STATUS, CUSTOM_PART_NO, wafer_visual_inspect, comp_code,FLAG,QTECH_CREATED_BY,QTECH_CREATED_DATE, QTECH_LASTUPDATE_BY, QTECH_LASTUPDATE_DATE from CustomerOItbl_test  where (customershortname = 'AA' or customershortname is null)   order by id desc ")
    
End Sub

Private Sub Command6_Click()

    'GC
    On Error Resume Next

    Dim FName

    '帅选文件
    CommonDialog2.Filter = "CSV文件(*.csv)|*.csv|EXCEL文件(*.xlsx)|*.xlsx|EXCEL文件(*.xls)|*.xls"
    
    CommonDialog2.ShowOpen
    '得到文件名
    FName = CommonDialog2.filename

    If FName <> "" Then
        Text3.text = FName

    End If

End Sub

Private Sub UploadGC()
   Exit Sub

    '读取CSV
    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    customerTemp = "GC"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    Dim dirName  As String

    Dim filename As String

    '获取文件名
    If InStrRev(Trim(Text3.text), "\") > 0 Then
        strFileName = Mid(Trim(Text3.text), InStrRev(Trim(Text3.text), "\") + 1)
        dirName = Mid$(Trim(Text3.text), 1, InStrRev(Trim(Text3.text), "\"))

    End If

    Dim con As New ADODB.Connection

    Dim rs  As New ADODB.Recordset

    con.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
    rs.Open "Select * From " & "[" & strFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
    Dim i            As Integer

    Dim j            As Integer

    Dim id           As Long

    Dim TEMP         As String

    Dim SumCount     As Integer

    Dim GCHeaderFlag As Boolean

    Dim str01        As String

    Dim str03        As String

    SumCount = 0
    rs.MoveFirst
        
    GCHeaderFlag = False
        
    For i = 0 To rs.RecordCount - 1
        TEMP = ""
        id = 0
        
        '付值
        gcHeaderTemp.Created_By = gUserName
        gcDetailTemp.ITEM = IIf(IsNull(rs.Fields(0).Value), "", rs.Fields(0).Value)
        gcHeaderTemp.po_no = IIf(IsNull(rs.Fields(1).Value), "", rs.Fields(1).Value)
        gcHeaderTemp.SUPPLIER = rs.Fields(2).Value
        gcHeaderTemp.ShipTo = rs.Fields(3).Value
        gcHeaderTemp.Fab_Device = rs.Fields(4).Value
        gcHeaderTemp.Customer_Device = rs.Fields(5).Value
        gcHeaderTemp.GC_Version = rs.Fields(6).Value
        gcDetailTemp.Marking_Lot_ID = IIf(IsNull(rs.Fields(7).Value), "", rs.Fields(7).Value)
   
        str01 = rs.Fields(8).Value
            
        If InStr(str01, "月") > 0 Then
            
            str03 = Replace(str01, "月", "-")
            str03 = Replace(str03, "日", "")
            str03 = Year(DATE) & "-" & str03
            gcHeaderTemp.GC_Date = str03
            
        Else
            
            gcHeaderTemp.GC_Date = rs.Fields(8).Value
            
        End If
        If InStr(rs.Fields(9).Value, ".") > 0 Then
            gcHeaderTemp.Lot_id = Left(rs.Fields(9).Value, InStr(rs.Fields(9).Value, ".") - 1)
        Else
            gcHeaderTemp.Lot_id = rs.Fields(9).Value
        End If
        gcDetailTemp.Lot_id = gcHeaderTemp.Lot_id
        'gcHeaderTemp.Lot_id = rs.Fields(9).Value
        'gcDetailTemp.Lot_id = rs.Fields(9).Value
        
        gcDetailTemp.wafer_id = rs.Fields(10).Value
        gcDetailTemp.Good_Die_Qty = CInt(rs.Fields(11).Value)
        gcHeaderTemp.WO_NO = rs.Fields(12).Value
        gcHeaderTemp.Ship_Out = IIf(IsNull(rs.Fields(14).Value), "", rs.Fields(14).Value)
            
        '2015-02-03 jiayunadd check shipOut
        '如果是COG的，则不可以为空
            
        If Left(gcHeaderTemp.Lot_id, 3) = "GXS" Then
            If gcHeaderTemp.Ship_Out = "" Then
                MsgBox "GC COG，最后一列发货地址不可以有空！"
                Exit Sub
                
            End If
                
        End If
            
        '2013-12-05 jiayun add
        '判断wo是否为空
            
        If Trim(gcHeaderTemp.WO_NO) = "" Then
            
            MsgBox "WO_NO有空值，请确认！"
            Exit Sub

        End If
            
        '2012-11-07 jiayun 修改Good_Die_Qty 根据市场部规则
  
        gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
            
        '2013-12-27 jiayun add
            
        If gcDetailTemp.Good_Die_Qty <= 0 Then
            MsgBox "请确认客户机种对应的Die数是否有维护好！"
            Exit Sub

        End If
            
        '2012-11-05 jiayun 修改 GC
            
        '判断lotID在Header表中是否已存在
            
        If (JudgeGCHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            '2013-12-05 jiayun add 如果lotid,wo_no 已存在，则查询出id
            '当lotid有隔行时，则查询上次的id
                
            id = GetGCLotIDWOId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
            '2013-01-11 jiayun add 客户简称
                
            If id = 0 Then
                MsgBox "DB主表ID生成失败1，请联系资讯！"
                Exit Sub
                
            Else
                
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
                
            End If
              
        End If
            
        '判断lotID在Detail表中是否已存在
            
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "GC 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
               
        Else
            '上传到Detail表中
            
            '2012-11-05 jiayun 修改 GCT
                   
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
                   
            If id = 0 Then
                MsgBox "DB主表ID生成失败2，请联系资讯！"
                Exit Sub
                
            Else
                Call AddGCDetail(gcDetailTemp, customerTemp, id)
                SumCount = SumCount + 1
                    
            End If
                
        End If
            
        rs.MoveNext
        
    Next i
        
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If

End Sub

Private Sub UploadGCBP()
'读取CSV
Dim source_batch_id_Temp As String
Dim customerTemp         As String
Dim bomRS2               As New ADODB.Recordset

customerTemp = "GC"
'上传OI的CSV
'处理文件名
If Text3.text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub

End If

Dim dirName  As String
Dim filename As String

'获取文件名
If InStrRev(Trim(Text3.text), "\") > 0 Then
    strFileName = Mid(Trim(Text3.text), InStrRev(Trim(Text3.text), "\") + 1)
    dirName = Mid$(Trim(Text3.text), 1, InStrRev(Trim(Text3.text), "\"))

End If

Dim con As New ADODB.Connection
Dim rs  As New ADODB.Recordset

con.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
rs.Open "Select * From " & "[" & strFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
Dim i            As Integer
Dim j            As Integer
Dim id           As Long
Dim TEMP         As String
Dim SumCount     As Integer
Dim GCHeaderFlag As Boolean
Dim str01        As String
Dim str03        As String
Dim gc9606wafer  As String

gc9606wafer = "123456789ABCDEFGHIJKLMNOP"
SumCount = 0
rs.MoveFirst
GCHeaderFlag = False
gUpID = Get_OracleStr("select PO_ITEM_SEQ.nextval from dual")
 
For i = 0 To rs.RecordCount - 1
    TEMP = ""
    id = 0
    '付值
    gcHeaderTemp.Created_By = gUserName
    gcDetailTemp.ITEM = IIf(IsNull(rs.Fields(0).Value), "", rs.Fields(0).Value)
    gcHeaderTemp.po_no = IIf(IsNull(rs.Fields(1).Value), "", rs.Fields(1).Value)
    
    If ComCbBand.text = "保税" And Left(gcHeaderTemp.po_no, 2) <> "HK" Then
        MsgBox "WO中PO以" & Left(gcHeaderTemp.po_no, 2) & "开头,不是保税机种,请确认", vbInformation, "提示"
        Exit Sub
    End If
    If ComCbBand.text = "非保税" And Left(gcHeaderTemp.po_no, 2) <> "SH" Then
        MsgBox "WO中PO以" & Left(gcHeaderTemp.po_no, 2) & "开头,不是非保税机种,请确认", vbInformation, "提示"
        Exit Sub
    End If
    
    gcHeaderTemp.SUPPLIER = rs.Fields(2).Value
    gcHeaderTemp.ShipTo = rs.Fields(3).Value
    gcHeaderTemp.Fab_Device = rs.Fields(4).Value
    gcHeaderTemp.Customer_Device = rs.Fields(5).Value
    gcHeaderTemp.GC_Version = rs.Fields(6).Value
    'gcDetailTemp.Marking_Lot_ID = IIf(IsNull(Rs.fields(7).Value), "", Rs.fields(7).Value)
    str01 = rs.Fields(7).Value
    If InStr(str01, "月") > 0 Then
        str03 = Replace(str01, "月", "-")
        str03 = Replace(str03, "日", "")
        str03 = Year(DATE) & "-" & str03
        gcHeaderTemp.GC_Date = str03
    Else
        gcHeaderTemp.GC_Date = rs.Fields(7).Value

    End If
    If InStr(rs.Fields(8).Value, ".") > 0 Then
        gcHeaderTemp.Lot_id = Left(rs.Fields(8).Value, InStr(rs.Fields(8).Value, ".") - 1)
    Else
        gcHeaderTemp.Lot_id = rs.Fields(8).Value
    End If
    gcDetailTemp.Lot_id = gcHeaderTemp.Lot_id
        
    'gcHeaderTemp.Lot_id = rs.Fields(8).Value
    'gcDetailTemp.Lot_id = rs.Fields(8).Value
    
    
    gcDetailTemp.wafer_id = rs.Fields(9).Value
    gcDetailTemp.Good_Die_Qty = CInt(rs.Fields(10).Value)
    gcHeaderTemp.WO_NO = rs.Fields(11).Value
    ' gcHeaderTemp.Ship_Out = IIf(IsNull(Rs.fields(14).Value), "", Rs.fields(14).Value)
    If Mid(gcHeaderTemp.Customer_Device, 1, InStr(gcHeaderTemp.Customer_Device, "-") - 1) = "GC9606" Then
        'gcHeaderTemp.Customer_Device = gcHeaderTemp.Customer_Device & "AA"
        gcHeaderTemp.GC_Version = gcHeaderTemp.GC_Version & "AA"

        '     twgcwafer = ""
        '     twgcwafer = Right(gcDetailTemp.Wafer_id, 2)
        '     If Mid(twgcwafer, 1, 1) = 0 Then
        '         twgcwafer = Mid(twgcwafer, 2, 1)
        '     End If
        '     gcwaferpart = Mid(gc9606wafer, twgcwafer, 1)
        '     gcDetailTemp.Marking_Lot_ID = "GC9606" & Mid(gcDetailTemp.Lot_ID, 1, 7) & gcwaferpart & Mid(Year(Now), 3, 2) & DatePart("ww", Now)
    End If

    '        If gcHeaderTemp.CUSTOMER_DEVICE = "GC802-4" Then
    '
    '            gcHeaderTemp.GC_Version = gcHeaderTemp.GC_Version & "AAAA"
    '
    '        End If
    '
    If Mid(gcHeaderTemp.Customer_Device, 1, InStr(gcHeaderTemp.Customer_Device, "-") - 1) = "GC9608" Then
        gcHeaderTemp.GC_Version = gcHeaderTemp.GC_Version & "AA"

    End If

    '2015-02-03 jiayunadd check shipOut
    '如果是COG的，则不可以为空
    If Left(gcHeaderTemp.Lot_id, 3) = "GXS" Then
        If gcHeaderTemp.Ship_Out = "" Then
            MsgBox "GC COG，最后一列发货地址不可以有空！"
            Exit Sub

        End If

    End If

    '2013-12-05 jiayun add
    '判断wo是否为空
    If Trim(gcHeaderTemp.WO_NO) = "" Then
        MsgBox "WO_NO有空值，请确认！"
        Exit Sub

    End If
    
    If gcHeaderTemp.Customer_Device = "" Then
        MsgBox "GC WO中，Customer_Device不可为空，请确认Wo!"
        Exit Sub
    End If

    If gcHeaderTemp.Lot_id = "" Then
        MsgBox "GC WO中，Fab Lot Id不可为空，请确认Wo!"
        Exit Sub
    End If
    
    If gcHeaderTemp.Fab_Device = "" Then
        MsgBox "GC WO中，Fab Device不可为空，请确认Wo!"
        Exit Sub
    End If
    
    If gcHeaderTemp.po_no = "" Then
        MsgBox "GC WO中，po no不可为空，请确认Wo!"
        Exit Sub
    End If
    
    If gcDetailTemp.wafer_id = "" Then
        MsgBox "GC WO中，Wafer id不可为空，请确认Wo!"
        Exit Sub
    End If
        
    

    '2012-11-07 jiayun 修改Good_Die_Qty 根据市场部规则
    ' gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
    '2013-12-27 jiayun add
    If gcDetailTemp.Good_Die_Qty <= 0 Then
        MsgBox "请确认客户机种对应的Die数是否有维护好！"
        Exit Sub

    End If

    '2012-11-05 jiayun 修改 GC
    '判断lotID在Header表中是否已存在
    If (JudgeGCHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)) Then
        If GCHeaderFlag = False Then

            '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
        End If

        '2013-12-05 jiayun add 如果lotid,wo_no 已存在，则查询出id
        '当lotid有隔行时，则查询上次的id
        id = GetGCLotIDWOId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)
    Else
        '上传到Header表中
        '取目前DB最大的ID号
        id = GetMaxID()
        '2013-01-11 jiayun add 客户简称
        If id = 0 Then
            MsgBox "DB主表ID生成失败1，请联系资讯！"
            Exit Sub
        Else
            'Call AddGCHeader(gcHeaderTemp, ID, customerTemp)
            Call AddGCHeader_new(gcHeaderTemp, id, customerTemp, gUpID)
            GCHeaderFlag = True

        End If

    End If

    '判断lotID在Detail表中是否已存在
    'Set bomRS2 = JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id)
    'MsgBox (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_id))
    Set bomRS2 = GetIdCount(gcDetailTemp.wafer_id, gcDetailTemp.Lot_id)
    'MsgBox (bomRS2.RecordCount)
    If (bomRS2.RecordCount > 1) Then
        MsgBox "GC 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在二次上传版本，无需上传!"
    Else
        '上传到Detail表中
        '2012-11-05 jiayun 修改 GCT
        ' 2017/08/03 tw "waferId & '+' "
        gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
        If Mid(gcHeaderTemp.Customer_Device, 1, InStr(gcHeaderTemp.Customer_Device, "-") - 1) = "GC9606" Then
            '                    ' 2017/08/03 TW 2次markingId沿用前回
            '                    Set bomRS2 = GetLastMI(gcDetailTemp.wafer_id, gcDetailTemp.lot_id)
            '
            '                    gcDetailTemp.Marking_Lot_ID = IIf(IsNull(bomRS2.fields("productid").Value), "", bomRS2.fields("productid").Value)
            'gcDetailTemp.Marking_Lot_ID = "GC9606" & Mid(gcDetailTemp.ITEM, 1, 7) & Right(("0" & gcDetailTemp.Wafer_id), 2) & Mid(Year(Now), 3, 2) & DatePart("ww", Now)
            
            gcDetailTemp.Marking_Lot_ID = "GC9606" & "\\" & Mid(gcDetailTemp.ITEM, 1, 7) & Right(("0" & gcDetailTemp.wafer_id), 2) & "\\" & Mid(Year(Now), 3, 2) & DatePart("ww", Now)

        End If

        If Left$(gcDetailTemp.wafer_id, 1) = "0" Then
            gcDetailTemp.ITEM = Get_OracleStr("select max(substrateid) || '+' as substrateid from mappingdatatest where wafer_id in ('" & Right(gcDetailTemp.wafer_id, 1) & "', '0' || '" & Right(gcDetailTemp.wafer_id, 1) & "') and lotid = '" & gcDetailTemp.Lot_id & "'")
        Else
            gcDetailTemp.ITEM = Get_OracleStr("select max(substrateid) || '+' as substrateid from mappingdatatest where wafer_id = '" & gcDetailTemp.wafer_id & "' and lotid = '" & gcDetailTemp.Lot_id & "'")

        End If

        If id = 0 Then
            MsgBox "DB主表ID生成失败2，请联系资讯！"
            Exit Sub
        Else
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1

        End If

    End If

    rs.MoveNext
Next i

If SumCount > 0 Then
    MsgBox "已成功上传" & SumCount & "笔！"
    If ExportExcel(gcHeaderTemp, customerTemp) = True Then
        Call SentMesToPMC(gcHeaderTemp)
    Else
        MsgBox "邮件未正常发送", vbInformation, "提示"
        ' StrFileName = Trim(Text3.Text)
        ' Call SentMesToPMC(gcHeaderTemp)
    End If
End If

End Sub

Private Function GetGCWLT(txtTemp As String) As String
    GetGCWLT = "F"
        
    Dim CusDevice As String

    Dim GCVersion As String
        
    Dim FName     As String

    Dim Nextline  As String

    FName = Trim(Text3.text)
    Open FName For Input As #2
        
    Do Until EOF(2)
        Line Input #2, Nextline
        
        If UCase(Left(Trim(Nextline), 4)) <> "ITEM" And Replace(Replace(Nextline, " ", ""), ",", "") <> "" Then
             
            Dim bid

            bid = Split(Nextline, ",")
                
            CusDevice = bid(5)
            GCVersion = bid(6)
                
            '判断是不是WLT
                
            If CusDevice = "GC0312-3" And Right(GCVersion, 1) = "C" Then
                GetGCWLT = "T"
            
            Else
                GetGCWLT = "F"

            End If

            Close #2
            Exit Function

        End If
        
    Loop
    Close #2
        
End Function

Private Sub UploadGCNew()
Dim source_batch_id_Temp As String
Dim customerTemp         As String
Dim cusPTTemp            As String
Dim gcVerTemp            As String
Dim gcVerLastTemp        As String

customerTemp = "GC"
If Text3.text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub

End If

If GetGCWLT(Trim(Text3.text)) = "T" Then
    UploadGCWLTNew
    Exit Sub

End If

Dim i            As Integer
Dim j            As Integer
Dim id           As Long
Dim TEMP         As String
Dim SumCount     As Integer
Dim GCHeaderFlag As Boolean
Dim str01        As String
Dim str03        As String
Dim FabTem       As String
Dim gc9606wafer  As String
Dim gcwaferpart  As String
Dim pjgcwafer    As String

gc9606wafer = "123456789ABCDEFGHIJKLMNOP"
SumCount = 0
GCHeaderFlag = False
Dim k        As Integer
Dim FName    As String
Dim Nextline As String

gUpID = Get_OracleStr("select PO_ITEM_SEQ.nextval from dual")


FName = Trim(Text3.text)
Open FName For Input As #1

Do Until EOF(1)
    Line Input #1, Nextline
    cusPTTemp = ""
    gcVerTemp = ""
    gcVerLastTemp = ""
    If UCase(Left(Trim(Nextline), 4)) <> "ITEM" And Replace(Replace(Nextline, " ", ""), ",", "") <> "" Then
        Dim bid

        bid = Split(Nextline, ",")
        id = 0
        '付值
        gcHeaderTemp.Created_By = gUserName
        gcDetailTemp.ITEM = Trim(bid(0))
        gcHeaderTemp.po_no = Trim(bid(1))
        If ComCbBand.text = "保税" And Left(Trim(bid(1)), 2) <> "HK" Then
            MsgBox "WO中PO以" & Left(Trim(bid(1)), 2) & "开头,不是保税机种,请确认", vbInformation, "提示"
            Exit Sub
        End If
        If ComCbBand.text = "非保税" And Left(Trim(bid(1)), 2) <> "SH" Then
            MsgBox "WO中PO以" & Left(Trim(bid(1)), 2) & "开头,不是非保税机种,请确认", vbInformation, "提示"
            Exit Sub
        End If
        gcHeaderTemp.SUPPLIER = Trim(bid(2))
        gcHeaderTemp.ShipTo = Trim(bid(3))
        gcHeaderTemp.Fab_Device = Trim(bid(4))
        gcHeaderTemp.Customer_Device = Trim(bid(5))
        cusPTTemp = Trim(gcHeaderTemp.Customer_Device)
        gcHeaderTemp.GC_Version = Trim(bid(6))
        gcVerTemp = Trim(UCase(gcHeaderTemp.GC_Version))
        ' 20180115  tangwei
        gcHeaderTemp.SecondFlag = Trim(UCase(gcHeaderTemp.GC_Version))
        If InStr(Trim(bid(9)), ".") > 0 Then
            gcHeaderTemp.Lot_id = Left(Trim(bid(9)), InStr(Trim(bid(9)), ".") - 1)
        Else
            gcHeaderTemp.Lot_id = Trim(bid(9))
        End If
        gcDetailTemp.Lot_id = gcHeaderTemp.Lot_id
        gcDetailTemp.TAX_TYPE = gTax
        gcHeaderTemp.LotProperty = Trim(bid(14))
        gcHeaderTemp.LotOwner = Trim(bid(15))
        gcHeaderTemp.Telephone = Trim(bid(16))
        gcHeaderTemp.Flow = Trim(bid(18))
        gcHeaderTemp.htdevice = ""
        If gcHeaderTemp.po_no = "" Then
            MsgBox "GC WO中，po no不可为空，请确认Wo!"
            Exit Sub
        End If
        If gcHeaderTemp.SUPPLIER = "" Then
            MsgBox "GC WO中，SUPPLIER不可为空，请确认Wo!"
            Exit Sub
        End If
        If gcHeaderTemp.ShipTo = "" Then
            MsgBox "GC WO中，ShipTo不可为空，请确认Wo!"
            Exit Sub
        End If
        If gcHeaderTemp.Fab_Device = "" Then
            MsgBox "GC WO中，Fab Device不可为空，请确认Wo!"
            Exit Sub
        End If
        If cusPTTemp = "" Then
            MsgBox "GC WO中，Customer_Device不可为空，请确认Wo!"
            Exit Sub
        End If
        If gcHeaderTemp.GC_Version = "" Then
            MsgBox "GC WO中，GCVersion不可为空，请确认Wo!"
            Exit Sub
        End If
        
        If gcHeaderTemp.Lot_id = "" Then
            MsgBox "GC WO中，Fab Lot Id不可为空，请确认Wo!"
            Exit Sub
        End If
  
        '2015-04-27 jiayun add 第三位系统自动带
        '2015-11-17 jiayun add  GC2145
        If cusPTTemp = "GC2145-3" Then
            If Left(bid(9), 1) = "H" Then
                gcVerLastTemp = "G"
            ElseIf Left(bid(9), 1) = "S" Then
                gcVerLastTemp = "G"
            ElseIf Left(bid(9), 1) = "E" Then
                gcVerLastTemp = "F"

            End If

        ElseIf cusPTTemp = "GC5005-3" Then
            If Mid(gcVerTemp, 2, 1) = "A" Then
                gcVerLastTemp = "D"
            ElseIf Mid(gcVerTemp, 2, 1) = "D" Then
                gcVerLastTemp = "A"

            End If

        ElseIf cusPTTemp = "GC2375-3" Then
            If Mid(gcHeaderTemp.Lot_id, 1, 1) = "E" Then
                gcVerLastTemp = "A"
            Else
                gcVerLastTemp = "F"

            End If

            ' ElseIf cusPTTemp = "GC9606-2.5" Then
            '  gcVerLastTemp = "AIAA"
            
        ElseIf cusPTTemp = "GC2375H-3" Then
            If Trim(bid(11)) = "5877" Then '12寸
                gcVerLastTemp = "B"
            Else
                gcVerLastTemp = "D" '8寸
            End If
        Else
            gcVerLastTemp = GetGCVerLastChar(cusPTTemp)

        End If

        If gcVerLastTemp <> "" Then
            gcHeaderTemp.GC_Version = gcVerTemp & gcVerLastTemp
            If cusPTTemp = "GC032A-3" Then
                gcHeaderTemp.GC_Version = gcVerTemp & "L"

            End If

            '2015-08-20 jiayun add 处理 GC0409-3
            FabTem = Left(UCase(Trim(gcHeaderTemp.Fab_Device)), 5)
            If FabTem = "P6418" Then
                gcHeaderTemp.GC_Version = gcVerTemp & "A"
            ElseIf FabTem = "P6820" Then
                gcHeaderTemp.GC_Version = gcVerTemp & "B"
            ElseIf FabTem = "P7238" Then
                gcHeaderTemp.GC_Version = gcVerTemp & "E"

            End If

        Else
            If cusPTTemp = "GC1004-3" Then
                If Mid(gcVerTemp, 1, 1) = "A" Or Mid(gcVerTemp, 1, 1) = "B" Or Mid(gcVerTemp, 1, 1) = "C" Or Mid(gcVerTemp, 1, 1) = "D" Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "A"
                Else
                    gcHeaderTemp.GC_Version = gcVerTemp & "B"

                End If

            ElseIf cusPTTemp = "GC0329-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "D"
                ElseIf Len(gcVerTemp) = 3 Then
                    gcHeaderTemp.GC_Version = gcVerTemp
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

                'CCS ADD 20160707
            ElseIf cusPTTemp = "GC5025-3" Then
                If Len(gcVerTemp) = 2 Then
                    ' 20180131 修改D-H
                    gcHeaderTemp.GC_Version = gcVerTemp & "H"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

                'pj add 20161014
            ElseIf cusPTTemp = "CC2601-3" Then
                If Len(gcVerTemp) = 3 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "A"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "CC2601-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "A"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC4623-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "C"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC02M0-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "A"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC1054-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "B"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC2053-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "B"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC1603-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "B"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC4633-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "B"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

                'CCS ADD 20160902
            ElseIf cusPTTemp = "GC033A-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "B"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC030A-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "N"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC6153-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "B"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC2063-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "B"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC2905-3" Or cusPTTemp = "GC8034-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "B"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

                'CCS ADD 20160930
            ElseIf cusPTTemp = "GC23A5-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "A"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC5035-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "D"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC032A-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "L"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC02M0B-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "A"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC9606-2.5" Then
                If Len(gcVerTemp) = 1 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "A"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC9606-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "A"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC9606-4" Then
                If Len(gcVerTemp) = 1 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "AIAA"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC802-2.5" Then
                If Len(gcVerTemp) = 1 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "AAA"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

                ' add by tony 20171208
            ElseIf cusPTTemp = "GC9608-2.5" Then
                If Len(gcVerTemp) = 1 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "A"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC2385-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "B"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC1066-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "B"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC2375H-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "B"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC1009-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "C"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            ElseIf cusPTTemp = "GC1034-3" Then
                If Len(gcVerTemp) = 2 Then
                    gcHeaderTemp.GC_Version = gcVerTemp & "B"
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            Else
                '判断长度是否为3 ，如果是，则按市场部的来上传，否则提提示错误
                If Len(gcVerTemp) = 3 Then
                    gcHeaderTemp.GC_Version = gcVerTemp
                Else
                    MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
                    Exit Sub

                End If

            End If

        End If

        gcDetailTemp.Marking_Lot_ID = Trim(bid(7))
        str01 = Trim(bid(8))
        If InStr(str01, "月") > 0 Then
            str03 = Replace(str01, "月", "-")
            str03 = Replace(str03, "日", "")
            str03 = Year(DATE) & "-" & str03
            gcHeaderTemp.GC_Date = str03
        Else
            gcHeaderTemp.GC_Date = bid(8)

        End If

        gcDetailTemp.wafer_id = Trim(bid(10))
        gcDetailTemp.Good_Die_Qty = Trim(bid(11))
        gcHeaderTemp.WO_NO = Trim(bid(13))
        gcHeaderTemp.Ship_Out = Trim(bid(17))
                
        If gcDetailTemp.Marking_Lot_ID = "" Then
            MsgBox "GC WO中，Marking_Lot_ID不可为空，请确认Wo!"
            Exit Sub
        End If
        If Len(gcDetailTemp.Marking_Lot_ID) <> 5 Then
            MsgBox "GC WO中，Marking_Lot_ID必须为5码，请确认Wo!"
            Exit Sub
        End If
        If gcHeaderTemp.GC_Date = "" Then
            MsgBox "GC WO中，Date不可为空，请确认Wo!"
            Exit Sub
        End If
        If gcDetailTemp.wafer_id = "" Then
            MsgBox "GC WO中，Wafer_id不可为空，请确认Wo!"
            Exit Sub
        End If
     '   MsgBox gcDetailTemp.Good_Die_Qty
'        If gcDetailTemp.Good_Die_Qty = "" Then
         '   MsgBox "GC WO中，Good_Die_Qty不可为空，请确认Wo!"
         '   Exit Sub
     '   End If
        If gcHeaderTemp.WO_NO = "" Then
            MsgBox "GC WO中，WO_NO 不可为空，请确认Wo!"
            Exit Sub
        End If
        
        'ccs add cusPTTemp <> "GC2145-3" Then
        If cusPTTemp <> "GC2145-3" Then
            If cusPTTemp <> "GC5025-3" Then
                If cusPTTemp <> "GC1034-3" Then
                    If cusPTTemp <> "GC2355C-3" Then
                        If cusPTTemp <> "GC1064R-3" Then
                            If cusPTTemp <> "GC0409-3" Then
                                If cusPTTemp <> "GC033A-3" Then
                                    If cusPTTemp <> "GC9606-3" Then
                                        If cusPTTemp <> "GC23A5-3" Then
                                            If cusPTTemp <> "CC2601-3" Then
                                                If cusPTTemp <> "GC032A-3" Then
                                                    If cusPTTemp <> "GC030A-3" Then
                                                        If cusPTTemp <> "GC1066-3" Then
                                                            If cusPTTemp <> "GC9606-3" Then
                                                                If cusPTTemp <> "GC2385-3" Then
                                                                    If cusPTTemp <> "GC2365-3" Then
                                                                        If cusPTTemp <> "GC2375-3" Then
                                                                            If cusPTTemp <> "GC1024-3" Then
                                                                                If cusPTTemp <> "GC5005-3" Then
                                                                                    If cusPTTemp <> "GC2375A-3" Then
                                                                                        If cusPTTemp <> "GC8024-3" Then
                                                                                            If cusPTTemp <> "GC9606-4" Then
                                                                                                If cusPTTemp <> "GC9606-2.5" Then
                                                                                                    If cusPTTemp <> "GC2033-3" Then
                                                                                                        If cusPTTemp <> "GC1066C-3" Then
                                                                                                            If cusPTTemp <> "GC2023A-3" Then
                                                                                                                If cusPTTemp <> "GC9608-2.5" Then
                                                                                                                    If cusPTTemp <> "GC2375H-3" Then
                                                                                                                        If cusPTTemp <> "GC8034-3" Then
                                                                                                                            If cusPTTemp <> "GC2905-3" Then
                                                                                                                                If cusPTTemp <> "GC02M0-3" Then
                                                                                                                                    If cusPTTemp <> "GC1009-3" Then
                                                                                                                                        gcHeaderTemp.TradeType = Trim(bid(15))

                                                                                                                                    End If

                                                                                                                                End If

                                                                                                                            End If

                                                                                                                        End If

                                                                                                                    End If

                                                                                                                End If

                                                                                                            End If

                                                                                                        End If

                                                                                                    End If

                                                                                                End If

                                                                                            End If

                                                                                        End If

                                                                                    End If

                                                                                End If

                                                                            End If

                                                                        End If

                                                                    End If

                                                                End If

                                                            End If

                                                        End If

                                                    End If

                                                End If

                                            End If

                                        End If

                                    End If

                                End If

                            End If

                        End If

                    End If

                End If

            End If

        End If

        If InStr(cusPTTemp, "-") <> 0 Then
            If Left(cusPTTemp, InStr(1, cusPTTemp, "-") - 1) = "GC9102" Then
                If gcHeaderTemp.Ship_Out = "" Then
                    MsgBox "GC COG，最后一列发货地址不可以有空！"
                    Exit Sub

                End If

            End If

        End If

        If Trim(gcHeaderTemp.WO_NO) = "" Then
            MsgBox "WO_NO有空值，请确认！"
            Exit Sub

        End If

        '2012-11-07 jiayun 修改Good_Die_Qty 根据市场部规则
        If Trim(gcHeaderTemp.Customer_Device) <> "GC2385-3" Then
            gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
            '2015-08-20 jiayun add 处理 GC0409-3
            If Trim(gcHeaderTemp.Customer_Device) = "GC5025-3" Then
                If chkGC_GC50253.Value = 1 Then
                    gcDetailTemp.Good_Die_Qty = 2212

                End If

            End If

            If Trim(gcHeaderTemp.Customer_Device) = "GC0409-3" Then
                FabTem = Left(UCase(Trim(gcHeaderTemp.Fab_Device)), 5)
                If FabTem = "P6418" Then
                    gcDetailTemp.Good_Die_Qty = 5192
                ElseIf FabTem = "P6820" Then
                    gcDetailTemp.Good_Die_Qty = 11994
                ElseIf FabTem = "P7238" Then
                    gcDetailTemp.Good_Die_Qty = 5191 '5211

                End If

            ElseIf Trim(gcHeaderTemp.Customer_Device) = "GC2145-3" Then
                'jiayun modify 2016-05-18
                If Left(gcHeaderTemp.Lot_id, 1) = "H" Then
                    'gcDetailTemp.Good_Die_Qty = 1676
                    gcDetailTemp.Good_Die_Qty = 1684
                ElseIf Left(gcHeaderTemp.Lot_id, 1) = "E" Then
                    gcDetailTemp.Good_Die_Qty = 3920

                End If

            ElseIf cusPTTemp = "GC2375-3" Then
                If Mid(gcHeaderTemp.Lot_id, 1, 1) = "E" Then
                    gcDetailTemp.Good_Die_Qty = 5877
                Else
                    gcDetailTemp.Good_Die_Qty = 2547

                End If

            ElseIf Trim(gcHeaderTemp.Customer_Device) = "CC2601-3" Then
                gcDetailTemp.Good_Die_Qty = 2341
            ElseIf Trim(gcHeaderTemp.Customer_Device) = "GC9606-4" Then
                gcDetailTemp.Good_Die_Qty = 2648
 
            ElseIf Trim(gcHeaderTemp.Customer_Device) = "GC2375H-3" Then
            'merry 20191125   2375H对应8寸和12寸两种，共用同一种客户机种，短期对策按实际填写的值上传
                gcDetailTemp.Good_Die_Qty = Trim(bid(11))
            End If

        End If

        '2013-12-27 jiayun add
        ' 2018-6-6
        If cusPTTemp = "GC2905-3" Then
            gcDetailTemp.Good_Die_Qty = 2786

        End If

        If cusPTTemp = "GC8034-3" Then
            gcDetailTemp.Good_Die_Qty = 1385

        End If

        If cusPTTemp = "GC1009-3" Then
            gcDetailTemp.Good_Die_Qty = 10350

        End If

        If gcDetailTemp.Good_Die_Qty <= 0 Then
            MsgBox "请确认客户机种对应的Die数是否有维护好！"
            Exit Sub

        End If

        Set oiRS = GetGCPT_C(cusPTTemp)
        If (oiRS.RecordCount > 0) Then
            gcHeaderTemp.Customer_Device = oiRS.Fields("CUSTOMERPTNew").Value

        End If

        '2012-11-05 jiayun 修改 GC
        '判断lotID在Header表中是否已存在
        If (JudgeGCHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)) Then
            id = GetGCLotIDWOId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
            '2013-01-11 jiayun add 客户简称
            If id = 0 Then
                MsgBox "DB主表ID生成失败1，请联系资讯！"
                Exit Sub
            Else
               ' Call AddGCHeader(gcHeaderTemp, ID, customerTemp)
                Call AddGCHeader_new(gcHeaderTemp, id, customerTemp, gUpID)
                GCHeaderFlag = True

            End If

        End If

        '判断lotID在Detail表中是否已存在
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id) And gUserName <> "07885") Then
            MsgBox "GC 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
        Else
            '上传到Detail表中
            '2012-11-05 jiayun 修改 GCT
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
            If InStr(cusPTTemp, "-") <> 0 Then
                If Mid(cusPTTemp, 1, InStr(cusPTTemp, "-") - 1) = "GC9606" Then
                    pjgcwafer = ""
                    pjgcwafer = Right(gcDetailTemp.ITEM, 2)
                    If Mid(pjgcwafer, 1, 1) = 0 Then
                        pjgcwafer = Mid(pjgcwafer, 2, 1)

                    End If

                    gcwaferpart = Mid(gc9606wafer, pjgcwafer, 1)
                    ' gcDetailTemp.Marking_Lot_ID = "GC9606" & Mid(gcDetailTemp.ITEM, 1, 7) & gcwaferpart & Mid(Year(Now), 3, 2) & DatePart("ww", Now)
                    ' 2017-11-31 tangwei 6+7+1+2+2  改成6+7+2+2+2
                    gcDetailTemp.Marking_Lot_ID = "GC9606" & "\\" & Mid(gcDetailTemp.ITEM, 1, 7) & Right(("0" & gcDetailTemp.wafer_id), 2) & "\\" & Mid(Year(Now), 3, 2) & DatePart("ww", Now)

                End If

            End If

            If id = 0 Then
                MsgBox "DB主表ID生成失败2，请联系资讯！"
                Exit Sub
            Else
                Call AddGCDetail(gcDetailTemp, customerTemp, id)
                SumCount = SumCount + 1

            End If

        End If

    End If

Loop
Close #1
If SumCount > 0 Then
    MsgBox "已成功上传" & SumCount & "笔！"
    If ExportExcel(gcHeaderTemp, customerTemp) = True Then
        Call SentMesToPMC(gcHeaderTemp)
    Else
        ' StrFileName = Trim(Text3.Text)
        ' Call SentMesToPMC(gcHeaderTemp)
        MsgBox "邮件未正常发送", vbInformation, "提示"
    End If
End If

End Sub

Private Sub UploadGCNewWLDT()
    '2015-04-28 jiayun add WLDT

    '读取CSV
    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim cusPTTemp            As String

    Dim gcVerTemp            As String

    Dim gcVerLastTemp        As String

    Dim waferIdTemp          As String

    Dim wo_HT_Temp           As String

    wo_HT_Temp = "WONO_" & Replace(Replace(Replace(Format(Now, "YYYY-MM-DD HH:MM:SS"), "-", ""), ":", ""), " ", "")

    customerTemp = "GC"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    Dim i            As Integer

    Dim j            As Integer

    Dim id           As Long

    Dim TEMP         As String

    Dim SumCount     As Integer

    Dim GCHeaderFlag As Boolean

    Dim str01        As String

    Dim str03        As String

    SumCount = 0
        
    GCHeaderFlag = False

    Dim k        As Integer
        
    Dim FName    As String

    Dim Nextline As String
    
    gUpID = Get_OracleStr("select PO_ITEM_SEQ.nextval from dual")
    FName = Trim(Text3.text)
    Open FName For Input As #1
        
    Do Until EOF(1)
        Line Input #1, Nextline
        cusPTTemp = ""
        gcVerTemp = ""
        gcVerLastTemp = ""
        waferIdTemp = ""
              
        If UCase(Left(Trim(Nextline), 2)) <> "NO" And Replace(Replace(Nextline, " ", ""), ",", "") <> "" Then

            Dim bid

            bid = Split(Nextline, ",")
             
            id = 0
        
            '付值
            gcHeaderTemp.Created_By = gUserName
            gcDetailTemp.ITEM = Trim(bid(0))
            gcHeaderTemp.po_no = Trim(bid(6))
            If ComCbBand.text = "保税" And Left(Trim(bid(6)), 2) <> "HK" Then
                MsgBox "WO中PO以" & Left(Trim(bid(6)), 2) & "开头,不是保税机种,请确认", vbInformation, "提示"
                Exit Sub
            End If
            If ComCbBand.text = "非保税" And Left(Trim(bid(6)), 2) <> "SH" Then
                MsgBox "WO中PO以" & Left(Trim(bid(6)), 2) & "开头,不是非保税机种,请确认", vbInformation, "提示"
                Exit Sub
            End If
        
            gcHeaderTemp.SUPPLIER = Trim(bid(1))
            gcHeaderTemp.ShipTo = Trim(bid(2))
            gcHeaderTemp.Fab_Device = Trim(bid(3))
            
            gcHeaderTemp.Customer_Device = Trim(bid(4)) & "-3"
            cusPTTemp = Trim(gcHeaderTemp.Customer_Device)
            gcHeaderTemp.GC_Version = Trim(bid(5))
            gcVerTemp = Trim(UCase(gcHeaderTemp.GC_Version))
            
            waferIdTemp = Trim(bid(10)) & Right("0" & Trim(bid(11)), 2)
            
            gcDetailTemp.Marking_Lot_ID = GetGCWLDMaringCode(waferIdTemp)
   
            str01 = Trim(bid(9))
            
            If InStr(str01, "月") > 0 Then
            
                str03 = Replace(str01, "月", "-")
                str03 = Replace(str03, "日", "")
                str03 = Year(DATE) & "-" & str03
                gcHeaderTemp.GC_Date = str03
            
            Else
            
                gcHeaderTemp.GC_Date = bid(9)
            
            End If
            If InStr(Trim(bid(10)), ".") > 0 Then
                gcHeaderTemp.Lot_id = Left(Trim(bid(10)), InStr(Trim(bid(10)), ".") - 1)
            Else
                gcHeaderTemp.Lot_id = Trim(bid(10))
            End If
            gcDetailTemp.Lot_id = gcHeaderTemp.Lot_id
            
            
            'gcHeaderTemp.Lot_id = Trim(bid(10))
            'gcDetailTemp.Lot_id = Trim(bid(10))
            
            gcDetailTemp.wafer_id = Trim(bid(11))
            gcDetailTemp.Good_Die_Qty = CInt(Trim(bid(12)))
            gcHeaderTemp.WO_NO = Trim(wo_HT_Temp)
            gcHeaderTemp.Ship_Out = Trim(bid(16))
            gcHeaderTemp.Flow = Trim(bid(17))
            If Trim(bid(18)) = "" Then
                MsgBox "厂内机种不能为空，请确认！"
                Exit Sub
            End If
            gcHeaderTemp.htdevice = Trim(bid(18))

            
            If Left(gcHeaderTemp.Lot_id, 3) = "GXS" Then
                If gcHeaderTemp.Ship_Out = "" Then
                    MsgBox "GC COG，最后一列发货地址不可以有空！"
                    Exit Sub
                
                End If
                
            End If
            

            If Trim(gcHeaderTemp.WO_NO) = "" Then
            
                MsgBox "WO_NO有空值，请确认！"
                Exit Sub

            End If
 
            If cusPTTemp = "" Then
                MsgBox "GC WO中，Customer_Device不可为空，请确认Wo!"
                Exit Sub
            End If
            
            If gcHeaderTemp.Fab_Device = "" Then
                MsgBox "GC WO中，Fab Device不可为空，请确认Wo!"
                Exit Sub
            End If
 
            If gcHeaderTemp.GC_Version = "" Then
                MsgBox "GC WO中，GC Version不可为空，请确认Wo!"
                Exit Sub
            End If

          '  If Get_SqlserverCnt("SELECT  IMAGER_CUSTOMER_REV  FROM erpbase..tblCustomerOI  Where SOURCE_BATCH_ID='" & gcHeaderTemp.Lot_id & "' and  Left(IMAGER_CUSTOMER_REV, 2)<>'" & Left(gcVerTemp, 2) & "'") > 0 Then
                
            If Get_SqlserverCnt("SELECT  a.IMAGER_CUSTOMER_REV  FROM erpbase..tblCustomerOI a ,ERPBASE..tblmappingData b  WHERE    a.SOURCE_BATCH_ID=b.LOTID AND   b.FILENAME=convert(VARCHAR(20),a.id) AND  convert(INT,b.WAFER_ID) =" & gcDetailTemp.wafer_id & "   AND a.SOURCE_BATCH_ID='" & gcHeaderTemp.Lot_id & "'  and  Left(IMAGER_CUSTOMER_REV, 2)<>'" & Left(gcVerTemp, 2) & "'") > 0 Then
                MsgBox gcHeaderTemp.Lot_id & " " & gcDetailTemp.wafer_id & "# 二级代码与WLA不一致", vbInformation, "提示"
                Exit Sub
            End If

    
            If gcHeaderTemp.po_no = "" Then
                MsgBox "GC WO中，po no不可为空，请确认Wo!"
                Exit Sub
            End If
            
            If gcHeaderTemp.GC_Date = "" Then
                MsgBox "GC WO中，Fab-Out Date不可为空，请确认Wo!"
                Exit Sub
            End If
                        
            If gcHeaderTemp.Lot_id = "" Then
                MsgBox "GC WO中，Fab Lot Id不可为空，请确认Wo!"
                Exit Sub
            End If
            
             If gcDetailTemp.wafer_id = "" Then
                MsgBox "GC WO中，Wafer id不可为空，请确认Wo!"
                Exit Sub
            End If
            'If gcDetailTemp.Good_Die_Qty = "" Then
            '    MsgBox "GC WO中，Gross Dies不可为空，请确认Wo!"
            '    Exit Sub
            'End If
  
            gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
            
            If gcDetailTemp.Good_Die_Qty <= 0 Then
                MsgBox "请确认客户机种对应的Die数是否有维护好！"
                Exit Sub

            End If
            

            If (JudgeGCHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)) Then
            
                If GCHeaderFlag = False Then


                End If
                
    
                
                id = GetGCLotIDWOId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)
                
            Else
      
                id = GetMaxID()
      
                If id = 0 Then
                    MsgBox "DB主表ID生成失败1，请联系资讯！"
                    Exit Sub
                
                Else
                    
                    'Call AddGCHeader(gcHeaderTemp, ID, customerTemp)
                    Call AddGCHeader_new(gcHeaderTemp, id, customerTemp, gUpID)
                    GCHeaderFlag = True
                
                End If
              
            End If
            
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & "+" & Right(("0" & gcDetailTemp.wafer_id), 2)
            
            If (JudgeGCDetailIdWLD(gcDetailTemp.Lot_id, gcDetailTemp.ITEM)) Then
                MsgBox "GC 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.ITEM & "已存在，无需上传!"
               
            Else
      
                   
                If id = 0 Then
                    MsgBox "DB主表ID生成失败2，请联系资讯！"
                    Exit Sub
                
                Else
                    Call AddGCDetail(gcDetailTemp, customerTemp, id)
                    SumCount = SumCount + 1
                    
                End If
                
            End If
 
        End If
        
    Loop
    Close #1
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"
        
        If ExportExcel(gcHeaderTemp, customerTemp) = True Then
            Call SentMesToPMC(gcHeaderTemp)
        Else
            ' StrFileName = Trim(Text3.Text)
            ' Call SentMesToPMC(gcHeaderTemp)
            MsgBox "邮件未正常发送", vbInformation, "提示"
        End If

    End If

End Sub

Private Sub AddGCDetail(mapTemp As GCDetail, customerTemp As String, oiKeyId As Long)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
'On Error GoTo DealError
         
Cnn.BeginTrans

cmdStr = "insert into mappingDataTest(id,substrateid,SUBSTRATETYPE,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename )" & _
         " values( mappingData_SEQ.Nextval,'" & mapTemp.ITEM & "','" & gTax & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & ",0,'" & customerTemp & "','Y','" & gUserName & "',sysdate," & oiKeyId & ")"
                                                               
                                                               
cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values('" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & ",0,'" & customerTemp & "','Y','" & gUserName & "',GETDATE()," & oiKeyId & ")"
                                                               
                                         
AddSql (cmdStr)
AddSql2 (cmdStr2)


If customerTemp = "MG" Then

cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingDataMG] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values('" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & ",0,'" & customerTemp & "','Y','" & gUserName & "',GETDATE()," & oiKeyId & ")"
         
 AddSql2 (cmdStr2)

End If

' 作数据备份
Dim cmdStr3 As String

cmdStr3 = " insert into wo_data(id, po_num,po_item,source_batch_id,source_mtrl_num,mtrl_num,mtrl_desc,test_mtrl_num,test_mtrl_desc,mpn,mpn_desc,source_mtrl_sloc, " & _
           " mtrl_num_mtrlgrp,probe_ship_part_type,offshore_asm_company,offshore_test_company,current_wafer_qty,die_qty,design_id,country_of_fab,fab_conv_id,fab_excr_id,reticle_level_71, " & _
           " reticle_level_72,reticle_level_73,wafer_size,imager_customer_rev,chromaticity,micro_lens_shift,temperature_spec,prb_containment_type,fabrication_facility,prb_excr_id,batch_comment_probe, " & _
           " assy_process_id,dark_bond_pad_assy,assy_serial_type,sticky_backs_to_save,optical_quality,encoded_mark_id,planned_laser_scribe,package_lid_type,package_type,pb_free_package,target_waf_thickness, " & _
           " reliability_sampling,lot_priority,wafer_box_type,test_site,assembly_facility,batch_comment_assy,tst_process_id,elec_special_test,box_type,protective_film_apld,shipping_mst_260,shipping_mst_level, " & _
           " t_price,ship_comment,batch_comment_test,created_date,created_time,unit_price,ref_po,ref_po_item,country_of_assembly,micron_material,date_code,ship_site,special_process_lot,lot_status,custom_part_no, " & _
           " flag,qtech_created_by,qtech_created_date,qtech_lastupdate_by,qtech_lastupdate_date,customershortname,downqty,invflag,wafer_visual_inspect,comp_code,eqdatacode,jobno,zx_fromsite,zx_invoice, SUBSTRATEID, SUBSTRATETYPE,PRODUCTID,MICRONLOTID,PASSBINCOUNT,FAILBINCOUNT,WAFER_ID,TIME_STATMP)   " & _
           " select   ct.id,ct.po_num,ct.po_item,ct.source_batch_id,ct.source_mtrl_num,ct.mtrl_num,ct.mtrl_desc,ct.test_mtrl_num,ct.test_mtrl_desc,ct.mpn,ct.mpn_desc,ct.source_mtrl_sloc,ct.mtrl_num_mtrlgrp, " & _
           " ct.probe_ship_part_type,ct.offshore_asm_company,ct.offshore_test_company,ct.current_wafer_qty,ct.die_qty,ct.design_id,ct.country_of_fab,ct.fab_conv_id,ct.fab_excr_id,ct.reticle_level_71,ct.reticle_level_72, " & _
           " ct.reticle_level_73,ct.wafer_size,ct.imager_customer_rev,ct.chromaticity,ct.micro_lens_shift,ct.temperature_spec,ct.prb_containment_type,ct.fabrication_facility,ct.prb_excr_id,ct.batch_comment_probe, " & _
           " ct.assy_process_id,ct.dark_bond_pad_assy,ct.assy_serial_type,ct.sticky_backs_to_save,ct.optical_quality,ct.encoded_mark_id,ct.planned_laser_scribe,ct.package_lid_type,ct.package_type,ct.pb_free_package, " & _
           " ct.target_waf_thickness,ct.reliability_sampling,ct.lot_priority,ct.wafer_box_type,ct.test_site,ct.assembly_facility,ct.batch_comment_assy,ct.tst_process_id,ct.elec_special_test,ct.box_type, " & _
           " ct.protective_film_apld,ct.shipping_mst_260,ct.shipping_mst_level,ct.t_price,ct.ship_comment,ct.batch_comment_test,ct.created_date,ct.created_time,ct.unit_price,ct.ref_po,ct.ref_po_item, " & _
           " ct.country_of_assembly,ct.micron_material,ct.date_code,ct.ship_site,ct.special_process_lot,ct.lot_status, " & _
           " ct.custom_part_no,ct.flag,ct.qtech_created_by,ct.qtech_created_date,ct.qtech_lastupdate_by,ct.qtech_lastupdate_date,ct.customershortname,ct.downqty,ct.invflag,ct.wafer_visual_inspect, " & _
           " ct.comp_code,ct.eqdatacode,ct.jobno,ct.zx_fromsite,ct.zx_invoice,mt.SUBSTRATEID, mt.SUBSTRATETYPE,mt.PRODUCTID,mt.MICRONLOTID,mt.PASSBINCOUNT,mt.FAILBINCOUNT,mt.WAFER_ID,sysdate from CustomerOItbl_test ct, MAPPINGDATATEST mt  where mt.substrateid =  '" & mapTemp.ITEM & "' and to_char(ct.id) = mt.filename"


AddSql (cmdStr3)
 
Cnn.CommitTrans

End Sub

Private Sub UploadGCWLTNew()

    '读取CSV
    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim wo_HT_Temp           As String

    wo_HT_Temp = "WONO_" & Replace(Replace(Replace(Format(Now, "YYYY-MM-DD HH:MM:SS"), "-", ""), ":", ""), " ", "")

    customerTemp = "GC"

    Dim i            As Integer

    Dim j            As Integer

    Dim id           As Long

    Dim TEMP         As String

    Dim SumCount     As Integer

    Dim GCHeaderFlag As Boolean

    Dim str01        As String

    Dim str03        As String

    SumCount = 0
        
    GCHeaderFlag = False

    Dim k        As Integer
        
    Dim FName    As String

    Dim Nextline As String

    FName = Trim(Text3.text)
    Open FName For Input As #3
        
    Do Until EOF(3)
        Line Input #3, Nextline
        
        If UCase(Left(Trim(Nextline), 4)) <> "ITEM" And Replace(Replace(Nextline, " ", ""), ",", "") <> "" Then

            Dim bid

            bid = Split(Nextline, ",")
             
            id = 0
        
            '付值
            gcHeaderTemp.Created_By = gUserName
            gcDetailTemp.ITEM = bid(0)
            gcHeaderTemp.po_no = bid(1)
            gcHeaderTemp.SUPPLIER = bid(2)
            gcHeaderTemp.ShipTo = bid(3)
            gcHeaderTemp.Fab_Device = bid(4)
            
            gcHeaderTemp.Customer_Device = bid(5)
            gcHeaderTemp.GC_Version = bid(6)
            gcDetailTemp.Marking_Lot_ID = bid(7)
   
            str01 = bid(8)
            
            If InStr(str01, "月") > 0 Then
            
                str03 = Replace(str01, "月", "-")
                str03 = Replace(str03, "日", "")
                str03 = Year(DATE) & "-" & str03
                gcHeaderTemp.GC_Date = str03
            
            Else
            
                gcHeaderTemp.GC_Date = bid(8)
            
            End If
            
            gcHeaderTemp.Lot_id = bid(9)
            gcDetailTemp.Lot_id = bid(9)
            gcDetailTemp.wafer_id = bid(10)
            gcDetailTemp.Good_Die_Qty = CInt(bid(11))
            gcDetailTemp.REMARK = "WLT"
            gcHeaderTemp.WO_NO = wo_HT_Temp
            gcHeaderTemp.Ship_Out = bid(13)
            
            '2012-11-07 jiayun 修改Good_Die_Qty 根据市场部规则
  
            gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
            
            '2013-12-27 jiayun add
            
            If gcDetailTemp.Good_Die_Qty <= 0 Then
                MsgBox "请确认客户机种对应的Die数是否有维护好！"
                Exit Sub

            End If
            
            '2012-11-05 jiayun 修改 GC
            
            '判断lotID在Header表中是否已存在
            
            If (JudgeGCHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)) Then
            
                If GCHeaderFlag = False Then

                    '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
                End If
                
                '2013-12-05 jiayun add 如果lotid,wo_no 已存在，则查询出id
                '当lotid有隔行时，则查询上次的id
                
                id = GetGCLotIDWOId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)
                
            Else
                '上传到Header表中
                '取目前DB最大的ID号
                id = GetMaxID()
                '2013-01-11 jiayun add 客户简称
                
                If id = 0 Then
                    MsgBox "DB主表ID生成失败1，请联系资讯！"
                    Exit Sub
                
                Else
                
                    Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                    GCHeaderFlag = True
                
                End If
              
            End If
            
            '判断lotID在Detail表中是否已存在
            
            '            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_ID)) Then
            '               MsgBox "GC 这笔：" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_ID & "已存在，无需上传!"
            '
            '            Else
            '上传到Detail表中
            
            '2012-11-05 jiayun 修改 GCT
                   
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & "-" & Right(("0" & gcDetailTemp.wafer_id), 2)
                   
            If id = 0 Then
                MsgBox "DB主表ID生成失败2，请联系资讯！"
                Exit Sub
                
            Else
                Call AddGCWLTDetail(gcDetailTemp, customerTemp, id)
                SumCount = SumCount + 1
                    
            End If
             
            '            End If
 
        End If
        
    Loop
    Close #3
        
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If

End Sub

Private Sub UploadEQ()

    '读取CSV
    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    customerTemp = "EQ"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    Dim dirName  As String

    Dim filename As String

    '获取文件名
    If InStrRev(Trim(Text3.text), "\") > 0 Then
        strFileName = Mid(Trim(Text3.text), InStrRev(Trim(Text3.text), "\") + 1)
        dirName = Mid$(Trim(Text3.text), 1, InStrRev(Trim(Text3.text), "\"))

    End If

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    Dim con          As New ADODB.Connection

    Dim rs           As New ADODB.Recordset

    ' con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
    ' Rs.open "Select * From " & "[" & strFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
    Dim i            As Integer

    Dim j            As Integer

    Dim id           As Long

    Dim TEMP         As String

    Dim SumCount     As Integer

    Dim GCHeaderFlag As Boolean

    Dim str01        As String

    Dim str03        As String

    SumCount = 0
    'Rs.MoveFirst
        
    GCHeaderFlag = False

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count

        ' For i = 0 To Rs.RecordCount - 1
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
            TEMP = ""
            id = 0
        
            '付值
            gcHeaderTemp.Created_By = gUserName
            
            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If

            'gcDetailTemp.ITEM = IIf(IsNull(Rs.fields(0).Value), "", Rs.fields(0).Value)
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If

            ' gcHeaderTemp.PO_NO = IIf(IsNull(Rs.fields(1).Value), "", Rs.fields(1).Value)
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If

            ' gcHeaderTemp.Supplier = Rs.fields(2).Value
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If

            ' gcHeaderTemp.ShipTo = Rs.fields(3).Value
            If j = 5 Then
                gcHeaderTemp.FAB_Device2 = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            ' gcHeaderTemp.FAB_Device2 = IIf(IsNull(Rs.fields(4).Value), "", Rs.fields(4).Value)
            If j = 16 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If

            'gcHeaderTemp.FAB_Device = IIf(IsNull(Rs.fields(15).Value), "", Rs.fields(15).Value)
         
            ' gcHeaderTemp.Customer_Device = IIf(IsNull(Rs.fields(5).Value), "", Rs.fields(5).Value)
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If

            ' gcHeaderTemp.GC_Version = IIf(IsNull(Rs.fields(6).Value), "", Rs.fields(6).Value)
            If j = 8 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If

            'gcDetailTemp.Marking_Lot_ID = IIf(IsNull(Rs.fields(7).Value), "", Rs.fields(7).Value)
            'gcHeaderTemp.GC_Date = Rs.fields(7).Value
            If j = 9 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)

            End If

            ' gcHeaderTemp.Lot_ID = Rs.fields(8).Value
            If j = 9 Then
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If

            ' gcDetailTemp.Lot_ID = Rs.fields(8).Value
            If j = 10 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If

            ' gcDetailTemp.Wafer_id = Rs.fields(9).Value
            If j = 11 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If

            'gcDetailTemp.Good_Die_Qty = CInt(Rs.fields(10).Value)
            If j = 12 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If

            'gcHeaderTemp.WO_NO = IIf(IsNull(Rs.fields(11).Value), "", Rs.fields(11).Value)
            If j = 13 Then
                gcHeaderTemp.remarkTemp = Trim(tempVal)

            End If

            ' gcHeaderTemp.remarkTemp = IIf(IsNull(Rs.fields(12).Value), "", Rs.fields(12).Value)
            If j = 14 Then
                gcHeaderTemp.DATE_CODE = Trim(tempVal)

            End If

            ' gcHeaderTemp.Date_Code = IIf(IsNull(Rs.fields(13).Value), "", Rs.fields(13).Value)
            If j = 15 Then
                gcHeaderTemp.Marking_Lot_ID1 = Trim(tempVal)

            End If

            'gcHeaderTemp.Marking_Lot_ID1 = IIf(IsNull(Rs.fields(14).Value), "", Rs.fields(14).Value)
            If j = 16 Then
                gcHeaderTemp.Marking_Lot_ID2 = Trim(tempVal)

            End If

            ' gcHeaderTemp.Marking_Lot_ID2 = IIf(IsNull(Rs.fields(15).Value), "", Rs.fields(15).Value)
            ' gcDetailTemp.Marking_Lot_ID = IIf(IsNull(Rs.fields(14).Value), "", Rs.fields(14).Value) & " " & IIf(IsNull(Rs.fields(15).Value), "", Rs.fields(15).Value)
            If j = 17 Then
                gcHeaderTemp.Veqdatecode = Trim(tempVal)

            End If

            'gcHeaderTemp.Veqdatecode = IIf(IsNull(Rs.fields(16).Value), "", Rs.fields(16).Value)
            
            '2013-12-05 jiayun add
            '判断wo是否为空
            
            ' If Trim(gcHeaderTemp.WO_NO) = "" Then
            
            ' MsgBox "WO_NO有空值，请确认！"
            ' Exit Sub

            '  End If
            
            '2012-11-07 jiayun 修改Good_Die_Qty 根据市场部规则
  
            'gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
            
            '2013-12-27 jiayun add
            
            '            If gcDetailTemp.Good_Die_Qty <= 0 Then
            '                    MsgBox "请确认客户机种对应的Die数是否有维护好！"
            '                    Exit Sub
            '            End If
            
            '2012-11-05 jiayun 修改 GC
            
            '判断lotID在Header表中是否已存在
        Next j
        
        
         ' Changed by: Project Administrator at: 2019/8/30-9:15:35 on machine: DESKTOP-MSUG5JD NPI: 余康生要求增加换行符\\
        'gcDetailTemp.Marking_Lot_ID = gcHeaderTemp.Marking_Lot_ID1 & gcHeaderTemp.Marking_Lot_ID2
            
        gcDetailTemp.Marking_Lot_ID = gcHeaderTemp.Marking_Lot_ID1 & "\\" & gcHeaderTemp.Marking_Lot_ID2
        
        If (JudgeEQHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO, gcHeaderTemp.po_no)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            '2013-12-05 jiayun add 如果lotid,wo_no 已存在，则查询出id
            '当lotid有隔行时，则查询上次的id
                
            id = GetGCLotIDWOId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
            '2013-01-11 jiayun add 客户简称
                
            If id = 0 Then
                MsgBox "DB主表ID生成失败1，请联系资讯！"
                Exit Sub
                
            Else
                
                Call AddEQHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
                
            End If
              
        End If
            
        '判断lotID在Detail表中是否已存在
            
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "GC 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
               
        Else
            '上传到Detail表中
            
            '2012-11-05 jiayun 修改 GCT
                   
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
                   
            If id = 0 Then
                MsgBox "DB主表ID生成失败2，请联系资讯！"
                Exit Sub
                
            Else
                Call AddGCDetail(gcDetailTemp, customerTemp, id)
                SumCount = SumCount + 1
                    
            End If
                
        End If
            
        'Rs.MoveNext
        
    Next i
        
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If

End Sub

Private Sub UploadEQ_ShippingRequest()

    Dim customerTemp As String

    Dim SumCount     As Integer

    customerTemp = "EQ"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    Set VBExcel = CreateObject("excel.application")
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(Text3.text)
    Set xlSheet = xlBook.Worksheets(1)

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 23 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If
    
    Dim i     As Integer

    Dim j     As Integer

    Dim id    As Long

    Dim TEMP  As String

    Dim temp2 As String

    SumCount = 0

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 2 To xlSheet.Range("A1").CurrentRegion.Columns.count

            strChar = Chr(96 + j)

            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            eqISHeaderTemp.Created_By = gUserName

            If j = 2 Then
                eqISHeaderTemp.SUBCONPO = Trim(tempVal)

            End If
            
            If j = 3 Then
                eqISHeaderTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 4 Then
                eqISHeaderTemp.Quantity = Trim(tempVal)

            End If
            
            If j = 5 Then
                eqISHeaderTemp.devicetemp = Trim(tempVal)

            End If
            
            If j = 6 Then
                eqISHeaderTemp.SPATemp = Trim(tempVal)

            End If

            '------
            If j = 7 Then
                eqISHeaderTemp.CSD = Trim(tempVal)

            End If
            
            If j = 8 Then
                eqISHeaderTemp.lot = Trim(tempVal)

            End If
            
            If j = 9 Then
                '
                eqISHeaderTemp.DATECODE1 = Trim(tempVal)
             
            End If
            
            If j = 10 Then
                eqISHeaderTemp.DELIVERYNAME = Trim(tempVal)

            End If
            
            If j = 11 Then
                eqISHeaderTemp.DELIVERYADDRESS = Trim(tempVal)
                
            End If

            '--------
            If j = 12 Then
                eqISHeaderTemp.WAREHOUSE = Trim(tempVal)

            End If
            
            If j = 13 Then
                eqISHeaderTemp.LOCATION = Trim(tempVal)
                
            End If
            
            If j = 14 Then
                eqISHeaderTemp.MODEOFDELIVERY = Trim(tempVal)

            End If
            
            If j = 15 Then
                eqISHeaderTemp.dateCodeTemp = Trim(tempVal)

            End If
            
            If j = 16 Then
                eqISHeaderTemp.SO = Trim(tempVal)

            End If
            
            If j = 17 Then
                eqISHeaderTemp.CARRIERNOTES = Trim(tempVal)

            End If
            
            If j = 18 Then
                eqISHeaderTemp.LINE = Trim(tempVal)

            End If
            
            If j = 19 Then
                eqISHeaderTemp.SCHEDULELINE = Trim(tempVal)

            End If
            
            If j = 20 Then
                eqISHeaderTemp.CUSTPN = Trim(tempVal)
                
            End If

            If j = 21 Then
                eqISHeaderTemp.COUNTRYANDNAMEOFDISTRIBUTOR = Trim(tempVal)
                
            End If

            If j = 22 Then
                eqISHeaderTemp.CUSTOMER = Trim(tempVal)
                
            End If

            If j = 23 Then
                eqISHeaderTemp.CUSTOMERPO = Trim(tempVal)
                
            End If
            
            '----------------------
        
        Next j
    
        If (JudgeEQISShippingRequest(eqISHeaderTemp.SO, eqISHeaderTemp.LINE, eqISHeaderTemp.SCHEDULELINE, eqISHeaderTemp.CUSTOMERPO, eqISHeaderTemp.dateCodeTemp, eqISHeaderTemp.DATECODE1, eqISHeaderTemp.lot, eqISHeaderTemp.devicetemp, eqISHeaderTemp.SUBCONPO)) Then
               
            MsgBox "已存在，无需上传!"
                
        Else
       
            Call AddEQISHeader_ShippingRequest(eqISHeaderTemp)
            SumCount = SumCount + 1
              
        End If

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If

End Sub

Private Sub UploadEQ_IS()

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = "EQ"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 30 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            ' strChar = Chr(96 + j)
        
            If j > 26 Then
                strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
            Else
                strChar = Chr(96 + j)

            End If
             
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            eqISHeaderTemp.Created_By = gUserName

            If j = 1 Then
                eqISHeaderTemp.Created_Datetime = Trim(tempVal)

            End If
            
            If j = 2 Then
                eqISHeaderTemp.Vendor = Trim(tempVal)

            End If
            
            If j = 3 Then
                eqISHeaderTemp.Process = Trim(tempVal)

            End If
            
            If j = 4 Then
                eqISHeaderTemp.ORDERTYPE = Trim(tempVal)

            End If
            
            If j = 5 Then
                eqISHeaderTemp.ESR_No = Trim(tempVal)

            End If

            '------
            If j = 6 Then
                eqISHeaderTemp.AssemblyDateCode = Trim(tempVal)

            End If
            
            If j = 7 Then
                eqISHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 8 Then
                '                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                eqISHeaderTemp.WO_NO = Trim(tempVal)
             
            End If
            
            If j = 9 Then
                eqISHeaderTemp.WorkOrder_PartNo = Trim(tempVal)

            End If
            
            If j = 10 Then
                eqISHeaderTemp.device = Trim(tempVal)
                
            End If

            '--------
            If j = 11 Then
                eqISHeaderTemp.WaferQTY = Trim(tempVal)

            End If
            
            If j = 12 Then
                eqISHeaderTemp.AssyQty = Trim(tempVal)
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
                
            End If
            
            If j = 13 Then
                eqISHeaderTemp.PACKAGE = Trim(tempVal)

            End If
            
            If j = 14 Then
                eqISHeaderTemp.FabLotNo = Trim(tempVal)

            End If
            
            If j = 15 Then
                eqISHeaderTemp.TSM_A = Trim(tempVal)

            End If

            '--------------------
            If j = 16 Then
                eqISHeaderTemp.TSM_B = Trim(tempVal)

            End If
            
            If j = 17 Then
                eqISHeaderTemp.TSM_C = Trim(tempVal)

            End If
            
            If j = 18 Then
                eqISHeaderTemp.TSM_D = Trim(tempVal)

            End If
            
            If j = 19 Then
                eqISHeaderTemp.BondingDiagram = Trim(tempVal)

            End If
            
            If j = 20 Then
                eqISHeaderTemp.CompleteLotno = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)
                
            End If
            
            '----------------------
            
            If j = 21 Then
                eqISHeaderTemp.Remarks = Trim(tempVal)

            End If

            If j = 22 Then
                eqISHeaderTemp.MarketingPartNumber = Trim(tempVal)

            End If

            If j = 23 Then
                eqISHeaderTemp.SPA = Trim(tempVal)

            End If

            If j = 24 Then
                eqISHeaderTemp.DATECODE = Trim(tempVal)

            End If

            If j = 25 Then
                eqISHeaderTemp.DieID = Trim(tempVal)

            End If
            
            '---------------------
              
            If j = 26 Then
                eqISHeaderTemp.LabelFormat = Trim(tempVal)

            End If

            If j = 27 Then
                eqISHeaderTemp.waferid = Trim(tempVal)
                gcDetailTemp.wafer_id = Trim(tempVal)
                  
            End If

            If j = 28 Then
                eqISHeaderTemp.SPADESC = Trim(tempVal)

            End If

            If j = 29 Then
                eqISHeaderTemp.Attention = Trim(tempVal)

            End If

            If j = 30 Then
                eqISHeaderTemp.CompanyName = Trim(tempVal)

            End If
        
        Next j
        
        gcDetailTemp.Marking_Lot_ID = eqISHeaderTemp.TSM_A & "\\" & eqISHeaderTemp.TSM_B
        
    
        If (JudgeEQISHeaderId(eqISHeaderTemp.po_no, eqISHeaderTemp.WO_NO, eqISHeaderTemp.CompleteLotno)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetEQISLotIDPOId(eqISHeaderTemp.CompleteLotno, eqISHeaderTemp.po_no)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddEQISHeader(eqISHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '    '判断lotID在Detail表中是否已存在
        '
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "EQ 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"

        Else
            '    '上传到Detail表中

            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)

            Call AddEQDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1

        End If
    
        ' 明细表下次再改------------------------
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If

    '------------------
    '读取CSV
    'Dim source_batch_id_Temp As String
    'Dim customerTemp As String
    '
    'customerTemp = "EQ"
    '
    ''上传OI的CSV
    ''处理文件名
    'If Text3.Text = "" Then
    '    MsgBox "先选择待上传的文件"
    '    Exit Sub
    'End If
    'Dim dirName As String
    'Dim FileName As String
    '
    ''获取文件名
    '    If InStrRev(Trim(Text3.Text), "\") > 0 Then
    '        strFileName = Mid(Trim(Text3.Text), InStrRev(Trim(Text3.Text), "\") + 1)
    '        dirName = Mid$(Trim(Text3.Text), 1, InStrRev(Trim(Text3.Text), "\"))
    '    End If
    '
    'Dim con As New ADODB.Connection
    'Dim Rs As New ADODB.Recordset
    '
    '
    '        con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
    '        Rs.open "Select * From " & "[" & strFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
    '
    '        Dim i As Integer
    '        Dim j As Integer
    '        Dim id As Long
    '        Dim temp As String
    '        Dim SumCount As Integer
    '        Dim GCHeaderFlag As Boolean
    '        Dim str01 As String
    '        Dim str03 As String
    '        SumCount = 0
    '        Rs.MoveFirst
    '
    '        GCHeaderFlag = False
    '
    '        For i = 0 To Rs.RecordCount - 1
    '            temp = ""
    '            id = 0
    '
    '            '付值
    '            gcHeaderTemp.Created_By = gUserName
    '            gcDetailTemp.item = IIf(IsNull(Rs.fields(0).Value), "", Rs.fields(0).Value)
    '            gcHeaderTemp.PO_NO = IIf(IsNull(Rs.fields(1).Value), "", Rs.fields(1).Value)
    '            gcHeaderTemp.Supplier = Rs.fields(2).Value
    '            gcHeaderTemp.ShipTo = Rs.fields(3).Value
    '            gcHeaderTemp.FAB_Device = Rs.fields(4).Value
    '            gcHeaderTemp.Customer_Device = Rs.fields(5).Value
    '            gcHeaderTemp.GC_Version = IIf(IsNull(Rs.fields(6).Value), "", Rs.fields(6).Value)
    '            'gcDetailTemp.Marking_Lot_ID = IIf(IsNull(Rs.fields(7).Value), "", Rs.fields(7).Value)
    '            gcHeaderTemp.GC_Date = Rs.fields(7).Value
    '
    '
    '            gcHeaderTemp.Lot_ID = Rs.fields(8).Value
    '            gcDetailTemp.Lot_ID = Rs.fields(8).Value
    '            gcDetailTemp.Wafer_ID = Rs.fields(9).Value
    '            gcDetailTemp.Good_Die_Qty = CInt(Rs.fields(10).Value)
    '            gcHeaderTemp.WO_NO = Rs.fields(11).Value
    '            gcHeaderTemp.remarkTemp = Rs.fields(12).Value
    '            gcHeaderTemp.Date_Code = Rs.fields(13).Value
    '            gcHeaderTemp.Marking_Lot_ID1 = Rs.fields(14).Value
    '            gcHeaderTemp.Marking_Lot_ID2 = Rs.fields(15).Value
    '            gcDetailTemp.Marking_Lot_ID = Rs.fields(14).Value & " " & Rs.fields(15).Value
    '
    '
    '
    '            '2013-12-05 jiayun add
    '            '判断wo是否为空
    '
    '            If Trim(gcHeaderTemp.WO_NO) = "" Then
    '
    '                MsgBox "WO_NO有空值，请确认！"
    '                Exit Sub
    '
    '            End If
    '
    '            '2012-11-07 jiayun 修改Good_Die_Qty 根据市场部规则
    '
    '            'gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
    '
    '            '2013-12-27 jiayun add
    '
    ''            If gcDetailTemp.Good_Die_Qty <= 0 Then
    ''                    MsgBox "请确认客户机种对应的Die数是否有维护好！"
    ''                    Exit Sub
    ''            End If
    '
    '
    '            '2012-11-05 jiayun 修改 GC
    '
    '            '判断lotID在Header表中是否已存在
    '
    '            If (JudgeEQHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO, gcHeaderTemp.PO_NO)) Then
    '
    '                If GCHeaderFlag = False Then
    '        '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
    '                End If
    '
    '                '2013-12-05 jiayun add 如果lotid,wo_no 已存在，则查询出id
    '                '当lotid有隔行时，则查询上次的id
    '
    ''                id = GetGCLotIDWOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)
    '
    '            Else
    '            '上传到Header表中
    '                '取目前DB最大的ID号
    '                id = GetMaxID()
    '                '2013-01-11 jiayun add 客户简称
    '
    '                If id = 0 Then
    '                    MsgBox "DB主表ID生成失败1，请联系资讯！"
    '                    Exit Sub
    '
    '                Else
    '
    '
    '                    Call AddEQHeader(gcHeaderTemp, id, customerTemp)
    '                    GCHeaderFlag = True
    '
    '                End If
    '
    '            End If
    '
    '
    '            '判断lotID在Detail表中是否已存在
    '
    '            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_ID)) Then
    '               MsgBox "GC 这笔：" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_ID & "已存在，无需上传!"
    '
    '            Else
    '            '上传到Detail表中
    '
    '                   '2012-11-05 jiayun 修改 GCT
    '
    '
    '                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_ID), 2)
    '
    '
    '                If id = 0 Then
    '                    MsgBox "DB主表ID生成失败2，请联系资讯！"
    '                    Exit Sub
    '
    '                Else
    '                    Call AddGCDetail(gcDetailTemp, customerTemp, id)
    '                    SumCount = SumCount + 1
    '
    '                End If
    '
    '
    '            End If
    '
    '
    '            Rs.MoveNext
    '
    '        Next i
    '
    '
    '        If SumCount > 0 Then
    '            MsgBox "已成功上传" & SumCount & "笔！"
    '        End If

End Sub

Private Sub UploadMC()

    '读取CSV
    '2013-12-17 jiayun add MC
    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    customerTemp = "MC"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    Dim dirName  As String

    Dim filename As String

    '获取文件名
    If InStrRev(Trim(Text3.text), "\") > 0 Then
        strFileName = Mid(Trim(Text3.text), InStrRev(Trim(Text3.text), "\") + 1)
        dirName = Mid$(Trim(Text3.text), 1, InStrRev(Trim(Text3.text), "\"))

    End If

    Dim con As New ADODB.Connection

    Dim rs  As New ADODB.Recordset

    con.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
    rs.Open "Select * From " & "[" & strFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
    Dim i            As Integer

    Dim j            As Integer

    Dim id           As Long

    Dim TEMP         As String

    Dim SumCount     As Integer

    Dim GCHeaderFlag As Boolean

    SumCount = 0
    rs.MoveFirst
        
    GCHeaderFlag = False
        
    For i = 0 To rs.RecordCount - 1
        TEMP = ""
        id = 0
        
        '付值
        gcHeaderTemp.Created_By = gUserName
        gcDetailTemp.ITEM = IIf(IsNull(rs.Fields(0).Value), "", rs.Fields(0).Value)
        gcHeaderTemp.po_no = Trim(IIf(IsNull(rs.Fields(1).Value), "", rs.Fields(1).Value))
        gcHeaderTemp.SUPPLIER = Trim(rs.Fields(2).Value)
        gcHeaderTemp.ShipTo = Trim(rs.Fields(3).Value)
        gcHeaderTemp.Fab_Device = Trim(rs.Fields(4).Value)
        gcHeaderTemp.Customer_Device = Trim(rs.Fields(5).Value)
        gcHeaderTemp.GC_Version = Trim(IIf(IsNull(rs.Fields(6).Value), "", rs.Fields(6).Value))
        gcDetailTemp.Marking_Lot_ID = Trim(IIf(IsNull(rs.Fields(7).Value), "", rs.Fields(7).Value))
        gcHeaderTemp.GC_Date = rs.Fields(8).Value
        gcHeaderTemp.Lot_id = Trim(rs.Fields(9).Value)
        gcDetailTemp.Lot_id = Trim(rs.Fields(9).Value)
        gcDetailTemp.wafer_id = Trim(rs.Fields(10).Value)
        gcDetailTemp.Good_Die_Qty = CInt(rs.Fields(11).Value)
        gcHeaderTemp.WO_NO = Trim(IIf(IsNull(rs.Fields(12).Value), "", rs.Fields(12).Value))
            
        gcHeaderTemp.TradeType = Trim(IIf(IsNull(rs.Fields(15).Value), "", rs.Fields(15).Value))
            
        '判断lotID在Header表中是否已存在
            
        If (JudgeMCHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            '2013-12-05 jiayun add 如果lotid,wo_no 已存在，则查询出id
            '当lotid有隔行时，则查询上次的id
                
            id = GetMCLotIDWOId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
            '2013-01-11 jiayun add 客户简称
                
            If id = 0 Then
                MsgBox "DB主表ID生成失败1，请联系资讯！"
                Exit Sub
                
            Else
                
                Call AddGCHeader(gcHeaderTemp, id, customerTemp)
                GCHeaderFlag = True
                
            End If
              
        End If
            
        '判断lotID在Detail表中是否已存在
            
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "GC 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
               
        Else
            '上传到Detail表中
                   
            '                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_ID), 2)
                
            gcDetailTemp.ITEM = gcDetailTemp.wafer_id
                 
            gcDetailTemp.wafer_id = Trim(Right(gcDetailTemp.wafer_id, 2))
                   
            If id = 0 Then
                MsgBox "DB主表ID生成失败2，请联系资讯！"
                Exit Sub
                
            Else
                Call AddGCDetail(gcDetailTemp, customerTemp, id)
                SumCount = SumCount + 1
                    
            End If
                
        End If
            
        rs.MoveNext
        
    Next i
        
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If

End Sub

'2014-02-10 jiayun add
Private Sub UploadNormalCustomer(customerNameTemp As String)

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    Dim tax                  As String

    customerTemp = customerNameTemp

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    If InStr(Text3.text, "A-") > 0 Then
        tax = "A"
    Else
        tax = "B"

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i         As Integer

    Dim j         As Integer

    Dim id        As Long

    Dim TEMP      As String

    Dim temp2     As String

    Dim tempVal   As String

    Dim mCodetemp As String

    Dim yTemp     As String

    Dim mTemp     As String

    Dim charTemp  As Long

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                
                If customerTemp = "MG" Then
                    
                    yTemp = Right(Year(DATE), 1)
                    mTemp = GetMonthChar(Month(DATE))
                    charTemp = GetHWMonthMaxQty()
                    
                    mCodetemp = yTemp + mTemp + ToNumberSystem26(charTemp)
                    gcDetailTemp.Marking_Lot_ID = mCodetemp
                    
                End If
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
            
                If customerNameTemp = "MR" Then
                    gcDetailTemp.wafer_id = Right(Trim(tempVal), 2)
                
                Else
            
                    If IsNumeric(Trim(tempVal)) = False Then
                        MsgBox "WaferId类型不对，请核对要上传的源文档 !"
                        Exit Sub
                        
                    Else
                         
                        gcDetailTemp.wafer_id = Trim(tempVal)
                         
                    End If
                
                End If
                
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 15 Then 'ccs add 20161130
                gcHeaderTemp.Ship_Out = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j
    
        gcHeaderTemp.TAXTYPE = tax
         
        If (JudgePOHeaderIdNew(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetPOLotIDPOIdNew(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        '2014-03-04 jiayun add  CN Wo  不用抛数据到Mapping表

        If customerNameTemp = "CN" Then
            SumCount = SumCount + 1
      
        ElseIf customerNameTemp = "GT" Then
            SumCount = SumCount + 1
      
        Else
    
            If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
                MsgBox "这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
               
            Else
                '上传到Detail表中
                   
                If customerNameTemp = "CN" Then
                    gcDetailTemp.ITEM = gcDetailTemp.wafer_id
             
                ElseIf customerNameTemp = "MR" Then
                   
                    gcDetailTemp.ITEM = gcDetailTemp.Lot_id & "-" & Right(("0" & gcDetailTemp.wafer_id), 2)
                
                Else
                   
                    gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
                   
                End If

                Call AddGCDetail(gcDetailTemp, customerTemp, id)
                   
                SumCount = SumCount + 1
              
            End If
            
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

'2016-01-08 jiayun add  MPS

Private Sub UploadMPSCustomer(customerNameTemp As String)

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = customerNameTemp

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i         As Integer

    Dim j         As Integer

    Dim id        As Long

    Dim TEMP      As String

    Dim temp2     As String

    Dim tempVal   As String

    Dim mCodetemp As String

    Dim yTemp     As String

    Dim mTemp     As String

    Dim charTemp  As Long

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                
                If customerTemp = "MG" Then
                    
                    yTemp = Right(Year(DATE), 1)
                    mTemp = GetMonthChar(Month(DATE))
                    charTemp = GetHWMonthMaxQty()
                    
                    mCodetemp = yTemp + mTemp + ToNumberSystem26(charTemp)
                    gcDetailTemp.Marking_Lot_ID = mCodetemp
                    
                End If
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
            
                If customerNameTemp = "MR" Then
                    gcDetailTemp.wafer_id = Right(Trim(tempVal), 2)
                
                Else
            
                    If IsNumeric(Trim(tempVal)) = False Then
                        MsgBox "WaferId类型不对，请核对要上传的源文档 !"
                        Exit Sub
                        
                    Else
                         
                        gcDetailTemp.wafer_id = Trim(tempVal)
                         
                    End If
                
                End If
                
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.Ship_Out = Trim(tempVal)

            End If
        
            If j = 14 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j

        If (JudgePOHeaderIdNew(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetPOLotIDPOIdNew(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddMPSHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        '2014-03-04 jiayun add  CN Wo  不用抛数据到Mapping表

        If customerNameTemp = "CN" Then
            SumCount = SumCount + 1
      
        ElseIf customerNameTemp = "GT" Then
            SumCount = SumCount + 1
      
        Else
    
            If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
                MsgBox "这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
               
            Else
                '上传到Detail表中
                   
                If customerNameTemp = "CN" Then
                    gcDetailTemp.ITEM = gcDetailTemp.wafer_id
             
                ElseIf customerNameTemp = "MR" Then
                   
                    gcDetailTemp.ITEM = gcDetailTemp.Lot_id & "-" & Right(("0" & gcDetailTemp.wafer_id), 2)
                
                Else
                   
                    gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
                   
                End If

                Call AddGCDetail(gcDetailTemp, customerTemp, id)
                   
                SumCount = SumCount + 1
              
            End If
            
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub UploadNormalCustomer77(customerNameTemp As String)

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = customerNameTemp

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i         As Integer

    Dim j         As Integer

    Dim id        As Long

    Dim TEMP      As String

    Dim temp2     As String

    Dim tempVal   As String

    Dim mCodetemp As String

    Dim yTemp     As String

    Dim mTemp     As String

    Dim charTemp  As Long

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                
                If customerTemp = "MG" Then
                    
                    yTemp = Right(Year(DATE), 1)
                    mTemp = GetMonthChar(Month(DATE))
                    charTemp = GetHWMonthMaxQty()
                    
                    mCodetemp = yTemp + mTemp + ToNumberSystem26(charTemp)
                    gcDetailTemp.Marking_Lot_ID = mCodetemp
                    
                End If
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
            
                If customerNameTemp = "MR" Then
                    gcDetailTemp.wafer_id = Right(Trim(tempVal), 2)
                
                Else
            
                    '                        If IsNumeric(Trim(tempVal)) = False Then
                    '                         MsgBox "WaferId类型不对，请核对要上传的源文档 !"
                    '                         Exit Sub
                    '
                    '                        Else
                         
                    gcDetailTemp.wafer_id = Trim(tempVal)
                         
                    'End If
                
                End If
                
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j

        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        '2014-03-04 jiayun add  CN Wo  不用抛数据到Mapping表

        If customerNameTemp = "CN" Then
            SumCount = SumCount + 1
      
        ElseIf customerNameTemp = "GT" Then
            SumCount = SumCount + 1
      
        Else
    
            If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
                MsgBox "这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
               
            Else
                '上传到Detail表中
                   
                If customerNameTemp = "CN" Then
                    gcDetailTemp.ITEM = gcDetailTemp.wafer_id
             
                ElseIf customerNameTemp = "MR" Then
                   
                    gcDetailTemp.ITEM = gcDetailTemp.Lot_id & "-" & Right(("0" & gcDetailTemp.wafer_id), 2)
                
                Else
                   
                    gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
                   
                End If

                Call AddGCDetail(gcDetailTemp, customerTemp, id)
                   
                SumCount = SumCount + 1
              
            End If
            
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

'2015-09-11 jiayun add 56
Private Sub UploadNormalCustomer56(customerNameTemp As String)

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = customerNameTemp

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i              As Integer

    Dim j              As Integer

    Dim id             As Long

    Dim TEMP           As String

    Dim temp2          As String

    Dim tempVal        As String

    Dim mCodetemp      As String

    Dim yTemp          As String

    Dim mTemp          As String

    Dim charTemp       As Long

    Dim waferAllDieQty As Long

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
        waferAllDieQty = 0
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                
                If customerTemp = "MG" Then
                    
                    yTemp = Right(Year(DATE), 1)
                    mTemp = GetMonthChar(Month(DATE))
                    charTemp = GetHWMonthMaxQty()
                    
                    mCodetemp = yTemp + mTemp + ToNumberSystem26(charTemp)
                    gcDetailTemp.Marking_Lot_ID = mCodetemp
                    
                End If
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
            
                If customerNameTemp = "MR" Then
                    gcDetailTemp.wafer_id = Right(Trim(tempVal), 2)
                
                Else
            
                    If IsNumeric(Trim(tempVal)) = False Then
                        MsgBox "WaferId类型不对，请核对要上传的源文档 !"
                        Exit Sub
                        
                    Else
                         
                        gcDetailTemp.wafer_id = Trim(tempVal)
                         
                    End If
                
                End If
                
            End If
            
            If j = 12 Then
                'gcDetailTemp.Good_Die_Qty = Trim(tempVal)
                waferAllDieQty = CLng(Trim(tempVal))
                
            End If
            
            If j = 13 Then
                gcDetailTemp.Good_Die_Qty = 0
             
                gcDetailTemp.NG_Die_Qty = 0
                 
            End If
            
            If j = 14 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
               
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j

        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        '2014-03-04 jiayun add  CN Wo  不用抛数据到Mapping表

        If customerNameTemp = "CN" Then
            SumCount = SumCount + 1
      
        ElseIf customerNameTemp = "GT" Then
            SumCount = SumCount + 1
          
        ElseIf customerNameTemp = "56" Then
            SumCount = SumCount + 1
            MsgBox "Mapping没有上传, 请单独上传"
      
        Else
    
            If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
                MsgBox "这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
               
            Else
                '上传到Detail表中
                   
                If customerNameTemp = "CN" Then
                    gcDetailTemp.ITEM = gcDetailTemp.wafer_id
             
                ElseIf customerNameTemp = "MR" Then
                   
                    gcDetailTemp.ITEM = gcDetailTemp.Lot_id & "-" & Right(("0" & gcDetailTemp.wafer_id), 2)
                
                Else
                   
                    gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
                   
                End If

                Call Add56Detail(gcDetailTemp, customerTemp, id)
                   
                SumCount = SumCount + 1
              
            End If
            
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

'2014-02-10 jiayun add
Private Sub UploadNormalCustomerZL(customerNameTemp As String)

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = customerNameTemp

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i         As Integer

    Dim j         As Integer

    Dim id        As Long

    Dim TEMP      As String

    Dim temp2     As String

    Dim tempVal   As String

    Dim mCodetemp As String

    Dim yTemp     As String

    Dim mTemp     As String

    Dim charTemp  As Long

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                
                If customerTemp = "MG" Then
                    
                    yTemp = Right(Year(DATE), 1)
                    mTemp = GetMonthChar(Month(DATE))
                    charTemp = GetHWMonthMaxQty()
                    
                    mCodetemp = yTemp + mTemp + ToNumberSystem26(charTemp)
                    gcDetailTemp.Marking_Lot_ID = mCodetemp
                    
                End If
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
            
                If IsNumeric(Trim(tempVal)) = False Then
                    MsgBox "WaferId类型不对，请核对要上传的源文档 !"
                    Exit Sub
               
                Else
               
                    gcDetailTemp.wafer_id = Trim(tempVal)
                
                End If
                
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcDetailTemp.NG_Die_Qty = CLng(Trim(tempVal)) - gcDetailTemp.Good_Die_Qty

            End If
            
            If j = 14 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j

        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        '2014-03-04 jiayun add  CN Wo  不用抛数据到Mapping表

        If customerNameTemp = "CN" Then
            SumCount = SumCount + 1
      
        ElseIf customerNameTemp = "GT" Then
            SumCount = SumCount + 1
      
        Else
    
            If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
                MsgBox "这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
               
            Else
                '上传到Detail表中
                   
                If customerNameTemp = "CN" Then
                    gcDetailTemp.ITEM = gcDetailTemp.wafer_id
                   
                Else
                   
                    gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
                   
                End If

                Call AddGCDetailZL(gcDetailTemp, customerTemp, id)
                   
                SumCount = SumCount + 1
              
            End If
            
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

'2015-04-08 jiayun add
Private Sub UploadNormalCustomerCS(customerNameTemp As String)

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = customerNameTemp

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i         As Integer

    Dim j         As Integer

    Dim id        As Long

    Dim TEMP      As String

    Dim temp2     As String

    Dim tempVal   As String

    Dim mCodetemp As String

    Dim yTemp     As String

    Dim mTemp     As String

    Dim charTemp  As Long

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                
                If customerTemp = "MG" Then
                    
                    yTemp = Right(Year(DATE), 1)
                    mTemp = GetMonthChar(Month(DATE))
                    charTemp = GetHWMonthMaxQty()
                    
                    mCodetemp = yTemp + mTemp + ToNumberSystem26(charTemp)
                    gcDetailTemp.Marking_Lot_ID = mCodetemp
                    
                End If
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
            
                If IsNumeric(Trim(tempVal)) = False Then
                    MsgBox "WaferId类型不对，请核对要上传的源文档 !"
                    Exit Sub
               
                Else
               
                    gcDetailTemp.wafer_id = Trim(tempVal)
                
                End If
                
            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcDetailTemp.NG_Die_Qty = CLng(Trim(tempVal)) - gcDetailTemp.Good_Die_Qty

            End If
            
            If j = 14 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 15 Then
                gcHeaderTemp.DATE_CODE = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j

        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddCSHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        '2014-03-04 jiayun add  CN Wo  不用抛数据到Mapping表

        If customerNameTemp = "CN" Then
            SumCount = SumCount + 1
      
        ElseIf customerNameTemp = "GT" Then
            SumCount = SumCount + 1
      
        Else
    
            If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
                MsgBox "这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
               
            Else
                '上传到Detail表中
                   
                If customerNameTemp = "CN" Then
                    gcDetailTemp.ITEM = gcDetailTemp.wafer_id
                   
                Else
                   
                    gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
                   
                End If

                Call AddGCDetailZL(gcDetailTemp, customerTemp, id)
                   
                SumCount = SumCount + 1
              
            End If
            
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

'2014-09-17 jiayun add
Private Sub UploadQR(customerNameTemp As String)

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = customerNameTemp

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcDetailTemp.NG_Die_Qty = Trim(tempVal) - gcDetailTemp.Good_Die_Qty

            End If
            
            If j = 14 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j

        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        '2014-03-04 jiayun add  CN Wo  不用抛数据到Mapping表

        If customerNameTemp = "CN" Then
            SumCount = SumCount + 1
      
        ElseIf customerNameTemp = "SI" Then
            SumCount = SumCount + 1
      
        Else
    
            If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
                MsgBox "这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
               
            Else
                '上传到Detail表中
                   
                If customerNameTemp = "CN" Then
                    gcDetailTemp.ITEM = gcDetailTemp.wafer_id
                   
                Else
                   
                    gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
                   
                End If

                Call AddQRDetail(gcDetailTemp, customerTemp, id)
                   
                SumCount = SumCount + 1
              
            End If
            
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

'2015-09-07 jiayun add  QR第二次回来
Private Sub UploadQRV2(customerNameTemp As String)

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = customerNameTemp

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcDetailTemp.NG_Die_Qty = Trim(tempVal) - gcDetailTemp.Good_Die_Qty

            End If
            
            If j = 14 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j

        If (JudgeQR2HeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetQR2LotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddQR2Header(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在

        If (JudgeQR2DetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
       
        Else
            '上传到Detail表中

            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)

            Call AddQR2Detail(gcDetailTemp, customerTemp, id)
           
            SumCount = SumCount + 1
      
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub UploadHY()

    '读取CSV
    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    customerTemp = "HY"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    Dim dirName  As String

    Dim filename As String

    '获取文件名
    If InStrRev(Trim(Text3.text), "\") > 0 Then
        strFileName = Mid(Trim(Text3.text), InStrRev(Trim(Text3.text), "\") + 1)
        dirName = Mid$(Trim(Text3.text), 1, InStrRev(Trim(Text3.text), "\"))

    End If

    Dim con As New ADODB.Connection

    Dim rs  As New ADODB.Recordset

    con.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
    rs.Open "Select * From " & "[" & strFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
    Dim i            As Integer

    Dim j            As Integer

    Dim id           As Long

    Dim TEMP         As String

    Dim SumCount     As Integer

    Dim GCHeaderFlag As Boolean

    SumCount = 0
    rs.MoveFirst
        
    GCHeaderFlag = False
        
    For i = 0 To rs.RecordCount - 1
        TEMP = ""
        
        '付值
        gcHeaderTemp.Created_By = gUserName
        gcDetailTemp.ITEM = IIf(IsNull(rs.Fields(0).Value), "", rs.Fields(0).Value)
        gcHeaderTemp.po_no = Trim(IIf(IsNull(rs.Fields(1).Value), "", rs.Fields(1).Value))
        gcHeaderTemp.SUPPLIER = Trim(IIf(IsNull(rs.Fields(2).Value), "", rs.Fields(2).Value))
        gcHeaderTemp.ShipTo = Trim(IIf(IsNull(rs.Fields(3).Value), "", rs.Fields(3).Value))
        gcHeaderTemp.Fab_Device = Trim(IIf(IsNull(rs.Fields(4).Value), "", rs.Fields(4).Value))
        gcHeaderTemp.Customer_Device = Trim(rs.Fields(5).Value)
        gcHeaderTemp.GC_Version = Trim(rs.Fields(6).Value)
        gcDetailTemp.Marking_Lot_ID = Trim(IIf(IsNull(rs.Fields(7).Value), "", rs.Fields(7).Value))
        gcHeaderTemp.GC_Date = rs.Fields(8).Value
        gcHeaderTemp.Lot_id = Trim(rs.Fields(9).Value)
        gcDetailTemp.Lot_id = Trim(rs.Fields(9).Value)
        gcDetailTemp.wafer_id = Trim(rs.Fields(10).Value)
        gcDetailTemp.Good_Die_Qty = CInt(rs.Fields(11).Value)
        gcHeaderTemp.WO_NO = Trim(IIf(IsNull(rs.Fields(12).Value), "", rs.Fields(12).Value))
        gcHeaderTemp.TradeType = Trim(IIf(IsNull(rs.Fields(15).Value), "", rs.Fields(15).Value))
            
        '2012-11-07 jiayun 修改Good_Die_Qty 根据市场部规则
        'gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(gcHeaderTemp.Customer_Device, gcDetailTemp.Good_Die_Qty)
   
        '2012-11-05 jiayun 修改 GC
            
        '判断lotID在Header表中是否已存在
            
        If (JudgeGCHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
            '2013-01-11 jiayun add 客户简称
                
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
            
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "HY 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
               
        Else
            '上传到Detail表中
            
            '2012-11-05 jiayun 修改 GCT
                   
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
                   
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1
              
        End If
            
        rs.MoveNext
        
    Next i
        
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If

End Sub

Private Sub UploadHT()

    '读取CSV
    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    customerTemp = "HT"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    Dim dirName  As String

    Dim filename As String

    '获取文件名
    If InStrRev(Trim(Text3.text), "\") > 0 Then
        strFileName = Mid(Trim(Text3.text), InStrRev(Trim(Text3.text), "\") + 1)
        dirName = Mid$(Trim(Text3.text), 1, InStrRev(Trim(Text3.text), "\"))

    End If

    Dim con As New ADODB.Connection

    Dim rs  As New ADODB.Recordset

    con.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
    rs.Open "Select * From " & "[" & strFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
    Dim i            As Integer

    Dim j            As Integer

    Dim id           As Long

    Dim TEMP         As String

    Dim SumCount     As Integer

    Dim GCHeaderFlag As Boolean

    SumCount = 0
    rs.MoveFirst
        
    GCHeaderFlag = False
        
    For i = 0 To rs.RecordCount - 1
        TEMP = ""
        
        '付值
        gcHeaderTemp.Created_By = gUserName
        gcDetailTemp.ITEM = rs.Fields(0).Value
        gcHeaderTemp.po_no = IIf(IsNull(rs.Fields(1).Value), "", rs.Fields(1).Value)
        gcHeaderTemp.SUPPLIER = rs.Fields(2).Value
        gcHeaderTemp.ShipTo = rs.Fields(3).Value
        gcHeaderTemp.Fab_Device = rs.Fields(4).Value
        gcHeaderTemp.Customer_Device = rs.Fields(5).Value
        gcHeaderTemp.GC_Version = rs.Fields(6).Value
        gcDetailTemp.Marking_Lot_ID = rs.Fields(7).Value
        gcHeaderTemp.GC_Date = rs.Fields(8).Value
        gcHeaderTemp.Lot_id = rs.Fields(9).Value
        gcDetailTemp.Lot_id = rs.Fields(9).Value
        gcDetailTemp.wafer_id = rs.Fields(10).Value
        gcDetailTemp.Good_Die_Qty = CInt(rs.Fields(11).Value)
        gcHeaderTemp.WO_NO = rs.Fields(12).Value
            
        gcHeaderTemp.TradeType = rs.Fields(15).Value
            
        '2012-11-07 jiayun 修改Good_Die_Qty 根据市场部规则
  
        'gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(gcHeaderTemp.Customer_Device, gcDetailTemp.Good_Die_Qty)
            
        '2012-11-05 jiayun 修改 GC
            
        '判断lotID在Header表中是否已存在
            
        If (JudgeGCHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.WO_NO)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
            '2013-01-11 jiayun add 客户简称
                
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
            
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "HT 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
               
        Else
            '上传到Detail表中
            
            '2012-11-05 jiayun 修改 GCT
                   
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
                   
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1
              
        End If
            
        rs.MoveNext
        
    Next i
        
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If

End Sub

Private Sub UploadSX36()

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = "36"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i          As Integer

    Dim j          As Integer

    Dim id         As Long

    Dim TEMP       As String

    Dim temp2      As String

    Dim tempVal    As String

    Dim yearpart   As String

    Dim monthpart  As String

    Dim lotpart    As String

    Dim waferpart1 As String

    Dim wfnum      As String

    Dim wfpart     As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            '            If j = 8 Then
            ''                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
            '                gcDetailTemp.Marking_Lot_ID = GetSXCodeID()
            '
            '            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j
    
        yearpart = Right(Year(Now), 1)
        monthpart = Month(Now)

        If monthpart = "10" Then
            monthpart = "A"
        ElseIf monthpart = "11" Then
            monthpart = "B"
        ElseIf monthpart = "12" Then
            monthpart = "C"

        End If

        lotpart = Mid(gcHeaderTemp.Lot_id, 3, 6)
        waferpart1 = "123456789ABCDEFGHJKLMNPQR"

        If Mid(gcDetailTemp.wafer_id, 1, 1) = 0 Then
            wfnum = Mid(gcDetailTemp.wafer_id, 2, 1)
        Else
            wfnum = gcDetailTemp.wafer_id

        End If

        wfpart = Mid(waferpart1, wfnum, 1)
        gcDetailTemp.Marking_Lot_ID = yearpart & monthpart & lotpart & wfpart
    
        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "SX 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
       
        Else
            '上传到Detail表中
           
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
           
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1
      
        End If

        yearpart = ""
        monthpart = ""
        lotpart = ""
        wfnum = ""
        wfpart = ""
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub UploadHJ()

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = "HJ"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i              As Integer

    Dim j              As Integer

    Dim id             As Long

    Dim TEMP           As String

    Dim temp2          As String

    Dim tempVal        As String
   
    Dim customerPTTemp As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
        customerPTTemp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                customerPTTemp = Trim(tempVal)
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
             
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                '                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                gcDetailTemp.Marking_Lot_ID = GetSXCodeID()
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j
    
        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, customerPTTemp)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, customerPTTemp)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "SX 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
       
        Else
            '上传到Detail表中
           
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
           
            If UCase(Trim(customerPTTemp)) = "OV02A" Then
                gcDetailTemp.Marking_Lot_ID = GetSX8CodeID(Trim(gcDetailTemp.Lot_id), Trim(gcDetailTemp.wafer_id))

            End If
           
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1
      
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub UploadSX()

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = "SX"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i              As Integer

    Dim j              As Integer

    Dim id             As Long

    Dim TEMP           As String

    Dim temp2          As String

    Dim tempVal        As String

    Dim customerPTTemp As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
        customerPTTemp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                customerPTTemp = Trim(tempVal)
                gcHeaderTemp.Customer_Device = Trim(tempVal)
          
            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = GetSXCodeID()

            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j
    
        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, customerPTTemp)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, customerPTTemp)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "SX 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
       
        Else
            '上传到Detail表中
           
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
           
            '2016-01-18 更新SX 的OV02A的MarkingCode
           
            If UCase(Trim(customerPTTemp)) = "OV02A" Then
                gcDetailTemp.Marking_Lot_ID = GetSX8CodeID(Trim(gcDetailTemp.Lot_id), Trim(gcDetailTemp.wafer_id))

            End If
           
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1
      
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub Upload59()

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = "59"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                'gcDetailTemp.Marking_Lot_ID = GetSXCodeID()
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j
    
        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "59 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
       
        Else
            '上传到Detail表中
           
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
           
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1
      
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub UploadZX()

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = "ZX"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)

            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j
    
        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "ZX 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
       
        Else
            '上传到Detail表中
           
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
           
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1
      
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub UploadOT()

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = "OT"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)

            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j
    
        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "OT 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
       
        Else
            '上传到Detail表中
           
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
           
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1
      
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub Upload95FPC()

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = "95FPC"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 3 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcHeaderTemp.po_no = Trim(tempVal)
            ElseIf j = 2 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)
            
            ElseIf j = 3 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            
            End If
        
        Next j
     
        Dim a As String
     
        id = GetMaxID()
        gcDetailTemp.wafer_id = "95FPC" & id
        Call AddGCHeader95FPC(gcHeaderTemp, id, "95") '插入主LOT号表
        GCHeaderFlag = True
           
        Call AddGCDetail95FPC(gcDetailTemp, customerTemp, id) '插入明细表
        SumCount = SumCount + 1
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub UploadRD()

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = "RD"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i           As Integer

    Dim j           As Integer

    Dim id          As Long

    Dim TEMP        As String

    Dim temp2       As String

    Dim tempVal     As String

    Dim making_code As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)

            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

                If Len(gcDetailTemp.wafer_id) < 2 Then
                    gcDetailTemp.wafer_id = "0" & gcDetailTemp.wafer_id

                End If

            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
               
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j
    
        If gcHeaderTemp.Customer_Device = "RDA2203" Then
            making_code = "B" & Right(gcHeaderTemp.Lot_id, 2) & Chr(gcDetailTemp.wafer_id + 64)
        ElseIf gcHeaderTemp.Customer_Device = "RDA2035" Then
            making_code = "C" & Right(gcHeaderTemp.Lot_id, 2) & Chr(gcDetailTemp.wafer_id + 64)
        ElseIf gcHeaderTemp.Customer_Device = "RDA2205" Or gcHeaderTemp.Customer_Device = "RDA2515" Or gcHeaderTemp.Customer_Device = "RDA2215" Or gcHeaderTemp.Customer_Device = "RDA2213" Or gcHeaderTemp.Customer_Device = "RDA2503" Then
            making_code = "RDA" & Right(gcHeaderTemp.Customer_Device, 4) & Right(gcHeaderTemp.Lot_id, 4) & gcDetailTemp.wafer_id & "ES"

        End If
    
        gcDetailTemp.Marking_Lot_ID = making_code
    
        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "RD 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
       
        Else
            '上传到Detail表中
           
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
           
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1
      
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub UploadDN()

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    Dim dnRemark             As String

    customerTemp = "DN"
    dnRemark = ""

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        dnRemark = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)

            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 14 Then
                dnRemark = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j

        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "RD 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
       
        Else
            '上传到Detail表中
           
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
           
            Call AddDNDetail(gcDetailTemp, customerTemp, id, dnRemark)
            SumCount = SumCount + 1
      
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub UploadPT()

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = "PT"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)

            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j
    
        If (JudgePTHeaderId(gcHeaderTemp.Lot_id)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "PT 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
       
        Else
            '上传到Detail表中
           
            '           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_ID), 2)
            '2013-03-04 jiayun modify
            gcDetailTemp.ITEM = gcDetailTemp.wafer_id
           
            gcDetailTemp.wafer_id = Right$(Trim(gcDetailTemp.wafer_id), 2)
           
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1
      
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub UploadBD()

    '2013-06-17 jiayun add BD
    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = "BD"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i              As Integer

    Dim j              As Integer

    Dim id             As Long

    Dim TEMP           As String

    Dim temp2          As String

    Dim tempVal        As String
   
    Dim PShortNameTemp As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        PShortNameTemp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)

            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 14 Then
                PShortNameTemp = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j
    
        '2013-12-05 jiayun add 校验po号是否为空
    
        If Trim(gcHeaderTemp.po_no) = "" Then
            MsgBox "PO_NO不允许为空值，请确认！", vbInformation, "提示"
            Exit Sub
    
        End If
    
        If (JudgePTHeaderId(gcHeaderTemp.Lot_id)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddBDHeader(gcHeaderTemp, id, customerTemp, PShortNameTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "BD 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
       
        Else
            '上传到Detail表中
           
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
           
            '2013-03-04 jiayun modify
            '           gcDetailTemp.item = gcDetailTemp.Wafer_ID
           
            gcDetailTemp.wafer_id = Right$(Trim(gcDetailTemp.wafer_id), 2)
           
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1
      
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub UploadSY()

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = "SY"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)

            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
        
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
            
        Next j
    
        If (JudgePTHeaderId(gcHeaderTemp.Lot_id)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "PT 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
       
        Else
            '上传到Detail表中
           
            '           gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_ID), 2)
            '2013-03-04 jiayun modify
            gcDetailTemp.ITEM = gcDetailTemp.wafer_id
           
            gcDetailTemp.wafer_id = Right$(Trim(gcDetailTemp.wafer_id), 2)
           
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1
      
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub UploadSX34()

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = "34"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                '                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                gcDetailTemp.Marking_Lot_ID = GetSXCodeID()
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If
            
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
        
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j
    
        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "SX 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
       
        Else
            '上传到Detail表中
           
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
           
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1
      
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub UploadSX32()

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    customerTemp = "32"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
    
    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 16 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    Dim i       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim TEMP    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        '查询一行的值
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

            TEMP = ""
        
            '付值
            gcHeaderTemp.Created_By = gUserName

            If j = 1 Then
                gcDetailTemp.ITEM = Trim(tempVal)

            End If
            
            If j = 2 Then
                gcHeaderTemp.po_no = Trim(tempVal)

            End If
            
            If j = 3 Then
                gcHeaderTemp.SUPPLIER = Trim(tempVal)

            End If
            
            If j = 4 Then
                gcHeaderTemp.ShipTo = Trim(tempVal)

            End If
            
            If j = 5 Then
                gcHeaderTemp.Fab_Device = Trim(tempVal)

            End If
            
            If j = 6 Then
                gcHeaderTemp.Customer_Device = Trim(tempVal)

            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then
                '                gcDetailTemp.Marking_Lot_ID = Trim(tempVal)
                gcDetailTemp.Marking_Lot_ID = GetSXCodeID()
             
            End If
            
            If j = 9 Then
                gcHeaderTemp.GC_Date = Trim(tempVal)

            End If
            
            If j = 10 Then
                gcHeaderTemp.Lot_id = Trim(tempVal)
                gcDetailTemp.Lot_id = Trim(tempVal)

            End If
            
            If j = 11 Then
                gcDetailTemp.wafer_id = Trim(tempVal)

            End If
            
            '            If j = 12 Then
            '                gcDetailTemp.Good_Die_Qty = Trim(tempVal)
            '            End If
            '
            '            If j = 13 Then
            '                gcHeaderTemp.WO_NO = Trim(tempVal)
            '            End If
            '
            '
            If j = 12 Then
                gcDetailTemp.Good_Die_Qty = Trim(tempVal)

            End If
            
            If j = 13 Then
                gcDetailTemp.NG_Die_Qty = CLng(Trim(tempVal)) - gcDetailTemp.Good_Die_Qty

            End If
            
            If j = 14 Then
                gcHeaderTemp.WO_NO = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)

            End If
        
        Next j
    
        If (JudgeSXHeaderId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)) Then
            
            If GCHeaderFlag = False Then

                '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
            End If
                
            id = GetSXLotIDPOId(gcHeaderTemp.Lot_id, gcHeaderTemp.po_no, gcHeaderTemp.Customer_Device)
                
        Else
            '上传到Header表中
            '取目前DB最大的ID号
            id = GetMaxID()
       
            Call AddGCHeader(gcHeaderTemp, id, customerTemp)
            GCHeaderFlag = True
              
        End If
            
        '判断lotID在Detail表中是否已存在
    
        If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id)) Then
            MsgBox "SX 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
       
        Else
            '上传到Detail表中
           
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)
           
            Call AddGCDetail(gcDetailTemp, customerTemp, id)
            SumCount = SumCount + 1
      
        End If
     
        '上传到DB
NextRecord2:

    Next i
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
        
End Sub

Private Sub Command7_Click()

    If ComCbBand.text = "" Then
        MsgBox "请选择保税非保税", vbInformation, "提示"
        Exit Sub
    End If
    
    If ComCbBand.ListIndex = 0 Then
        gTax = "A"
    Else
        gTax = "B"
    End If

    If CmbCustomer.text = "" Then
        MsgBox "请先选择客户！"
        Exit Sub

    End If
    If chkMsgAppend.Value = 1 And txtMsg.text = "" Then
        MsgBox "请填入邮件正文补充信息,否则请取消复选框选项" & vbCrLf & "{填入:快递单号等信息或补充事项说明" & vbCrLf & "联系电话....", vbInformation, "提醒"
        Exit Sub

    End If
    If CmbCustomer.text = "GC" Then
        UploadGCNew

    ElseIf CmbCustomer.text = "GC_BUMPING" Then
        UploadGCBP

    ElseIf CmbCustomer.text = "GC_WLD/T" Then
        UploadGCNewWLDT

    ElseIf CmbCustomer.text = "SX" Then
        UploadSX

    ElseIf CmbCustomer.text = "HJ" Then
        UploadHJ

    ElseIf CmbCustomer.text = "59" Then
        Upload59

    ElseIf CmbCustomer.text = "36" Then
        UploadSX36

    ElseIf CmbCustomer.text = "34" Then
        UploadSX34

    ElseIf CmbCustomer.text = "32" Then
        UploadSX32

    ElseIf CmbCustomer.text = "PT" Then
        UploadPT

    ElseIf CmbCustomer.text = "SY" Then
        UploadSY

    ElseIf CmbCustomer.text = "RD" Then
        UploadRD

    ElseIf CmbCustomer.text = "DN" Then
        UploadDN

    ElseIf CmbCustomer.text = "BD" Then
        UploadBD

    ElseIf CmbCustomer.text = "ZX" Then
        UploadZX

    ElseIf CmbCustomer.text = "HY" Then
        UploadHY

    ElseIf CmbCustomer.text = "HT" Then
        UploadHT

    ElseIf CmbCustomer.text = "OT" Then
        UploadOT

    ElseIf CmbCustomer.text = "MC" Then
        UploadMC

    ElseIf CmbCustomer.text = "GT" Then
        Call UploadNormalCustomer("GT")

    ElseIf CmbCustomer.text = "MG" Then
        Call UploadNormalCustomer("MG")

    ElseIf CmbCustomer.text = "LX" Then
        Call UploadNormalCustomer("LX")

    ElseIf CmbCustomer.text = "HH" Then
        Call UploadNormalCustomer("HH")

    ElseIf CmbCustomer.text = "CN" Then
        Call UploadNormalCustomer("CN")

    ElseIf CmbCustomer.text = "KT" Then
        Call UploadNormalCustomer("KT")

    ElseIf CmbCustomer.text = "HD" Then
        Call UploadNormalCustomer("HD")

    ElseIf CmbCustomer.text = "RS" Then
        Call UploadNormalCustomer("RS")

    ElseIf CmbCustomer.text = "AM" Then
        Call UploadNormalCustomer("AM")

    ElseIf CmbCustomer.text = "ZL" Then
        Call UploadNormalCustomerZL("ZL")

    ElseIf CmbCustomer.text = "SD" Then
        Call UploadNormalCustomer("SD")

    ElseIf CmbCustomer.text = "RO" Then
        Call UploadNormalCustomer("RO")

    ElseIf CmbCustomer.text = "YW" Then
        Call UploadNormalCustomer("YW")

    ElseIf CmbCustomer.text = "MR" Then
        Call UploadNormalCustomer("MR")

    ElseIf CmbCustomer.text = "XA" Then
        Call UploadNormalCustomer("XA")

    ElseIf CmbCustomer.text = "37" Then
        Call UploadNormalCustomer("37")

    ElseIf CmbCustomer.text = "69" Then
        Call UploadNormalCustomer("69")

    ElseIf CmbCustomer.text = "80" Then
        Call UploadNormalCustomer("80")

    ElseIf CmbCustomer.text = "81" Then
        Call UploadNormalCustomer("81")

    ElseIf CmbCustomer.text = "87" Then
        Call UploadNormalCustomer("87")

    ElseIf CmbCustomer.text = "88" Then
        Call UploadNormalCustomer("88")

    ElseIf CmbCustomer.text = "77" Then
        Call UploadNormalCustomer77("77")

    ElseIf CmbCustomer.text = "64" Then
        Call UploadNormalCustomer("64")

    ElseIf CmbCustomer.text = "79" Then
        Call UploadNormalCustomer("79")

    ElseIf CmbCustomer.text = "78" Then
        Call UploadNormalCustomer("78")

    ElseIf CmbCustomer.text = "68" Then
        Call UploadMPSCustomer("68")

    ElseIf CmbCustomer.text = "BJ128" Then
        Call UploadMPSCustomer("BJ128")

    ElseIf CmbCustomer.text = "HK006" Then
        Call UploadMPSCustomer("HK006")

    ElseIf CmbCustomer.text = "70" Then
        Call UploadMPSCustomer("70")

    ElseIf CmbCustomer.text = "45" Then
        Call UploadNormalCustomer("45")

    ElseIf CmbCustomer.text = "50" Then
        Call UploadNormalCustomer("50")

    ElseIf CmbCustomer.text = "56" Then
        Call UploadNormalCustomer56("56")

    ElseIf CmbCustomer.text = "49" Then
        Call UploadNormalCustomer("49")

    ElseIf CmbCustomer.text = "XW" Then
        Call UploadNormalCustomer("XW")

    ElseIf CmbCustomer.text = "B1" Then
        Call UploadNormalCustomer("B1")

    ElseIf CmbCustomer.text = "SL" Then
        Call UploadNormalCustomer("SL")

    ElseIf CmbCustomer.text = "30" Then
        Call UploadNormalCustomer("30")

    ElseIf CmbCustomer.text = "33" Then
        Call UploadNormalCustomer("33")

    ElseIf CmbCustomer.text = "57" Then
        Call UploadNormalCustomer("57")

    ElseIf CmbCustomer.text = "94" Then
        Call UploadNormalCustomer("94")

    ElseIf CmbCustomer.text = "93" Then
        Call UploadNormalCustomer("93")

    ElseIf CmbCustomer.text = "95" Then
        Call UploadNormalCustomer("95")

    ElseIf CmbCustomer.text = "95FPC" Then
        Call Upload95FPC

    ElseIf CmbCustomer.text = "55" Then
        Call UploadNormalCustomer56("55")

    ElseIf CmbCustomer.text = "54" Then
        Call UploadNormalCustomer("54")

    ElseIf CmbCustomer.text = "60" Then
        Call UploadNormalCustomer("60")

    ElseIf CmbCustomer.text = "61" Then
        Call UploadNormalCustomer("61")

    ElseIf CmbCustomer.text = "YX" Then
        Call UploadNormalCustomer("YX")

    ElseIf CmbCustomer.text = "QR" Then
        Call UploadQR("QR")

    ElseIf CmbCustomer.text = "QR2" Then
        Call UploadQRV2("QR")

    ElseIf CmbCustomer.text = "GD" Then
        Call UploadNormalCustomer("GD")

    ElseIf CmbCustomer.text = "EQ" Then
        UploadEQ

        '2015-03-18 jiayun add
    ElseIf CmbCustomer.text = "EQ_IS" Then
        UploadEQ_IS
    ElseIf CmbCustomer.text = "EQ_ShippingRequest" Then
        UploadEQ_ShippingRequest

    ElseIf CmbCustomer.text = "CS" Then
        Call UploadNormalCustomerCS("CS")

    Else

    End If
    
    
    
    
    
    

End Sub

Private Function GetGCGoodDieQty(productNameTemp As String, dieQtyTemp As Long) As Integer
    GetGCGoodDieQty = 0

    Set updateRS = GetWO_GC_Die(productNameTemp)

    If updateRS.RecordCount > 0 Then
        GetGCGoodDieQty = CInt(updateRS.Fields("dieqty").Value)

    End If

End Function


Private Function GetGCVerLastChar(ptTemp As String) As String
    '2013-12-26 jiayun add
    '根据Gc pt 查询数量

    GetGCVerLastChar = ""

    Set updateRS = GetWO_GC_Ver(ptTemp)

    If updateRS.RecordCount > 0 Then

        GetGCVerLastChar = CStr(updateRS.Fields("Gcversion").Value)

    Else

        GetGCVerLastChar = ""

    End If

End Function

Private Sub Command8_Click()

    If CmbCustomer.text = "" Then
        MsgBox "请先选择客户！"
        Exit Sub

    End If

    If CmbCustomer.text = "EQ_IS" Then

        ExporToExcel ("  select po_num as PO_NO, ship_site as Supplier,test_site as Ship_To, fab_conv_id as FAB_Device, mpn_desc as Customer_Device," & " imager_customer_rev as GC_Version,created_date as GC_Date,source_batch_id  as Lot_ID, mtrl_num   As WO_NO , probe_ship_part_type as 贸易类型 " & " From CustomerOItbl_test  where CustomerShortName = 'EQ' AND MTRL_DESC LIKE 'IS%'order by id ")
               
    ElseIf CmbCustomer.text = "EQ_ShippingRequest" Then

        ExporToExcel ("SELECT * FROM EQ_SHIPPING_REQUEST")
 
    ElseIf CmbCustomer.text = "95FPC" Then
        ExporToExcel ("  select po_num as PO_NO, ship_site as Supplier,test_site as Ship_To, fab_conv_id as FAB_Device, mpn_desc as Customer_Device," & " imager_customer_rev as GC_Version,created_date as GC_Date,source_batch_id  as Lot_ID, mtrl_num   As WO_NO , probe_ship_part_type as 贸易类型 " & " From CustomerOItbl_test a where CustomerShortName = '95'   and a.source_batch_id = '95FPC' order by id ")

    End If

End Sub

Private Sub Command9_Click()

    If CmbCustomer.text = "" Then
        MsgBox "请先选择客户！"
        Exit Sub

    End If

    If CmbCustomer.text = "GC_WLD/T" Then

        ExporToExcel (" select substrateid as Item ,productid as Marking_Lot_ID ,lotid as Lot_ID ,wafer_id ,passbincount as Good_Die_Qty " & " from  mappingDataTest where  CustomerShortName = '" & CmbCustomer.text & "' and remark='WLT' order by id")
 
    ElseIf CmbCustomer.text = "EQ_IS" Then

        ExporToExcel (" select substrateid as Item ,productid as Marking_Lot_ID ,lotid as Lot_ID ,wafer_id ,passbincount as Good_Die_Qty " & " from  mappingDataTest where  CustomerShortName = 'EQ'  order by id")
               
    ElseIf CmbCustomer.text = "95FPC" Then

        ExporToExcel (" select substrateid as Item ,productid as Marking_Lot_ID ,lotid as Lot_ID ,wafer_id ,passbincount as Good_Die_Qty " & " from  mappingDataTest a where  CustomerShortName = '95' and  a.lotid = '95FPC' order by id")

    End If

    ' ExporToExcel (" select substrateid as Item ,productid as Marking_Lot_ID ,lotid as Lot_ID ,wafer_id ,passbincount as Good_Die_Qty " & _
    '               " from  mappingDataTest where  CustomerShortName = '" & CmbCustomer.Text & "' order by id")
 
End Sub

Private Sub Form_Load()

    Com.flags = &H80200

    ComSI.flags = &H80200

    CmbCustomer.AddItem ("GC")
    CmbCustomer.AddItem ("GC_WLD/T")
    CmbCustomer.AddItem ("GC_BUMPING")
    'CmbCustomer.AddItem ("SX")
    CmbCustomer.AddItem ("HJ")

    CmbCustomer.AddItem ("PT")
    CmbCustomer.AddItem ("SY")
    CmbCustomer.AddItem ("RD")
    CmbCustomer.AddItem ("DN")
    CmbCustomer.AddItem ("BD")
    CmbCustomer.AddItem ("BJ128")
    CmbCustomer.AddItem ("ZX")
    CmbCustomer.AddItem ("HY")
    CmbCustomer.AddItem ("HT")
    CmbCustomer.AddItem ("OT")
    CmbCustomer.AddItem ("MC")
    '2014-09-17 jiayun modify si 改为GT
    CmbCustomer.AddItem ("GT")

    CmbCustomer.AddItem ("CN")
    CmbCustomer.AddItem ("KT")
    CmbCustomer.AddItem ("HD")

    CmbCustomer.AddItem ("RS")
    CmbCustomer.AddItem ("SD")

    CmbCustomer.AddItem ("QR")
    CmbCustomer.AddItem ("QR2")

    CmbCustomer.AddItem ("MG")
    CmbCustomer.AddItem ("LX")
    CmbCustomer.AddItem ("GD")
    CmbCustomer.AddItem ("AM")
    CmbCustomer.AddItem ("EQ")
    CmbCustomer.AddItem ("EQ_IS")
    CmbCustomer.AddItem ("EQ_ShippingRequest")
    CmbCustomer.AddItem ("ZL")
    CmbCustomer.AddItem ("YW")
    CmbCustomer.AddItem ("RO")
    CmbCustomer.AddItem ("MR")
    CmbCustomer.AddItem ("CS")

    CmbCustomer.AddItem ("36")
    CmbCustomer.AddItem ("34")
    CmbCustomer.AddItem ("33")

    CmbCustomer.AddItem ("32")
    CmbCustomer.AddItem ("45")
    CmbCustomer.AddItem ("50")
    CmbCustomer.AddItem ("60")

    CmbCustomer.AddItem ("30")
    CmbCustomer.AddItem ("55")
    CmbCustomer.AddItem ("54")
    CmbCustomer.AddItem ("56")
    CmbCustomer.AddItem ("57")
    CmbCustomer.AddItem ("49")
    CmbCustomer.AddItem ("59")
    CmbCustomer.AddItem ("64")
    CmbCustomer.AddItem ("61")

    CmbCustomer.AddItem ("68")
    CmbCustomer.AddItem ("HK006")
    CmbCustomer.AddItem ("70")
    CmbCustomer.AddItem ("69")
    CmbCustomer.AddItem ("80")
    CmbCustomer.AddItem ("81")
    CmbCustomer.AddItem ("87")
    CmbCustomer.AddItem ("88")
    CmbCustomer.AddItem ("94")
    CmbCustomer.AddItem ("93")
    CmbCustomer.AddItem ("95")
    CmbCustomer.AddItem ("95FPC") '委外西安
    CmbCustomer.AddItem ("B1")
    CmbCustomer.AddItem ("TW058")

    CmbCustomer.AddItem ("XW")

    CmbCustomer.AddItem ("YX")

    CmbCustomer.AddItem ("37")
    CmbCustomer.AddItem ("77")
    CmbCustomer.AddItem ("78")

    CmbCustomer.AddItem ("XA")
    CmbCustomer.AddItem ("HH")
    CmbCustomer.AddItem ("SL")

    Combo1.AddItem ("AA")
    Combo1.AddItem ("自购")
    Combo1.AddItem ("CN")
    
    InitData

End Sub


Private Sub SentMesToPMC(dT As GCHeader)
'发送邮件给计划部
Dim strSentTo(100) As String
Dim strSentCC(20)  As String
Dim strSentTitle   As String
Dim strSentText    As String
Dim dirtemp        As String
Dim strTemp        As String
Dim i              As Integer
Dim strBand        As String
Dim CUSTOMER_CODE  As String
Dim REMARK As String

If gUserName = "07885" Then

     Exit Sub
End If

If CmbCustomer.text = "GC" Or CmbCustomer.text = "GC_BUMPING" Or CmbCustomer.text = "GC_WLD/T" Then
    CUSTOMER_CODE = "GC"
Else
    CUSTOMER_CODE = CmbCustomer.text
End If


'If bBonded = True Then
'    strBand = "保税"
'Else
'    strBand = "非保税"'

'End If

i = 0
dirtemp = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\邮件接收\SentTo_Upload.cfg"

strSentTitle = "已上传订单," & ComCbBand.text & ",客户代码:" & CUSTOMER_CODE & ",客户机种:" & dT.Customer_Device & ",请注意接收"
    If CmbCustomer.text = "GC_WLD/T" Then
        REMARK = "WLA回货"
    ElseIf CmbCustomer.text = "GC_BUMPING" Then
        REMARK = "Bumping上传"
    Else
        REMARK = ""
    End If
    
strSentText = "内勤部" & strRealName & ",工号:" & gUserName & ", " & REMARK & "  已上传订单," & ComCbBand.text & ",客户代码:" & CUSTOMER_CODE & ",客户机种:" & dT.Customer_Device & " 明细见附件" & vbCrLf
strSentText = strSentText & txtMsg.text

Open dirtemp For Input As #1

Do While Not EOF(1)
    Line Input #1, strTemp
    strSentTo(i) = Trim$(strTemp)
    i = i + 1
Loop
Close #1

strSentCC(0) = "Sharon.dong_ks@ht-tech.com"
If SentMes(strSentTitle, strSentText, strSentTo, strFileName, strSentCC) = True Then
    MsgBox "邮件已发送", vbInformation, Me.Caption
Else
    MsgBox "邮件发送失败", vbCritical, Me.Caption

End If

End Sub


Private Sub InitData()
Dim strSql As String
gUpID = ""
gBackID = ""

Select Case gUserName

    Case "15507"
        strRealName = "王丹"

    Case "16452"
        strRealName = "顾妍"

    Case "18035"
        strRealName = "潘葆芸"

    Case "7433", "07433"
        strRealName = "刘璐"

    Case "14117"
        strRealName = "蒋芹"

    Case "8240", "08240"
        strRealName = "刘忻"

    Case "16368"
        strRealName = "刘明"

    Case "12089"
        strRealName = "张强"

    Case "12725"
        strRealName = "全娟敏"

    Case "15236"
        strRealName = "何丹萍"

    Case "16642"
        strRealName = "吴芳"

    Case "07885"
        strRealName = "管理员"

    Case "18420"
        strRealName = "徐晴奋"

    Case "18697"
        strRealName = "王媛"

    Case "18252"
        strRealName = "宋美娟"

    Case "18881"
        strRealName = "吉滢铭"

    Case "19400"
        strRealName = "黄婷"
    
End Select

strSql = "select EmpName from XTW..employee where empno = '" & gUserName & "'"
strRealName = Get_SqlStr2(strSql)

End Sub


Private Sub AddGCHeader_new(mapTemp As GCHeader, id As Long, customerTemp As String, gUpID As String)
'增加抛到SqlServer
Dim cmdStr As String
Dim cmdStr2 As String
                                   
'On Error GoTo DealError
         
Cnn.BeginTrans

cmdStr = "insert into CustomerOItbl_test(id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code,probe_ship_part_type,jobno, chromaticity, reliability_sampling, lot_priority, wafer_box_type,WAFER_VISUAL_INSPECT,Flow,MTRL_DESC) values (" & _
" " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.ShipTo & "','" & mapTemp.Fab_Device & "'," & _
"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & gUserName & "',sysdate,'" & mapTemp.Ship_Out & "','" & mapTemp.TradeType & "','" & mapTemp.TAXTYPE & "', '" & mapTemp.SecondFlag & "', '" & mapTemp.LotProperty & "', '" & mapTemp.LotOwner & "', '" & mapTemp.Telephone & "', '" & gUpID & "','" & mapTemp.Flow & "','" & mapTemp.htdevice & "')"


cmdStr2 = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code,probe_ship_part_type,jobno,chromaticity,reliability_sampling, lot_priority, wafer_box_type,WAFER_VISUAL_INSPECT,Flow,MTRL_DESC) values (" & _
" " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.ShipTo & "','" & mapTemp.Fab_Device & "'," & _
"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & gUserName & "',GETDATE(),'" & mapTemp.Ship_Out & "' ,'" & mapTemp.TradeType & "','" & mapTemp.TAXTYPE & "', '" & mapTemp.SecondFlag & "', '" & mapTemp.LotProperty & "', '" & mapTemp.LotOwner & "', '" & mapTemp.Telephone & "', '" & gUpID & "','" & mapTemp.Flow & "','" & mapTemp.htdevice & "')"


                                               
AddSql (cmdStr)
AddSql2 (cmdStr2)
 
Cnn.CommitTrans

'Exit Sub
'
'DealError:
'
'Cnn.RollbackTrans

End Sub


Private Function ExportExcel(dT As GCHeader, customerTemp As String) As Boolean

On Error GoTo Ert

Dim xlsApp     As Excel.Application
Dim xlsBook    As Excel.Workbook
Dim xlsSheet   As Excel.Worksheet
Dim i          As Long
Dim j          As Long
Dim iCnt       As Integer
Dim strFileSeq As String, strPartName As String
Dim rs         As New ADODB.Recordset
Dim REMARK     As String
Dim strWOPath As String
Dim strWOPath_Loc As String


ExportExcel = False
Set rs.ActiveConnection = OraConnect
    
    If CmbCustomer.text = "GC_WLD/T" Then '回货WO
        REMARK = "WLA"
  
        rs.Source = "select row_number() over(ORDER BY  bb.lotid,bb.substrateid) as 序号,case bb.substratetype when 'A' then '保税' else '非保税' end as 是否保税, bb.customershortname as 客户代码, " & _
           "       aa.Fab_conv_id as FAB机种,aa.mpn_desc as 客户机种," & _
           "       '' as NPI负责人员, '' as 厂内机种, " & _
           "       aa.po_num as PO_NUM, " & _
           "       bb.lotid as LOT_ID, " & _
           "       bb.wafer_id as WAFER_NO, " & _
           "       bb.substrateid as WAFERID, " & _
           "       bb.passbincount as GOOD_DIES, " & _
           "       bb.failbincount as NG_DIES, " & _
           "       bb.passbincount + bb.failbincount as GROSS_DIES, " & _
           "       bb.productid as Marking, " & _
           "       aa.Imager_Customer_Rev as 二级代码, " & _
           "       bb.qtech_created_by as 上传人员,bb.qtech_created_date as 上传时间,  bb.qtech_lastupdate_by as 更新人员, bb.qtech_lastupdate_date as 更新时间 ,'" & REMARK & "' as WLA,aa.flow as 形式,aa.MTRL_DESC as 厂内机种" & _
           "  from customeroitbl_test aa " & _
           " inner join mappingdatatest bb " & _
           "    on to_char(aa.id) = bb.filename " & _
           "   and aa.wafer_visual_inspect = '" & gUpID & "' and aa.customershortname = '" & customerTemp & "' " & _
           "   group by  bb.customershortname,aa.Fab_conv_id, aa.mpn_desc,aa.po_num,bb.lotid,bb.wafer_id,bb.substrateid,bb.passbincount,bb.failbincount,bb.productid,aa.Imager_Customer_Rev ,bb.substratetype,bb.qtech_created_by,bb.qtech_created_date,bb.qtech_lastupdate_by,bb.qtech_lastupdate_date,aa.flow ,aa.MTRL_DESC"

    Else '首次WO
         REMARK = ""
        rs.Source = "select row_number() over(ORDER BY  bb.lotid,bb.substrateid) as 序号,case bb.substratetype when 'A' then '保税' else '非保税' end as 是否保税, bb.customershortname as 客户代码, " & _
          "       aa.Fab_conv_id as FAB机种,aa.mpn_desc as 客户机种," & _
          "       '' as NPI负责人员, '' as 厂内机种, " & _
          "       aa.po_num as PO_NUM, " & _
          "       bb.lotid as LOT_ID, " & _
          "       bb.wafer_id as WAFER_NO, " & _
          "       bb.substrateid as WAFERID, " & _
          "       bb.passbincount as GOOD_DIES, " & _
          "       bb.failbincount as NG_DIES, " & _
          "       bb.passbincount + bb.failbincount as GROSS_DIES, " & _
          "       bb.productid as Marking, " & _
          "       aa.Imager_Customer_Rev as 二级代码, " & _
          "       bb.qtech_created_by as 上传人员,bb.qtech_created_date as 上传时间,  bb.qtech_lastupdate_by as 更新人员, bb.qtech_lastupdate_date as 更新时间 ,'" & REMARK & "' as WLA,aa.flow as 形式" & _
          "  from customeroitbl_test aa " & _
          " inner join mappingdatatest bb " & _
          "    on to_char(aa.id) = bb.filename " & _
          "   and aa.wafer_visual_inspect = '" & gUpID & "' and aa.customershortname = '" & customerTemp & "' " & _
          "   group by  bb.customershortname,aa.Fab_conv_id, aa.mpn_desc,aa.po_num,bb.lotid,bb.wafer_id,bb.substrateid,bb.passbincount,bb.failbincount,bb.productid,aa.Imager_Customer_Rev ,bb.substratetype,bb.qtech_created_by,bb.qtech_created_date,bb.qtech_lastupdate_by,bb.qtech_lastupdate_date,aa.flow "
   
    End If
    
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount = 0 Then
    MsgBox "查询不到订单信息, 此次上传失败, 请重新确认,再次上传", vbCritical, "警告"
    Exit Function

End If

iCnt = rs.RecordCount
Set xlsApp = CreateObject("Excel.Application")
Set xlsBook = xlsApp.Workbooks.Add
Set xlsSheet = xlsBook.Worksheets(1)
xlsSheet.name = "WO"
With xlsApp
    .Rows(1).Font.Bold = True

End With

For j = 1 To rs.Fields.count
    xlsSheet.Cells(1, j) = ("" & rs(j - 1).name)
Next
rs.MoveFirst

For i = 2 To rs.RecordCount + 1
    For j = 1 To rs.Fields.count
        xlsSheet.Cells(i, j) = ("" & rs(j - 1).Value)
    Next j

    rs.MoveNext
Next i

rs.Close
'---------------------
rs.Source = "select row_number() over(ORDER BY  bb.lotid) as 序号, '' AS 料号,  bb.customershortname as 客户 ,'' as 型号,  bb.lotid AS  LOT,count(distinct bb.wafer_id) AS 数量    " & _
" from customeroitbl_test aa inner join mappingdatatest bb on to_char(aa.id) = bb.filename  " & _
" and aa.wafer_visual_inspect = '" & gUpID & "' and aa.customershortname = '" & customerTemp & "' " & _
" group by  bb.customershortname,bb.lotid "
 
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText



If xlsBook.Worksheets.count = 1 Then
xlsBook.Worksheets.Add after:=xlsBook.Worksheets(1)
End If
Set xlsSheet = xlsBook.Worksheets(2)
xlsSheet.name = "标签"
With xlsApp
    .Rows(1).Font.Bold = True

End With

For j = 1 To rs.Fields.count
    xlsSheet.Cells(1, j) = ("" & rs(j - 1).name)
Next
rs.MoveFirst
For i = 2 To rs.RecordCount + 1
    For j = 1 To rs.Fields.count
        xlsSheet.Cells(i, j) = ("" & rs(j - 1).Value)
    Next j

    rs.MoveNext
Next i

rs.Close


'---------------------
Set rs = Nothing
xlsBook.Worksheets(1).Activate
xlsApp.Visible = True

strWOPath_Loc = "D:\已上传WO"
If Dir(strWOPath_Loc, vbDirectory) = "" Then
    MkDir strWOPath_Loc
End If

strFileName = strWOPath_Loc & "\" & customerTemp & "_" & iCnt & "片" & "-" & Format(Now, "YYYYMMDD-HHMMSS") & ".xlsx"
xlsBook.SaveAs strFileName
Set xlsApp = Nothing

ExportExcel = True


strWOPath = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\已上传\" & customerTemp
If Dir(strWOPath, vbDirectory) = "" Then
    MkDir strWOPath

End If

strWOPath = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\已上传\" & customerTemp & "\" & Replace(dT.Customer_Device, "/", "")
If Dir(strWOPath, vbDirectory) = "" Then
    MkDir strWOPath

End If

Call CopyFileToFtp(Text3.text, strWOPath & "\")
Call CopyFileToFtp(strFileName, strWOPath & "\")

Exit Function
Ert:
If Not (xlsApp Is Nothing) Then
    Set xlsApp = Nothing

End If

End Function
























