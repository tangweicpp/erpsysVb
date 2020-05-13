VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm37Portal 
   Caption         =   "SemTech Portal TX"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   12960
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   $"Frm37Portal.frx":0000
      TabPicture(0)   =   "Frm37Portal.frx":0013
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fps(0)"
      Tab(0).Control(1)=   "CmdWaferRecQuery"
      Tab(0).Control(2)=   "CmdWaferRecOut"
      Tab(0).Control(3)=   "ChkAll"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "PO commit"
      TabPicture(1)   =   "Frm37Portal.frx":002F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(1)=   "CmdPOComit"
      Tab(1).Control(2)=   "CmdPOComitOut"
      Tab(1).Control(3)=   "fps(1)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "PO Start"
      TabPicture(2)   =   "Frm37Portal.frx":004B
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fps(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "CmdPOStart"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "CmdPOStartOut"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "PO Complete"
      TabPicture(3)   =   "Frm37Portal.frx":0067
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label4"
      Tab(3).Control(1)=   "Label3"
      Tab(3).Control(2)=   "fps(3)"
      Tab(3).Control(3)=   "CmdPOCompleteOut"
      Tab(3).Control(4)=   "CmdPOComplete"
      Tab(3).Control(5)=   "Txt37BillNo"
      Tab(3).Control(6)=   "CmbPoType"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "PO Ship"
      TabPicture(4)   =   "Frm37Portal.frx":0083
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ComboShipICI"
      Tab(4).Control(1)=   "CmdPOShipOut"
      Tab(4).Control(2)=   "CmdPOShip"
      Tab(4).Control(3)=   "Txt37BillNoShip"
      Tab(4).Control(4)=   "fps(4)"
      Tab(4).Control(5)=   "Label5"
      Tab(4).Control(6)=   "Label1"
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "PO Recommit"
      TabPicture(5)   =   "Frm37Portal.frx":009F
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command2"
      Tab(5).Control(1)=   "Command1"
      Tab(5).Control(2)=   "fps(5)"
      Tab(5).ControlCount=   3
      Begin VB.ComboBox ComboShipICI 
         Height          =   315
         ItemData        =   "Frm37Portal.frx":00BB
         Left            =   -73440
         List            =   "Frm37Portal.frx":00C5
         Style           =   2  'Dropdown List
         TabIndex        =   140
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox CmbPoType 
         Height          =   315
         ItemData        =   "Frm37Portal.frx":00D7
         Left            =   -73440
         List            =   "Frm37Portal.frx":00E1
         Style           =   2  'Dropdown List
         TabIndex        =   138
         Top             =   600
         Width           =   1695
      End
      Begin VB.Frame Frame7 
         Caption         =   "Commit文件上传厂内"
         Height          =   1095
         Left            =   -74640
         TabIndex        =   128
         Top             =   1080
         Width           =   10455
         Begin VB.CommandButton CmdPOComOut 
            Caption         =   "导出"
            Height          =   465
            Left            =   8520
            TabIndex        =   142
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            Height          =   495
            Left            =   3720
            MultiLine       =   -1  'True
            TabIndex        =   135
            Top             =   6300
            Width           =   4935
         End
         Begin VB.CommandButton Command20 
            Caption         =   ".."
            Height          =   495
            Left            =   9000
            TabIndex        =   134
            Top             =   6300
            Width           =   375
         End
         Begin VB.CommandButton Command21 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   4080
            TabIndex        =   133
            Top             =   7140
            Width           =   1335
         End
         Begin VB.CommandButton Command22 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   6960
            TabIndex        =   132
            Top             =   7140
            Width           =   1335
         End
         Begin VB.CommandButton Command27 
            Caption         =   "上传DB"
            Height          =   465
            Left            =   7080
            TabIndex        =   131
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton Command28 
            Caption         =   ".."
            Height          =   495
            Left            =   6000
            TabIndex        =   130
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox Text8 
            Enabled         =   0   'False
            Height          =   375
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   129
            Top             =   480
            Width           =   5295
         End
         Begin MSComDlg.CommonDialog CommonDialog5 
            Left            =   5880
            Top             =   5700
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComDlg.CommonDialog CommonDialog7 
            Left            =   4920
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            MaxFileSize     =   10000
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的CSV："
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   6
            Left            =   3720
            TabIndex        =   137
            Top             =   5940
            Width           =   1545
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的csv："
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   8
            Left            =   600
            TabIndex        =   136
            Top             =   240
            Width           =   1500
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "导出"
         Height          =   360
         Left            =   -71520
         TabIndex        =   126
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "查询"
         Height          =   360
         Left            =   -74280
         TabIndex        =   125
         Top             =   600
         Width           =   990
      End
      Begin VB.CheckBox ChkAll 
         Height          =   255
         Left            =   -62760
         TabIndex        =   123
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton CmdPOShipOut 
         Caption         =   "导出"
         Height          =   360
         Left            =   -63360
         TabIndex        =   120
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton CmdPOShip 
         Caption         =   "查询"
         Height          =   360
         Left            =   -65160
         TabIndex        =   119
         Top             =   600
         Width           =   990
      End
      Begin VB.TextBox Txt37BillNoShip 
         Height          =   375
         Left            =   -69600
         TabIndex        =   118
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox Txt37BillNo 
         Height          =   375
         Left            =   -69960
         TabIndex        =   116
         Top             =   600
         Width           =   3735
      End
      Begin VB.CommandButton CmdPOComplete 
         Caption         =   "查询"
         Height          =   360
         Left            =   -65640
         TabIndex        =   114
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton CmdPOCompleteOut 
         Caption         =   "导出"
         Height          =   360
         Left            =   -63840
         TabIndex        =   113
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton CmdPOStartOut 
         Caption         =   "导出"
         Height          =   360
         Left            =   3360
         TabIndex        =   111
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton CmdPOStart 
         Caption         =   "查询"
         Height          =   360
         Left            =   720
         TabIndex        =   110
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton CmdPOComit 
         Caption         =   "查询"
         Height          =   360
         Left            =   -74280
         TabIndex        =   109
         Top             =   480
         Width           =   990
      End
      Begin VB.CommandButton CmdPOComitOut 
         Caption         =   "导出"
         Height          =   360
         Left            =   -71640
         TabIndex        =   108
         Top             =   480
         Width           =   990
      End
      Begin VB.CommandButton CmdWaferRecOut 
         Caption         =   "导出"
         Height          =   360
         Left            =   -71640
         TabIndex        =   106
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton CmdWaferRecQuery 
         Caption         =   "查询"
         Height          =   360
         Left            =   -74280
         TabIndex        =   105
         Top             =   600
         Width           =   990
      End
      Begin VB.Frame Frame3 
         Caption         =   "WO上传"
         Height          =   2535
         Left            =   -74040
         TabIndex        =   55
         Top             =   1380
         Width           =   7095
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   495
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   60
            Top             =   840
            Width           =   4935
         End
         Begin VB.CommandButton Command6 
            Caption         =   ".."
            Height          =   495
            Left            =   6120
            TabIndex        =   59
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton Command7 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   1200
            TabIndex        =   58
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command8 
            Caption         =   "导出主表"
            Height          =   480
            Left            =   3720
            TabIndex        =   57
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CommandButton Command9 
            Caption         =   "导出明细表"
            Height          =   480
            Left            =   5400
            TabIndex        =   56
            Top             =   1680
            Width           =   1095
         End
         Begin MSComDlg.CommonDialog CommonDialog2 
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
            Index           =   2
            Left            =   840
            TabIndex        =   61
            Top             =   480
            Width           =   1545
         End
      End
      Begin VB.TextBox TxtCustomer 
         Height          =   375
         Left            =   -73080
         TabIndex        =   54
         Top             =   900
         Width           =   2415
      End
      Begin VB.TextBox TxtPO 
         Height          =   375
         Left            =   -68640
         TabIndex        =   53
         Top             =   900
         Width           =   2415
      End
      Begin VB.TextBox TxtPOItem 
         Height          =   375
         Left            =   -64440
         TabIndex        =   52
         Top             =   900
         Width           =   2415
      End
      Begin VB.TextBox TxtLotId 
         Height          =   375
         Left            =   -73080
         TabIndex        =   51
         Top             =   1520
         Width           =   2415
      End
      Begin VB.TextBox TxtMpn 
         Height          =   375
         Left            =   -68640
         TabIndex        =   50
         Top             =   1500
         Width           =   2415
      End
      Begin VB.TextBox TxtMpnDesc 
         Height          =   375
         Left            =   -64440
         TabIndex        =   49
         Top             =   1520
         Width           =   2415
      End
      Begin VB.TextBox TxtWaferQty 
         Height          =   375
         Left            =   -73080
         TabIndex        =   48
         Top             =   2140
         Width           =   2415
      End
      Begin VB.TextBox TxtDieQty 
         Height          =   375
         Left            =   -68640
         TabIndex        =   47
         Top             =   2140
         Width           =   2415
      End
      Begin VB.TextBox TxtDesign 
         Height          =   375
         Left            =   -64440
         TabIndex        =   46
         Top             =   2140
         Width           =   2415
      End
      Begin VB.TextBox TxtCountryFab 
         Height          =   375
         Left            =   -73080
         TabIndex        =   45
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox TxtImageRev 
         Height          =   375
         Left            =   -68640
         TabIndex        =   44
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox TxtFFacility 
         Height          =   375
         Left            =   -64440
         TabIndex        =   43
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox TxtMarkId 
         Height          =   375
         Left            =   -73080
         TabIndex        =   42
         Top             =   3380
         Width           =   2415
      End
      Begin VB.TextBox TxtLotPriority 
         Height          =   375
         Left            =   -68640
         TabIndex        =   41
         Top             =   3380
         Width           =   2415
      End
      Begin VB.TextBox TxtFilmApld 
         Height          =   375
         Left            =   -64440
         TabIndex        =   40
         Top             =   3380
         Width           =   2415
      End
      Begin VB.TextBox TxtShip260 
         Height          =   375
         Left            =   -73080
         TabIndex        =   39
         Top             =   4000
         Width           =   2415
      End
      Begin VB.TextBox TxtShipLevel 
         Height          =   375
         Left            =   -68640
         TabIndex        =   38
         Top             =   4000
         Width           =   2415
      End
      Begin VB.TextBox TxtMicMaterial 
         Height          =   375
         Left            =   -64440
         TabIndex        =   37
         Top             =   4000
         Width           =   2415
      End
      Begin VB.TextBox TxtShipSite 
         Height          =   375
         Left            =   -73080
         TabIndex        =   36
         Top             =   4620
         Width           =   2415
      End
      Begin VB.TextBox TxtLotStatus 
         Height          =   375
         Left            =   -68640
         TabIndex        =   35
         Top             =   4620
         Width           =   2415
      End
      Begin VB.CommandButton CmdSaveOI 
         Caption         =   "保存"
         Height          =   480
         Left            =   -71040
         TabIndex        =   34
         Top             =   5580
         Width           =   1335
      End
      Begin VB.CommandButton CmdClearOI 
         Caption         =   "清空"
         Height          =   480
         Left            =   -68160
         TabIndex        =   33
         Top             =   5580
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mapping_XML"
         Height          =   2295
         Left            =   -74040
         TabIndex        =   27
         Top             =   1380
         Width           =   9015
         Begin VB.CommandButton Command10 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   4080
            TabIndex        =   31
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton Command11 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   1200
            TabIndex        =   30
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton Command12 
            Caption         =   ".."
            Height          =   495
            Left            =   6120
            TabIndex        =   29
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox Text4 
            Enabled         =   0   'False
            Height          =   495
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   840
            Width           =   4935
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
            TabIndex        =   32
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.TextBox TxtCustomerName 
         Height          =   375
         Left            =   -73440
         TabIndex        =   26
         Top             =   780
         Width           =   2415
      End
      Begin VB.ComboBox CmbCustomer 
         Height          =   315
         ItemData        =   "Frm37Portal.frx":00F3
         Left            =   -73320
         List            =   "Frm37Portal.frx":00F5
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   780
         Width           =   1695
      End
      Begin VB.Frame Frame5 
         Caption         =   "Mapping上传"
         Height          =   3015
         Left            =   -74040
         TabIndex        =   19
         Top             =   4260
         Width           =   7095
         Begin VB.TextBox TxtSI 
            Enabled         =   0   'False
            Height          =   975
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   480
            Width           =   5295
         End
         Begin VB.CommandButton Command13 
            Caption         =   ".."
            Height          =   495
            Left            =   6120
            TabIndex        =   22
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton Command14 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   1200
            TabIndex        =   21
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command15 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   4080
            TabIndex        =   20
            Top             =   1680
            Width           =   1335
         End
         Begin MSComDlg.CommonDialog ComSI 
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
            Caption         =   "选择待上传的map："
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   840
            TabIndex        =   24
            Top             =   240
            Width           =   1560
         End
      End
      Begin VB.TextBox TxtCustomerID 
         Height          =   375
         Left            =   -73200
         TabIndex        =   18
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox TxtItem 
         Height          =   375
         Left            =   -69480
         TabIndex        =   17
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox TxtPONO 
         Height          =   375
         Left            =   -65280
         TabIndex        =   16
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox TxtSupplier 
         Height          =   375
         Left            =   -61680
         TabIndex        =   15
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox TxtFabDevice 
         Height          =   375
         Left            =   -69480
         TabIndex        =   14
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox TxtCustomerDevice 
         Height          =   375
         Left            =   -65280
         TabIndex        =   13
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox TxtWaferVersion 
         Height          =   375
         Left            =   -61680
         TabIndex        =   12
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox TxtMarkingLotID 
         Height          =   375
         Left            =   -73200
         TabIndex        =   11
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox TxtLotId2 
         Height          =   375
         Left            =   -65280
         TabIndex        =   9
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox TxtWaferID2 
         Height          =   375
         Left            =   -61680
         TabIndex        =   8
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox TxtTotailDie 
         Height          =   375
         Left            =   -73200
         TabIndex        =   7
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox TxtGoodDieQty 
         Height          =   375
         Left            =   -69480
         TabIndex        =   6
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox TxtWONO2 
         Height          =   375
         Left            =   -65280
         TabIndex        =   5
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox TxtShipTo 
         Height          =   375
         Left            =   -73200
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "保存"
         Height          =   600
         Left            =   -71040
         TabIndex        =   3
         Top             =   4320
         Width           =   1575
      End
      Begin VB.CommandButton ComdClear 
         Caption         =   "清空"
         Height          =   600
         Left            =   -67800
         TabIndex        =   2
         Top             =   4320
         Width           =   1575
      End
      Begin VB.CommandButton Command18 
         Caption         =   "退出"
         Height          =   600
         Left            =   -64920
         TabIndex        =   1
         Top             =   4320
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -69480
         TabIndex        =   10
         Top             =   2160
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   251985921
         CurrentDate     =   42173
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5295
         Index           =   0
         Left            =   -74640
         TabIndex        =   104
         Top             =   1320
         Width           =   13215
         _Version        =   524288
         _ExtentX        =   23310
         _ExtentY        =   9340
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
         SpreadDesigner  =   "Frm37Portal.frx":00F7
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   4815
         Index           =   1
         Left            =   -74640
         TabIndex        =   107
         Top             =   2280
         Width           =   11895
         _Version        =   524288
         _ExtentX        =   20981
         _ExtentY        =   8493
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
         SpreadDesigner  =   "Frm37Portal.frx":05BB
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5295
         Index           =   2
         Left            =   360
         TabIndex        =   112
         Top             =   1320
         Width           =   11895
         _Version        =   524288
         _ExtentX        =   20981
         _ExtentY        =   9340
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
         SpreadDesigner  =   "Frm37Portal.frx":0A7F
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5295
         Index           =   3
         Left            =   -74400
         TabIndex        =   115
         Top             =   1320
         Width           =   13575
         _Version        =   524288
         _ExtentX        =   23945
         _ExtentY        =   9340
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
         SpreadDesigner  =   "Frm37Portal.frx":0F43
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5295
         Index           =   4
         Left            =   -74520
         TabIndex        =   121
         Top             =   1320
         Width           =   13935
         _Version        =   524288
         _ExtentX        =   24580
         _ExtentY        =   9340
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
         SpreadDesigner  =   "Frm37Portal.frx":1407
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5295
         Index           =   5
         Left            =   -74640
         TabIndex        =   127
         Top             =   1320
         Width           =   11895
         _Version        =   524288
         _ExtentX        =   20981
         _ExtentY        =   9340
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
         SpreadDesigner  =   "Frm37Portal.frx":18CB
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO类型"
         Height          =   195
         Left            =   -74040
         TabIndex        =   141
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO类型"
         Height          =   195
         Left            =   -74040
         TabIndex        =   139
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出货单据编号："
         Height          =   195
         Left            =   -71040
         TabIndex        =   122
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出货单据编号："
         Height          =   195
         Left            =   -71400
         TabIndex        =   117
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户："
         Height          =   195
         Left            =   -73680
         TabIndex        =   103
         Top             =   1020
         Width           =   540
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Po_Num："
         Height          =   195
         Left            =   -69480
         TabIndex        =   102
         Top             =   1020
         Width           =   765
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Po_Item："
         Height          =   195
         Left            =   -65280
         TabIndex        =   101
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source_Batch_Id："
         Height          =   195
         Left            =   -74520
         TabIndex        =   100
         Top             =   1620
         Width           =   1410
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mpn："
         Height          =   195
         Left            =   -69240
         TabIndex        =   99
         Top             =   1620
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mpn_Desc："
         Height          =   195
         Left            =   -65400
         TabIndex        =   98
         Top             =   1620
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current_Wafer_Qty："
         Height          =   195
         Left            =   -74760
         TabIndex        =   97
         Top             =   2220
         Width           =   1635
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Die_Qty："
         Height          =   195
         Left            =   -69480
         TabIndex        =   96
         Top             =   2340
         Width           =   765
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Design_Id："
         Height          =   195
         Left            =   -65400
         TabIndex        =   95
         Top             =   2220
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country_Of_Fab："
         Height          =   195
         Left            =   -74520
         TabIndex        =   94
         Top             =   2820
         Width           =   1395
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imager_Customer_Rev："
         Height          =   195
         Left            =   -70560
         TabIndex        =   93
         Top             =   2820
         Width           =   1845
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fabrication_Facility："
         Height          =   195
         Left            =   -66000
         TabIndex        =   92
         Top             =   2820
         Width           =   1560
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encoded_Mark_Id："
         Height          =   195
         Left            =   -74640
         TabIndex        =   91
         Top             =   3420
         Width           =   1470
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot_Priority："
         Height          =   195
         Left            =   -69720
         TabIndex        =   90
         Top             =   3420
         Width           =   1005
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Protective_Film_Apld："
         Height          =   195
         Left            =   -66120
         TabIndex        =   89
         Top             =   3420
         Width           =   1680
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping_Mst_260："
         Height          =   195
         Left            =   -74640
         TabIndex        =   88
         Top             =   4140
         Width           =   1485
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping_Mst_Level："
         Height          =   195
         Left            =   -70320
         TabIndex        =   87
         Top             =   4140
         Width           =   1590
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Micron_Material："
         Height          =   195
         Left            =   -65760
         TabIndex        =   86
         Top             =   4140
         Width           =   1305
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ship_Site："
         Height          =   195
         Left            =   -74040
         TabIndex        =   85
         Top             =   4740
         Width           =   840
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot_Status："
         Height          =   195
         Left            =   -69720
         TabIndex        =   84
         Top             =   4740
         Width           =   960
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注："
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -73920
         TabIndex        =   83
         Top             =   4140
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excel模板格式为：WaferId LotId ProductId 良品数 不良数"
         Height          =   195
         Left            =   -73440
         TabIndex        =   82
         Top             =   4500
         Width           =   4395
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户："
         Height          =   195
         Left            =   -74040
         TabIndex        =   81
         Top             =   900
         Width           =   540
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户"
         Height          =   195
         Left            =   -73800
         TabIndex        =   80
         Top             =   780
         Width           =   360
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请先选择客户代码，然后再上传WO或Mapping"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -71280
         TabIndex        =   79
         Top             =   900
         Width           =   3570
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -71520
         TabIndex        =   78
         Top             =   900
         Width           =   90
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GC客户WLT，MG客户上传时，先上传WO，后再上传Mapping。"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -73680
         TabIndex        =   77
         Top             =   7260
         Width           =   4860
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码："
         Height          =   195
         Left            =   -74280
         TabIndex        =   76
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item："
         Height          =   195
         Left            =   -70320
         TabIndex        =   75
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO NO："
         Height          =   195
         Left            =   -66240
         TabIndex        =   74
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier："
         Height          =   195
         Left            =   -62640
         TabIndex        =   73
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To："
         Height          =   195
         Left            =   -74160
         TabIndex        =   72
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FAB Device："
         Height          =   195
         Left            =   -70560
         TabIndex        =   71
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Device："
         Height          =   195
         Left            =   -66960
         TabIndex        =   70
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "wafer Version："
         Height          =   195
         Left            =   -63000
         TabIndex        =   69
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marking Lot ID："
         Height          =   195
         Left            =   -74640
         TabIndex        =   68
         Top             =   2280
         Width           =   1275
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date："
         Height          =   195
         Left            =   -70320
         TabIndex        =   67
         Top             =   2280
         Width           =   675
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot ID："
         Height          =   195
         Left            =   -66240
         TabIndex        =   66
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wafer ID："
         Height          =   195
         Left            =   -62640
         TabIndex        =   65
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Die Qty："
         Height          =   195
         Left            =   -74520
         TabIndex        =   64
         Top             =   3000
         Width           =   1275
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Good Die Qty："
         Height          =   195
         Left            =   -70680
         TabIndex        =   63
         Top             =   3000
         Width           =   1275
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WO NO："
         Height          =   195
         Left            =   -66240
         TabIndex        =   62
         Top             =   2880
         Width           =   795
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注：文件生成路径 C:\SemTechReport"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   960
      TabIndex        =   124
      Top             =   7800
      Width           =   3030
   End
End
Attribute VB_Name = "Frm37Portal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
   ' E_ID = 1                 'id
    E_Event = 1                'Key
    E_PONumber                  'PO
    E_POLineItem               'PO item
    E_PT                '料号
    E_QTY               '数量
    E_LotID            'LotID
     E_OK                     '选择
    E_End
    


    
End Enum


Private Enum E_FPS1          'Detail
   ' E_ID = 1                 'id
    E_Event = 1               'Key
    E_PONumber                'PO
    E_POLineItem              'PO item
    E_Pt2                '料号
    E_ComDate            '日期
    E_LotID
    E_End
    
End Enum

Private Enum E_FPS2          'Detail
   ' E_ID = 1                 'id
    E_Event = 1               'Key
    E_PONumber                'PO
    E_POLineItem              'PO item
    E_Pt2                '料号
    E_StartDate            '日期
    E_LotNumber           'LotID
    E_End
    
End Enum


Private Enum E_FPS3          'Detail
   ' E_ID = 1                 'id
    E_Event = 1               'Key
    E_PONumber                'PO
    E_POLineItem              'PO item
    E_Pt2                '料号
    E_STATUS
    E_QTY
    E_GDQty
    E_NGQty
    E_LotNumber           'LotID
    E_End
    
End Enum


Private Enum E_FPS4          'Detail
   ' E_ID = 1                 'id
    E_Event = 1               'Key
    E_PONumber                'PO
    E_POLineItem              'PO item
    E_Pt2                '料号
    E_MPlant
    E_SPlant
    E_MNumber
    E_SLotNumber
    E_STATUS
    E_GDQty
    E_SDate
    E_Origin
    E_DateCode
    E_LotNumber
    E_End
    
End Enum



Private Enum E_FPS5          'Detail
   ' E_ID = 1                 'id
    E_Event = 1               'Key
    E_PONumber                'PO
    E_POLineItem              'PO item
    E_Pt2                '料号
    E_ComDate            '日期
    E_LotID
    E_End
    
End Enum






Dim listRS As New ADODB.Recordset
Dim list2RS As New ADODB.Recordset


Public g_Date           As String



Private Sub CmdAdd_Click()
'新增
Dim tempKey As String
Dim tempValue As String
Dim getValue As String
Dim otherValue As String

Dim sqlTemp As String

tempKey = UCase(Trim(txtdelNote.Text))
tempValue = Trim(txtawb.Text)
getValue = CombMo.Text
otherValue = Trim(TxtPackage.Text)

'判断是否已输入
 If tempKey = "" Or getValue = "" Then
    MsgBox "输入完整后，再提交！", vbInformation, "友情提示"
    Exit Sub
 
 End If


 
sqlTemp = " insert into  tblsetpf(fieldName,fieldValue,resultValue,other,flag,createby,createdate) values ('" & tempKey & "','" & tempValue & "','" & getValue & "','" & otherValue & "','Y','Auto',sysdate)"
AddSql (sqlTemp)

 MsgBox "添加成功!", vbInformation, "友情提示"
 
ShowData_Where

End Sub

Private Sub CmdOut_Click()
Dim tempBillNo As String
tempBillNo = UCase(Trim(TxtBillNo.Text))

If tempBillNo = "" Then
    MsgBox "请输入单据编号!", vbInformation, "友情提示"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNo2(tempBillNo)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号维护过的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If
    


 Dim sqlTemp As String

 sqlTemp = "SELECT [PO_num] ,[PO_Item] ,[Previous_Batch_ID] ,[Previous_Mtrl_Num],[Batch_ID] ,[Mtrl_num] ,[Mtrl_desc] ,[Mtrl_num_Mtrlgrp] " & _
           " ,[Output_Qty] ,[Consumed_Qty] ,[Reject_Qty] ,[Current_Wafer_Qty] ,[Film_Frame_Qty] ,[Optical_Quality] ,[Country_of_Assembly] " & _
           " ,[Offshore_ASM_Company] ,[Asm_Containment_type] ,[Date_code] ,[asm_conv_id] ,[asm_excr_id] ,[assembly_facility] " & _
           " ,[Country_of_Test],[Offshore_TEST_Company] ,[Tst_Containment_type] ,[Tst_Program_rev] ,[Created_date] ,[Created_time],b.Del_Note,b.AWB,b.Weight,b.Package " & _
           " FROM [erpdata].[dbo].[GRdetailHistory] a,[erpdata].[dbo].[GRdetailSetUp] b " & _
           " Where a.单据编号 = b.单据编号 and a.单据编号='" + tempBillNo + "' "

  SqlServerExporToExcel (sqlTemp)


End Sub

Private Sub CmdSaver_Click()
'保存到SqlServer中

Dim tempBillNo As String
Dim tempdelNote As String
Dim tempAwb As String

Dim tempWeight As Single
Dim tempPackage As Integer

Dim cmdStrSql As String

tempBillNo = ""
tempdelNote = ""
tempAwb = ""

tempBillNo = UCase(Trim(TxtBillNo.Text))
tempBillNo = Replace(tempBillNo, vbCrLf, "")
tempBillNo = Replace(tempBillNo, vbCr, "")
tempBillNo = Replace(tempBillNo, vbLf, "")


tempdelNote = UCase(Trim(txtdelNote.Text))
tempdelNote = Replace(tempdelNote, vbCrLf, "")
tempdelNote = Replace(tempdelNote, vbCr, "")
tempdelNote = Replace(tempdelNote, vbLf, "")


tempAwb = UCase(Trim(txtawb.Text))
tempAwb = Replace(tempAwb, vbCrLf, "")
tempAwb = Replace(tempAwb, vbCr, "")
tempAwb = Replace(tempAwb, vbLf, "")


If tempBillNo = "" Or tempdelNote = "" Or tempAwb = "" Or Trim(txtWeight.Text) = "" Or Trim(TxtPackage.Text) = "" Then
    MsgBox "请输入完整资料!", vbInformation, "友情提示"
    Exit Sub
End If



tempWeight = CSng(Trim(txtWeight.Text))
tempWeight = Replace(tempWeight, vbCrLf, "")
tempWeight = Replace(tempWeight, vbCr, "")
tempWeight = Replace(tempWeight, vbLf, "")


tempPackage = CInt(UCase(Trim(TxtPackage.Text)))
tempPackage = Replace(tempPackage, vbCrLf, "")
tempPackage = Replace(tempPackage, vbCr, "")
tempPackage = Replace(tempPackage, vbLf, "")


'2013-11-21 判断单据编号 是否存在

  Dim judgeEmp As Boolean
  judgeEmp = JudgeGRBillNo(tempBillNo)

    If judgeEmp = False Then
    
     MsgBox "这单据编号还没生成GR，暂时不可以维护相关信息!", vbInformation, "友情提示"
     Exit Sub
     
    End If
    
   '是否已维护过
    judgeEmp = JudgeGRBillNo2(tempBillNo)
     If judgeEmp = True Then
    
     MsgBox "这单据编号已维护过，不可再次维护，请确认!", vbInformation, "友情提示"
     Exit Sub
     
    End If
    

    

cmdStrSql = " insert into [erpdata].[dbo].[GRdetailSetUp](单据编号,Del_Note,AWB,[Weight],Package) values('" & tempBillNo & "'," & _
             " '" & tempdelNote & "','" & tempAwb & "'," & tempWeight & "," & tempPackage & " )"



AddSql2 (cmdStrSql)

MsgBox "保存信息成功 !", vbInformation, "提示"


End Sub

Private Sub CmdSend_Click()
'发送

Dim tempBillNo As String
tempBillNo = UCase(Trim(TxtBillNo.Text))

If tempBillNo = "" Then
    MsgBox "请输入单据编号!", vbInformation, "友情提示"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNo2(tempBillNo)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号维护过的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If


'    SaveFileSend
    SaveFileSendTest

End Sub

Private Sub Combo2_Change()
TxtBillNoGC.SetFocus

End Sub

Private Sub Combo2_Click()
TxtBillNoGC.SetFocus
End Sub

Private Sub ChkAll_Click()


Dim i As Integer
    If ChkAll.Value = 1 Then
        For i = 1 To fps(0).MaxRows
            With fps(0)
                .Row = i
                .Col = E_FPS0.E_OK
                .Text = 1
            End With
        Next i
        
    ElseIf ChkAll.Value = 0 Then
        For i = 1 To fps(0).MaxRows
            With fps(0)
                .Row = i
                .Col = E_FPS0.E_OK
                .Text = 0
            End With
        Next i
        
    End If



End Sub

Private Sub CmdPOComit_Click()
'明细数据
Dim i As Integer
Dim waferIdTemp As String
Dim woType As String


    Set listRS = GetFps37POCommit()


If listRS.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认", vbInformation, "友情提示"
    Exit Sub
    
Else


    fps(1).MaxRows = listRS.RecordCount
    For i = 0 To listRS.RecordCount - 1
         


         
         With fps(1)
                 .Row = i + 1
                 .Col = E_FPS1.E_Event
                 .Text = CStr(listRS.Fields(0).Value)
                 
                .Row = i + 1
                 .Col = E_FPS1.E_PONumber
                .Text = "" & CStr(IIf(IsNull(listRS.Fields(1).Value), "", listRS.Fields(1).Value))
                 
                
                
                  .Row = i + 1
                 .Col = E_FPS1.E_POLineItem
                .Text = "" & CStr(IIf(IsNull(listRS.Fields(2).Value), "", listRS.Fields(2).Value))
                 
                  .Row = i + 1
                 .Col = E_FPS1.E_Pt2
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(3).Value), "", listRS.Fields(3).Value))
                 
                 
                  .Row = i + 1
                 .Col = E_FPS1.E_ComDate
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(4).Value), "", listRS.Fields(4).Value))
                 
                 
                   .Row = i + 1
                 .Col = E_FPS1.E_LotID
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(5).Value), "", listRS.Fields(5).Value))
                 
                 
                
        
        End With
    
NextRecord:
       
        listRS.MoveNext

    Next


End If


End Sub

'Private Sub Command1_Click()
'
'
''明细数据
'Dim i As Integer
'Dim waferidTemp As String
'Dim woType As String
''ST
'woType = Mid(woTemp, 2, 2)
'
''If (customerTemp = "AA" Or customerTemp = "AA(ON)") And (woType = "ST" Or woType = "ET") Then
''
''    'Set listRS = GetFpsAARTWo(strwhereTemp, customerTemp, woType)
''
''Else
''
''    'Set listRS = GetFps(strwhereTemp, customerTemp)
''
''End If
'
'
'If listRS.RecordCount <= 0 Then
'    MsgBox "明细表中没有相关数据，请确认"
'    Exit Sub
'
'Else
'
'    '2014-11-12 jiayun add
'
'    fps(0).MaxRows = listRS.RecordCount
'    For i = 0 To listRS.RecordCount - 1
'
'         waferidTemp = CStr(listRS.fields(1).Value)
'
'
'
'
'
'         With fps(0)
'                 .Row = i + 1
'                 .Col = E_FPS0.E_ID
'                 .Text = i + 1
'
'                .Row = i + 1
'                 .Col = E_FPS0.E_WaferID
'                .Text = CStr(listRS.fields(1).Value)
'
'
'                 .Row = i + 1
'                 .Col = E_FPS0.E_CompleteFlag
'                .Text = ""
'
'                  .Row = i + 1
'                 .Col = E_FPS0.E_TotalDie
'                 .Text = CStr(listRS.fields(3).Value)
'
'                  .Row = i + 1
'                 .Col = E_FPS0.E_GoodDie
'                 .Text = CStr(listRS.fields(4).Value)
'
'
'                  .Row = i + 1
'                 .Col = E_FPS0.E_WaferLot
'                 .Text = CStr(listRS.fields(5).Value)
'
'                  .Row = i + 1
'                 .Col = E_FPS0.E_MarkingCode
'                 .Text = "" & listRS.fields(6).Value
'
'
'                 .Row = i + 1
'                 .Col = E_FPS0.E_OK
'                .Text = CStr("1")
'
'
'
'        End With
'
'NextRecord:
'
'        listRS.MoveNext
'
'    Next
'
'
'End If
'
'
'
'
'
'
'
'End Sub

Private Sub CmdPOComitOut_Click()




Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

'On Error GoTo ErrHandler
'查询报表名的序号

fileNo = Format(Now, "YYYYMMDD_HHMM") & "HTKS_PO_ACK"

'20151106_0846HTKS_RECEIPT.csv

Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path37 & "\" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "Event,PO Number,PO Line Item Number,Production Order Number,Commit Date,LotID(厂内),PO交期(厂内),客户机种(厂内),华天机种(厂内) ,Bag#(厂内)" & vbCrLf
    '明细数据
    

    
  strDatas = strDatas + ",,,,,,," & vbCrLf

  
strSql = "select distinct 'COMMIT',d.po_num ,d.po_item,d.mpn,a.planenddate,d.source_batch_id,d.date_code,d.mpn_desc,substr(a.product,3,7), " & _
 " d.mtrl_num from ib_wohistory a, ib_waferlist b ,mappingdatatest c ,customeroitbl_test d ,mfgorder e,container f where " & _
 " a.OrderName = b.OrderName And c.SubstrateId = b.waferid And to_char(d.id) = c.FileName And e.mfgordername = a.OrderName " & _
 " and e.mfgorderid = f.mfgorderid and d.customershortname = '37' and a.erpcreatedate > sysdate - 3 and d.po_num is not null order by a.planenddate desc  "




    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    
    If Cnn.State = 0 Then
    ConOracle
    End If
    
    
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then
    
   MsgBox "明细表中没有相关数据，请确认", vbInformation, "友情提示"
    
     Exit Sub
     
     End If
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        For j = 0 To rs.Fields.Count - 1

             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
'    '发邮件
'    Dim strRecipient    As String
'    Dim strRecipientCC  As String
'
'    strRecipient = "jiayun.zhang@qtechglobal.com"
'    strRecipientCC = "wanli.ma@qtechglobal.com"
'
'    Call MailDetailHD("HD 发货报表", strRecipient, g_Path & "\" & "HD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
'
'    '把发送记录保存到DB中
    
'    Dim sqlTemp2 As String
'
'    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'HD') "
'
'    Call AddSql2(sqlTemp2)
'
'    MsgBox "发送成功！", vbInformation, "友情提示"
'
'
'ErrHandler:
'    Set FSO = Nothing


MsgBox "导出成功！", vbInformation, "友情提示"





End Sub

Private Sub CmdPOComOut_Click()


 ExporToExcel ("   select 'COMMIT' as Event,ponum as PONumber ,poline as POLineItemNumber, productpt as ProductionOrderNumber, commitdate as  CommitDate ,createdate  from  SemtechPortalCommit order by createdate desc")


End Sub

Private Sub CmdPOComplete_Click()

Dim tempBillNo As String
Dim custNameTemp As String


If CmbPoType.Text = "ICI" Then

POComleteICI

Exit Sub

End If




tempBillNo = UCase(Trim(Txt37BillNo.Text))


If tempBillNo = "" Then
    MsgBox "请输入出货单据编号!", vbInformation, "友情提示"
    Exit Sub
End If


 Dim judgeEmp As Boolean

judgeEmp = JudgeSemtechBillNo(tempBillNo)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If





'明细数据
Dim i As Integer
Dim waferIdTemp As String
Dim woType As String
Dim baofeiQty As Long
Dim outQty As Long
Dim sendTimeTemp As Long
Dim maxBillNoTemp As String


    Set listRS = GetFps37POComplete(tempBillNo)


If listRS.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认"
    Exit Sub
    
Else


    fps(3).MaxRows = listRS.RecordCount
    For i = 0 To listRS.RecordCount - 1
         baofeiQty = 0
         outQty = 0
         sendTimeTemp = 0
         maxBillNoTemp = ""
         
         With fps(3)
                 .Row = i + 1
                 .Col = E_FPS3.E_Event
                 .Text = CStr(listRS.Fields(0).Value)
                 
                .Row = i + 1
                 .Col = E_FPS3.E_PONumber
                .Text = CStr(listRS.Fields(1).Value)
                
                
                  .Row = i + 1
                 .Col = E_FPS3.E_POLineItem
                 .Text = CStr(listRS.Fields(2).Value)
                 
                  .Row = i + 1
                 .Col = E_FPS3.E_Pt2
                 .Text = CStr(listRS.Fields(3).Value)
                 
                 
                
                 
                 .Row = i + 1
                 .Col = E_FPS3.E_QTY
                 .Text = CStr(listRS.Fields(5).Value)
                 
                   .Row = i + 1
                 .Col = E_FPS3.E_GDQty
                 .Text = CStr(listRS.Fields(6).Value)
                 
                .Row = i + 1
                 .Col = E_FPS3.E_NGQty
                 '查询质量有没有报废
                 baofeiQty = GetQty37Baofei(CStr(listRS.Fields(8).Value))
                 If baofeiQty = 0 Then
                 .Text = ""
                 Else
                   .Text = CStr(baofeiQty)
                 End If
                 
                 
                  .Row = i + 1
                 .Col = E_FPS3.E_STATUS
                 '判断有没有出完
                 
                 outQty = GetQty37OutQty(CStr(listRS.Fields(8).Value))
                 
                 If outQty + baofeiQty = CLng(listRS.Fields(5).Value) Then
                 
                 '判断今天有几次发货，如果有多次，则当最后一次才显示X
                 sendTimeTemp = GetQty37OutTimes(CStr(listRS.Fields(8).Value))
                 
                        If sendTimeTemp > 1 Then
                        
                        '判断这个发货单号是否为最大的发货单号，是则为X ,否则为0
                        maxBillNoTemp = GetQty37OutMaxBill(CStr(listRS.Fields(8).Value))
                        
                                If maxBillNoTemp = UCase(Trim(Txt37BillNo.Text)) Then
                                 .Text = "X"
                                Else
                                .Text = ""
                                End If
        
                        End If
                 
                 
                 
                 Else
                  .Text = ""
                 
                 End If
                 
                 
                  .Row = i + 1
                 .Col = E_FPS3.E_LotNumber
                 .Text = CStr(listRS.Fields(8).Value)
                 
                 
                
        
        End With
    
NextRecord:
       
        listRS.MoveNext

    Next


End If






End Sub



Private Sub POComleteICI()

Dim tempBillNo As String
Dim custNameTemp As String


'明细数据
Dim i As Integer
Dim waferIdTemp As String
Dim woType As String
Dim baofeiQty As Long
Dim outQty As Long
Dim sendTimeTemp As Long
Dim maxBillNoTemp As String


    Set listRS = GetFps37POCompleteICI()


If listRS.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认"
    Exit Sub
    
Else


    fps(3).MaxRows = listRS.RecordCount
    For i = 0 To listRS.RecordCount - 1
         baofeiQty = 0
         outQty = 0
         sendTimeTemp = 0
         maxBillNoTemp = ""
         
         With fps(3)
                 .Row = i + 1
                 .Col = E_FPS3.E_Event
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(0).Value), "", listRS.Fields(0).Value))
                 
                .Row = i + 1
                 .Col = E_FPS3.E_PONumber
                .Text = "" & CStr(IIf(IsNull(listRS.Fields(1).Value), "", listRS.Fields(1).Value))
                
                
                  .Row = i + 1
                 .Col = E_FPS3.E_POLineItem
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(2).Value), "", listRS.Fields(2).Value))
                 
                  .Row = i + 1
                 .Col = E_FPS3.E_Pt2
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(3).Value), "", listRS.Fields(3).Value))
                 
                 
                
                 
                 .Row = i + 1
                 .Col = E_FPS3.E_QTY
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(5).Value), "", listRS.Fields(5).Value))
                 
                   .Row = i + 1
                 .Col = E_FPS3.E_GDQty
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(6).Value), "", listRS.Fields(6).Value))
                 
                .Row = i + 1
                 .Col = E_FPS3.E_NGQty
                 '查询NG Die Qty
                 baofeiQty = GetQty37NGQty(CStr(listRS.Fields(8).Value))
                 If baofeiQty = 0 Then
                 .Text = ""
                 Else
                   .Text = CStr(baofeiQty)
                 End If
                 
                 
                  .Row = i + 1
                 .Col = E_FPS3.E_STATUS
                 
            
                 
                 
                 '判断有没有出完
                 
                 outQty = GetQty37OutQtyICI(CStr(listRS.Fields(8).Value))

                 If outQty + baofeiQty >= CLng(listRS.Fields(5).Value) Then

'                 '判断今天有几次发货，如果有多次，则当最后一次才显示X
'                 sendTimeTemp = GetQty37OutTimes(CStr(listRS.fields(8).Value))
'
'                        If sendTimeTemp > 1 Then
'
'                        '判断这个发货单号是否为最大的发货单号，是则为X ,否则为0
'                        maxBillNoTemp = GetQty37OutMaxBill(CStr(listRS.fields(8).Value))
'
'                                If maxBillNoTemp = UCase(Trim(Txt37BillNo.Text)) Then
'                                 .Text = "X"
'                                Else
'                                .Text = ""
'                                End If
'
'                        End If
'
                     .Text = "X"


                 Else
                  .Text = ""

                 End If
                 
                 
                  .Row = i + 1
                 .Col = E_FPS3.E_LotNumber
                 .Text = CStr(listRS.Fields(8).Value)
                 
                 
                
        
        End With
    
NextRecord:
       
        listRS.MoveNext

    Next


End If






End Sub


Private Sub CmdPOCompleteOut_Click()


Dim tempBillNo As String
Dim custNameTemp As String
Dim baofeiQty As Long
Dim outQty  As Long
Dim allQty As Long



'Dim tempBillNo As String
'Dim custNameTemp As String


If CmbPoType.Text = "ICI" Then

CmdPOCompleteOutICI

Exit Sub

End If




tempBillNo = UCase(Trim(Txt37BillNo.Text))


If tempBillNo = "" Then
    MsgBox "请输入出货单据编号!", vbInformation, "友情提示"
    Exit Sub
End If

 Dim judgeEmp As Boolean

judgeEmp = JudgeSemtechBillNo(tempBillNo)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If



Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer

Dim sendTimeTemp As Long
Dim maxBillNoTemp As String



Dim lotIdTemp As String

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

'On Error GoTo ErrHandler
'查询报表名的序号

fileNo = Format(Now, "YYYYMMDD_HHMM") & "HTKS_COMPLETE"

'20151106_0846HTKS_RECEIPT.csv

Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path37 & "\" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "Event,PO Number,PO Line Item Number,Production Order Number,Order Close Indicator,Quantity,Yield Quantity,Scrap Quantity" & vbCrLf
    '明细数据
    

    
  strDatas = strDatas + ",,,,,," & vbCrLf


'
'strSql = "  select distinct  'START' as wostart,c.po_num,c.po_item,c.mpn,to_char(a.erpcreatedate,'YYYYMMDD'),c.SOURCE_MTRL_SLOC  from ib_wohistory a ,ib_waferlist b ,customeroitbl_test c ,mappingdata37 d" & _
'" Where b.OrderName = a.OrderName and a.customer='37' and a.erpcreatedate>to_Date('2016-03-26','YYYY-MM-DD') and c.source_batch_id=b.waferlot " & _
'" and d.batch=c.source_batch_id and d.purchaseno=c.po_num and a.lot_status is null "




strSql = " SELECT  'COMPLETE' as Event ,c.PO_NUM,c.PO_ITEM,c.MPN,'' as OrderClose,c.CURRENT_WAFER_QTY,COUNT(b.流程卡编号) as Quantity,'' as ScrapQuantity, " & _
" a.工单号 FROM  [ERPdata].[dbo].[tblStockMove] a, [ERPdata].[dbo].[tblStockMovesub] b ,[ERPBASE].[dbo].[tblCustomerOI] c " & _
" where a.单据编号='" + tempBillNo + "' and a.客户代码='37' " & _
" and b.单据编号=a.单据编号 and b.工单号=a.工单号 " & _
" and c.CUSTOMERSHORTNAME='37' and c.SOURCE_BATCH_ID=a.工单号 and c.PO_NUM <>'' " & _
" group by   c.PO_NUM,c.PO_ITEM,c.MPN,c.CURRENT_WAFER_QTY, a.工单号 "



    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    
If INIadoCon2.State = 0 Then
INIConnectSTART2
End If
    
    
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        baofeiQty = 0
        outQty = 0
        allQty = 0
        allQty = CLng(rs.Fields(5).Value)
        lotIdTemp = CStr(rs.Fields(rs.Fields.Count - 1).Value)
        baofeiQty = GetQty37Baofei(lotIdTemp)
        outQty = GetQty37OutQty(lotIdTemp)
        
          sendTimeTemp = 0
         maxBillNoTemp = ""
         
        
        For j = 0 To rs.Fields.Count - 2
              
             If j = 7 Then
             '报废处理

                If baofeiQty = 0 Then
                 strColData = strColData + "" + ","
                Else
                  strColData = strColData + CStr(baofeiQty) + ","
                End If
                
            ElseIf j = 4 Then
            
              If outQty + baofeiQty = allQty Then
              
              
                 '判断今天有几次发货，如果有多次，则当最后一次才显示X
                 sendTimeTemp = GetQty37OutTimes(lotIdTemp)
                 
                        If sendTimeTemp > 1 Then
                        
                        '判断这个发货单号是否为最大的发货单号，是则为X ,否则为0
                        maxBillNoTemp = GetQty37OutMaxBill(lotIdTemp)
                        
                                If maxBillNoTemp = UCase(Trim(Txt37BillNo.Text)) Then
                                 strColData = strColData + "X" + ","
                                Else
                                 strColData = strColData + "" + ","
                                End If
        
                        End If
              
              
              
             
                
                
                
              Else
               strColData = strColData + "" + ","
              
              End If
             
             
             Else
             
             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
             
             End If
           
        Next
        
        
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
'    '发邮件
'    Dim strRecipient    As String
'    Dim strRecipientCC  As String
'
'    strRecipient = "jiayun.zhang@qtechglobal.com"
'    strRecipientCC = "wanli.ma@qtechglobal.com"
'
'    Call MailDetailHD("HD 发货报表", strRecipient, g_Path & "\" & "HD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
'
'    '把发送记录保存到DB中
    
'    Dim sqlTemp2 As String
'
'    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'HD') "
'
'    Call AddSql2(sqlTemp2)
'
'    MsgBox "发送成功！", vbInformation, "友情提示"
'
'
'ErrHandler:
'    Set FSO = Nothing


MsgBox "导出成功！", vbInformation, "友情提示"






End Sub



Private Sub CmdPOCompleteOutICI()


Dim tempBillNo As String
Dim custNameTemp As String
Dim baofeiQty As Long
Dim outQty  As Long
Dim allQty As Long



'Dim tempBillNo As String
'Dim custNameTemp As String





Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer

Dim sendTimeTemp As Long
Dim maxBillNoTemp As String



Dim lotIdTemp As String

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

'On Error GoTo ErrHandler
'查询报表名的序号

fileNo = Format(Now, "YYYYMMDD_HHMM") & "HTKS_COMPLETE"

'20151106_0846HTKS_RECEIPT.csv

Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path37 & "\" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "Event,PO Number,PO Line Item Number,Production Order Number,Order Close Indicator,Quantity,Yield Quantity,Scrap Quantity,LotID(厂内)" & vbCrLf
    '明细数据
    

    
  strDatas = strDatas + ",,,,,,," & vbCrLf




'strSql = " SELECT  'COMPLETE' as Event ,c.PO_NUM,c.PO_ITEM,c.MPN,'' as OrderClose,c.CURRENT_WAFER_QTY,COUNT(b.流程卡编号) as Quantity,'' as ScrapQuantity, " & _
'" a.工单号 FROM  [ERPdata].[dbo].[tblStockMove] a, [ERPdata].[dbo].[tblStockMovesub] b ,[ERPBASE].[dbo].[tblCustomerOI] c " & _
'" where a.单据编号='" + tempBillNo + "' and a.客户代码='37' " & _
'" and b.单据编号=a.单据编号 and b.工单号=a.工单号 " & _
'" and c.CUSTOMERSHORTNAME='37' and c.SOURCE_BATCH_ID=a.工单号 and c.PO_NUM <>'' " & _
'" group by   c.PO_NUM,c.PO_ITEM,c.MPN,c.CURRENT_WAFER_QTY, a.工单号 "
'


'strSql = " select X.Event,X.po_num,X.po_item,X.mpn,X.OrderClose,X.die_qty,sum(X.moveinqty) as qty,X.ScrapQuantity,X.source_batch_id from ( " & _
'" select distinct a.containername,'COMPLETE' as Event , wo.po_num, wo.po_item, wo.mpn, '' as OrderClose, wo.die_qty, a.moveinqty, '' as ScrapQuantity, wo.source_batch_id " & _
'"  from a_wiplothistory  a, a_wiplotdetailshistory b, container conn, mfgorder mfg, ib_wohistory ibwo, a_lotwafers  waf, customeroitbl_test wo " & _
'" where a.specname = '5272' and a.creationtimestamp > to_date('2016-04-20', 'YYYY-MM-DD') " & _
'"   and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername " & _
'"   and waf.waferscribenumber = substr(conn.containername, 1, instr(conn.containername, '-') - 1)" & _
'"   and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername" & _
'"   and wo.customershortname = '37' " & _
'"   and wo.source_batch_id =  waf.wafernumber   and  a.creationtimestamp>sysdate-7 ) X group by X.Event,X.po_num,X.po_item,X.mpn,X.OrderClose,X.die_qty,X.ScrapQuantity,X.source_batch_id "
'
'
   
'    strSql = " select X.Event,X.po_num,X.po_item,X.mpn,X.OrderClose,X.die_qty,sum(X.moveinqty) as qty,X.ScrapQuantity,X.source_batch_id from ( " & _
'" select distinct a.containername,'COMPLETE' as Event , wo.po_num, wo.po_item, wo.mpn, '' as OrderClose, wo.die_qty, a.moveoutqty as moveinqty, '' as ScrapQuantity, wo.source_batch_id " & _
'"  from a_wiplothistory  a, a_wiplotdetailshistory b, container conn, mfgorder mfg, ib_wohistory ibwo, a_lotwafers  waf, customeroitbl_test wo " & _
'" where a.specname = '5270'  and a.productname ='18X37025B000BR'  and a.creationtimestamp > to_date('2016-04-20', 'YYYY-MM-DD') " & _
'"   and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername " & _
'"   and waf.waferscribenumber = substr(conn.containername, 1, instr(conn.containername, '-') - 1)" & _
'"   and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername" & _
'"    and wo.customershortname = '37' " & _
'"   and wo.source_batch_id = waf.wafernumber and  a.creationtimestamp>sysdate-7 ) X group by X.Event,X.po_num,X.po_item,X.mpn,X.OrderClose,X.die_qty,X.ScrapQuantity,X.source_batch_id order by  X.po_num "


strSql = "select X.Event, X.po_num, X.po_item, X.mpn, X.OrderClose, x.die_qty*(select  max(rownum) from customeroitbl_test  ct where ct.po_num = X.po_num group by ct.po_num) as dieqty," & _
" sum(X.qty) as qty, X.ScrapQuantity, X.source_batch_id from (select 'COMPLETE' as Event, qty, b.po_num, b.po_item, b.mpn, '' as OrderClose, b.die_qty, '' as ScrapQuantity," & _
" b.source_batch_id from Cus37_5272Qty a, customeroitbl_test b, mappingdatatest c where b.id = c.filename and (c.substrateid = substr(containername, 1, instr(containername, '-A') - 1) or " & _
" c.substrateid = substr(containername, 1, instr(containername, '-A') - 1))) x group by X.Event, X.po_num, X.po_item, X.mpn, X.OrderClose, X.die_qty, X.ScrapQuantity, X.source_batch_id " & _
" order by X.po_num "

   


    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    
If Cnn.State = 0 Then
ConOracle
End If
    
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        baofeiQty = 0
        outQty = 0
        allQty = 0
        allQty = CLng(rs.Fields(5).Value)
        lotIdTemp = CStr(rs.Fields(rs.Fields.Count - 1).Value)
        baofeiQty = GetQty37NGQty(lotIdTemp)
        outQty = GetQty37OutQtyICI(lotIdTemp)
        
          sendTimeTemp = 0
         maxBillNoTemp = ""
         
        
        For j = 0 To rs.Fields.Count - 1
              
             If j = 7 Then
             '报废处理

                If baofeiQty = 0 Then
                 strColData = strColData + "" + ","
                Else
                  strColData = strColData + CStr(baofeiQty) + ","
                End If
                
            ElseIf j = 4 Then
            
              If outQty + baofeiQty >= allQty Then
              
              
'                 '判断今天有几次发货，如果有多次，则当最后一次才显示X
'                 sendTimeTemp = GetQty37OutTimes(lotidTemp)
'
'                        If sendTimeTemp > 1 Then
'
'                        '判断这个发货单号是否为最大的发货单号，是则为X ,否则为0
'                        maxBillNoTemp = GetQty37OutMaxBill(lotidTemp)
'
'                                If maxBillNoTemp = UCase(Trim(Txt37BillNo.Text)) Then
'                                 strColData = strColData + "X" + ","
'                                Else
'                                 strColData = strColData + "" + ","
'                                End If
'
'                        End If
              
              
              
             strColData = strColData + "X" + ","
                
                
                
              Else
               strColData = strColData + "" + ","
              
              End If
             
             
             Else
             
             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
             
             End If
           
        Next
        
        
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    


MsgBox "导出成功！", vbInformation, "友情提示"






End Sub



Private Sub CmdPOShip_Click()


Dim tempBillNo As String
Dim custNameTemp As String




If ComboShipICI.Text = "ICI" Then

CmdPOShipICI

Exit Sub

End If




tempBillNo = UCase(Trim(Txt37BillNoShip.Text))


If tempBillNo = "" Then
    MsgBox "请输入出货单据编号!", vbInformation, "友情提示"
    Exit Sub
End If


 Dim judgeEmp As Boolean

judgeEmp = JudgeSemtechBillNo(tempBillNo)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If



'明细数据
Dim i As Integer
Dim waferIdTemp As String
Dim woType As String
Dim baofeiQty As Long
Dim outQty As Long
Dim lotIdTemp As String
Dim allQty As String

Dim sendTimeTemp As Long
Dim maxBillNoTemp As String


    Set listRS = GetFps37POShip(tempBillNo)


If listRS.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认"
    Exit Sub
    
Else


    fps(4).MaxRows = listRS.RecordCount
    For i = 0 To listRS.RecordCount - 1
         baofeiQty = 0
         outQty = 0
         allQty = 0
         lotIdTemp = CStr(listRS.Fields(13).Value)
         
           sendTimeTemp = 0
         maxBillNoTemp = ""

         
         With fps(4)
                 .Row = i + 1
                 .Col = E_FPS4.E_Event
                 .Text = CStr(listRS.Fields(0).Value)
                 
                .Row = i + 1
                 .Col = E_FPS4.E_PONumber
                .Text = CStr(listRS.Fields(1).Value)
                
                
                  .Row = i + 1
                 .Col = E_FPS4.E_POLineItem
                 .Text = CStr(listRS.Fields(2).Value)
                 
                  .Row = i + 1
                 .Col = E_FPS4.E_Pt2
                 .Text = CStr(listRS.Fields(3).Value)
                 
                                 
                  .Row = i + 1
                 .Col = E_FPS4.E_MPlant
                 .Text = CStr(listRS.Fields(4).Value)
                 
                  .Row = i + 1
                 .Col = E_FPS4.E_SPlant
                 .Text = CStr(listRS.Fields(5).Value)
                 
                  .Row = i + 1
                 .Col = E_FPS4.E_MNumber
                 .Text = CStr(listRS.Fields(6).Value)
                 
                  .Row = i + 1
                 .Col = E_FPS4.E_SLotNumber
                 .Text = CStr(listRS.Fields(7).Value)
                 
                .Row = i + 1
                 .Col = E_FPS4.E_STATUS
                 
                 '查有没有出完
                 
                 baofeiQty = GetQty37Baofei(lotIdTemp)
                 
                 outQty = GetQty37OutQty(lotIdTemp)
                 
                 allQty = GetQty37OIAllQty(lotIdTemp, CStr(listRS.Fields(1).Value))
                 
                 If baofeiQty + outQty = allQty Then
                 
                 
                 
                 '判断今天有几次发货，如果有多次，则当最后一次才显示X
                 sendTimeTemp = GetQty37OutTimes(lotIdTemp)
                 
                        If sendTimeTemp > 1 Then
                        
                        '判断这个发货单号是否为最大的发货单号，是则为X ,否则为0
                        maxBillNoTemp = GetQty37OutMaxBill(lotIdTemp)
                        
                                If maxBillNoTemp = UCase(Trim(Txt37BillNoShip.Text)) Then
                                 .Text = "X"
                                Else
                                .Text = ""
                                End If
        
                        End If
                 
                 
    
                 
                 
                 Else
                  .Text = ""
                 
                 End If
                 
            
                  .Row = i + 1
                 .Col = E_FPS4.E_GDQty
                 .Text = CStr(listRS.Fields(9).Value)
                 
                .Row = i + 1
                 .Col = E_FPS4.E_SDate
                 .Text = CStr(listRS.Fields(10).Value)
                 
                  .Row = i + 1
                 .Col = E_FPS4.E_Origin
                 .Text = CStr(listRS.Fields(11).Value)
                 
                 .Row = i + 1
                 .Col = E_FPS4.E_DateCode
                 '把日期转为Oracle周
                   .Text = GetDateToCode(CStr(listRS.Fields(12).Value))

                    .Row = i + 1
                 .Col = E_FPS4.E_LotNumber
                 .Text = CStr(listRS.Fields(13).Value)
                        
                 

'                 .Row = i + 1
'                 .Col = E_FPS4.E_Qty
'                 .Text = CStr(listRS.fields(5).Value)
'
'                   .Row = i + 1
'                 .Col = E_FPS4.E_GDQty
'                 .Text = CStr(listRS.fields(6).Value)
'
'                .Row = i + 1
'                 .Col = E_FPS3.E_NGQty
'                 '查询质量有没有报废
'                 baofeiQty = GetQty37Baofei(CStr(listRS.fields(8).Value))
'                 If baofeiQty = 0 Then
'                 .Text = ""
'                 Else
'                   .Text = CStr(baofeiQty)
'                 End If
'
'
'                  .Row = i + 1
'                 .Col = E_FPS4.E_Status
'                 '判断有没有出完
'
'                 outQty = GetQty37OutQty(CStr(listRS.fields(8).Value))
'
'                 If outQty + baofeiQty = CLng(listRS.fields(5).Value) Then
'                 .Text = "X"
'                 Else
'                  .Text = ""
'
'                 End If
'
'
'                  .Row = i + 1
'                 .Col = E_FPS4.E_LotNumber
'                 .Text = CStr(listRS.fields(8).Value)
'
                 
                
        
        End With
    
NextRecord:
       
        listRS.MoveNext

    Next


End If





End Sub



Private Sub CmdPOShipICI()


Dim tempBillNo As String
Dim custNameTemp As String


'明细数据
Dim i As Integer
Dim waferIdTemp As String
Dim woType As String
Dim baofeiQty As Long
Dim outQty As Long
Dim lotIdTemp As String
Dim allQty As String

Dim sendTimeTemp As Long
Dim maxBillNoTemp As String


    Set listRS = GetFps37POShipICI()


If listRS.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认"
    Exit Sub
    
Else


    fps(4).MaxRows = listRS.RecordCount
    For i = 0 To listRS.RecordCount - 1
         baofeiQty = 0
         outQty = 0
         allQty = 0
         lotIdTemp = "" & CStr(IIf(IsNull(listRS.Fields(13).Value), "", listRS.Fields(13).Value))
         
           sendTimeTemp = 0
         maxBillNoTemp = ""

         
         With fps(4)
                 .Row = i + 1
                 .Col = E_FPS4.E_Event
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(0).Value), "", listRS.Fields(0).Value))
                 
                .Row = i + 1
                 .Col = E_FPS4.E_PONumber
                .Text = "" & CStr(IIf(IsNull(listRS.Fields(1).Value), "", listRS.Fields(1).Value))
                
                
                  .Row = i + 1
                 .Col = E_FPS4.E_POLineItem
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(2).Value), "", listRS.Fields(2).Value))
                 
                  .Row = i + 1
                 .Col = E_FPS4.E_Pt2
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(3).Value), "", listRS.Fields(3).Value))
                 
                                 
                  .Row = i + 1
                 .Col = E_FPS4.E_MPlant
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(4).Value), "", listRS.Fields(4).Value))
                 
                  .Row = i + 1
                 .Col = E_FPS4.E_SPlant
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(5).Value), "", listRS.Fields(5).Value))
                 
                  .Row = i + 1
                 .Col = E_FPS4.E_MNumber
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(6).Value), "", listRS.Fields(6).Value))
                 
                  .Row = i + 1
                 .Col = E_FPS4.E_SLotNumber
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(7).Value), "", listRS.Fields(7).Value))
                 
                .Row = i + 1
                 .Col = E_FPS4.E_STATUS
                 
                 '查有没有出完
                 
                 baofeiQty = GetQty37NGQty(lotIdTemp)
                 
                 outQty = GetQty37OutQtyICI(lotIdTemp)
                 
                 allQty = GetQty37OIAllQty(lotIdTemp, ("" & CStr(IIf(IsNull(listRS.Fields(1).Value), "", listRS.Fields(1).Value))))
                 
                 If baofeiQty + outQty = allQty Then
                 
                 
'                 '判断今天有几次发货，如果有多次，则当最后一次才显示X
'                 sendTimeTemp = GetQty37OutTimes(lotidTemp)
'
'                        If sendTimeTemp > 1 Then
'
'                        '判断这个发货单号是否为最大的发货单号，是则为X ,否则为0
'                        maxBillNoTemp = GetQty37OutMaxBill(lotidTemp)
'
'                                If maxBillNoTemp = UCase(Trim(Txt37BillNoShip.Text)) Then
'                                 .Text = "X"
'                                Else
'                                .Text = ""
'                                End If
'
'                        End If
                 
                   .Text = "X"
    
                 
                 
                 Else
                  .Text = ""
                 
                 End If
                 
            
                  .Row = i + 1
                 .Col = E_FPS4.E_GDQty
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(9).Value), "", listRS.Fields(9).Value))
                 
                .Row = i + 1
                 .Col = E_FPS4.E_SDate
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(10).Value), "", listRS.Fields(10).Value))
                 
                  .Row = i + 1
                 .Col = E_FPS4.E_Origin
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(11).Value), "", listRS.Fields(11).Value))
                 
                 .Row = i + 1
                 .Col = E_FPS4.E_DateCode
                 '把日期转为Oracle周
                   .Text = "" & CStr(IIf(IsNull(listRS.Fields(12).Value), "", listRS.Fields(12).Value))

                    .Row = i + 1
                 .Col = E_FPS4.E_LotNumber
                 .Text = "" & CStr(IIf(IsNull(listRS.Fields(13).Value), "", listRS.Fields(13).Value))
                        
                 

'                 .Row = i + 1
'                 .Col = E_FPS4.E_Qty
'                 .Text = CStr(listRS.fields(5).Value)
'
'                   .Row = i + 1
'                 .Col = E_FPS4.E_GDQty
'                 .Text = CStr(listRS.fields(6).Value)
'
'                .Row = i + 1
'                 .Col = E_FPS3.E_NGQty
'                 '查询质量有没有报废
'                 baofeiQty = GetQty37Baofei(CStr(listRS.fields(8).Value))
'                 If baofeiQty = 0 Then
'                 .Text = ""
'                 Else
'                   .Text = CStr(baofeiQty)
'                 End If
'
'
'                  .Row = i + 1
'                 .Col = E_FPS4.E_Status
'                 '判断有没有出完
'
'                 outQty = GetQty37OutQty(CStr(listRS.fields(8).Value))
'
'                 If outQty + baofeiQty = CLng(listRS.fields(5).Value) Then
'                 .Text = "X"
'                 Else
'                  .Text = ""
'
'                 End If
'
'
'                  .Row = i + 1
'                 .Col = E_FPS4.E_LotNumber
'                 .Text = CStr(listRS.fields(8).Value)
'
                 
                
        
        End With
    
NextRecord:
       
        listRS.MoveNext

    Next


End If





End Sub


Private Sub CmdPOShipOut_Click()



Dim tempBillNo As String
Dim custNameTemp As String
Dim baofeiQty As Long
Dim outQty  As Long
Dim allQty As Long
Dim sendTimeTemp As Long
Dim maxBillNoTemp As String



If ComboShipICI.Text = "ICI" Then

CmdPOShipOutICI

Exit Sub

End If



tempBillNo = UCase(Trim(Txt37BillNoShip.Text))


If tempBillNo = "" Then
    MsgBox "请输入出货单据编号!", vbInformation, "友情提示"
    Exit Sub
End If



 Dim judgeEmp As Boolean

judgeEmp = JudgeSemtechBillNo(tempBillNo)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If


Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer

Dim lotIdTemp As String

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

'On Error GoTo ErrHandler
'查询报表名的序号

fileNo = Format(Now, "YYYYMMDD_HHMM") & "HTKS_SHIP"

'20151106_0846HTKS_RECEIPT.csv

Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path37 & "\" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "Event,PO Number,PO Line Item Number,Production Order Number,Manufacturing Plant,Ship to Plant,Material Number,Semtech Lot Number,Order Close Indicator,Total Ship Quantity,Current Ship Date,Country of Origin,Date Code,Vendor Site,Assembly Revision Level,Test Program Number,Test Program  Revision,Retest,Lot Combine,NCMR,Vendor Lot Number,Vendor Lot Number 2,Vendor Lot Number 3,Vendor Lot Number 4,Vendor Lot Number 5,Waiver ID,PO Number 1,Semtech Lot Number 1,PO Quantity 1,PO Number 2,Semtech Lot Number 2,PO Quantity 2" & vbCrLf
    '明细数据
    

    
  strDatas = strDatas + ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,," & vbCrLf


'
'strSql = "  select distinct  'START' as wostart,c.po_num,c.po_item,c.mpn,to_char(a.erpcreatedate,'YYYYMMDD'),c.SOURCE_MTRL_SLOC  from ib_wohistory a ,ib_waferlist b ,customeroitbl_test c ,mappingdata37 d" & _
'" Where b.OrderName = a.OrderName and a.customer='37' and a.erpcreatedate>to_Date('2016-03-26','YYYY-MM-DD') and c.source_batch_id=b.waferlot " & _
'" and d.batch=c.source_batch_id and d.purchaseno=c.po_num and a.lot_status is null "




'strSql = " SELECT  'COMPLETE' as Event ,c.PO_NUM,c.PO_ITEM,c.MPN,'' as OrderClose,c.CURRENT_WAFER_QTY,COUNT(b.流程卡编号) as Quantity,'' as ScrapQuantity, " & _
'" a.工单号 FROM  [ERPdata].[dbo].[tblStockMove] a, [ERPdata].[dbo].[tblStockMovesub] b ,[ERPBASE].[dbo].[tblCustomerOI] c " & _
'" where a.单据编号='" + tempBillNo + "' and a.客户代码='37' " & _
'" and b.单据编号=a.单据编号 and b.工单号=a.工单号 " & _
'" and c.CUSTOMERSHORTNAME='37' and c.SOURCE_BATCH_ID=a.工单号 and c.PO_NUM <>'' " & _
'" group by   c.PO_NUM,c.PO_ITEM,c.MPN,c.CURRENT_WAFER_QTY, a.工单号 "



strSql = "  SELECT  'SHIP' as Event ,c.PO_NUM,c.PO_ITEM,c.MPN," & _
" substring(ltrim(c.offshore_asm_company),1,4) as MPlant,substring(ltrim(c.offshore_test_company),1,4) as SPlant," & _
" c.MPN_DESC,c.source_mtrl_sloc,'' as OrderClose,COUNT(b.流程卡编号) as Quantity," & _
" CONVERT(char(8), a.操作日期, 112) as sdate,'CN' AS COrigin,CONVERT(varchar(100), d.ERPCREATEDATE, 23) as datecode ,a.工单号 " & _
" FROM  [erpdata].[dbo].[tblStockMove] a, [erpdata].[dbo].[tblStockMovesub] b ,[ERPBASE].[dbo].[tblCustomerOI] c ,[erpdata].[dbo].[tblTSVworkorder] d,[erpdata].[dbo].[tblTSVwaferlist] e " & _
" where a.单据编号='" + tempBillNo + "'and a.客户代码='37' and b.单据编号=a.单据编号 and b.工单号=a.工单号 " & _
" and c.CUSTOMERSHORTNAME='37' and c.SOURCE_BATCH_ID=a.工单号 and c.PO_NUM <>'' " & _
" and d.ORDERNAME=e.ORDERNAME and e.WAFERLOT=a.工单号 and e.WAFERID=b.流程卡编号 " & _
" group by   c.PO_NUM,c.PO_ITEM,c.MPN,c.CURRENT_WAFER_QTY, a.工单号,c.offshore_asm_company,c.offshore_test_company," & _
" c.MPN_DESC , c.source_mtrl_sloc, CONVERT(Char(8), a.操作日期, 112), d.ERPCreateDate"




    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    
If INIadoCon2.State = 0 Then
INIConnectSTART2
End If
    
    
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        baofeiQty = 0
        outQty = 0
        allQty = 0
        allQty = CLng(rs.Fields(5).Value)
        lotIdTemp = Trim(CStr(rs.Fields(rs.Fields.Count - 1).Value))
        baofeiQty = GetQty37Baofei(lotIdTemp)
        outQty = GetQty37OutQty(lotIdTemp)
        
        sendTimeTemp = 0
         maxBillNoTemp = ""
        
        
        
        For j = 0 To rs.Fields.Count - 2
              
           
                
            If j = 8 Then
            
   
                 allQty = GetQty37OIAllQty(Trim(lotIdTemp), CStr(rs.Fields(1).Value))
                 
                 If baofeiQty + outQty = allQty Then
                 
                 
                 '判断今天有几次发货，如果有多次，则当最后一次才显示X
                 sendTimeTemp = GetQty37OutTimes(lotIdTemp)
                 
                        If sendTimeTemp > 1 Then
                        
                        '判断这个发货单号是否为最大的发货单号，是则为X ,否则为0
                        maxBillNoTemp = GetQty37OutMaxBill(lotIdTemp)
                        
                                If maxBillNoTemp = UCase(Trim(Txt37BillNoShip.Text)) Then
                                strColData = strColData + "X" + ","
                                Else
                                strColData = strColData + "" + ","
                                End If
        
                        End If
                 
                 
                 Else
                  strColData = strColData + "" + ","
                 
                 End If
                 
                 
              ElseIf j = 12 Then
            
                  strColData = strColData + GetDateToCode(CStr(rs.Fields(12).Value)) + ","

             Else
             
             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
             
             End If
           
        Next
        
        strColData = strColData + ",,,,,,,,,,,,,,,,,"
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
'    '发邮件
'    Dim strRecipient    As String
'    Dim strRecipientCC  As String
'
'    strRecipient = "jiayun.zhang@qtechglobal.com"
'    strRecipientCC = "wanli.ma@qtechglobal.com"
'
'    Call MailDetailHD("HD 发货报表", strRecipient, g_Path & "\" & "HD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
'
'    '把发送记录保存到DB中
    
'    Dim sqlTemp2 As String
'
'    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'HD') "
'
'    Call AddSql2(sqlTemp2)
'
'    MsgBox "发送成功！", vbInformation, "友情提示"
'
'
'ErrHandler:
'    Set FSO = Nothing


MsgBox "导出成功！", vbInformation, "友情提示"






End Sub



Private Sub CmdPOShipOutICI()



Dim tempBillNo As String
Dim custNameTemp As String
Dim baofeiQty As Long
Dim outQty  As Long
Dim allQty As Long
Dim sendTimeTemp As Long
Dim maxBillNoTemp As String



Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer

Dim lotIdTemp As String

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String
Dim recordQty As Long


'On Error GoTo ErrHandler
'查询报表名的序号

fileNo = Format(Now, "YYYYMMDD_HHMM") & "HTKS_SHIP"

'20151106_0846HTKS_RECEIPT.csv

Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path37 & "\" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "Event,PO Number,PO Line Item Number,Production Order Number,Manufacturing Plant,Ship to Plant,Material Number,Semtech Lot Number,Order Close Indicator,Total Ship Quantity,Current Ship Date,Country of Origin,Date Code,Vendor Site,Assembly Revision Level,Test Program Number,Test Program  Revision,Retest,Lot Combine,NCMR,Vendor Lot Number,Vendor Lot Number 2,Vendor Lot Number 3,Vendor Lot Number 4,Vendor Lot Number 5,Waiver ID,PO Number 1,Semtech Lot Number 1,PO Quantity 1,PO Number 2,Semtech Lot Number 2,PO Quantity 2" & vbCrLf
    '明细数据
    

    
  strDatas = strDatas + ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,," & vbCrLf




'strSql = " select X.Event,X.po_num,X.po_item,X.mpn,X.MPlant,X.SPlant,X.MPN_DESC,X.source_batch_id,'' as OrderClose,  sum(X.moveinqty) as qty, X.sdate, " & _
'"    'CN' AS COrigin,X.datecode, X.source_batch_id as 工单号 ,X.containername " & _
'"  from (select distinct a.containername, 'SHIP' as Event, wo.po_num,wo.po_item,wo.mpn,'' as OrderClose, wo.die_qty, a.moveinqty, " & _
'"  '' as ScrapQuantity, wo.source_batch_id, substr(ltrim(wo.offshore_asm_company),1,4) as MPlant, substr(ltrim(wo.offshore_test_company),1,4) as SPlant, " & _
'"  wo.mpn_desc, to_char(a.moveintimestamp,'YYYYMMDD') as sdate, to_char(ibwo.erpcreatedate,'YYWW') as datecode " & _
' "  from a_wiplothistory a , a_wiplotdetailshistory b, container conn, mfgorder mfg, ib_wohistory ibwo, a_lotwafers waf,customeroitbl_test wo " & _
'"   where a.specname = '5272' and a.creationtimestamp > to_date('2016-04-20', 'YYYY-MM-DD') and b.wiplothistoryid = a.wiplothistoryid " & _
'" and conn.containername = a.containername and waf.waferscribenumber = substr(conn.containername, 1,instr(conn.containername, '-') - 1) " & _
'"  and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername   " & _
'"   and wo.customershortname = '37' " & _
'"   and wo.source_batch_id = waf.wafernumber and a.containername like '%-A%'  and a.creationtimestamp  >sysdate-7  and a.containername<>'10001-A-01' " & _
'"  union select distinct a.containername,'SHIP' as Event, wo.po_num,wo.po_item, wo.mpn,'' as OrderClose,wo.die_qty,waf.ndpw, '' as ScrapQuantity, " & _
'" wo.source_batch_id, substr(ltrim(wo.offshore_asm_company), 1, 4) as MPlant,substr(ltrim(wo.offshore_test_company), 1, 4) as SPlant,wo.mpn_desc,to_char(a.moveintimestamp, 'YYYYMMDD') as sdate, to_char(ibwo.erpcreatedate, 'YYWW') as datecode " & _
'"  from a_wiplothistory a,a_wiplotdetailshistory b,container conn,mfgorder mfg,ib_wohistory  ibwo,a_lotwafers waf, customeroitbl_test wo " & _
'" where a.specname = '5272' and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername " & _
'"   and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername and wo.source_batch_id = waf.wafernumber and wo.customershortname = '37' " & _
'"   and a.containername not like '%-F%' and waf.containerid=conn.containerid and a.creationtimestamp > sysdate - 7 " & _
'" ) X " & _
'"  group by X.Event,  X.po_num,X.po_item, X.mpn, X.OrderClose,X.die_qty,X.ScrapQuantity,X.source_batch_id, X.MPlant, X.SPlant , X.mpn_desc, X.sdate, X.DateCode,X.containername  order by  X.containername "
'
'
    
'  strSql = "select X.Event,X.po_num,X.po_item,X.mpn,X.MPlant,X.SPlant,X.mpn_desc,X.source_batch_id,X.OrderClose,sum(X.qty) as qty,x.sdate,'CN' AS COrigin, " & _
'" X.date_code, X.source_batch_id as 工单号,' ' as containername from ( " & _
'" select 'SHIP' as Event  ,qty ,b.po_num,b.po_item, b.mpn,'' as OrderClose ,b.die_qty ,'' as ScrapQuantity,b. source_batch_id ," & _
'" substr(ltrim(b.offshore_asm_company),1,4) as MPlant,substr(ltrim(b.offshore_test_company),1,4) as SPlant,mpn_desc,to_char(sysdate,'YYYYMMDD') as sdate, " & _
'" b.Date_Code from  cus37_5272qtyNotMerg a , customeroitbl_test b " & _
'" where b.mtrl_num=substr( containername,1,instr(containername,'-A')-1) " & _
'" ) x group by X.Event,X.po_num,X.po_item,X.mpn,X.MPlant,X.SPlant,X.mpn_desc,X.OrderClose,X.die_qty,X.ScrapQuantity,X.source_batch_id,x.sdate,X.date_code" & _
'" union all select X.Event,X.po_num,X.po_item,X.mpn,X.MPlant,X.SPlant,X.mpn_desc,X.source_batch_id,X.OrderClose,sum(X.qty) as qty,x.sdate,'CN' AS COrigin," & _
'" X.date_code, X.source_batch_id as 工单号,X.containername from ( " & _
'" select 'SHIP' as Event  ,qty ,b.po_num,b.po_item, b.mpn,'' as OrderClose ,b.die_qty ,'' as ScrapQuantity,b. source_batch_id ," & _
'" substr(ltrim(b.offshore_asm_company),1,4) as MPlant,substr(ltrim(b.offshore_test_company),1,4) as SPlant,mpn_desc,to_char(sysdate,'YYYYMMDD') as sdate," & _
'" b.Date_Code , a.containername from  cus37_5272qtyMerg a , customeroitbl_test b " & _
'" where b.mtrl_num=substr( containername,1,instr(containername,'-A')-1) " & _
'" ) x group by X.Event,X.po_num,X.po_item,X.mpn,X.MPlant,X.SPlant,X.mpn_desc,X.OrderClose,X.die_qty,X.ScrapQuantity,X.source_batch_id,x.sdate,X.date_code,X.containername  "

 strSql = "select distinct X.Event,X.po_num,X.po_item,X.mpn,'3179' as MPlant,X.SPlant,X.mpn_desc,X.source_batch_id,X.OrderClose,X.qty ,x.sdate,'CN' AS COrigin, " & _
" X.date_code, X.source_batch_id as 工单号, X.containername from  cus37_5272qtydetails X  "



    
    
    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    
If Cnn.State = 0 Then
ConOracle
End If

    
    
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        baofeiQty = 0
        outQty = 0
        allQty = 0
        'allQty = CLng(Rs.fields(5).Value)
       ' lotIDTemp = Trim(CStr(Rs.fields(Rs.fields.Count - 2).Value))
        lotIdTemp = "" & CStr(IIf(IsNull(rs.Fields(rs.Fields.Count - 2).Value), " ", rs.Fields(rs.Fields.Count - 2).Value))
        baofeiQty = GetQty37NGQty(lotIdTemp)
        outQty = GetQty37OutQtyICI(lotIdTemp)
        
        sendTimeTemp = 0
         maxBillNoTemp = ""
        recordQty = CLng(rs.Fields(9).Value)
        
        
        
        For j = 0 To rs.Fields.Count - 3
              
           
                
            If j = 8 Then
            
   
                 allQty = GetQty37OIAllQty(Trim(lotIdTemp), "" & CStr(IIf(IsNull(rs.Fields(rs.Fields.Count - 2).Value), " ", rs.Fields(rs.Fields.Count - 2).Value)))
                 
'                 If baofeiQty + outQty = allQty Then
                 
'
'                 '判断今天有几次发货，如果有多次，则当最后一次才显示X
'                 sendTimeTemp = GetQty37OutTimes(lotidTemp)
'
'                        If sendTimeTemp > 1 Then
'
'                        '判断这个发货单号是否为最大的发货单号，是则为X ,否则为0
'                        maxBillNoTemp = GetQty37OutMaxBill(lotidTemp)
'
'                                If maxBillNoTemp = UCase(Trim(Txt37BillNoShip.Text)) Then
'                                strColData = strColData + "X" + ","
'                                Else
'                                strColData = strColData + "" + ","
'                                End If
'
'                        End If
                        
                   If baofeiQty + outQty >= allQty Then
                    
                         If recordQty < 20000 And Right(CStr(rs.Fields(14).Value), 5) <> "-A-01" Then
                                 strColData = strColData + "X" + ","
                        
                         Else
                         strColData = strColData + "" + ","
                        
                        End If
                 
                 Else
                     strColData = strColData + "" + ","
                 
                 End If
                 
                 
              'ElseIf j = 12 Then
            
                 'strColData = strColData + "" & CStr(IIf(IsNull(listRS.fields(12).Value), "NULL", listRS.fields(12).Value)) + ","

             Else
             
             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
             
             End If
           
        Next
        
        'jiayun add merge lotid
        Dim containerLotTemp As String
        Dim lotidCount As Integer
        
        Dim poNumber1 As String
        Dim semtechLot1 As String
        Dim poQuantity1 As Long
        Dim poNumber2 As String
        Dim semtechLot2 As String
        Dim poQuantity2 As String
        
        
        
        
        containerLotTemp = rs.Fields(14).Value
        
        lotidCount = Get37MerLotCounts(containerLotTemp)
        
        If lotidCount > 1 Then
        
'        poNumber2 = Rs.fields(1).Value
'        semtechLot2 = Rs.fields(7).Value
'        poQuantity2 = allQty
        
        poNumber2 = ""
        semtechLot2 = ""
        poQuantity2 = ""
        
                
                Set list2RS = Get37MergeLotDetails(containerLotTemp)
                
                If list2RS.RecordCount > 0 Then
                
                 poNumber1 = list2RS.Fields(1).Value
                 semtechLot1 = list2RS.Fields(0).Value
                 poQuantity1 = CLng(list2RS.Fields(2).Value)
                 
                list2RS.Close
                Set list2RS = Nothing
                End If
       
        strColData = strColData + ",,,,," & "Y," & ",,,,,,," & poNumber1 & "," & semtechLot1 & "," & poQuantity1 & "," & poNumber2 & "," & semtechLot2 & "," & poQuantity2
       
        Else
        
        
        strColData = strColData + ",,,,,,,,,,,,,,,,,"
        
        
        End If
        
        
        
        strColData = strColData + ",,,,,,,,,,,,,,,,,"
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    '根据市场需求，如果此批有被合到别的批次上面，这批次就不显示了；所以调用函数来进行数据整理
    strRowData = Get_ShipInfoZL(strRowData)
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
'    '发邮件
'    Dim strRecipient    As String
'    Dim strRecipientCC  As String
'
'    strRecipient = "jiayun.zhang@qtechglobal.com"
'    strRecipientCC = "wanli.ma@qtechglobal.com"
'
'    Call MailDetailHD("HD 发货报表", strRecipient, g_Path & "\" & "HD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
'
'    '把发送记录保存到DB中
    
'    Dim sqlTemp2 As String
'
'    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'HD') "
'
'    Call AddSql2(sqlTemp2)
'
'    MsgBox "发送成功！", vbInformation, "友情提示"
'
'
'ErrHandler:
'    Set FSO = Nothing


MsgBox "导出成功！", vbInformation, "友情提示"






End Sub


'数据重新整理函数
Private Function Get_ShipInfoZL(strINfo As String) As String
Dim i               As Long
Dim j               As Integer
Dim strDataInfo     As String
Dim strRowTmp()     As String
Dim strColTmp()     As String
Dim strColTmp1()    As String
Dim strPo           As String
Dim strKey          As String
Dim strQty          As String
Dim b_flag          As Boolean

    strDataInfo = ""
    If Trim(strINfo) = "" Then Exit Function    '没有数据就退出
    strRowTmp() = Split(strINfo, vbCrLf)        '首先数据按vbCrlf进行拆分得到每行，然后对每行进行分析，数据整合
    For i = 0 To UBound(strRowTmp)
        b_flag = False
        strPo = ""
        strKey = ""
        strQty = ""
        If Trim(strRowTmp(i)) = "" Then Exit For    '如果行数据为空表示到最后了，就退出For循环
        '拆分行，得到每列的值
        strColTmp() = Split(strRowTmp(i), ",")
        '看18列是否有标志是Y的，外循环不比较Y的,这里主要是和内循环中带Y的进行比较
        If Trim(strColTmp(18)) <> "Y" Then
            '如果外循环中行没有Y,就记录下PO,LOT
            strPo = Trim$(strColTmp(1))
            strKey = Trim$(strColTmp(7))
            strQty = Trim$(strColTmp(9))
            For j = 0 To UBound(strRowTmp)
                strColTmp1() = Split(strRowTmp(j), ",")
                If Trim$(strColTmp1(18)) = "Y" Then
                    '表示外循环中的PO和Lot已经被合并到此批次中，外循环中的就要取消行
                    If strPo = Trim(strColTmp1(26)) And strKey = Trim(strColTmp1(27)) And strQty = Trim$(strColTmp1(28)) Then
                        b_flag = True
                        Exit For
                    End If
                End If
            Next j
        End If
        If b_flag = False Then  '表示没有被合并过
            strDataInfo = strDataInfo + strRowTmp(i) + vbCrLf
        End If
    Next i
    '返回到赋值变量中
    Get_ShipInfoZL = strDataInfo
    
End Function

Private Sub CmdPOStart_Click()


'明细数据
Dim i As Integer
Dim waferIdTemp As String
Dim woType As String


    Set listRS = GetFps37POStart()


If listRS.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认", vbInformation, "友情提示"
    Exit Sub
    
Else


    fps(2).MaxRows = listRS.RecordCount
    For i = 0 To listRS.RecordCount - 1
         


         
         With fps(2)
                 .Row = i + 1
                 .Col = E_FPS2.E_Event
                 .Text = CStr(listRS.Fields(0).Value)
                 
                .Row = i + 1
                 .Col = E_FPS2.E_PONumber
                .Text = CStr(IIf(IsNull(listRS.Fields(1).Value), "", listRS.Fields(1).Value))
                
                
                  .Row = i + 1
                 .Col = E_FPS2.E_POLineItem
                 .Text = CStr(IIf(IsNull(listRS.Fields(2).Value), "", listRS.Fields(2).Value))
                 
                  .Row = i + 1
                 .Col = E_FPS2.E_Pt2
                 .Text = CStr(IIf(IsNull(listRS.Fields(3).Value), "", listRS.Fields(3).Value))
                 
                 
                  .Row = i + 1
                 .Col = E_FPS2.E_StartDate
                 .Text = CStr(IIf(IsNull(listRS.Fields(4).Value), "", listRS.Fields(4).Value))
                 
                .Row = i + 1
                 .Col = E_FPS2.E_LotNumber
                 .Text = CStr(IIf(IsNull(listRS.Fields(5).Value), "", listRS.Fields(5).Value))
                 
                
        
        End With
    
NextRecord:
       
        listRS.MoveNext

    Next


End If


End Sub

Private Sub CmdPOStartOut_Click()



Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

'On Error GoTo ErrHandler
'查询报表名的序号

fileNo = Format(Now, "YYYYMMDD_HHMM") & "HTKS_PO_START"

'20151106_0846HTKS_RECEIPT.csv

Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path37 & "\" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "Event,PO Number,PO Line Item Number,Production Order Number,Start Date,Vendor Lot Number,客户机种(厂内),华天料号(厂内),Bag#(厂内)" & vbCrLf
    '明细数据
    

    
  strDatas = strDatas + ",,,,," & vbCrLf



'strSql = "  select distinct  'START' as wostart,c.po_num,c.po_item,c.mpn,to_char(a.erpcreatedate,'YYYYMMDD'),c.source_batch_id  from ib_wohistory a ,ib_waferlist b ,customeroitbl_test c ,mappingdata37 d" & _
'" Where b.OrderName = a.OrderName and a.customer='37' and a.erpcreatedate>to_Date('2016-03-26','YYYY-MM-DD') and c.source_batch_id=b.waferlot " & _
'" and d.batch=c.source_batch_id and d.purchaseno=c.po_num and a.lot_status is null "


strSql = "  select distinct 'START',d.po_num ,d.po_item,d.mpn,a.erpcreatedate,f.firstname ,d.mpn_desc,a.product,d.mtrl_num from ib_wohistory a, " & _
" ib_waferlist b ,mappingdatatest c ,customeroitbl_test d ,mfgorder e,container f where a.ordername = b.ordername and c.substrateid = b.waferid " & _
" and to_char(d.id) = c.filename and e.mfgordername = a.ordername and e.mfgorderid = f.mfgorderid and d.customershortname = '37'" & _
" and a.erpcreatedate > sysdate - 3 and d.po_num is not null order by a.erpcreatedate desc "
  
  


    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    
    If Cnn.State = 0 Then
    ConOracle
    End If
    
    
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then
    
     MsgBox "明细表中没有相关数据，请确认", vbInformation, "友情提示"
  
    
     Exit Sub
    
    End If
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        For j = 0 To rs.Fields.Count - 1

             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
'    '发邮件
'    Dim strRecipient    As String
'    Dim strRecipientCC  As String
'
'    strRecipient = "jiayun.zhang@qtechglobal.com"
'    strRecipientCC = "wanli.ma@qtechglobal.com"
'
'    Call MailDetailHD("HD 发货报表", strRecipient, g_Path & "\" & "HD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
'
'    '把发送记录保存到DB中
    
'    Dim sqlTemp2 As String
'
'    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'HD') "
'
'    Call AddSql2(sqlTemp2)
'
'    MsgBox "发送成功！", vbInformation, "友情提示"
'
'
'ErrHandler:
'    Set FSO = Nothing


MsgBox "导出成功！", vbInformation, "友情提示"




End Sub

Private Sub CmdWaferRecOut_Click()

'2016-03-29 添加有没有打钩

Dim strLotIDTemp As String
Dim m As Integer

strLotIDTemp = ""

With fps(0)

For m = 1 To .MaxRows

    .Row = m
    .Col = 7
    If .Text = "1" Then

   
    .Row = m
    .Col = 6
    
     strLotIDTemp = strLotIDTemp & .Text & "','"
    
    End If

Next m

End With



If strLotIDTemp = "" Then
 
 MsgBox "请先选择LotId !"
 Exit Sub
 
 Else
 
 strLotIDTemp = Mid(strLotIDTemp, 1, Len(strLotIDTemp) - 3)
 
 End If



Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

'On Error GoTo ErrHandler
'查询报表名的序号

fileNo = Format(Now, "YYYYMMDD_HHMM") & "HTKS_RECEIPT"

'20151106_0846HTKS_RECEIPT.csv

Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path37 & "\" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "Event,PO Number,PO Line Item Number,Material Number,Quantity,Fab Lot Number" & vbCrLf
    '明细数据
    
  strDatas = strDatas + ",,,,," & vbCrLf

  
  strSql = "  select 'PO_RECEIPT' as Event, '' as purchaseno,'' as purchaseorderlineitem,a.devicename,a.wf,a.batch from MAPPINGDATA37 a  where a.status=0    and   a.batch in ('" & strLotIDTemp & "') order by a.qtech_created_date "
       
       
    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    
    If Cnn.State = 0 Then
    ConOracle
    End If
    
    
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        For j = 0 To rs.Fields.Count - 1

             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
'    '发邮件
'    Dim strRecipient    As String
'    Dim strRecipientCC  As String
'
'    strRecipient = "jiayun.zhang@qtechglobal.com"
'    strRecipientCC = "wanli.ma@qtechglobal.com"
'
'    Call MailDetailHD("HD 发货报表", strRecipient, g_Path & "\" & "HD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
'
'    '把发送记录保存到DB中
    
'    Dim sqlTemp2 As String
'
'    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'HD') "
'
'    Call AddSql2(sqlTemp2)
'
'    MsgBox "发送成功！", vbInformation, "友情提示"
'
'
'ErrHandler:
'    Set FSO = Nothing


MsgBox "导出成功！", vbInformation, "友情提示"

End Sub

Private Sub CmdWaferRecQuery_Click()



'明细数据
Dim i As Integer
Dim waferIdTemp As String
Dim woType As String


    Set listRS = GetFps37WaferRec()


If listRS.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认"
    Exit Sub
    
Else


    fps(0).MaxRows = listRS.RecordCount
    For i = 0 To listRS.RecordCount - 1
         


         
         With fps(0)
                 .Row = i + 1
                 .Col = E_FPS0.E_Event
                 .Text = CStr(listRS.Fields(0).Value)
                 
                .Row = i + 1
                 .Col = E_FPS0.E_PONumber
                .Text = CStr(listRS.Fields(1).Value)
                
                
                  .Row = i + 1
                 .Col = E_FPS0.E_POLineItem
                 .Text = CStr(listRS.Fields(2).Value)
                 
                  .Row = i + 1
                 .Col = E_FPS0.E_PT
                 .Text = CStr(listRS.Fields(3).Value)
                 
                 
                  .Row = i + 1
                 .Col = E_FPS0.E_QTY
                 .Text = CStr(listRS.Fields(4).Value)
                 
                  .Row = i + 1
                 .Col = E_FPS0.E_LotID
                 .Text = "" & listRS.Fields(5).Value
                  
        
        End With
    
NextRecord:
       
        listRS.MoveNext

    Next


End If



End Sub

Private Sub IniFpsWaferRec()
    With fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
        .MaxRows = 0
        
        '砞竚Α
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        
         .Col = E_FPS0.E_OK
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        
        
        .SetText E_FPS0.E_Event, 0, "Event"
        .SetText E_FPS0.E_PONumber, 0, "PO Number"
        .SetText E_FPS0.E_POLineItem, 0, "PO Line Item Number"
        .SetText E_FPS0.E_PT, 0, "Material Number"
        .SetText E_FPS0.E_QTY, 0, "Quantity"
        .SetText E_FPS0.E_LotID, 0, "Fab Lot Number"
         .SetText E_FPS0.E_OK, 0, "选择"
        
        
        .ColWidth(E_FPS0.E_Event) = 15
        .ColWidth(E_FPS0.E_PONumber) = 15
        .ColWidth(E_FPS0.E_POLineItem) = 15
        .ColWidth(E_FPS0.E_PT) = 15
        .ColWidth(E_FPS0.E_QTY) = 15
        .ColWidth(E_FPS0.E_LotID) = 15
        .ColWidth(E_FPS0.E_OK) = 10

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
         .Col = E_FPS0.E_OK
        .Lock = False

    
        .ReDraw = True
    End With
    
    
    

End Sub


Private Sub IniFpsPOCommit()
    With fps(1)
        .ReDraw = False
        .MaxCols = E_FPS1.E_End - 1
        .MaxRows = 0
        
        '砞竚Α
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        
        .SetText E_FPS1.E_Event, 0, "Event"
        .SetText E_FPS1.E_PONumber, 0, "PO Number"
        .SetText E_FPS1.E_POLineItem, 0, "PO Line Item Number"
        .SetText E_FPS1.E_Pt2, 0, "Production Order Number"
        .SetText E_FPS1.E_ComDate, 0, "Commit Date"
        .SetText E_FPS1.E_LotID, 0, "LotID"

        .ColWidth(E_FPS1.E_Event) = 15
        .ColWidth(E_FPS1.E_PONumber) = 15
        .ColWidth(E_FPS1.E_POLineItem) = 15
        .ColWidth(E_FPS1.E_Pt2) = 15
        .ColWidth(E_FPS1.E_ComDate) = 15
         .ColWidth(E_FPS1.E_LotID) = 15


        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
    
        .ReDraw = True
    End With
    
    
    

End Sub


Private Sub IniFpsPOReCommit()
    With fps(5)
        .ReDraw = False
        .MaxCols = E_FPS5.E_End - 1
        .MaxRows = 0
        
        '砞竚Α
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        
        .SetText E_FPS5.E_Event, 0, "Event"
        .SetText E_FPS5.E_PONumber, 0, "PO Number"
        .SetText E_FPS5.E_POLineItem, 0, "PO Line Item Number"
        .SetText E_FPS5.E_Pt2, 0, "Production Order Number"
        .SetText E_FPS5.E_ComDate, 0, "Commit Date"
        .SetText E_FPS5.E_LotID, 0, "LotID"

        .ColWidth(E_FPS5.E_Event) = 15
        .ColWidth(E_FPS5.E_PONumber) = 15
        .ColWidth(E_FPS5.E_POLineItem) = 15
        .ColWidth(E_FPS5.E_Pt2) = 15
        .ColWidth(E_FPS5.E_ComDate) = 15
         .ColWidth(E_FPS5.E_LotID) = 15


        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
    
        .ReDraw = True
    End With
    
    
    

End Sub



Private Sub IniFpsPOStart()
    With fps(2)
        .ReDraw = False
        .MaxCols = E_FPS2.E_End - 1
        .MaxRows = 0
        
        '砞竚Α
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        
        .SetText E_FPS2.E_Event, 0, "Event"
        .SetText E_FPS2.E_PONumber, 0, "PO Number"
        .SetText E_FPS2.E_POLineItem, 0, "PO Line Item Number"
        .SetText E_FPS2.E_Pt2, 0, "Production Order Number"
        .SetText E_FPS2.E_StartDate, 0, "Start Date"
        .SetText E_FPS2.E_LotNumber, 0, "Vendor Lot Number"

        .ColWidth(E_FPS2.E_Event) = 15
        .ColWidth(E_FPS2.E_PONumber) = 15
        .ColWidth(E_FPS2.E_POLineItem) = 15
        .ColWidth(E_FPS2.E_Pt2) = 15
        .ColWidth(E_FPS2.E_StartDate) = 15
        .ColWidth(E_FPS2.E_LotNumber) = 15

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
    
        .ReDraw = True
    End With
    
    
    

End Sub


Private Sub IniFpsPOComplete()
    With fps(3)
        .ReDraw = False
        .MaxCols = E_FPS3.E_End - 1
        .MaxRows = 0
        
        '砞竚Α
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        
        .SetText E_FPS3.E_Event, 0, "Event"
        .SetText E_FPS3.E_PONumber, 0, "PO Number"
        .SetText E_FPS3.E_POLineItem, 0, "PO Line Item Number"
        .SetText E_FPS3.E_Pt2, 0, "Production Order Number"
        .SetText E_FPS3.E_STATUS, 0, "Order Close Indicator"
        .SetText E_FPS3.E_QTY, 0, "Quantity"
        .SetText E_FPS3.E_GDQty, 0, "Yield Quantity"
        .SetText E_FPS3.E_NGQty, 0, "Scrap Quantity"
        .SetText E_FPS3.E_LotNumber, 0, "LotID"


        .ColWidth(E_FPS3.E_Event) = 15
        .ColWidth(E_FPS3.E_PONumber) = 15
        .ColWidth(E_FPS3.E_POLineItem) = 15
        .ColWidth(E_FPS3.E_Pt2) = 15
        .ColWidth(E_FPS3.E_STATUS) = 8
        .ColWidth(E_FPS3.E_QTY) = 8
         .ColWidth(E_FPS3.E_GDQty) = 8
        .ColWidth(E_FPS3.E_NGQty) = 8
         .ColWidth(E_FPS3.E_LotNumber) = 15
   
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
    
        .ReDraw = True
    End With
    
    
    

End Sub




Private Sub IniFpsPOShip()
    With fps(4)
        .ReDraw = False
        .MaxCols = E_FPS4.E_End - 1
        .MaxRows = 0
        
        '砞竚Α
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        
        .SetText E_FPS4.E_Event, 0, "Event"
        .SetText E_FPS4.E_PONumber, 0, "PO Number"
        .SetText E_FPS4.E_POLineItem, 0, "PO Line Item Number"
        .SetText E_FPS4.E_Pt2, 0, "Production Order Number"
        .SetText E_FPS4.E_MPlant, 0, "Manufacturing Plant"
        .SetText E_FPS4.E_SPlant, 0, "Ship to Plant"
        .SetText E_FPS4.E_MNumber, 0, "Material Number"
        .SetText E_FPS4.E_SLotNumber, 0, "Semtech Lot Number"
        .SetText E_FPS4.E_STATUS, 0, "Order Close Indicator"
        
        .SetText E_FPS4.E_GDQty, 0, "Total Ship Quantity"
        .SetText E_FPS4.E_SDate, 0, "Current Ship Date"
        .SetText E_FPS4.E_Origin, 0, "Country of Origin"
        .SetText E_FPS4.E_DateCode, 0, "Date Code"
        .SetText E_FPS4.E_LotNumber, 0, "LotID"
        


        .ColWidth(E_FPS4.E_Event) = 10
        .ColWidth(E_FPS4.E_PONumber) = 10
        .ColWidth(E_FPS4.E_POLineItem) = 10
        .ColWidth(E_FPS4.E_Pt2) = 10
        
         .ColWidth(E_FPS4.E_MPlant) = 10
        .ColWidth(E_FPS4.E_SPlant) = 10
         .ColWidth(E_FPS4.E_MNumber) = 10
        .ColWidth(E_FPS4.E_SLotNumber) = 10
         .ColWidth(E_FPS4.E_STATUS) = 10
        
        
        
        
        .ColWidth(E_FPS4.E_GDQty) = 10
        .ColWidth(E_FPS4.E_SDate) = 10
         .ColWidth(E_FPS4.E_Origin) = 10
        .ColWidth(E_FPS4.E_DateCode) = 10
        .ColWidth(E_FPS4.E_LotNumber) = 10
   
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
    
        .ReDraw = True
    End With
    
    
    

End Sub



Private Sub SaveFileSendTest()
Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer
Dim rs          As New ADODB.Recordset

On Error GoTo ErrHandler
    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path_GR & "\" & "QT_FG_CSP_" & Format(g_Date, "YYYYMMDD") & "_" & Format(g_Date, "hhmmss") & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "PO_num,PO_Item,Previous_Batch_ID,Previous_Mtrl_Num,Batch_ID,mtrl_num,mtrl_desc,Mtrl_num_Mtrlgrp,Output_Qty,Consumed_Qty,Reject_Qty,Current_Wafer_Qty" & _
                    ",Film_Frame_Qty,Optical_Quality,Country_of_Assembly,Offshore_ASM_Company,Asm_Containment_type,Date_code,asm_conv_id,asm_excr_id,assembly_facility,Country_of_Test,Offshore_TEST_Company" & _
                    ",Tst_Containment_type,Tst_Program_rev,Created_date,Created_time,Del_Note,AWB,weight(kgs),package" & vbCrLf
    '明细数据
    strSql = "SELECT [PO_num] ,[PO_Item] ,[Previous_Batch_ID] ,[Previous_Mtrl_Num],[Batch_ID] ,[Mtrl_num] ,[Mtrl_desc] ,[Mtrl_num_Mtrlgrp] " & _
           " ,[Output_Qty] ,[Consumed_Qty] ,[Reject_Qty] ,[Current_Wafer_Qty] ,[Film_Frame_Qty] ,[Optical_Quality] ,[Country_of_Assembly] " & _
           " ,[Offshore_ASM_Company] ,[Asm_Containment_type] ,[Date_code] ,[asm_conv_id] ,[asm_excr_id] ,[assembly_facility] " & _
           " ,[Country_of_Test],[Offshore_TEST_Company] ,[Tst_Containment_type] ,[Tst_Program_rev] ,[Created_date] ,[Created_time],b.Del_Note,b.AWB,b.Weight,b.Package " & _
           " FROM [erpdata].[dbo].[GRdetailHistory] a,[erpdata].[dbo].[GRdetailSetUp] b " & _
           " Where a.单据编号 = b.单据编号 and a.单据编号='" + UCase(Trim(TxtBillNo.Text)) + "' "

    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    For i = 1 To rs.RecordCount
        strColData = ""
        For j = 0 To rs.Fields.Count - 1
            If j = 26 Then
             strColData = strColData + Format(g_Date, "hh:mm:ss") + ","
            Else
             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
            
            End If
        
           
        Next
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    MsgBox ("发送成功！")
    
    LogFile.Close
    Set LogFile = Nothing
    
ErrHandler:
    Set FSO = Nothing
End Sub


Private Sub SaveFileSendSX()
Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

On Error GoTo ErrHandler
'查询报表名的序号

fileNo = GetGC_FileNo("SX")

Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "SX_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "NO,供货方,客户,产品名称,客户订单号,客户Lot,WaferNo,GoodDieQty,BadDieQty,Yield,出货日期,LaserMark,箱号,备注" & vbCrLf
    '明细数据
  
  strSql = " select cast([NO] as int) as NO,[Sub_Name] as [供货方],[Ship_To]as [客户] ,[Customer_Device] as [产品名称],[PO_NO] as [客户订单号]," & _
          " [FAB_Lot_ID]as[客户Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[出货日期], " & _
          " [Wafer_Mark]as [Laser Mark],CartonNO as [箱号], [Remark] as [备注] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + UCase(Trim(TxtBillNoGC.Text)) + "' order by 1  "
          
          
           
    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        For j = 0 To rs.Fields.Count - 1

             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '发邮件
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    Call MailDetailSX("SX 发货报表", strRecipient, g_Path & "\" & "SX_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
    Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'SX') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub

'Private Sub SaveFileSendSX()
'Dim FSO         As New FileSystemObject
'Dim LogFile     As TextStream
'Dim strDatas    As String
'Dim strRowData  As String
'Dim strColData  As String
'Dim strSql      As String
'Dim i           As Integer, j           As Integer
'
'Dim maxRow As Integer
'
'Dim Rs          As New ADODB.Recordset
'
'Dim fileNo As String
'
'On Error GoTo ErrHandler
''查询报表名的序号
'
'fileNo = GetGC_FileNo("SX")
'
'Dim kk As String
'
'    '创建文件
'    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "SX_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
'    '写数据
'    strDatas = ""
'    '头数据
'    strDatas = "NO,供货方,客户,产品名称,客户订单号,客户Lot,WaferNo,GoodDieQty,BadDieQty,Yield,出货日期,LaserMark,箱号,备注" & vbCrLf
'    '明细数据
'
'  strSql = " select cast([NO] as int) as NO,[Sub_Name] as [供货方],[Ship_To]as [客户] ,[Customer_Device] as [产品名称],[PO_NO] as [客户订单号]," & _
'          " [FAB_Lot_ID]as[客户Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[出货日期], " & _
'          " [Wafer_Mark]as [Laser Mark],CartonNO as [箱号], [Remark] as [备注] " & _
'          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + UCase(Trim(TxtBillNoGC.Text)) + "' order by 1  "
'
'
'
'    strRowData = ""
'    If Rs.State = adStateOpen Then Rs.Close
'    If INIadoCon.State <> adStateOpen Then
'        INIConnectSTART
'    End If
'    Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
'    If Rs.EOF Then Exit Sub
'
'    maxRow = Rs.RecordCount
'
'    For i = 1 To Rs.RecordCount
'        strColData = ""
'        For j = 0 To Rs.fields.Count - 1
'
'             strColData = strColData + Trim("" & Rs.fields(j).Value) + ","
'
'        Next
'
'        If i = maxRow Then
'          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
'
'        Else
'
'        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
'
'        End If
'
'        Rs.MoveNext
'    Next
'    strDatas = strDatas + strRowData '数据连接
'    '写入文件
'    LogFile.WriteLine (strDatas)
'
'    LogFile.Close
'    Set LogFile = Nothing
'
'
'    '发邮件
'    Dim strRecipient    As String
'    Dim strRecipientCC  As String
'
'    strRecipient = "jiayun.zhang@qtechglobal.com"
'    strRecipientCC = "wanli.ma@qtechglobal.com"
'
'    Call MailDetailSX("SX 发货报表", strRecipient, g_Path & "\" & "SX_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
'
'    '把发送记录保存到DB中
'
'    Dim sqltemp2 As String
'
'    sqltemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'SX') "
'
'    Call AddSql2(sqltemp2)
'
'    MsgBox "发送成功！", vbInformation, "友情提示"
'
'
'ErrHandler:
'    Set FSO = Nothing
'End Sub

Private Sub SaveFileSendBD()
Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

On Error GoTo ErrHandler
'查询报表名的序号

fileNo = GetGC_FileNo("BD")

Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "BD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "NO,供货方,客户,产品名称,客户订单号,客户Lot,WaferNo,GoodDieQty,BadDieQty,Yield,出货日期,LaserMark,箱号,备注" & vbCrLf
    '明细数据
  
  strSql = " select cast([NO] as int) as NO,[Sub_Name] as [供货方],[Ship_To]as [客户] ,[Customer_Device] as [产品名称],[PO_NO] as [客户订单号]," & _
          " [FAB_Lot_ID]as[客户Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[出货日期], " & _
          " [Wafer_Mark]as [Laser Mark],CartonNO as [箱号], [Remark] as [备注] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + UCase(Trim(TxtBillNoGC.Text)) + "' order by 1  "
          
          
           
    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        For j = 0 To rs.Fields.Count - 1

             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '发邮件
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    Call MailDetailSX("BD 发货报表", strRecipient, g_Path & "\" & "BD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
    Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'BD') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub


Private Sub SaveFileSendHD()
Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

On Error GoTo ErrHandler
'查询报表名的序号

fileNo = GetGC_FileNo("HD")

Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "HD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "NO,供货方,客户,版本,产品名称,客户订单号,客户Lot,WaferNo,GoodDieQty,NGDieQty,ShipmentGoodDie,Yield,出货日期,备注" & vbCrLf
    '明细数据
  
  strSql = " select cast([NO] as int) as NO,[Sub_Name] as [供货方],[Ship_To]as [客户] ,[Fab_Device] as [版本],[Customer_Device] as [产品名称],[PO_NO] as [客户订单号]," & _
          " [FAB_Lot_ID]as[客户Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[Gross_Dies]as [NGDieQty],[NG_Die]as [ShipmentGoodDie] ,[Yield] ,[PACK_Out_Date]as[出货日期], " & _
          " [Remark] as [备注] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + UCase(Trim(TxtBillNoGC.Text)) + "' order by 1  "
          
          
           
    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        For j = 0 To rs.Fields.Count - 1

             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '发邮件
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    Call MailDetailHD("HD 发货报表", strRecipient, g_Path & "\" & "HD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
    Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'HD') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub



Private Sub SaveFileSendGC()
Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer
Dim waferidMain As String
Dim waferPT As String
Dim waferVer As String
Dim waferVerResult As String

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

On Error GoTo ErrHandler
'查询报表名的序号

fileNo = GetGC_FileNo("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PP_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "NO,Sub Name,Ship To,Fab Device,Customer Device,PO NO,WO,GC Version,Invoice NO,PACK-Out Date,PACK Lot ID,FAB Lot ID" & _
               ",Wafer ID,Wafer Mark,Gross Dies,Pass Dies,NG Die,Yield,Remark,System CartonNO,PACK Device,CartonNO,MaskType" & vbCrLf
    '明细数据
    strSql = "select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion, cast([NO] as int),[Sub_Name],[Ship_To],[Fab_Device],[Customer_Device],[PO_NO] " & _
           " ,[WO],[GC_Version],[Invoice_NO],[PACK_Out_Date],[PACK_Lot_ID],[FAB_Lot_ID] " & _
           " ,[Wafer_ID],[Wafer_Mark],[Gross_Dies],[Pass_Dies],[NG_Die],[Yield] " & _
           " ,[Remark],[System_CartonNO],[PACK_Device],[CartonNO],[MaskType] " & _
           " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a  " & _
           " Where a.单据编号='" + UCase(Trim(TxtBillNoGC.Text)) + "'  order by 4 "

    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        
            waferidMain = Trim("" & rs.Fields(0).Value) & "-A"
            
            waferPT = Trim("" & rs.Fields(1).Value)
            
            waferVer = Trim("" & rs.Fields(2).Value)
            
            waferVerResult = GetGCOutRpt_Ver(waferidMain, waferPT, waferVer)
            
        
        For j = 3 To rs.Fields.Count - 1
             
             If j = 10 Then
             
             strColData = strColData + waferVerResult + ","
             
             Else
             
             
             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
             
             End If
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '发邮件
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    Call MailDetail("GC 发货报表", strRecipient, g_Path & "\" & "PP_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub



Private Sub SaveFileSend()
'Excel附件

Dim strSql      As String
Dim i           As Integer, j           As Integer
Dim rs          As New ADODB.Recordset
Dim RsD         As New ADODB.Recordset
Dim xlApp       As New Excel.Application
Dim xlBook      As Excel.Workbook
Dim xlSheet     As Excel.Worksheet
Dim currentSheetRow As Long

Dim txtHeaderTemp As String



On Error GoTo ErrHandle
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlBook = xlApp.Workbooks().Add
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Name = "GrData"
    xlSheet.Activate
    xlApp.Visible = False
'
'
'    '第一行标题
'    xlSheet.Cells(1, 1) = "PO_num"
'    xlSheet.Cells(1, 2) = "PO_Item"
'    xlSheet.Cells(1, 3) = "Previous_Batch_ID"
'    xlSheet.Cells(1, 4) = "Previous_Mtrl_Num"
'    xlSheet.Cells(1, 5) = "Batch_ID"
'    xlSheet.Cells(1, 6) = "mtrl_num"
'    xlSheet.Cells(1, 7) = "mtrl_desc"
'    xlSheet.Cells(1, 8) = "Mtrl_num_Mtrlgrp"
'    xlSheet.Cells(1, 9) = "Output_Qty"
'    xlSheet.Cells(1, 10) = "Consumed_Qty"
'
'    xlSheet.Cells(1, 11) = "Reject_Qty"
'    xlSheet.Cells(1, 12) = "Current_Wafer_Qty"
'
'    xlSheet.Cells(1, 13) = "Film_Frame_Qty"
'    xlSheet.Cells(1, 14) = "Optical_Quality"
'    xlSheet.Cells(1, 15) = "Country_of_Assembly"
'    xlSheet.Cells(1, 16) = "Offshore_ASM_Company"
'
'    xlSheet.Cells(1, 17) = "Asm_Containment_type"
'
'    xlSheet.Cells(1, 18) = "Date_code"
'    xlSheet.Cells(1, 19) = "asm_conv_id"
'
'    xlSheet.Cells(1, 20) = "asm_excr_id"
'    xlSheet.Cells(1, 21) = "assembly_facility"
'    xlSheet.Cells(1, 22) = "Country_of_Test"
'    xlSheet.Cells(1, 23) = "Offshore_TEST_Company"
'
'    xlSheet.Cells(1, 24) = "Tst_Containment_type"
'    xlSheet.Cells(1, 25) = "Tst_Program_rev"
'    xlSheet.Cells(1, 26) = "Created_date"
'    xlSheet.Cells(1, 27) = "Created_time"
'
'    xlSheet.Cells(1, 28) = "Del_Note"
'    xlSheet.Cells(1, 29) = "AWB"
'    xlSheet.Cells(1, 30) = "weight(kgs)"
'    xlSheet.Cells(1, 31) = "package"
    
    txtHeaderTemp = "PO_num,PO_Item,Previous_Batch_ID,Previous_Mtrl_Num,Batch_ID,mtrl_num,mtrl_desc,Mtrl_num_Mtrlgrp,Output_Qty,Consumed_Qty,Reject_Qty,Current_Wafer_Qty" & _
                    " Film_Frame_Qty,Optical_Quality,Country_of_Assembly,Asm_Containment_type,Date_code,asm_conv_id,asm_excr_id,assembly_facility,Country_of_Test,Offshore_TEST_Company" & _
                    " Tst_Containment_type,Tst_Program_rev,Created_date,Created_time,Del_Note,AWB,weight(kgs),package" & vbCrLf
       xlSheet.Cells(1, 1) = txtHeaderTemp
    
Dim tempBillNo As String
tempBillNo = UCase(Trim(TxtBillNo.Text))

 Dim sqlTemp As String

 strSql = "SELECT [PO_num] ,[PO_Item] ,[Previous_Batch_ID] ,[Previous_Mtrl_Num],[Batch_ID] ,[Mtrl_num] ,[Mtrl_desc] ,[Mtrl_num_Mtrlgrp] " & _
           " ,[Output_Qty] ,[Consumed_Qty] ,[Reject_Qty] ,[Current_Wafer_Qty] ,[Film_Frame_Qty] ,[Optical_Quality] ,[Country_of_Assembly] " & _
           " ,[Offshore_ASM_Company] ,[Asm_Containment_type] ,[Date_code] ,[asm_conv_id] ,[asm_excr_id] ,[assembly_facility] " & _
           " ,[Country_of_Test],[Offshore_TEST_Company] ,[Tst_Containment_type] ,[Tst_Program_rev] ,[Created_date] ,[Created_time],b.Del_Note,b.AWB,b.Weight,b.Package " & _
           " FROM [erpdata].[dbo].[GRdetailHistory] a,[erpdata].[dbo].[GRdetailSetUp] b " & _
           " Where a.单据编号 = b.单据编号 and a.单据编号='" + tempBillNo + "' "


    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
    INIConnectSTART
    End If

    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
'     xlSheet.Range("a2:K" & Rs.RecordCount + 1).NumberFormatLocal = "@"
     currentSheetRow = rs.RecordCount + 1
    For i = 2 To rs.RecordCount + 1
        For j = 0 To rs.Fields.Count - 1
            xlSheet.Cells(i, j + 1) = Trim("" & rs.Fields(j).Value)
        Next
        rs.MoveNext
    Next

'
 

  
'    xlSheet.SaveAs g_Path_GR & "\" & Format(g_Date, "YYYY-MM-DD hhmmss") & "WipReport.xls"
    
    xlSheet.SaveAs g_Path_GR & "\" & "QT_FG_CSP_" & Format(g_Date, "YYYYMMDD") & "_" & Format(g_Date, "hhmmss") & ".csv"
    
    
    xlBook.Close
    
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    rs.Close
    Set rs = Nothing
    
    g_IsShouldSend = True
    
    Exit Sub
ErrHandle:
    Set xlApp = Nothing  '"ユ临北倒Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing

End Sub



Private Sub Command1_Click()
'po recomit




'明细数据
Dim i As Integer
Dim waferIdTemp As String
Dim woType As String


    Set listRS = GetFps37POCommit()


If listRS.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认"
    Exit Sub
    
Else


    fps(5).MaxRows = listRS.RecordCount
    For i = 0 To listRS.RecordCount - 1
         


         
         With fps(5)
                 .Row = i + 1
                 .Col = E_FPS5.E_Event
                 .Text = CStr(listRS.Fields(0).Value)
                 
                .Row = i + 1
                 .Col = E_FPS5.E_PONumber
                .Text = CStr(listRS.Fields(1).Value)
                
                
                  .Row = i + 1
                 .Col = E_FPS5.E_POLineItem
                 .Text = CStr(listRS.Fields(2).Value)
                 
                  .Row = i + 1
                 .Col = E_FPS5.E_Pt2
                 .Text = CStr(listRS.Fields(3).Value)
                 
                 
                  .Row = i + 1
                 .Col = E_FPS5.E_ComDate
                 .Text = CStr(listRS.Fields(4).Value)
                 
                   .Row = i + 1
                 .Col = E_FPS5.E_LotID
                 .Text = CStr(listRS.Fields(5).Value)
                 
                
        
        End With
    
NextRecord:
       
        listRS.MoveNext

    Next


End If




End Sub

Private Sub Command2_Click()

'po recomit


Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer
Dim outQty As Integer
Dim lotIdTemp As String
Dim potemp As String
Dim allQty As Integer

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

'On Error GoTo ErrHandler
'查询报表名的序号

fileNo = Format(Now, "YYYYMMDD_HHMM") & "HTKS_RECOMMIT"

'20151106_0846HTKS_RECEIPT.csv

Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path37 & "\" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "Event,PO Number,PO Line Item Number,Production Order Number,Commit Date,LotID(厂内),PO交期(厂内),客户机种(厂内),华天机种(厂内),上次Commit Date(厂内),剩余片数(厂内)" & vbCrLf
    '明细数据

  
strSql = " select 'RECOMMIT'as Event,po_num,po_item, mpn,' ' as commitdate,a.source_batch_id ,a.date_code,a.mpn_desc,b.qtechptno ,c.commitDate, 0 from customeroitbl_test a ,TBLTsvNpiProduct b ,SemtechPortalCommit c where a.customershortname='37' and a.qtech_created_date>to_date('2016-03-26','YYYY-MM-DD') " & _
" and a.lot_status is null and a.po_num is not null  and ( b.customerptno1=a.mpn_desc  or b.customerptno2=a.mpn_desc )  and c.ponum=a.po_num  and  c.productPt=a.MPN order by a.id "

    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    
    If Cnn.State = 0 Then
    ConOracle
    End If
    
    
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
          lotIdTemp = CStr(rs.Fields(5).Value)
          potemp = CStr(rs.Fields(1).Value)
        
        For j = 0 To rs.Fields.Count - 1

            ' strColData = strColData + Trim("" & Rs.fields(j).Value) + ","
             '2016-04-07 jiayun add 剩余片数
             
             
             
             If j = rs.Fields.Count - 1 Then
                   allQty = GetQty37OIAllQty(lotIdTemp, potemp)
                   outQty = GetQty37OutQty(lotIdTemp)
                   
                   strColData = strColData + CStr((allQty - outQty)) + ","
                   
             
             Else
             
                strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
             
             End If
             
             
             
             
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
'    '发邮件
'    Dim strRecipient    As String
'    Dim strRecipientCC  As String
'
'    strRecipient = "jiayun.zhang@qtechglobal.com"
'    strRecipientCC = "wanli.ma@qtechglobal.com"
'
'    Call MailDetailHD("HD 发货报表", strRecipient, g_Path & "\" & "HD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
'
'    '把发送记录保存到DB中
    
'    Dim sqlTemp2 As String
'
'    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'HD') "
'
'    Call AddSql2(sqlTemp2)
'
'    MsgBox "发送成功！", vbInformation, "友情提示"
'
'
'ErrHandler:
'    Set FSO = Nothing


MsgBox "导出成功！", vbInformation, "友情提示"







End Sub

Private Sub Command27_Click()
'把commit文件上传到厂内系统中，供下次commit使用

Dim userid As String
userid = UCase(gUserName)

'Semteck 晶圆数据上传


'读取CSV
Dim source_batch_id_Temp As String
Dim customerTemp As String
Dim cusPTTemp As String
Dim gcVerTemp As String
Dim gcVerLastTemp As String
Dim waferIDList As String


'上传OI的CSV
'处理文件名
If Text8.Text = "" Then

    
        MsgBox "先选择待上传的文件!", vbInformation, "友情提示"
        
        
    Exit Sub
End If



        Dim i As Integer
        Dim j As Integer
        Dim id As Long
        Dim temp As String
        Dim SumCount As Integer
        Dim GCHeaderFlag As Boolean
        Dim str01 As String
        Dim str03 As String
        Dim FabTem As String
        Dim qtyTemp As Long
        Dim cmdStr2 As String
        
        SumCount = 0
 
        

        Dim k As Integer
        
        Dim FName As String
        Dim Nextline As String
        FName = Trim(Text8.Text)
        Open FName For Input As #1
        
        Do Until EOF(1)
        Line Input #1, Nextline
              
              
             If UCase(Left(Trim(Nextline), 6)) = "COMMIT" Then
             Dim BID
             BID = Split(Nextline, ",")
             
    

                   
                     Dim ponumTemp As String
                     Dim polineTemp As String
                     Dim ptTemp As String
                     Dim commitDate As String
                     
                     ponumTemp = BID(1)
                     polineTemp = BID(2)
                     ptTemp = BID(3)
                     commitDate = BID(4)
                     
         
           cmdStr2 = " insert into SemtechPortalCommit(ponum,poline,productPt,commitDate,flag,createby,createdate) values ('" & ponumTemp & "','" & polineTemp & "','" & ptTemp & "','" & commitDate & "','Y','" & userid & "',sysdate) "
                 
          AddSql (cmdStr2)
                                 
          SumCount = SumCount + 1
         
            
        End If
        
        Loop
        Close #1
        
        
        
        If SumCount > 0 Then
            MsgBox "已成功上传" & SumCount & "笔！", vbInformation, "友情提示"
        End If
        


End Sub

Private Sub Command28_Click()

'GC
On Error Resume Next
Dim FName
    '帅选文件
    CommonDialog7.Filter = "CSV文件(*.csv)|*.csv|EXCEL文件(*.xlsx)|*.xlsx|EXCEL文件(*.xls)|*.xls"
    
    CommonDialog7.ShowOpen
    '得到文件名
    FName = CommonDialog7.filename
    If FName <> "" Then
       Text8.Text = FName
    End If
    


End Sub

Private Sub Form_Activate()
Txt37BillNo.SetFocus
Txt37BillNoShip.SetFocus
End Sub

Private Sub Form_Load()
IniFpsWaferRec

IniFpsPOCommit

IniFpsPOReCommit

IniFpsPOStart

IniFpsPOComplete

IniFpsPOShip


'txtKey.Text = "PROTECTIVE_FILM_APLD"
'TxtAttri.Text = "BB栏"
'
' With fps(0)
'        .ReDraw = False
'        .MaxCols = E_FPS0.E_End - 1
'        .MaxRows = 0
'
'        '砞竚Α
'        .DAutoHeadings = False
'        .DAutoCellTypes = False
'        .DAutoSizeCols = DAutoSizeColsNone
'
'        .Col = -1
'        .Row = -1
'        .Lock = True
'
'
'        .OperationMode = OperationModeNormal
'        .TypeVAlign = TypeVAlignCenter
'        .SelForeColor = &HFF8080
'
'
'
'        .SetText E_FPS0.E_Key, 0, "字段名"
'        .SetText E_FPS0.E_Value, 0, "字段值"
'        .SetText E_FPS0.E_getValue, 0, "是否贴膜"
'        .SetText E_FPS0.E_otherValue, 0, "备注"
'
'
'        .ColWidth(E_FPS0.E_Key) = 20
'        .ColWidth(E_FPS0.E_Value) = 15
'        .ColWidth(E_FPS0.E_getValue) = 15
'        .ColWidth(E_FPS0.E_otherValue) = 25
'
'
'
'        .RowHeight(0) = 20
'        .RowHeight(-1) = 15
'
'
'
'
'        .ReDraw = True
'    End With
'
'
'ShowData_Where


'Combo2.AddItem ("GC")
'Combo2.AddItem ("SX")
'Combo2.AddItem ("HJ")
'
'Combo2.AddItem ("HD")
'Combo2.AddItem ("BD")
'Combo2.AddItem ("45")


End Sub


Private Sub ShowData_Where()
'Set reportRS = GetpfData()
'
'With fps(0)
'        .MaxRows = 0
'        If reportRS.RecordCount > 0 Then
'            Set .DataSource = reportRS
'
'        End If
'End With

End Sub


Private Sub GCCmdOut_Click()


Dim tempBillNo As String
Dim custNameTemp As String


tempBillNo = UCase(Trim(TxtBillNoGC.Text))
custNameTemp = UCase(Trim(Combo2.Text))

If tempBillNo = "" Or custNameTemp = "" Then
    MsgBox "请选择客户代码，输入单据编号!", vbInformation, "友情提示"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGC(tempBillNo, custNameTemp)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If
    


 Dim sqlTemp As String
      
 If custNameTemp = "GC" Then
           
sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [Sub Name],[Ship_To]as [Ship To] ,[Fab_Device]as [Fab Device] ,[Customer_Device] as [Customer Device],[PO_NO] as [PO NO]," & _
          " [WO],[GC_Version]as [GC Version],[Invoice_NO]as [Invoice NO] ,[PACK_Out_Date]as[PACK-Out Date],[PACK_Lot_ID]as[PACK Lot ID],[FAB_Lot_ID]as[FAB Lot ID] ," & _
          " [Wafer_ID]as [Wafer ID],[Wafer_Mark]as [Wafer Mark],[Gross_Dies]as [Gross Dies],[Pass_Dies]as [Pass Dies],[NG_Die]as [NG Die] ,[Yield] ," & _
          " [Remark] , [System_CartonNO]as [System CartonNO], [PACK_Device]as [PACK Device], [CartonNO]as [CartonNO], [MaskType] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + tempBillNo + "' order by 1  "
                 
ElseIf custNameTemp = "SX" Or custNameTemp = "HJ" Then

sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [供货方],[Ship_To]as [客户] ,[Customer_Device] as [产品名称],[PO_NO] as [客户订单号]," & _
          " [FAB_Lot_ID]as[客户Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[出货日期], " & _
          " [Wafer_Mark]as [Laser Mark],CartonNO as [箱号], [Remark] as [备注] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + tempBillNo + "' order by 1  "
          
ElseIf custNameTemp = "BD" Then

sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [供货方],[Ship_To]as [客户] ,[Customer_Device] as [产品名称],[PO_NO] as [客户订单号]," & _
          " [FAB_Lot_ID]as[客户Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[出货日期], " & _
          " [Wafer_Mark]as [Laser Mark],CartonNO as [箱号], [Remark] as [备注] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + tempBillNo + "' order by 1  "
          
          
ElseIf custNameTemp = "HD" Then

sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [供货方],[Ship_To]as [客户] ,[Fab_Device] as [版本],[Customer_Device] as [产品名称],[PO_NO] as [客户订单号]," & _
          " [FAB_Lot_ID]as[客户Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[Gross_Dies]as [NGDieQty],[NG_Die]as [ShipmentGoodDie] ,[Yield] ,[PACK_Out_Date]as[出货日期], " & _
          "  [Remark] as [备注] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + tempBillNo + "' order by 1  "
End If

  SqlServerExporToExcel (sqlTemp)

End Sub

Private Sub GCCmdSend_Click()



'发送
Dim tempBillNo As String
Dim custNameTemp As String

tempBillNo = UCase(Trim(TxtBillNoGC.Text))
custNameTemp = UCase(Trim(Combo2.Text))


If tempBillNo = "" Or custNameTemp = "" Then
    MsgBox "请选择客户代码，输入单据编号!", vbInformation, "友情提示"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGC(tempBillNo, custNameTemp)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If

If custNameTemp = "GC" Then

SaveFileSendGC

ElseIf custNameTemp = "SX" Or custNameTemp = "HJ" Then
SaveFileSendSX

ElseIf custNameTemp = "BD" Then
SaveFileSendBD


ElseIf custNameTemp = "HD" Then
SaveFileSendHD


End If


    
End Sub

Private Sub TxtPackage_KeyPress(KeyAscii As Integer)
Dim Numbers As String
Numbers = "1234567890" + Chr(8) + Chr(46) + Chr(13)
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
CmdSaver.SetFocus
End If
End Sub

Private Sub TxtWeight_KeyPress(KeyAscii As Integer)

Dim Numbers As String
Numbers = "1234567890" + Chr(8) + Chr(46) + Chr(13)
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
TxtPackage.SetFocus
End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

 Select Case PreviousTab
      Case 4
       Txt37BillNo.SetFocus
      Case 3
         Txt37BillNoShip.SetFocus
   End Select
   
   
End Sub

Private Sub Txt37BillNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CmdPOComplete_Click
End If
End Sub

Private Sub Txt37BillNoShip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CmdPOShip_Click
End If

End Sub
