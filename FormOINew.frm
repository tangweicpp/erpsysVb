VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmOINew 
   Caption         =   "上传客户来料与订单信息"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
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
   ScaleHeight     =   8895
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab DataCombo1 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Aptina及CN 导入"
      TabPicture(0)   =   "FormOINew.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label25"
      Tab(0).Control(1)=   "lbl12"
      Tab(0).Control(2)=   "lbl23"
      Tab(0).Control(3)=   "Label11"
      Tab(0).Control(4)=   "DTPStop(2)"
      Tab(0).Control(5)=   "DTPStart(2)"
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(8)=   "Combo1"
      Tab(0).Control(9)=   "txtSubstrateID"
      Tab(0).Control(10)=   "Command35"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "通用版"
      TabPicture(1)   =   "FormOINew.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label24"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label26"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label28"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "CmbCustomer"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Check1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "SemTech/MPS/GULF资料上传"
      TabPicture(2)   =   "FormOINew.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label44"
      Tab(2).Control(1)=   "Label45"
      Tab(2).Control(2)=   "Label10"
      Tab(2).Control(3)=   "Frame6"
      Tab(2).Control(4)=   "Frame7"
      Tab(2).Control(5)=   "CmbCustomer37"
      Tab(2).Control(6)=   "CmbPoType"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "SemTech收料信息(ICI)"
      TabPicture(3)   =   "FormOINew.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CmdExit"
      Tab(3).Control(1)=   "CmdOK"
      Tab(3).Control(2)=   "Frame8"
      Tab(3).Control(3)=   "txtWaferID"
      Tab(3).Control(4)=   "Label1"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "SemTech二次PO上传(Excel)"
      TabPicture(4)   =   "FormOINew.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame11"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "手工创建新客户Mapping"
      TabPicture(5)   =   "FormOINew.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame4"
      Tab(5).Control(1)=   "TxtCustomerName"
      Tab(5).Control(2)=   "CmdSaveOI"
      Tab(5).Control(3)=   "CmdClearOI"
      Tab(5).Control(4)=   "Label21"
      Tab(5).Control(5)=   "Label22"
      Tab(5).Control(6)=   "Label23"
      Tab(5).ControlCount=   7
      Begin VB.CommandButton Command35 
         Caption         =   "删除"
         Height          =   360
         Left            =   -64680
         TabIndex        =   116
         Top             =   960
         Width           =   990
      End
      Begin VB.TextBox txtSubstrateID 
         Height          =   285
         Left            =   -66600
         TabIndex        =   114
         Top             =   1035
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "回货上传"
         Height          =   315
         Left            =   6840
         TabIndex        =   113
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Frame Frame11 
         Caption         =   ".csv数据"
         Height          =   5175
         Left            =   -73440
         TabIndex        =   96
         Top             =   960
         Width           =   11415
         Begin VB.TextBox Text9 
            Height          =   1335
            Left            =   360
            MultiLine       =   -1  'True
            TabIndex        =   101
            Top             =   960
            Width           =   9975
         End
         Begin VB.CommandButton Command34 
            Caption         =   "导出PO"
            Height          =   480
            Left            =   3960
            TabIndex        =   99
            Top             =   3000
            Width           =   1095
         End
         Begin VB.CommandButton Command33 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   1080
            TabIndex        =   98
            Top             =   3000
            Width           =   1335
         End
         Begin VB.CommandButton Command18 
            Caption         =   ".."
            Height          =   495
            Left            =   7320
            TabIndex        =   97
            Top             =   360
            Width           =   375
         End
         Begin MSComDlg.CommonDialog CommonDialog9 
            Left            =   6600
            Top             =   360
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
            Index           =   9
            Left            =   360
            TabIndex        =   100
            Top             =   480
            Width           =   1545
         End
      End
      Begin VB.ComboBox CmbPoType 
         Height          =   315
         ItemData        =   "FormOINew.frx":00A8
         Left            =   -72120
         List            =   "FormOINew.frx":00AA
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   5040
         Width           =   2655
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "取消"
         Height          =   480
         Left            =   -67920
         TabIndex        =   88
         Top             =   7920
         Width           =   1575
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "确定"
         Height          =   480
         Left            =   -70560
         TabIndex        =   87
         Top             =   7920
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mapping_XML"
         Height          =   2295
         Left            =   -72720
         TabIndex        =   78
         Top             =   1740
         Width           =   9015
         Begin VB.TextBox Text4 
            Enabled         =   0   'False
            Height          =   495
            Left            =   7560
            MultiLine       =   -1  'True
            TabIndex        =   82
            Top             =   9000
            Width           =   4935
         End
         Begin VB.CommandButton Command12 
            Caption         =   ".."
            Height          =   495
            Left            =   12840
            TabIndex        =   81
            Top             =   9000
            Width           =   375
         End
         Begin VB.CommandButton Command11 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   7920
            TabIndex        =   80
            Top             =   9720
            Width           =   1335
         End
         Begin VB.CommandButton Command10 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   10800
            TabIndex        =   79
            Top             =   9720
            Width           =   1335
         End
         Begin MSComDlg.CommonDialog CommonDialog3 
            Left            =   9720
            Top             =   8400
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
            Left            =   7560
            TabIndex        =   83
            Top             =   8640
            Width           =   1455
         End
      End
      Begin VB.TextBox TxtCustomerName 
         Height          =   375
         Left            =   -72120
         TabIndex        =   77
         Top             =   1140
         Width           =   2415
      End
      Begin VB.Frame Frame8 
         Caption         =   "报表"
         Height          =   6735
         Left            =   -65040
         TabIndex        =   58
         Top             =   1020
         Width           =   4575
         Begin VB.CommandButton Command32 
            Caption         =   "导出Excel"
            Height          =   480
            Left            =   1200
            TabIndex        =   89
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Frame Frame9 
            Caption         =   "走ERP发货的"
            Height          =   3015
            Left            =   31080
            TabIndex        =   68
            Top             =   9120
            Width           =   4335
            Begin VB.CommandButton Command29 
               Caption         =   "导出Excel"
               Height          =   480
               Left            =   95640
               TabIndex        =   70
               Top             =   29400
               Width           =   1215
            End
            Begin VB.TextBox TxtBillNoGC 
               Height          =   375
               Left            =   94800
               TabIndex        =   69
               Top             =   28680
               Width           =   3495
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单据编号："
               Height          =   195
               Left            =   94800
               TabIndex        =   71
               Top             =   28320
               Width           =   900
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "每一次出去的"
            Height          =   2655
            Left            =   31080
            TabIndex        =   60
            Top             =   6120
            Width           =   4335
            Begin VB.CommandButton Command30 
               Caption         =   "导出Excel"
               Height          =   480
               Left            =   95640
               TabIndex        =   62
               Top             =   21000
               Width           =   1215
            End
            Begin VB.ComboBox CusPT 
               Height          =   315
               Left            =   96120
               Style           =   2  'Dropdown List
               TabIndex        =   61
               Top             =   19200
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker DTP1 
               Height          =   375
               Index           =   0
               Left            =   96120
               TabIndex        =   63
               Top             =   19800
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               _Version        =   393216
               Format          =   291176449
               CurrentDate     =   41424
            End
            Begin MSComCtl2.DTPicker DTP2 
               Height          =   375
               Index           =   0
               Left            =   96120
               TabIndex        =   64
               Top             =   20400
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               _Version        =   393216
               Format          =   291176449
               CurrentDate     =   41424
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "结束日期："
               Height          =   195
               Left            =   94920
               TabIndex        =   67
               Top             =   20520
               Width           =   900
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "开始日期： "
               Height          =   195
               Left            =   94920
               TabIndex        =   66
               Top             =   19920
               Width           =   945
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "客户机种： "
               Height          =   195
               Left            =   94920
               TabIndex        =   65
               Top             =   19320
               Width           =   945
            End
         End
         Begin VB.CommandButton Command31 
            Caption         =   "导出Excel"
            Height          =   480
            Left            =   21720
            TabIndex        =   59
            Top             =   6120
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   22200
            TabIndex        =   72
            Top             =   4800
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   291176449
            CurrentDate     =   41424
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   22200
            TabIndex        =   73
            Top             =   5400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   291176449
            CurrentDate     =   41424
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   1680
            TabIndex        =   90
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   291176449
            CurrentDate     =   41424
         End
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   375
            Left            =   1680
            TabIndex        =   91
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   291176449
            CurrentDate     =   41424
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "开始日期： "
            Height          =   195
            Left            =   480
            TabIndex        =   93
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结束日期："
            Height          =   195
            Left            =   480
            TabIndex        =   92
            Top             =   1320
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结束日期："
            Height          =   195
            Left            =   21000
            TabIndex        =   75
            Top             =   5520
            Width           =   900
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "开始日期： "
            Height          =   195
            Left            =   21000
            TabIndex        =   74
            Top             =   4920
            Width           =   945
         End
      End
      Begin VB.TextBox txtWaferID 
         Height          =   6615
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         Top             =   1140
         Width           =   9615
      End
      Begin VB.ComboBox CmbCustomer37 
         Height          =   315
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   480
         Width           =   1695
      End
      Begin VB.Frame Frame7 
         Caption         =   "一次客户PO"
         Height          =   2775
         Left            =   -73800
         TabIndex        =   38
         Top             =   5400
         Width           =   7095
         Begin VB.TextBox Text8 
            Enabled         =   0   'False
            Height          =   495
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   52
            Top             =   1200
            Width           =   5295
         End
         Begin VB.CommandButton Command28 
            Caption         =   ".."
            Height          =   495
            Left            =   6120
            TabIndex        =   51
            Top             =   1200
            Width           =   375
         End
         Begin VB.CommandButton Command27 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   840
            TabIndex        =   50
            Top             =   1920
            Width           =   1335
         End
         Begin VB.CommandButton Command24 
            Caption         =   "导出PO"
            Height          =   480
            Left            =   4080
            TabIndex        =   49
            Top             =   1920
            Width           =   1335
         End
         Begin VB.CommandButton Command22 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   6960
            TabIndex        =   42
            Top             =   7140
            Width           =   1335
         End
         Begin VB.CommandButton Command21 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   4080
            TabIndex        =   41
            Top             =   7140
            Width           =   1335
         End
         Begin VB.CommandButton Command20 
            Caption         =   ".."
            Height          =   495
            Left            =   9000
            TabIndex        =   40
            Top             =   6300
            Width           =   375
         End
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            Height          =   495
            Left            =   3720
            MultiLine       =   -1  'True
            TabIndex        =   39
            Top             =   6300
            Width           =   4935
         End
         Begin MSComDlg.CommonDialog CommonDialog5 
            Left            =   5880
            Top             =   5700
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComDlg.CommonDialog CommonDialog7 
            Left            =   6240
            Top             =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            MaxFileSize     =   10000
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的xls："
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   8
            Left            =   600
            TabIndex        =   53
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的CSV："
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   6
            Left            =   3720
            TabIndex        =   43
            Top             =   5940
            Width           =   1545
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Fab厂数据"
         Height          =   3975
         Left            =   -73800
         TabIndex        =   34
         Top             =   960
         Width           =   7095
         Begin VB.TextBox txtText10 
            Height          =   285
            Left            =   360
            TabIndex        =   103
            Top             =   1680
            Width           =   1575
         End
         Begin VB.CommandButton cmdQes 
            BackColor       =   &H00FFFF00&
            Caption         =   "条件查询"
            Height          =   585
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   3000
            Width           =   1335
         End
         Begin VB.TextBox Text7 
            Enabled         =   0   'False
            Height          =   495
            Left            =   360
            MultiLine       =   -1  'True
            TabIndex        =   47
            Top             =   840
            Width           =   4935
         End
         Begin VB.CommandButton Command26 
            Caption         =   ".."
            Height          =   495
            Left            =   5640
            TabIndex        =   46
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton Command25 
            Caption         =   "上传DB"
            Height          =   585
            Left            =   360
            TabIndex        =   45
            Top             =   3000
            Width           =   1335
         End
         Begin VB.CommandButton Command23 
            Caption         =   "导出Wafer"
            Height          =   585
            Left            =   4320
            TabIndex        =   44
            Top             =   3000
            Width           =   1335
         End
         Begin VB.CommandButton Command19 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   6960
            TabIndex        =   37
            Top             =   4380
            Width           =   1335
         End
         Begin VB.CommandButton Command17 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   4080
            TabIndex        =   36
            Top             =   4380
            Width           =   1335
         End
         Begin VB.CommandButton Command16 
            Caption         =   ".."
            Height          =   495
            Left            =   9000
            TabIndex        =   35
            Top             =   3660
            Width           =   375
         End
         Begin MSComDlg.CommonDialog CommonDialog4 
            Left            =   9240
            Top             =   2940
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            MaxFileSize     =   10000
         End
         Begin MSComDlg.CommonDialog CommonDialog6 
            Left            =   2520
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            MaxFileSize     =   10000
         End
         Begin MSComCtl2.DTPicker DTP1 
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   106
            Top             =   2400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   257622017
            CurrentDate     =   41424
         End
         Begin MSComCtl2.DTPicker DTP2 
            Height          =   375
            Index           =   1
            Left            =   2520
            TabIndex        =   108
            Top             =   2400
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   257622017
            CurrentDate     =   41424
         End
         Begin VB.Label lbl2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结束时间："
            Height          =   195
            Left            =   2520
            TabIndex        =   107
            Top             =   2160
            Width           =   900
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "开始时间："
            Height          =   195
            Left            =   360
            TabIndex        =   105
            Top             =   2160
            Width           =   900
         End
         Begin VB.Label lblDeviceName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DeviceName:"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   360
            TabIndex        =   104
            Top             =   1440
            Width           =   945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的CSV："
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   48
            Top             =   480
            Width           =   1545
         End
      End
      Begin VB.CommandButton CmdSaveOI 
         Caption         =   "保存"
         Height          =   480
         Left            =   -70800
         TabIndex        =   33
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton CmdClearOI 
         Caption         =   "清空"
         Height          =   480
         Left            =   -67920
         TabIndex        =   32
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Caption         =   "Mapping上传"
         Height          =   3975
         Left            =   960
         TabIndex        =   23
         Top             =   4200
         Width           =   7095
         Begin VB.CommandButton Command15 
            Caption         =   "导出报表"
            Height          =   600
            Left            =   4440
            TabIndex        =   27
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CommandButton Command14 
            Caption         =   "上传DB"
            Height          =   600
            Left            =   480
            TabIndex        =   26
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CommandButton Command13 
            Caption         =   ".."
            Height          =   495
            Left            =   5880
            TabIndex        =   25
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox TxtSI 
            Enabled         =   0   'False
            Height          =   1935
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   24
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的map："
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   28
            Top             =   240
            Width           =   1560
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormOINew.frx":00AC
         Left            =   -73320
         List            =   "FormOINew.frx":00AE
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Caption         =   "WO上传"
         Height          =   2775
         Left            =   960
         TabIndex        =   13
         Top             =   1440
         Width           =   7095
         Begin VB.CommandButton Command9 
            Caption         =   "导出明细表"
            Height          =   600
            Left            =   4440
            TabIndex        =   19
            Top             =   1920
            Width           =   1335
         End
         Begin VB.CommandButton Command8 
            Caption         =   "导出主表"
            Height          =   600
            Left            =   2400
            TabIndex        =   17
            Top             =   1920
            Width           =   1335
         End
         Begin VB.CommandButton Command7 
            Caption         =   "上传DB"
            Height          =   600
            Left            =   480
            TabIndex        =   16
            Top             =   1920
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   ".."
            Height          =   495
            Left            =   5880
            TabIndex        =   15
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   495
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   840
            Width           =   5295
         End
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   3000
            Top             =   240
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
            Index           =   2
            Left            =   480
            TabIndex        =   18
            Top             =   600
            Width           =   1545
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mapping_XML"
         Height          =   2415
         Left            =   -74040
         TabIndex        =   7
         Top             =   1620
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
         Begin VB.CommandButton Cmd 
            Caption         =   ".."
            Height          =   495
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
         Top             =   4380
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
      Begin MSDataListLib.DataCombo CmbCustomer 
         Height          =   315
         Left            =   1800
         TabIndex        =   31
         Top             =   1020
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DTPStart 
         Height          =   375
         Index           =   2
         Left            =   -69720
         TabIndex        =   111
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   257097729
         CurrentDate     =   41424
      End
      Begin MSComCtl2.DTPicker DTPStop 
         Height          =   375
         Index           =   2
         Left            =   -69720
         TabIndex        =   112
         Top             =   990
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   257884161
         CurrentDate     =   41424
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"FormOINew.frx":00B0
         Height          =   390
         Left            =   -67680
         TabIndex        =   115
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label lbl23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间："
         Height          =   195
         Left            =   -70680
         TabIndex        =   110
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lbl12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间："
         Height          =   195
         Left            =   -70680
         TabIndex        =   109
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO模板类型"
         Height          =   195
         Left            =   -73320
         TabIndex        =   94
         Top             =   5040
         Width           =   930
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注："
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -72600
         TabIndex        =   86
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excel模板格式为：WaferId LotId ProductId 良品数 不良数"
         Height          =   195
         Left            =   -72120
         TabIndex        =   85
         Top             =   4860
         Width           =   4395
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户："
         Height          =   195
         Left            =   -72720
         TabIndex        =   84
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫入二维码:"
         Height          =   195
         Left            =   -74520
         TabIndex        =   76
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户"
         Height          =   195
         Left            =   -73560
         TabIndex        =   55
         Top             =   540
         Width           =   360
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请先选择客户代码，然后再上传WO或Mapping"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -71160
         TabIndex        =   54
         Top             =   600
         Width           =   3570
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GC客户WLT，MG客户上传时，先上传WO，后再上传Mapping。"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1680
         TabIndex        =   30
         Top             =   8160
         Width           =   4860
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*请先选择客户代码，然后再上传WO或Mapping"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3840
         TabIndex        =   29
         Top             =   1020
         Width           =   3660
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户"
         Height          =   195
         Left            =   -73800
         TabIndex        =   22
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码"
         Height          =   195
         Left            =   1080
         TabIndex        =   20
         Top             =   1080
         Width           =   720
      End
   End
End
Attribute VB_Name = "FrmOINew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim mapTemp        As MapRecord

Dim gcHeaderTemp   As GCHeader

Dim eqISHeaderTemp As EQISHeader

Dim semtechFabTemp As SemtechFabDetail

Dim semPotemp      As SemtechPOHeader

Dim gcDetailTemp   As GCDetail

Dim gCnt           As Integer

'Dim SumCount As Integer
Dim ErrorInf       As String

Dim updateRS       As New ADODB.Recordset

Dim oiRS           As New ADODB.Recordset

Private Sub cmd_Click()

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

Private Sub CmdClearOI_Click()
    ClearData

End Sub

Private Sub ClearData()
    TxtCustomer.text = ""
    txtPO.text = ""
    txtPOItem.text = ""
    txtLotID.text = ""
    txtMPN.text = ""

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
    txtShipSite.text = ""
    txtLotStatus.text = ""

    TxtCustomer.SetFocus

End Sub

Private Sub cmdExit_Click()
    txtSubstrateID.text = ""

End Sub

Private Sub CmdOK_Click()
    '保存二维码信息

    Dim txtStr       As String

    Dim dirtemp      As String

    Dim cmdStr2      As String

    Dim fileNameTemp As String

    Dim msgTxtTemp   As String

    Dim msgTxtTemp2  As String

    Dim fwTemp       As String

    Dim dTemp        As String

    Dim lottemp      As String

    Dim dcTemp       As String

    Dim dqTemp       As String

    Dim snTemp       As String

    Dim bagTemp      As String

    Dim j            As Integer

    Dim waferStrTemp As String

    Dim sqlDB        As String

    fileNameTemp = ""
    msgTxtTemp = ""

    txtStr = txtSubstrateID.text

    msgTxtTemp = Replace(txtStr, vbCrLf, ";")

    ''1234,'456,'789'
    'msgTxtTemp2 = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & "," & Right(msgTxtTemp, Len(msgTxtTemp) - InStr(msgTxtTemp, ","))

    Dim bid

    bid = Split(msgTxtTemp, ";")

    Dim lotStr  As String

    Dim lotStr2 As String

    For i = 0 To UBound(bid) - 1
        lotStr = bid(i)
    
        '对字符串，进行拆开
    
        'SPZ14017G5HCN,uC3311Z.N.A04,GE8189,1606,387278,38944,58981.1
    
        Dim bid2

        bid2 = Split(lotStr, ",")
    
        fwTemp = bid2(0)
        dTemp = bid2(1)
        lottemp = bid2(2)
        dcTemp = bid2(3)
        dqTemp = bid2(4)
        bagTemp = bid2(5)
        snTemp = bid2(6)
    
        '判断
    
        If (Judge37FabDataICI(bagTemp)) Then

            'MsgBox "这笔：" & semtechFabTemp.PurchaseNo & "  " & semtechFabTemp.Batch & " 已存在，无需上传!", vbInformation, "友情提示"

        Else
       
            cmdStr2 = "insert into MAPPINGDATA37( devicename,batch,wf,flag,Qtech_Created_By,Qtech_Created_Date ," & " id,Maptype,Assydevicename,sn,datecode,Bagno) values (" & " '" & fwTemp & "', '" & lottemp & "','" & dqTemp & "','Y','" & gUserName & "',sysdate," & "  CUSTOMER37FabID_SEQ.nextval,'ICI','" & dTemp & "','" & snTemp & "','" & dcTemp & "','" & bagTemp & "')"
            
            AddSql (cmdStr2)
                
            ' waferidNoTemp = bidWaferID(n)
            'waferStrTemp = snTemp & "01"
                
            waferStrTemp = bagTemp
                          
            Call Add37FabICIDetail(CStr(snTemp), waferStrTemp, 1, CLng(dqTemp))
                
            SumCount = SumCount + 1
                
        End If
    
    Next i
   
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！", vbInformation, "友情提示"

    End If

End Sub

Private Sub cmdQes_Click()

    Dim beginTime As String

    Dim endTime   As String

    Dim sqlOra    As String

    Dim cmdOra    As String

    Dim DEV       As String

    ' 按照时间查找
    DEV = UCase(Trim(txtText10.text))
    beginTime = Format(DTP1(1).Value, "YYYY/MM/DD")
    endTime = Format(DTP2(1).Value, "YYYY/MM/DD")

    sqlOra = "select a.DEVICENAME ,a.BATCH, a.WF, a.WAFER_ID , a.qtech_created_date, a.qtech_created_by from MAPPINGDATA37 a, mappingdatatest b" & " where a.BATCH = b.lotid and a.qtech_created_date >= to_date('" + beginTime + "','YYYY/MM/DD') and a.qtech_created_date <=to_date('" + endTime + "','YYYY/MM/DD') "

    ' 按照机种查询
    If DEV <> "" Then
        sqlOra = sqlOra & " and a.DEVICENAME = '" & DEV & "'"

    End If

    ExporToExcel (sqlOra)

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
    oiRecordTemp.PoNum = Trim(txtPO.text)
    oiRecordTemp.PoItem = Trim(txtPOItem.text)
    oiRecordTemp.LOTID = Trim(txtLotID.text)
    oiRecordTemp.MPN = Trim(txtMPN.text)
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
    oiRecordTemp.ShipSite = Trim(txtShipSite.text)
    oiRecordTemp.LotStatus = Trim(txtLotStatus.text)
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

            'mapTemp.SUBSTRATEID = xn01.Attributes(1).nodeValue
            mapTemp.SUBSTRATEID = xn01.Attributes(3).nodeValue  ' add
            
            If Len(mapTemp.SUBSTRATEID) <> 11 Then

                mapTemp.SUBSTRATEID = xn01.Attributes(1).nodeValue
            End If
        
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
                    mapTemp.LOTID = xn02.Attributes(6).nodeValue    'add
                    
                    mapTemp.LOTID = Replace$(mapTemp.LOTID, ".", "")
                    mapTemp.PRODUCTID = xn02.Attributes(6).nodeValue
                    mapTemp.CreateDate = xn02.Attributes(8).nodeValue
                    mapTemp.CreateDate = xn02.Attributes(2).nodeValue   'add
                    
                    
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
    If CmbCustomer.text = "" Then
        MsgBox "请先选择客户！"

        Exit Sub

    End If

    Dim FName

    '筛选文件类型
    If CmbCustomer.text <> "KR002" Then
        ComSI.Filter = "map文件(*.map)|*.map|txt文件(*.txt)|*.txt|FIT_A0 文件(*.FIT_A0)|*.FIT_A0|PT1文件(*.PT1)|*.PT1|ASC文件(*.ASC)|*.ASC|"

    End If
    
    ComSI.ShowOpen
    '得到文件名
    FName = ComSI.filename

    If FName <> "" Then
        TxtSI.text = Replace(FName, Chr(0), ",")

    End If

End Sub

Private Sub Command14_Click()

    If CmbCustomer.text = "" Then
        MsgBox "请先选择客户！"

        Exit Sub

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
    
    If InStr(1, filename, ",") > 0 Then
        dirtemp = Split(filename, ",")
        
        For i = 1 To UBound(dirtemp)

            If CmbCustomer.text = "GT" Or CmbCustomer.text = "SI" Then
             
                UpMap (dirtemp(0) + "\" + dirtemp(i))
            
            ElseIf CmbCustomer.text = "HD" Then
                'HD客户
                UpMapHD (dirtemp(0) + "\" + dirtemp(i))
                 
            ElseIf CmbCustomer.text = "GC" Then
                'HD客户
                UpMapGCWlt (dirtemp(0) + "\" + dirtemp(i))
                 
            ElseIf CmbCustomer.text = "MG" Then
                'MG客户
                UpMapMG (dirtemp(0) + "\" + dirtemp(i))
                  
            ElseIf CmbCustomer.text = "56" Then
            
                UpMap56 (dirtemp(0) + "\" + dirtemp(i))
                
            ElseIf CmbCustomer.text = "TW058" Then
             
                UpMapTW058 (dirtemp(0) + "\" + dirtemp(i))
                
            ElseIf CmbCustomer.text = "KR001" Then
                UpMapKR001 (dirtemp(0) + "\" + dirtemp(i))
                
            ElseIf CmbCustomer.text = "KR002" Then
                UpMapKR002 (dirtemp(0) + "\" + dirtemp(i))
                
            ElseIf CmbCustomer.text = "AT71" Or CmbCustomer.text = "AH033" Then
                UpMapAT71 (dirtemp(0) + "\" + dirtemp(i))
                
            ElseIf CmbCustomer.text = "HW50" Then
                UpMapHW50 (dirtemp(0) + "\" + dirtemp(i))
                
            ElseIf CmbCustomer.text = "AB18" Then
                UpMapAB18 (dirtemp(0) + "\" + dirtemp(i))
            
            ElseIf CmbCustomer.text = "JS150" Then
                UpJS150 (dirtemp(0) + "\" + dirtemp(i))
            
            ElseIf CmbCustomer.text = "77" Then
                Up77 (dirtemp(0) + "\" + dirtemp(i))
                
            ElseIf CmbCustomer.text = "KR009" Then
                UpKR009 (dirtemp(0) + "\" + dirtemp(i))
                
            ElseIf CmbCustomer.text = "EU008" Then
                UpEU008 (dirtemp(0) + "\" + dirtemp(i))
            End If
            
        Next
        
    Else

        If CmbCustomer.text = "GT" Or CmbCustomer.text = "SI" Then
        
            UpMap (filename)
        
        ElseIf CmbCustomer.text = "HD" Then
            'HD客户
            UpMapHD (filename)
         
        ElseIf CmbCustomer.text = "AB18" Then
            'HD客户
            UpMapAB18 (filename)
         
        ElseIf CmbCustomer.text = "GC" Then
            'GC客户   2015-03-20 jiayun add
            UpMapGCWlt (filename)
         
        ElseIf CmbCustomer.text = "MG" Then
            UpMapMG (filename)
         
        ElseIf CmbCustomer.text = "56" Then
        
            UpMap56 (filename)
        
        ElseIf CmbCustomer.text = "TW058" Then
        
            UpMapTW058 (filename)
        
        ElseIf CmbCustomer.text = "KR001" Then
        
            UpMapKR001 (filename)
        ElseIf CmbCustomer.text = "KR002" Then
        
            UpMapKR002 (filename)
            
        ElseIf CmbCustomer.text = "AT71" Or CmbCustomer.text = "AH033" Then
        
            UpMapAT71 (filename)
            
        ElseIf CmbCustomer.text = "HW50" Then
            UpMapHW50 (filename)
        
        ElseIf CmbCustomer.text = "JS150" Then
            UpJS150 (filename)
            
        ElseIf CmbCustomer.text = "77" Then
            Up77 (filename)
        ElseIf CmbCustomer.text = "KR009" Then
            UpKR009 (filename)
        
        ElseIf CmbCustomer.text = "EU008" Then
            UpEU008 (filename)
        End If
        
    End If
    
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"
        SumCount = 0

    End If
    
    If ErrorInf <> "" Then
        MsgBox "上传失败的有:" + ErrorInf + "数据库中已存在！"

    End If

End Sub

Public Sub UpMap(dirtemp As String)

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

    Close #1
    ' 加入打标码. 20171218 tw
    mapTemp.PRODUCTID = GetGTCodeID()

    If (JudgeFlagStauts(mapTemp.SUBSTRATEID)) Then
        MsgBox "这笔：" & mapTemp.SUBSTRATEID & "已存在，无需上传!"
       
    Else
       
        Call AddMap(mapTemp, customerNameTemp)

    End If

End Sub

Private Sub Up37jono(dirtemp As String)

    Dim JOBNO As String

    Dim aer   As String

End Sub

Private Sub Up37POData(dirtemp As String)

    Dim source_batch_id_Temp As String

    Dim lotTypeTemp          As String

    Dim dirName              As String

    Dim filename             As String

    Set VBExcel = CreateObject("excel.application")

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(dirtemp)

    Set xlSheet = xlBook.Worksheets(1)

    Dim i              As Integer

    Dim j              As Integer

    Dim id             As Long

    Dim TEMP           As String

    Dim temp2          As String

    Dim tempVal        As String

    Dim dieQtyTemp     As Long

    Dim pcsQtemp       As Integer

    Dim dateFormatTemp As String

    Dim JOBNO          As String

    Dim Price          As String

    Dim price1         As Integer

    Dim unit           As String

    Dim unit1          As Integer

    Dim verDateStr     As String

    Dim shipAddResStr  As String

    Dim serviceStr     As String

    Dim plantStr       As String

    Dim keyRecord      As String

    Dim waferIdTemp    As String

    Dim po1lot1        As String

    Dim po1lot2        As String

    Dim po1lot3        As String

    verDateStr = ""
    shipAddResStr = ""
    serviceStr = ""
    plantStr = ""

    lotTypeTemp = "P"

    BCResultFlag = False

    For i = 2 To xlSheet.Range("a1").CurrentRegion.Rows.count
   
        For j = 1 To 46
             
            If j <= 26 Then
                strChar = Chr(96 + j)
            Else
                strChar = "A" & Chr(96 + j - 26)

            End If

            tempVal = xlSheet.Range(strChar & i).Value
                 
            If j = 9 Then
                semPotemp.PurchaseOrderNo = Trim(tempVal)

            End If
                 
            If j = 46 Then
                semPotemp.ProductionOrderNo = Trim(tempVal)
            End If
                 
            If j = 4 Then
                semPotemp.ShippingAddress = Trim(tempVal)
            End If
                
            If j = 20 Then
                semPotemp.CURRENCY = Trim(tempVal)
            End If
                
            If j = 10 Then
                semPotemp.ITEM = CInt(Trim(tempVal))
                  
            End If
                
            If j = 14 Then
                semPotemp.Quantity = CInt(Trim(tempVal))
                  
            End If
                
            If j = 15 Then
                semPotemp.UM = Trim(tempVal)
                  
            End If
                
            If j = 17 Then
                semPotemp.DelDate = Trim(tempVal)
                  
            End If
                 
            If j = 18 Then
                Price = Trim(tempVal)

            End If
                
            If j = 19 Then
                unit = Trim(tempVal)
                semPotemp.UnitPrice = Price / unit

            End If
                
            If j = 21 Then
                semPotemp.NetAmount = CLng(Trim(tempVal))
                  
            End If
                
            If j = 11 Then
                semPotemp.YourMaterialNumber = Trim(tempVal)
                  
            End If
            
            If j = 26 Then
                waferIdTemp = Trim(tempVal)
                  
            End If
                
            If j = 35 Then
                semPotemp.PartNumber = Trim(tempVal)
                  
            End If
                
            If j = 16 Then
                po1lot1 = Trim(tempVal)
                semPotemp.JOBNO = po1lot1

            End If
                
            If j = 37 Then
                po1lot3 = Trim(tempVal)
                
            End If
                
            If j = 41 Then
                po1lot2 = Trim(tempVal)
                semPotemp.LotNO = po1lot2

            End If
                
            If j = 43 Then
                semPotemp.WaferFAB = Trim(tempVal)
                  
            End If
                
            If j = 44 Then
                semPotemp.WaferREV = Trim(tempVal)

            End If
                
        Next j
    
Recordcontinue:

    Next i

    If Len(po1lot2) > 0 Then
        semPotemp.LotNO = po1lot2
    Else
        semPotemp.LotNO = po1lot3

    End If

    semPotemp.id = GetMaxID()

    semPotemp.QTECH_CREATED_BY = gUserName

    semPotemp.KeyStr = semPotemp.PurchaseOrderNo & "_" & semPotemp.LotNO

'    If Len(semPotemp.PurchaseOrderNo) < 1 Then
'        MsgBox "PO_NUM不能为空!"
'
'        Exit Sub
'
'    End If

    Dim bidWaferID

    Dim N             As Integer

    Dim waferidNoTemp As String

    Dim waferStrTemp  As String

    If Len(waferIdTemp) > 0 And InStr(waferIdTemp, "Wafer ID:") > 0 Then
        waferIdTemp = Replace(waferIdTemp, "Wafer ID:", "")
                     
        bidWaferID = Split(waferIdTemp, ",")
                     
        For N = 0 To UBound(bidWaferID)
            waferidNoTemp = Trim(bidWaferID(N))
            waferStrTemp = semPotemp.LotNO & Right("0" & waferidNoTemp, 2)
                          
           ' Call Add37POwaferDetail(semPotemp.waferlot, waferStrTemp, waferidNoTemp, semPotemp.PurchaseOrderNo)
            Call update37pojobno(semPotemp.WAFERLOT, waferStrTemp, waferidNoTemp, semPotemp.PurchaseOrderNo, semPotemp)
                          
        Next

    Else
       ' Call update37pojobno1(semPotemp.waferlot, waferStrTemp, waferidNoTemp, semPotemp.PurchaseOrderNo, semPotemp)
        Call Add37POwaferDetail(semPotemp.WAFERLOT, waferStrTemp, waferidNoTemp, semPotemp.PurchaseOrderNo)
                      
    End If
     
    xlBook.Close

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

End Sub

Private Sub Up37PODataICI(dirtemp As String)
    '上传资料

    Dim source_batch_id_Temp As String

    Dim lotTypeTemp          As String

    Dim dirName              As String

    Dim filename             As String

    '2012-06-27 jiayunzhang 修改读Excel的方式

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(dirtemp)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同

    '    If xlSheet.Range("N1").CurrentRegion.Columns.Count <> 14 Then
    '          MsgBox "转换后的Excel列数与原模板不一致，请确认PO格式是否有变化！", vbInformation, "提示"
    '          SumCount = SumCount - 1
    '          Exit Sub
    '
    '    End If

    Dim i              As Integer

    Dim j              As Integer

    Dim id             As Long

    Dim TEMP           As String

    Dim temp2          As String

    Dim tempVal        As String

    Dim dieQtyTemp     As Long

    Dim pcsQtemp       As Integer

    Dim dateFormatTemp As String

    Dim verDateStr     As String

    Dim shipAddResStr  As String

    Dim serviceStr     As String

    Dim plantStr       As String

    Dim keyRecord      As String

    Dim waferIdTemp    As String

    verDateStr = ""
    shipAddResStr = ""
    serviceStr = ""
    plantStr = ""

    '循环Excel
   
    lotTypeTemp = "P"

    BCResultFlag = False

    For i = 7 To 36
   
        For j = 1 To 13
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
                 
            If i = 7 And j = 9 Then
                semPotemp.PurchaseOrderNo = Trim(tempVal)
                GoTo Recordcontinue

            End If
                 
            If i = 8 And j = 9 Then
                semPotemp.ProductionOrderNo = Trim(tempVal)
                GoTo Recordcontinue

            End If
                 
            If i = 9 And j = 9 Then
                '3.Version 4.Date:
                verDateStr = Trim(tempVal)
                semPotemp.Version = Left(verDateStr, 1)
                semPotemp.DATE = Right(verDateStr, Len(verDateStr) - 1)
                GoTo Recordcontinue

            End If
                 
            If (i = 16 Or i = 17 Or i = 18 Or i = 19) And j = 4 Then
                '3.ship address
                shipAddResStr = shipAddResStr & Trim(tempVal) & " "
                GoTo Recordcontinue

            End If
                 
            If (i = 20) And j = 4 Then
                '3.ship address
                shipAddResStr = shipAddResStr & Trim(tempVal) & " "
                semPotemp.ShippingAddress = shipAddResStr
                GoTo Recordcontinue

            End If
                 
            If (i = 21) And j = 4 Then
                semPotemp.TermsPayment = Trim(tempVal)

            End If
                 
            If (i = 21) And j = 12 Then
                semPotemp.CURRENCY = Trim(tempVal)
                GoTo Recordcontinue

            End If
                
            If (i = 22) And j = 4 Then
                semPotemp.TermsDelivery = Trim(tempVal)
                GoTo Recordcontinue

            End If
                
            If (i = 23) And j = 3 Then
                semPotemp.FreightCarrier = Trim(tempVal)
                GoTo Recordcontinue

            End If
                
            If (i = 25) And j = 1 Then
                semPotemp.ITEM = CInt(Trim(tempVal))
                  
            End If
                
            'add lotid
                
            If (i = 25) And j = 2 Then
                semPotemp.MaterialDes = Trim(Left(Trim(tempVal), InStr(Trim(tempVal), " ")))
                  
                semPotemp.LotNO = Right(Trim(tempVal), Len(Trim(tempVal)) - InStr(Trim(tempVal), " "))
                  
            End If
                
            If (i = 25) And j = 6 Then
                'semPoTemp.Quantity = CLng(Left(Trim(tempVal), Len(Trim(tempVal)) - 5))
                semPotemp.UM = Right(Trim(tempVal), 2)
                  
            End If
                
            '                 If (i = 25) And j = 7 Then
            '                  semPoTemp.UM = Trim(tempVal)
            '
            '                End If
                
            If (i = 25) And j = 8 Then
                semPotemp.DelDate = Trim(tempVal)
                  
            End If
                 
            If (i = 25) And j = 11 Then
                semPotemp.UnitPrice = Trim(tempVal)
                  
                '处理单价
                  
                Dim tt1 As Double

                tt1 = CDbl(Left(Trim(tempVal), InStr(Trim(tempVal), "/") - 1))
                    
                Dim tt2 As Double

                tt2 = CDbl(Replace(Mid(Trim(tempVal), InStr(Trim(tempVal), "/") + 1, Len(Trim(tempVal))), "EA", ""))
                    
                Dim tt3 As Double

                tt3 = tt1 / tt2

                semPotemp.POPrice = tt3
                  
            End If
                
            If (i = 25) And j = 13 Then
                semPotemp.NetAmount = CLng(Trim(tempVal))
                GoTo Recordcontinue
                  
            End If
                
            If (i = 28) And j = 5 Then
                semPotemp.YourMaterialNumber = Trim(tempVal)
                GoTo Recordcontinue
                  
            End If
                
            If (i = 29) And j = 4 Then
                serviceStr = Trim(tempVal)
                semPotemp.TypeService = serviceStr
                GoTo Recordcontinue
                  
            End If
                
            If (i = 30) And j = 2 Then
                plantStr = Trim(tempVal)
                semPotemp.MfgPlant = Trim(Right(plantStr, Len(plantStr) - 10))
                  
                'semPoTemp.ReceivingPlant = Mid(plantStr, InStr(InStr(1, plantStr, "Receiving"), plantStr, ":") + 1, 25)
                  
                GoTo Recordcontinue
                  
            End If
                
            Dim splitStr

            'jiayun add bagno,datecode
            If (i = 31) And j = 2 Then
                
                splitStr = Split(Trim(tempVal), ";")
                
                If InStr(splitStr(0), "B#") > 0 Then
                
                    semPotemp.BagNo = Replace(splitStr(0), "B#", "")
                
                ElseIf InStr(splitStr(0), "D#") > 0 Then
                
                    semPotemp.DATECODE = Replace(splitStr(0), "D#", "")
                
                End If
                
                If InStr(splitStr(1), "B#") > 0 Then
                
                    semPotemp.BagNo = Replace(splitStr(1), "B#", "")
                
                ElseIf InStr(splitStr(1), "D#") > 0 Then
                
                    semPotemp.DATECODE = Replace(splitStr(1), "D#", "")
                
                End If

                GoTo Recordcontinue

            End If

            '                If (i = 30) And j = 2 Then
            '                  waferidTemp = Trim(tempVal)
            '                   GoTo Recordcontinue
            '
            '
            '                End If
                
            If (i = 36) And j = 1 Then
                semPotemp.PartNumber = Trim(tempVal)
                  
            End If
                
            If (i = 36) And j = 3 Then
                semPotemp.Quantity = CLng(Trim(tempVal))
                  
            End If
                
            If (i = 36) And j = 5 Then
                semPotemp.LotNumber = Trim(tempVal)
                  
            End If
                
            '                 If (i = 35) And j = 6 Then
            '                  semPoTemp.LotNumber = Trim(tempVal)
            '
            '                End If
                
            If (i = 36) And j = 7 Then
                semPotemp.WAFERLOT = Trim(tempVal)
                   
            End If
                
            If (i = 36) And j = 10 Then
                semPotemp.WaferFAB = Trim(tempVal)
                  
            End If
                
            If (i = 36) And j = 11 Then
                semPotemp.WaferREV = Trim(tempVal)
                GoTo Recordcontinue

            End If
                
        Next j
    
Recordcontinue:

    Next i

    semPotemp.id = GetMaxID()

    semPotemp.QTECH_CREATED_BY = gUserName

    semPotemp.KeyStr = semPotemp.PurchaseOrderNo & "_" & semPotemp.LotNO

    If (JudgeFlag37POHeader(semPotemp.KeyStr)) Then
        MsgBox "这笔：" & semPotemp.KeyStr & " 已存在，无需再次上传!"

        ' SumCount = SumCount - 1
        Exit Sub

    End If

    Call Add37POHeaderICI(semPotemp, "")
    
    '2016-04-05 jiayun add 把wo中的waferid 保存到表里
    
    '按waferid 所数据放到Mapping表里
    Dim bidWaferID

    Dim N             As Integer

    Dim waferidNoTemp As String

    Dim waferStrTemp  As String
                     
    bidWaferID = Split(waferIdTemp, ".")
                     
    '                      For n = 0 To UBound(bidWaferID)
    '                          waferidNoTemp = bidWaferID(n)
    waferStrTemp = semPotemp.LotNO & "01"
                          
    '                       Call Add37POwaferDetail1(semPoTemp.LotNo, waferStrTemp, 1, semPoTemp.PurchaseOrderNo)
    Call Add37POwaferDetail(semPotemp.LotNO, waferStrTemp, 1, semPotemp.PurchaseOrderNo)
                          
    '                      Next
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

End Sub

Private Sub Up68POData(dirtemp As String)
    '上传资料  MPS客户

    Dim source_batch_id_Temp As String

    Dim lotTypeTemp          As String

    Dim dirName              As String

    Dim filename             As String

    '2012-06-27 jiayunzhang 修改读Excel的方式

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(dirtemp)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同

    '    If xlSheet.Range("N1").CurrentRegion.Columns.Count <> 14 Then
    '          MsgBox "转换后的Excel列数与原模板不一致，请确认PO格式是否有变化！", vbInformation, "提示"
    '          SumCount = SumCount - 1
    '          Exit Sub
    '
    '    End If

    Dim i               As Integer

    Dim j               As Integer

    Dim id              As Long

    Dim TEMP            As String

    Dim temp2           As String

    Dim tempVal         As String

    Dim dieQtyTemp      As Long

    Dim pcsQtemp        As Integer

    Dim dateFormatTemp  As String

    Dim verDateStr      As String

    Dim shipAddResStr   As String

    Dim serviceStr      As String

    Dim plantStr        As String

    Dim keyRecord       As String

    Dim waferIdTemp     As String

    Dim recoredTemp     As String

    Dim lotIdTemp       As String

    Dim shitp           As String
  
    Dim insertFlag      As Boolean

    Dim insertBeginFlag As Boolean
  
    Dim insert34Flag    As Boolean
  
    Dim insertEndFlag   As Boolean

    verDateStr = ""
    shipAddResStr = ""
    serviceStr = ""
    plantStr = ""

    '循环Excel
   
    lotTypeTemp = "P"

    BCResultFlag = False

    insertFlag = False
    'For i = 1 To xlSheet.Range("A1").CurrentRegion.Rows.Count
 
    insertBeginFlag = False
    insertEndFlag = False
    insert34Flag = False
 
    For i = 1 To xlSheet.UsedRange.Rows.count
  
        If i >= 32 And (i - 32) Mod 7 = 0 Then
            insertBeginFlag = True

        End If
            
        If i >= 34 And (i - 34) Mod 7 = 0 Then
            insert34Flag = True

        End If
            
        If i >= 38 And (i - 38) Mod 7 = 0 Then
            insertEndFlag = True

        End If
  
        For j = 1 To 10
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
                 
            'po号
            If i = 1 And j = 9 Then
                semPotemp.PurchaseOrderNo = Trim(tempVal)
                GoTo Recordcontinue

            End If
                 
            '客户机种
            If i >= 32 And (i - 32) Mod 7 = 0 And j = 3 Then
                semPotemp.YourMaterialNumber = Trim(tempVal)
                  
                If semPotemp.YourMaterialNumber = "" Then
                  
                    GoTo RecordEndcontinue
                  
                End If
                  
                If (Not Judge68FabDieFlag(semPotemp.YourMaterialNumber)) Then

                    MsgBox "研发产品对照表中无此机种：" & semPotemp.YourMaterialNumber & " !", vbInformation, "友情提示"
        
                    Exit Sub

                End If
                  
            End If
                 
            '片数
            If i >= 32 And (i - 32) Mod 7 = 0 And j = 6 Then
                semPotemp.Quantity = Trim(tempVal)
                  
            End If
                 
            '日期
            If i >= 32 And (i - 32) Mod 7 = 0 And j = 7 Then
                semPotemp.DATE = Trim(tempVal)
                 
            End If
                 
            '单价
            If i >= 32 And (i - 32) Mod 7 = 0 And j = 10 Then
                semPotemp.UnitPrice = Trim(tempVal)
                GoTo Recordcontinue

            End If
                 
            'lotid
            If i >= 34 And (i - 34) Mod 7 = 0 And j = 4 Then
                semPotemp.LotNO = Trim(tempVal)
                 
            End If
                 
            'waferid
            If i >= 34 And (i - 34) Mod 7 = 0 And j = 5 Then
                semPotemp.waferIDList = Trim(tempVal)
                GoTo Recordcontinue

            End If
                 
            'Fab lotid
                 
            If i >= 38 And (i - 38) Mod 7 = 0 And j = 1 Then
                If Trim(tempVal) = "General Instruction:" Then
                    insertFlag = True
                     
                End If
                  
            End If

            If i >= 38 And (i - 38) Mod 7 = 0 And j = 3 Then
                recoredTemp = Trim(tempVal)

                lotIdTemp = Mid(recoredTemp, InStr(recoredTemp, "#") + 1)
                semPotemp.WAFERLOT = Left(lotIdTemp, InStr(lotIdTemp, ";") - 1)
                    
                shitp = Mid(lotIdTemp, InStr(LCase(lotIdTemp), "to") + 2)
                semPotemp.ShippingAddress = Trim(Left(shitp, InStr(LCase(shitp), "after") - 1))
                    
                If insertFlag = True Then
                    '插入DB
                    '看PO号,LotID号是否已经 存在，如果存在，则退出
                    
                    If (JudgeFlag68POHeader(semPotemp.PurchaseOrderNo, semPotemp.LotNO)) Then
                        MsgBox "这笔：" & semPotemp.PurchaseOrderNo & " 已存在，无需再次上传!"

                        ' SumCount = SumCount - 1
                        Exit Sub

                    End If
                    
                    semPotemp.id = GetMaxID()
                    semPotemp.QTECH_CREATED_BY = gUserName
                    Call Add68POHeader(semPotemp, UCase(Trim(CmbCustomer37.text)), semPotemp.waferIDList)
                          
                    'SumCount = SumCount + 1
                           
                    insertBeginFlag = False
                    insertEndFlag = False
                    insert34Flag = False
                           
                    'Waferid处理
                           
                    '按waferid 所数据放到Mapping表里
                    Dim bidWaferID

                    Dim N             As Integer

                    Dim waferidNoTemp As String

                    Dim waferStrTemp  As String

                    Dim waferDieQty   As Long
                            
                    bidWaferID = Split(semPotemp.waferIDList, ".")
                            
                    waferDieQty = Get68DieQty(semPotemp.YourMaterialNumber)
                            
                    For N = 0 To UBound(bidWaferID)
                        waferidNoTemp = bidWaferID(N)
                        waferStrTemp = semPotemp.WAFERLOT & Right(("0" & waferidNoTemp), 2)

                        Call Add68POwaferDetail(semPotemp.WAFERLOT, waferStrTemp, waferidNoTemp, UCase(Trim(CmbCustomer37.text)), waferDieQty, semPotemp.id)
                                 
                    Next

                End If
                                      
                GoTo Recordcontinue
                  
            End If
                 
            If i > 38 And j = 5 Then
                
                If InStr(Trim(tempVal), "Total Amount") > 0 Then
               
                    GoTo RecordEndcontinue
               
                End If
                
            End If

        Next j
    
Recordcontinue:

    Next i

    'semPoTemp.id = GetMaxID()
    '
    'semPoTemp.QTECH_CREATED_BY = gUserName
    '
    'semPoTemp.KeyStr = semPoTemp.PurchaseOrderNo & "_" & semPoTemp.LotNo
    '
    '    If (JudgeFlag37POHeader(semPoTemp.KeyStr)) Then
    '       MsgBox "这笔：" & semPoTemp.KeyStr & " 已存在，无需再次上传!"
    '         ' SumCount = SumCount - 1
    '       Exit Sub
    '    End If
    '
    '
    '    Call Add37POHeader(semPoTemp)
    '
    '    '2016-04-05 jiayun add 把wo中的waferid 保存到表里
    '
    '     '按waferid 所数据放到Mapping表里
    '                     Dim bidWaferID
    '                     Dim n As Integer
    '                     Dim waferidNoTemp As String
    '                     Dim waferStrTemp As String
    '
    '                     bidWaferID = Split(waferidTemp, ".")
    '
    '                      For n = 0 To UBound(bidWaferID)
    '                          waferidNoTemp = bidWaferID(n)
    '                          waferStrTemp = semPoTemp.WaferLot & waferidNoTemp
    '
    '                        Call Add37POwaferDetail(semPoTemp.WaferLot, waferStrTemp, waferidNoTemp, semPoTemp.PurchaseOrderNo)
    '
    '
    '                      Next
    
RecordEndcontinue:
     
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

End Sub

Public Sub UpMap56(dirtemp As String)

    Dim FLAG             As Integer

    Dim JudgeFlag        As Boolean

    Dim customerNameTemp As String

    Dim productaNameTenp As String

    Dim waferIDSeq       As String

    Dim allDieQty        As Long

    Dim goodDieQty       As Long

    Dim badDieQty        As Long

    Dim fileNameTemp     As String

    fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
    mapTemp.filename = fileNameTemp
    customerNameTemp = "56"
 
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
    
        If InStr(TextLine, "Lot id") > 0 Then
            mapTemp.LOTID = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))

        End If
    
        If InStr(TextLine, "Wafer ID") > 0 Then
            waferIDSeq = Right("0" & Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20)), 2)
            mapTemp.SUBSTRATEID = mapTemp.LOTID & waferIDSeq

        End If
    
        If InStr(TextLine, "Gross Dice") > 0 Then
            'qty
            allDieQty = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     
        End If
    
        If InStr(TextLine, "Good Dice") > 0 Then
            'qty
            mapTemp.PASSBINCOUNT = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     
            mapTemp.FailBinCount = CLng(allDieQty) - mapTemp.PASSBINCOUNT
    
        End If

        If InStr(TextLine, "Yield") > 0 Then
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

Public Sub UpMapKR001(dirtemp As String)

    Dim FLAG             As Integer

    Dim JudgeFlag        As Boolean

    Dim customerNameTemp As String

    Dim waferIdTemp      As String

    Dim waferIDSeq       As String

    Dim allDieQty        As Integer

    Dim goodDieQty       As Integer

    Dim badDieQty        As Integer

    Dim productTemp      As String

    Dim fileNameTemp     As String

    fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
    mapTemp.filename = fileNameTemp
    customerNameTemp = "KR001"
 
    Dim TextLine As String

    Open dirtemp For Input As #1

    Do While Not EOF(1)
        Line Input #1, TextLine

        If InStr(TextLine, "LOT_NO") > 0 Then
            mapTemp.LOTID = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
             mapTemp.LOTID = Replace(mapTemp.LOTID, "RS", "")
     
        End If

        If InStr(TextLine, "WAFER_ID") > 0 Then

            waferIdTemp = Trim(Mid(TextLine, InStr(TextLine, "WAFER_ID   :") + 12, 20))
     
            mapTemp.SUBSTRATEID = mapTemp.LOTID & waferIdTemp
     
        End If
    
        If InStr(TextLine, "TOTAL_DIE") > 0 Then

            mapTemp.TotalQty = Trim(Mid(TextLine, InStr(TextLine, "TOTAL_DIE  :") + 12, 10))
     
        End If
    
        If InStr(TextLine, "GOOD_DIE") > 0 Then
     
            mapTemp.PASSBINCOUNT = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
 
            mapTemp.FailBinCount = mapTemp.TotalQty - mapTemp.PASSBINCOUNT
    
        End If

        If InStr(TextLine, "TEST_TIME") > 0 Then
            GoTo ContinueFlag
    
        End If

    Loop

ContinueFlag:

    Close #1
            
    Call updateKR001Map(mapTemp.SUBSTRATEID, mapTemp.PASSBINCOUNT, mapTemp.FailBinCount)

End Sub

Public Sub UpMapKR002(dirtemp As String)
Dim FLAG             As Integer
Dim JudgeFlag        As Boolean
Dim customerNameTemp As String
Dim waferIdTemp      As String
Dim waferIDSeq       As String
Dim allDieQty        As Integer
Dim goodDieQty       As Integer
Dim badDieQty        As Integer
Dim productTemp      As String
Dim fileNameTemp     As String

fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
mapTemp.filename = fileNameTemp
customerNameTemp = "KR002"
Dim TextLine As String

Open dirtemp For Input As #1

Do While Not EOF(1)
    Line Input #1, TextLine
    'LotID
    If InStr(TextLine, "Lot") > 0 Then
        'lotid
        mapTemp.LOTID = Trim(Split(TextLine, ":")(1))

    End If

    'WaferID
    If InStr(TextLine, "Wafer") > 0 Then
     '   waferIdTemp = Trim(Mid(TextLine, InStr(TextLine, "WAFER_ID   :") + 12, 20))
        
        waferIdTemp = Right("00" & Trim(Split(TextLine, ":")(1)), 2)
        
        mapTemp.SUBSTRATEID = mapTemp.LOTID & waferIdTemp

    End If

    If InStr(TextLine, "A - Grade Die") > 0 Then
        mapTemp.PASSBINCOUNT = Trim(Mid(TextLine, InStr(TextLine, "Grade Die  :") + 16, 10))
        allDieQty = GetKRGrossDie(mapTemp.SUBSTRATEID)
        mapTemp.FailBinCount = allDieQty - mapTemp.PASSBINCOUNT

    End If

    If InStr(TextLine, "B - Grade Die") > 0 Then
        GoTo ContinueFlag

    End If

Loop
ContinueFlag:
Close #1    ' 关闭文件。
Call updateKR002Map(mapTemp.SUBSTRATEID, mapTemp.PASSBINCOUNT, mapTemp.FailBinCount)

End Sub

Public Sub UpMapAT71(dirtemp As String)

    Dim FLAG             As Integer

    Dim JudgeFlag        As Boolean

    Dim customerNameTemp As String

    Dim waferIdTemp      As String

    Dim sOra             As String

    Dim sSql             As String

    Dim rs               As ADODB.Recordset

    Dim waferIDSeq       As String

    Dim allDieQty        As Integer

    Dim goodDieQty       As Long

    Dim badDieQty        As Long
    
    Dim strLotID As String, strWaferID As String

    Dim productTemp      As String

    Dim fileNameTemp     As String

    fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
    mapTemp.filename = fileNameTemp
 
    'SI Mapping

    Dim TextLine   As String

    Dim aWaferID() As String

    Open dirtemp For Input As #1

    ' 打开文件。
    Do While Not EOF(1)
        ' 循环至文件尾。
        Line Input #1, TextLine

        aWaferID() = Split(TextLine)
        
        If UBound(aWaferID) = 0 Then
            If Get_OracleCnt("select * from mappingdatatest where lotid = '" & aWaferID(0) & "'") > 0 Then
                strLotID = aWaferID(0)
            End If
        End If

        If TextLine <> "" And UBound(aWaferID) <> 0 Then
            aWaferID(0) = Right(aWaferID(0), 2)
            strWaferID = strLotID & aWaferID(0)
            
            sOra = "select * from mappingdatatest where substrateid = '" & strWaferID & "'"
            Set rs = Get_OracleRs(sOra)

            If rs.RecordCount > 0 Then

                For i = 1 To UBound(aWaferID)

                    If IsNumeric(aWaferID(i)) Then
                
                        goodDieQty = goodDieQty + aWaferID(i)
                
                    End If
            
                Next
    
                sOra = "select (passbincount + failbincount) from mappingdatatest where substrateid = '" & strWaferID & "'"
        
                badDieQty = Get_OracleStr(sOra) - goodDieQty
        
                Call updateAT71Map(strWaferID, goodDieQty, badDieQty)
                goodDieQty = 0

            End If

        End If

    Loop

ContinueFlag:

    Close #1    ' 关闭文件。

End Sub

Public Sub UpMapTW058(dirtemp As String)

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

    ' 判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
       
    If (JudgeFlagStauts(mapTemp.SUBSTRATEID)) Then
        MsgBox "这笔：" & mapTemp.SUBSTRATEID & "已存在，无需上传!"
       
    Else
       
        Call AddMap(mapTemp, customerNameTemp)

    End If

End Sub

Public Sub UpMapHW50(dirtemp As String)

    Dim fileNameTemp As String

    fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
    mapTemp.filename = fileNameTemp
    customerNameTemp = "HW50"

    Dim TextLine   As String

    Dim sLotLen    As String

    Dim waferIDSeq As String
    
    Dim ss

    Open dirtemp For Input As #1

    ' 打开文件。
    Do While Not EOF(1)
        ' 循环至文件尾。
        Line Input #1, TextLine

        ' lot长度
        If InStr(TextLine, "Wafer Size") Then
            sLotLen = Mid(TextLine, InStr(TextLine, ":") + 2)

        End If
    
        If InStr(TextLine, "Wafer ID") > 0 Then
'            'mapTemp.LOTID = Trim(Mid(TextLine, InStr(TextLine, ":") + 2, sLotLen)) & "A"
'            mapTemp.LOTID = Trim(Mid(TextLine, InStr(TextLine, ":") + 2, sLotLen))
'            waferIDSeq = Trim(Mid(Mid(TextLine, InStr(TextLine, "_") + 1), sLotLen + 2))
'            mapTemp.SUBSTRATEID = mapTemp.LOTID & Right("0" & waferIDSeq, 2)
            
            ss = Split(TextLine, "-")
            mapTemp.LOTID = Trim(Replace(ss(0), "Wafer ID:", ""))
             waferIDSeq = ss(UBound(ss))
            mapTemp.SUBSTRATEID = mapTemp.LOTID & Right("0" & waferIDSeq, 2)
            
     
        End If
    
        If InStr(TextLine, "Pass Dice") > 0 Then
            'qty
            mapTemp.PASSBINCOUNT = Trim(Mid(TextLine, InStr(TextLine, ":") + 2))

        End If

        If InStr(TextLine, "Fail Dice") > 0 Then
    
            mapTemp.FailBinCount = Trim(Mid(TextLine, InStr(TextLine, ":") + 2))
        
            GoTo ContinueFlag
    
        End If

    Loop

ContinueFlag:

    Close #1    ' 关闭文件。
            
    Call updateHW50Map(mapTemp.SUBSTRATEID, mapTemp.PASSBINCOUNT, mapTemp.FailBinCount)

End Sub

Public Sub UpJS150(dirtemp As String)

    Dim strCode As String

    Dim strPart

    Dim strLotID    As String

    Dim strWaferID  As String

    Dim strGoodDies As String

    Open dirtemp For Input As #1

    Do While Not EOF(1)
        Line Input #1, strCode
    
        If InStr(UCase(strCode), "CUSTOMER") > 0 Then
            Close #1
            UpJS150P (dirtemp)
            Exit Sub

        End If
    
        If InStr(UCase(strCode), "LOT") > 0 Then
            strPart = Split(strCode)
            strLotID = strPart(UBound(strPart))

        End If
    
        If InStr(UCase(strCode), "WAFERID") > 0 Then
            strPart = Split(strCode)
            strWaferID = strPart(UBound(strPart))
        
        End If
    
        If InStr(UCase(strCode), "TOTAL") > 0 Then
            strPart = Split(strCode)
            strGoodDies = strPart(UBound(strPart))
        
        End If
    
        If strLotID <> "" And strWaferID <> "" And strGoodDies <> "" Then
            Call UpdateMapInfo(strLotID, strWaferID, strGoodDies)
            Close #1
            Exit Sub

        End If
    
    Loop

End Sub

Private Sub UpJS150P(dirtemp As String)

    Dim strCode As String

    Dim strPart

    Dim strLotID    As String

    Dim strWaferID  As String

    Dim strGoodDies As String

    strLotID = ""
    Open dirtemp For Input As #1

    Do While Not EOF(1)
        Line Input #1, strCode
 
        If strLotID = "" Then
            strPart = Split(dirtemp, "\")
            strLotID = Split(strPart(UBound(strPart)), "-")(0) & ".1"

        End If
        
        If strWaferID = "" Then
            strPart = Split(dirtemp, "\")
            strWaferID = Replace(Split(strPart(UBound(strPart)), "-")(1), ".txt", "")
        
        End If
    
        If InStr(UCase(strCode), "GOOD") > 0 Then
            strPart = Split(strCode, ":")
            strGoodDies = strPart(UBound(strPart))
        
        End If
    
        If strLotID <> "" And strWaferID <> "" And strGoodDies <> "" Then
            Call UpdateMapInfo3(strLotID, strWaferID, strGoodDies)
            Close #1
            Exit Sub

        End If
    
    Loop

End Sub

Public Sub Up77(dirtemp As String)

    Dim strCode As String

    Dim strPart

    Dim strLotID    As String

    Dim strWaferID  As String

    Dim strGoodDies As String

    Open dirtemp For Input As #1

    Do While Not EOF(1)
        Line Input #1, strCode
    
        If InStr(UCase(strCode), "LOT_NO") > 0 Then
            strPart = Split(Replace(strCode, " ", ""), ":")
            strLotID = Replace(strPart(1), "WAFER_ID", "")

        End If
    
        If InStr(UCase(strCode), "WAFER_ID") > 0 Then
            strPart = Split(Replace(strCode, " ", ""), ":")
            strWaferID = strPart(UBound(strPart))
        
        End If
    
        If InStr(UCase(strCode), "GOOD_DIE") > 0 Then
            strPart = Split(Replace(strCode, " ", ""), ":")
            strGoodDies = Replace(strPart(1), "TOTAL_DIE", "")
        
        End If
    
        If strLotID <> "" And strWaferID <> "" And strGoodDies <> "" Then
            Call UpdateMapInfo(strLotID, strWaferID, strGoodDies)
            Close #1
            Exit Sub

        End If
    
    Loop

End Sub

Public Sub UpKR009(dirtemp As String)

    Dim strCode As String

    Dim strPart

    Dim strLotID    As String

    Dim strWaferID  As String
    
    Dim strBadDies As String
    
    Dim strGoodDies As String

    Open dirtemp For Input As #1

    Do While Not EOF(1)
        Line Input #1, strCode

        If InStr(UCase(strCode), "WAFERID") > 0 Then
            strPart = Split(strCode, ":")
            strWaferID = Replace(strPart(UBound(strPart)), "-", "")
        
        End If
    
        If InStr(UCase(strCode), "#") > 0 Then
            strPart = Split(strCode, "-")
       
            If InStr(strCode, "D-") = 0 Then
                strGoodDies = strPart(UBound(strPart))
            Else
                strBadDies = strPart(UBound(strPart))
            End If
            
        End If
    
        If strWaferID <> "" And (strBadDies <> "" Or strGoodDies <> "") Then
            Call UpdateMapInfo2(strWaferID, strGoodDies, strBadDies)
            
            strWaferID = ""
            strBadDies = ""
            strGoodDies = ""
        End If
    
    Loop
    
    Close #1

End Sub

Public Sub UpEU008(dirtemp As String)

 Dim strCode As String

    Dim strPart

    Dim strLotID    As String

    Dim strWaferID  As String

    Dim strGoodDies As String

    Open dirtemp For Input As #1

    Do While Not EOF(1)
        Line Input #1, strCode

        If InStr(UCase(strCode), "LOTID:") > 0 Then
            strPart = Split(strCode, " ")
            strLotID = strPart(UBound(strPart)) & ".00"
        
        End If
    
        If InStr(UCase(strCode), "WAFERID:") > 0 Then
            strPart = Split(strCode, "-")
            strWaferID = strLotID & Left(strPart(UBound(strPart)), 2)
        
        End If
        
        If InStr(UCase(strCode), "PASS") > 0 Then
            strPart = Split(strCode)
            strGoodDies = strPart(UBound(strPart))
        
        End If
        
    
        If strWaferID <> "" And strGoodDies <> "" Then
            Call UpdateMapInfo2(strWaferID, strGoodDies, "")
            Close #1
            Exit Sub

        End If
    
    Loop

End Sub

Private Sub UpdateMapInfo(strLotID As String, strWaferID As String, strGoodDies As String)

    Dim strsql      As String

    Dim strSql2     As String

    Dim strWaferLot As String

    strWaferLot = strLotID & Right$("0" & strWaferID, 2)

    strsql = "update mappingdatatest set QTECH_LASTUPDATE_BY = '" & gUserRealName & "',QTECH_LASTUPDATE_DATE = sysdate,Passbincount = '" & strGoodDies & "', failbincount = (Passbincount + failbincount - '" & strGoodDies & "') where customershortname = '" & CmbCustomer.text & "' and lotid = '" & strLotID & "' and substrateid = '" & strWaferLot & "' "

    If AddSql(strsql) = 0 Then
        MsgBox strWaferLot & ":没有更新到Mapping, 请确认", vbExclamation, "警告"
    Else
    
        strSql2 = "update [ERPBASE].[dbo].[tblmappingData] set Passbincount = '" & strGoodDies & "', failbincount = (Passbincount + failbincount - '" & strGoodDies & "') where customershortname = '" & CmbCustomer.text & "' and lotid = '" & strLotID & "' and substrateid = '" & strWaferLot & "' "

        AddSql2 (strSql2)

        SumCount = SumCount + 1

    End If

End Sub

Private Sub UpdateMapInfo3(strLotID As String, _
                           strWaferID As String, _
                           strGoodDies As String)

    Dim strsql      As String

    Dim strSql2     As String

    Dim strWaferLot As String

    strWaferLot = strLotID & Right$("0" & strWaferID, 2)

    strsql = "update mappingdatatest set Passbincount = '" & strGoodDies & "', failbincount = (Passbincount + failbincount - '" & strGoodDies & "') where customershortname = '" & CmbCustomer.text & "' and lotid = '" & strLotID & "' and substrateid = '" & strWaferLot & "' "

    If AddSql(strsql) = 0 Then
        MsgBox strWaferLot & ":没有更新到Mapping, 请确认", vbExclamation, "警告"
    Else
    
        strSql2 = "update [ERPBASE].[dbo].[tblmappingData] set Passbincount = '" & strGoodDies & "', failbincount = (Passbincount + failbincount - '" & strGoodDies & "') where customershortname = '" & CmbCustomer.text & "' and lotid = '" & strLotID & "' and substrateid = '" & strWaferLot & "' "

        AddSql2 (strSql2)

        SumCount = SumCount + 1

    End If

End Sub

Private Sub UpdateMapInfo2(strWaferID As String, strGoodDies As String, strBadDies As String)

    Dim strsql  As String

    Dim strSql2 As String

    If strBadDies <> "" Then
        strsql = "update mappingdatatest set QTECH_LASTUPDATE_BY = '" & gUserRealName & "',QTECH_LASTUPDATE_DATE = sysdate,failbincount = '" & strBadDies & "',Passbincount =  (Passbincount + failbincount - '" & strBadDies & "') where customershortname = '" & CmbCustomer.text & "'  and substrateid = '" & strWaferID & "' "

        If AddSql(strsql) = 0 Then
            MsgBox strWaferLot & ":没有更新到Mapping, 请确认", vbExclamation, "警告"
        Else
    
            strSql2 = "update [ERPBASE].[dbo].[tblmappingData] set failbincount = '" & strBadDies & "',Passbincount =  (Passbincount + failbincount - '" & strBadDies & "') where customershortname = '" & CmbCustomer.text & "' and substrateid = '" & strWaferID & "' "

            AddSql2 (strSql2)

            SumCount = SumCount + 1

        End If

    Else
        strsql = "update mappingdatatest set QTECH_LASTUPDATE_BY = '" & gUserRealName & "',QTECH_LASTUPDATE_DATE = sysdate, failbincount = (Passbincount + failbincount - '" & strGoodDies & "') ,Passbincount = '" & strGoodDies & "'  where customershortname = '" & CmbCustomer.text & "'  and substrateid = '" & strWaferID & "' "

        If AddSql(strsql) = 0 Then
            MsgBox strWaferLot & ":没有更新到Mapping, 请确认", vbExclamation, "警告"
        Else
    
            strSql2 = "update [ERPBASE].[dbo].[tblmappingData] set  failbincount = (Passbincount + failbincount - '" & strGoodDies & "') ,Passbincount = '" & strGoodDies & "' where customershortname = '" & CmbCustomer.text & "' and substrateid = '" & strWaferID & "' "

            AddSql2 (strSql2)

            SumCount = SumCount + 1

        End If
        
    End If

End Sub

'2015-04-20 jiayun add MG

Public Sub UpMapMG(dirtemp As String)

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

Public Sub UpMapHD(dirtemp As String)

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
 
    'SI Mapping

    Dim TextLine As String

    Open dirtemp For Input As #1

    ' 打开文件。
    Do While Not EOF(1)
        ' 循环至文件尾。
        Line Input #1, TextLine

        '判断这行，是否要取资料，是则处理；否则下一行
        'LotID
        If InStr(TextLine, "Lot No") > 0 Then
            'lotid
            mapTemp.LOTID = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
   
        End If
    
        'WaferID
        If InStr(TextLine, "Wafer ID") > 0 Then
            'lotid
            ' mapTemp.lotID = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     
            'D02939-1
            waferIdTemp = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
            waferIdTemp = Mid(waferIdTemp, InStr(waferIdTemp, "-") + 1, 2)
     
            waferIDSeq = Right("0" & waferIdTemp, 2)
            mapTemp.SUBSTRATEID = mapTemp.LOTID & waferIDSeq
     
        End If
    
        If InStr(TextLine, "Total Pass") > 0 Then
            'qty
            mapTemp.PASSBINCOUNT = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
 
        End If
    
        If InStr(TextLine, "Total Fail") > 0 Then
            'qty
            mapTemp.FailBinCount = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
    
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

Public Sub UpMapAB18(dirtemp As String)

    Dim FLAG             As Integer

    Dim JudgeFlag        As Boolean

    Dim customerNameTemp As String

    Dim waferIdTemp      As String

    Dim qtysql           As String

    Dim waferIDSeq       As String

    Dim allDieQty        As Integer

    Dim goodDieQty       As Integer

    Dim badDieQty        As Integer

    Dim productTemp      As String

    Dim fileNameTemp     As String

    fileNameTemp = Mid(dirtemp, InStrRev(dirtemp, "\") + 1, Len(dirtemp) - InStrRev(dirtemp, "\") + 1)
    mapTemp.filename = fileNameTemp
    customerNameTemp = "HD"
 
    'SI Mapping

    Dim TextLine As String

    Open dirtemp For Input As #1

    ' 打开文件。
    Do While Not EOF(1)
        ' 循环至文件尾。
        Line Input #1, TextLine

        '判断这行，是否要取资料，是则处理；否则下一行
        'LotID
        If InStr(TextLine, "LOT ID") > 0 Then
            'lotid
            mapTemp.LOTID = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
            '     waferIDSeq = Trim(Mid(TextLine, InStrRev(TextLine, ":") + 1, Len(TextLine) - InStrRev(TextLine, ":")))
            '     mapTemp.SubstrateId = mapTemp.lotID & waferIDSeq
     
        End If
    
        'WaferID
        If InStr(TextLine, "WAFER ID") > 0 Then
            'lotid
            ' mapTemp.lotID = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     
            'D02939-1
            waferIdTemp = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
            waferIdTemp = Mid(waferIdTemp, InStr(waferIdTemp, "-") + 1, 2)
     
            waferIDSeq = Right("0" & waferIdTemp, 2)
            mapTemp.SUBSTRATEID = mapTemp.LOTID & waferIDSeq
     
        End If
    
        If InStr(TextLine, "PRODUCT ID") > 0 Then
            'lotid
            ' mapTemp.lotID = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     
            'D02939-1
            productTemp = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 20))
     
        End If
    
        If InStr(TextLine, "TESTED DIE") > 0 Then
            'qty
            mapTemp.TotalQty = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
     
        End If
    
        If InStr(TextLine, "PASS DIE") > 0 Then

            'qty
            Dim rs As ADODB.Recordset
    
            mapTemp.PASSBINCOUNT = Trim(Mid(TextLine, InStr(TextLine, ":") + 1, 10))
     
            qtysql = "select aa.passbincount + aa.failbincount  from mappingdatatest aa  WHERE aa.substrateid = '" & mapTemp.SUBSTRATEID & "' and rownum = 1  "
    
            mapTemp.FailBinCount = Get_OracleNo(qtysql) - mapTemp.PASSBINCOUNT

        End If

        If InStr(TextLine, "TEST_TIME") > 0 Then
            GoTo ContinueFlag
    
        End If

    Loop

ContinueFlag:

    Close #1    ' 关闭文件。

    Call updateAB18Map(mapTemp.SUBSTRATEID, mapTemp.PASSBINCOUNT, mapTemp.FailBinCount)

End Sub

Public Sub UpMapGCWlt(dirtemp As String)

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

Private Sub Command15_Click()

    Dim cust As String

    cust = CmbCustomer.text
    ExporToExcel ("select SUBSTRATEID, SUBSTRATETYPE, LOTID, PRODUCTID, CREATEDATE,MICRONLOTID, PASSBINCOUNT, FAILBINCOUNT, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE ,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE from mappingDataTest where customershortname  = '" + cust + "'  and qtech_created_date>sysdate-100  order by qtech_created_date desc , lotid, substrateid")

End Sub

Private Sub Command18_Click()

    CommonDialog9.flags = cdlOFNAllowMultiselect Or cdlOFNExplorer

    'GC
    On Error Resume Next

    Dim FName

    '帅选文件
    CommonDialog9.Filter = "CSV文件(*.csv)|*.csv"
    
    CommonDialog9.ShowOpen
    '得到文件名
    FName = CommonDialog9.filename

    If FName <> "" Then
      
        Text9.text = Replace(FName, Chr(0), ",")
        
    End If
    
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

Private Sub Command23_Click()

    If CmbCustomer37.text = "" Then
        MsgBox "请先选择客户代码", vbInformation, "提示"

        Exit Sub
    
    End If

    If CmbCustomer37.text = "37" Then

        ExporToExcel ("   select id, DEVICENAME,BATCH ,WF,PRICE ,CURRENCY ,SHIPPEDDT , PURCHASENO ,PURCHASEORDERLINEITEM ,INVOICE , MAWBNUMBER  ,DESTINATION  ,WAFER_ID   ,QTECH_CREATED_DATE " & " from  MAPPINGDATA37 a where  a.flag='Y' order by id desc ")
    ElseIf CmbCustomer37.text = "68" Or CmbCustomer37.text = "70" Then

        ExporToExcel ("   select id,  lotid as FabLot,substrateid,passbincount,failbincount, qtech_created_date " & " from  mappingdatatest a where customershortname in ('68','70') and a.flag='Y' order by id desc ")

    End If

End Sub

Private Sub Command24_Click()
    '导出PO

    If CmbCustomer37.text = "" Then
        MsgBox "请先选择客户代码", vbInformation, "提示"

        Exit Sub
    
    End If
    
    If CmbCustomer37.text = "37" And CmbPoType.text = "ICI" Then

        ExporToExcel (" select ID,PO_NUM as PurchaseOrderNo, MPN as ProductionOrderNo,  CREATED_DATE as PODate,  SHIPPING_MST_260 as Currency,  " & "  SHIP_SITE as ShippingAddress, COUNTRY_OF_ASSEMBLY as Termsofpayment,  PO_ITEM as Item, MPN_DESC as MaterialDescription,SOURCE_BATCH_ID as LotNo, SOURCE_MTRL_SLOC as WaferLot, " & "  DIE_QTY as Quantity, mtrl_num as BagNo,DATE_CODE as  DateCode, t_price as UnitPrice,  CURRENT_WAFER_QTY as NetAmount, " & "  SOURCE_MTRL_NUM as PartNumber, COUNTRY_OF_FAB as WaferFAB, IMAGER_CUSTOMER_REV as WaferREV " & "   from customeroitbl_test a  where  customershortname='37' and a.qtech_created_date>to_date('2016-03-26','YYYY-MM-DD') and a.flag='Y' order by id desc ")
  
    ElseIf CmbCustomer37.text = "37" And CmbPoType.text <> "ICI" Then

        ExporToExcel (" select ID,PO_NUM as PurchaseOrderNo, MPN as ProductionOrderNo,  CREATED_DATE as PODate,  SHIPPING_MST_260 as Currency,  " & "  SHIP_SITE as ShippingAddress, COUNTRY_OF_ASSEMBLY as Termsofpayment,  PO_ITEM as Item, MPN_DESC as MaterialDescription, SOURCE_MTRL_SLOC as LotNo, JOBNO, " & "  CURRENT_WAFER_QTY as Quantity, DATE_CODE as  DelDate, REF_PO as UnitPrice,  DIE_QTY as NetAmount, " & "  SOURCE_MTRL_NUM as PartNumber, SOURCE_BATCH_ID as WaferLot,COUNTRY_OF_FAB as WaferFAB, IMAGER_CUSTOMER_REV as WaferREV,mtrl_num as BagNo " & "   from customeroitbl_test a  where  customershortname='37' and a.qtech_created_date>to_date('2016-03-26','YYYY-MM-DD') and a.flag='Y' order by id desc ")
  
    ElseIf CmbCustomer37.text = "68" Or CmbCustomer37.text = "70" Then

        ExporToExcel (" select id, po_num as PurchaseOrderNo,a.mpn_desc as 客户机种,a.current_wafer_qty as WFRQTY ,  " & " created_date as StartDate,a.source_mtrl_sloc as MPSLot,a.reticle_level_71 as WaferID,source_batch_id as FabLot,ship_site as Shipto,REF_PO as  UnitPrice" & "   from customeroitbl_test a  where  customershortname in ('68','70') and a.qtech_created_date>to_date('2016-03-26','YYYY-MM-DD') and a.flag='Y' order by id desc ")
  
    End If

End Sub

Private Sub Command25_Click()
    'Semteck 晶圆数据上传

    '读取CSV
    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim cusPTTemp            As String

    Dim gcVerTemp            As String

    Dim gcVerLastTemp        As String

    Dim waferIDList          As String

    Dim customTemp           As String

    semtechFabTemp.custom_Temp = Trim(CmbCustomer37.text)

    If CmbCustomer37.text = "68" Or CmbCustomer37.text = "70" Or CmbCustomer37.text = "BJ128" Then
        UploadMPSFab

        Exit Sub
    
    End If
    
    If CmbCustomer37.text = "HK006" Then
        UploadMPSFab

        Exit Sub
    
    End If

    '上传OI的CSV
    '处理文件名
    If Text7.text = "" Then
    
        MsgBox "先选择待上传的文件!", vbInformation, "友情提示"
        
        Exit Sub

    End If

    Dim i             As Integer

    Dim j             As Integer

    Dim id            As Long

    Dim TEMP          As String

    Dim SumCount      As Integer

    Dim GCHeaderFlag  As Boolean

    Dim str01         As String

    Dim str03         As String

    Dim FabTem        As String

    Dim qtyTemp       As Long

    Dim bidWaferID()  As String

    Dim bidWaferID1   As String

    Dim waferidNoTemp As String

    SumCount = 0

    Dim k        As Integer
        
    Dim FName    As String

    Dim Nextline As String
        
    FName = Trim(Text7.text)
    ' First open file for check
    Open FName For Input As #1

    Do Until EOF(1)
        Line Input #1, Nextline
    
        ' 从第二行开始检查
        If UCase(Left(Trim(Nextline), 6)) <> "DEVICE" And Len(Nextline) > 15 Then
            bid = Split(Nextline, ",")
        
            semtechFabTemp.DeviceName = Trim(bid(0))
            semtechFabTemp.Batch = Trim(bid(1))
            semtechFabTemp.WF = CInt(Trim(bid(2)))
            semtechFabTemp.Price = CDbl(Trim(bid(3)))
            semtechFabTemp.CURRENCY = Trim(bid(4))
            semtechFabTemp.ShippedDt = CDate(Trim(bid(5)))
            semtechFabTemp.PurchaseNo = Trim(bid(6))
            semtechFabTemp.PurchaseOrderLineItem = Trim(bid(7))
            semtechFabTemp.Invoice = Trim(bid(8))
            semtechFabTemp.MAWBNumber = Trim(bid(9))
            semtechFabTemp.Destination = Trim(bid(10))
            semtechFabTemp.wafer_id = Trim(bid(11))
            semtechFabTemp.PoBatch = Trim(bid(12))
    
            waferIDList = semtechFabTemp.wafer_id

            ' Guard 1
            If (Not JudgeMPSBankPT(semtechFabTemp.DeviceName)) Then
            
                MsgBox "这机种：" & semtechFabTemp.DeviceName & " 在系统设定表中不存在，请联系市场部与研发部!", vbInformation, "友情提示"
                Close #1

                Exit Sub

            End If
    
            ' Guard 2
            If InStr(semtechFabTemp.wafer_id, ".") > 0 Then
                bidWaferID = Split(waferIDList, ".")
            
                If UBound(bidWaferID) + 1 <> Val(semtechFabTemp.WF) Then
                    MsgBox "Batch号为" & (semtechFabTemp.Batch) & "的条目, 片数与WaferID个数不一致, 请重新调整本次未上传任何batch", vbInformation, "友情提示"
                    Close #1

                    Exit Sub

                End If
            
                For k = 0 To UBound(bidWaferID)
                     
                    '根据po号，batch号，判断是否已经上传过
                    If (Judge37FabData(semtechFabTemp.Batch, bidWaferID(k))) Then
                        MsgBox "这笔：" & bidWaferID(k) & " 已存在，无需上传!", vbInformation, "友情提示"
                        Close #1

                        Exit Sub

                    End If

                Next
            
                For N = 0 To UBound(bidWaferID)
                    waferidNoTemp = bidWaferID(N)
                
                    If waferidNoTemp = "" Then
                        MsgBox "WaferId存在空值"
                        Close #1

                        Exit Sub
            
                    ElseIf Val(waferidNoTemp) > 25 Or Val(waferidNoTemp) < 1 Then
                        MsgBox "WaferId超出1-25范围"
                        Close #1

                        Exit Sub

                    End If
                
                    For KK = N + 1 To UBound(bidWaferID)
                        comparewaferid = bidWaferID(KK)
                    
                        If comparewaferid = waferidNoTemp Then
                            MsgBox "WaferId有重复"
                            Close #1

                            Exit Sub

                        End If

                    Next
                
                Next
            
            Else
                bidWaferID1 = waferIDList

                If (Judge37FabData(semtechFabTemp.Batch, bidWaferID1)) Then
                    MsgBox "这笔：" & bidWaferID1 & " 已存在，无需上传!", vbInformation, "友情提示"
                    Close #1

                    Exit Sub

                End If
            
                If Val(semtechFabTemp.WF) <> 1 Then
                    MsgBox "Batch号为" & (semtechFabTemp.Batch) & "的条目, 片数与WaferID个数不一致", vbInformation, "友情提示"
                    Close #1

                    Exit Sub

                End If
                       
                If bidWaferID1 = "" Then
                    MsgBox "Batch号为" & (semtechFabTemp.Batch) & "的条目, 片数与WaferID个数不一致", vbInformation, "友情提示"
                    Close #1

                    Exit Sub

                ElseIf Val(bidWaferID1) > 25 Or Val(bidWaferID1) < 1 Then
                    MsgBox "Batch号为" & (semtechFabTemp.Batch) & "的条目, 片数与WaferID个数不一致", vbInformation, "友情提示"
                    Close #1

                    Exit Sub

                End If

            End If
        
        End If

    Loop

    Close #1
       
    Open FName For Input As #1
        
    Do Until EOF(1)
        Line Input #1, Nextline
              
        If UCase(Left(Trim(Nextline), 6)) <> "DEVICE" Then
       
            bid = Split(Nextline, ",")
             
            id = 0
            qtyTemp = 0
        
            '付值
            semtechFabTemp.QTECH_CREATED_BY = gUserName
            
            semtechFabTemp.id = Get37FabMaxID()
            
            semtechFabTemp.DeviceName = Trim(bid(0))
            
            '根据客户机种，查die数
            
            '判断产品对照表上，有没有这个机种，没有则退出
            
            If (Not Judge37FabDieFlag(semtechFabTemp.DeviceName)) Then

                MsgBox "这机种：" & semtechFabTemp.DeviceName & " 在系统设定表中不存在，请联系市场部与研发部!", vbInformation, "友情提示"

                Exit Sub

            Else

                qtyTemp = Get37DieQty(semtechFabTemp.DeviceName)

            End If
            
            semtechFabTemp.Batch = Trim(bid(1))
            semtechFabTemp.WF = CInt(Trim(bid(2)))
            semtechFabTemp.Price = CDbl(Trim(bid(3)))
            semtechFabTemp.CURRENCY = Trim(bid(4))
            semtechFabTemp.ShippedDt = CDate(Trim(bid(5)))
            semtechFabTemp.PurchaseNo = Trim(bid(6))
            semtechFabTemp.PurchaseOrderLineItem = Trim(bid(7))
            semtechFabTemp.Invoice = Trim(bid(8))
            semtechFabTemp.MAWBNumber = Trim(bid(9))
            semtechFabTemp.Destination = Trim(bid(10))
            semtechFabTemp.wafer_id = Trim(bid(11))
            
            waferIDList = semtechFabTemp.wafer_id
        
            Dim waferStrTemp As String
                     
            If InStr(waferIDList, ".") > 0 Then
                bidWaferID = Split(waferIDList, ".")
            Else
                bidWaferID(0) = waferIDList

            End If

            For k = 0 To UBound(bidWaferID)
                '根据po号，batch号，判断是否已经上传过

                If Judge37FabData(semtechFabTemp.Batch, bidWaferID(k)) Then

                    MsgBox "这笔：" & semtechFabTemp.Batch & "  " & bidWaferID(k) & " 已存在，无需上传!", vbInformation, "友情提示"

                    Exit Sub

                End If

            Next

            Call Add37FabData(semtechFabTemp)
                    
            '按waferid 所数据放到Mapping表里
                     
            For N = 0 To UBound(bidWaferID)
                waferidNoTemp = bidWaferID(N)

                If Len(waferidNoTemp) <> 2 Then
                    waferStrTemp = semtechFabTemp.Batch & "0" & waferidNoTemp
                Else
                    waferStrTemp = semtechFabTemp.Batch & waferidNoTemp

                End If
                          
                Call Add37FabDetail(semtechFabTemp, waferStrTemp, waferidNoTemp, qtyTemp)
                          
            Next
                    
            SumCount = SumCount + 1
         
        End If
        
    Loop

    Close #1
        
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！", vbInformation, "友情提示"

    End If

End Sub

Private Sub UploadMPSFab()

    'Semteck 晶圆数据上传
    '读取CSV
    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim cusPTTemp            As String

    Dim gcVerTemp            As String

    Dim gcVerLastTemp        As String

    Dim waferIDList          As String

    Dim wafid()              As String

    Dim i                    As Integer

    Dim j                    As Integer

    Dim id                   As Long

    Dim TEMP                 As String

    Dim SumCount             As Integer

    Dim GCHeaderFlag         As Boolean

    Dim str01                As String

    Dim str03                As String

    Dim FabTem               As String

    Dim qtyTemp              As Long

    Dim k                    As Integer

    Dim FName                As String

    Dim Nextline             As String

    Dim bidWaferID()         As String

    Dim N                    As Integer

    Dim waferidNoTemp        As String

    Dim waferStrTemp         As String

    Dim bidWaferID1          As String

    Dim Message              As String

    Dim comparewaferid       As String

    Dim bid

    '上传OI的CSV
    '处理文件名
    If Text7.text = "" Then
        MsgBox "先选择待上传的文件!", vbInformation, "友情提示"

        Exit Sub

    End If
         
    SumCount = 0

    FName = Trim(Text7.text)

    ' First open file for check
    Open FName For Input As #1

    Do Until EOF(1)
        Line Input #1, Nextline
    
        ' 从第二行开始检查
        If UCase(Left(Trim(Nextline), 6)) <> "DEVICE" And Len(Nextline) > 15 Then
            bid = Split(Nextline, ",")
        
            semtechFabTemp.DeviceName = Trim(bid(0))
            semtechFabTemp.Batch = Trim(bid(1))
            semtechFabTemp.WF = CInt(Trim(bid(2)))
            semtechFabTemp.Price = CDbl(Trim(bid(3)))
            semtechFabTemp.CURRENCY = Trim(bid(4))
            semtechFabTemp.ShippedDt = CDate(Trim(bid(5)))
            semtechFabTemp.PurchaseNo = Trim(bid(6))
            semtechFabTemp.PurchaseOrderLineItem = Trim(bid(7))
            semtechFabTemp.Invoice = Trim(bid(8))
            semtechFabTemp.MAWBNumber = Trim(bid(9))
            semtechFabTemp.Destination = Trim(bid(10))
            semtechFabTemp.wafer_id = Trim(bid(11))
            semtechFabTemp.PoBatch = Trim(bid(12))
    
            waferIDList = semtechFabTemp.wafer_id

            ' Guard 1
            If (Not JudgeMPSBankPT(semtechFabTemp.DeviceName)) Then
            
                MsgBox "这机种：" & semtechFabTemp.DeviceName & " 在系统设定表中不存在，请联系市场部与研发部!", vbInformation, "友情提示"
                Close #1

                Exit Sub

            End If
    
            ' Guard 2
            If InStr(semtechFabTemp.wafer_id, ".") > 0 Then
                bidWaferID = Split(waferIDList, ".")
            
                If UBound(bidWaferID) + 1 <> Val(semtechFabTemp.WF) Then
                    MsgBox "Batch号为" & (semtechFabTemp.Batch) & "的条目, 片数与WaferID个数不一致, 请重新调整本次未上传任何batch", vbInformation, "友情提示"
                    Close #1

                    Exit Sub

                End If
            
                For k = 0 To UBound(bidWaferID)
                     
                    '根据po号，batch号，判断是否已经上传过
                    If (Judge37FabData(semtechFabTemp.Batch, bidWaferID(k))) Then
                        MsgBox "这笔：" & bidWaferID(k) & " 已存在，无需上传!", vbInformation, "友情提示"
                        Close #1

                        Exit Sub

                    End If

                Next
            
                For N = 0 To UBound(bidWaferID)
                    waferidNoTemp = bidWaferID(N)
                
                    If waferidNoTemp = "" Then
                        MsgBox "WaferId存在空值"
                    
                        Close #1

                        Exit Sub
            
                    ElseIf Val(waferidNoTemp) > 25 Or Val(waferidNoTemp) < 1 Then
                        MsgBox "WaferId超出1-25范围"
                        Close #1

                        Exit Sub

                    End If
                
                    For KK = N + 1 To UBound(bidWaferID)
                        comparewaferid = bidWaferID(KK)
                    
                        If comparewaferid = waferidNoTemp Then
                            MsgBox "WaferId有重复"
                            Close #1

                            Exit Sub

                        End If

                    Next
                
                Next
            
            Else
                bidWaferID1 = waferIDList

                If (Judge37FabData(semtechFabTemp.Batch, bidWaferID1)) Then
                    MsgBox "这笔：" & bidWaferID1 & " 已存在，无需上传!", vbInformation, "友情提示"
                    Close #1

                    Exit Sub

                End If
            
                If Val(semtechFabTemp.WF) <> 1 Then
                    MsgBox "Batch号为" & (semtechFabTemp.Batch) & "的条目, 片数与WaferID个数不一致", vbInformation, "友情提示"
                    Close #1

                    Exit Sub

                End If
                       
                If bidWaferID1 = "" Then
                    MsgBox "Batch号为" & (semtechFabTemp.Batch) & "的条目, 片数与WaferID个数不一致", vbInformation, "友情提示"
                    Close #1

                    Exit Sub

                ElseIf Val(bidWaferID1) > 25 Or Val(bidWaferID1) < 1 Then
                    MsgBox "Batch号为" & (semtechFabTemp.Batch) & "的条目, 片数与WaferID个数不一致", vbInformation, "友情提示"
                    Close #1

                    Exit Sub

                End If

            End If
        
        End If

    Loop

    Close #1

    ' Second open for upload
    Open FName For Input As #1
        
    Do Until EOF(1)
        Line Input #1, Nextline
             
        If UCase(Left(Trim(Nextline), 6)) <> "DEVICE" And Len(Nextline) > 15 Then
       
            bid = Split(Nextline, ",")
    
            id = 0
            qtyTemp = 0
        
            '付值
            semtechFabTemp.QTECH_CREATED_BY = gUserName
            semtechFabTemp.id = Get37FabMaxID()
            semtechFabTemp.DeviceName = Trim(bid(0))
            semtechFabTemp.Batch = Trim(bid(1))
            semtechFabTemp.WF = CInt(Trim(bid(2)))
            semtechFabTemp.Price = CDbl(Trim(bid(3)))
            semtechFabTemp.CURRENCY = Trim(bid(4))
            semtechFabTemp.ShippedDt = CDate(Trim(bid(5)))
            semtechFabTemp.PurchaseNo = Trim(bid(6))
            semtechFabTemp.PurchaseOrderLineItem = Trim(bid(7))
            semtechFabTemp.Invoice = Trim(bid(8))
            semtechFabTemp.MAWBNumber = Trim(bid(9))
            semtechFabTemp.Destination = Trim(bid(10))
            semtechFabTemp.wafer_id = Trim(bid(11))
            semtechFabTemp.PoBatch = Trim(bid(12))
            waferIDList = semtechFabTemp.wafer_id
            
            qtyTemp = Get68DieQty(semtechFabTemp.DeviceName)
      
            Call AddMPSFabData(semtechFabTemp)
        
            Dim sOra As String
        
            sOra = "select mes_dn_pkg.MES_BANK_MPS('" & semtechFabTemp.Batch & "','" & semtechFabTemp.wafer_id & "') from dual"
            Get_OracleRs (sOra)
        
            If InStr(semtechFabTemp.wafer_id, ".") > 0 Then
                bidWaferID = Split(waferIDList, ".")

                For N = 0 To UBound(bidWaferID)
                    waferidNoTemp = bidWaferID(N)

                    If Len(waferidNoTemp) <> 2 Then
                        waferStrTemp = semtechFabTemp.Batch & "0" & waferidNoTemp
                    Else
                        waferStrTemp = semtechFabTemp.Batch & waferidNoTemp

                    End If
                
                    Call AddMPSFabDetail(semtechFabTemp, waferStrTemp, waferidNoTemp, qtyTemp, Trim(CmbCustomer37.text))
                                 
                Next

            Else
                waferidNoTemp = waferIDList
            
                If Len(waferidNoTemp) <> 2 Then
                    waferStrTemp = semtechFabTemp.Batch & "0" & waferidNoTemp

                End If

                waferStrTemp = semtechFabTemp.Batch & waferidNoTemp
                          
                Call AddMPSFabDetail(semtechFabTemp, waferStrTemp, waferidNoTemp, qtyTemp, Trim(CmbCustomer37.text))

            End If
        
            SumCount = SumCount + 1

        End If
        
    Loop

    Close #1
        
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！", vbInformation, "友情提示"

    End If

End Sub

Private Sub Command26_Click()

    'GC
    On Error Resume Next

    Dim FName

    '帅选文件
    CommonDialog6.Filter = "CSV文件(*.csv)|*.csv|EXCEL文件(*.xlsx)|*.xlsx|EXCEL文件(*.xls)|*.xls"
    
    CommonDialog6.ShowOpen
    '得到文件名
    FName = CommonDialog6.filename

    If FName <> "" Then
        Text7.text = FName

    End If

End Sub

Private Sub Command27_Click()
    MsgBox "此处无法上传,请勿在此上传", vbInformation, "提示"
    Exit Sub

    '上传PO

    If CmbCustomer37.text = "" Then
        MsgBox "请先选择客户代码", vbInformation, "提示"

        Exit Sub
    
    End If

    If CmbCustomer37.text = "68" Or CmbCustomer37.text = "70" Then

        UploadMPS

        Exit Sub

    End If

    SumCount = 0
    ErrorInf = ""

    If Text8.text = "" Then
        MsgBox "请先选择待上传的文件", vbInformation, "提示"

        Exit Sub
    
    End If
    
    Dim filename As String

    filename = Text8.text

    Dim dirtemp()  As String

    Dim i          As Integer
    
    Dim poTypeTemp As String

    poTypeTemp = CmbPoType.text
    
    If poTypeTemp = "" Then
        MsgBox "请先选择PO模板类型!", vbInformation, "提示"

        Exit Sub

    End If
    
    If InStr(1, filename, ",") > 0 Then
        dirtemp = Split(filename, ",")
        
        For i = 1 To UBound(dirtemp)

            If poTypeTemp = "ICI" Then
                Up37PODataICI (dirtemp(0) + "\" + dirtemp(i))
            Else
                
                Up37POData (dirtemp(0) + "\" + dirtemp(i))
                
            End If

        Next
        
    Else

        If poTypeTemp = "ICI" Then
            Up37PODataICI (filename)
        Else
            Up37POData (filename)
        
        End If
        
    End If
    
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！", vbInformation, "提示"

    End If

End Sub

Private Sub UploadMPS()

    '上传PO

    SumCount = 0
    ErrorInf = ""

    If Text8.text = "" Then
        MsgBox "先选择待上传的文件"

        Exit Sub
    
    End If
    
    Dim filename As String

    filename = Text8.text

    Dim dirtemp() As String

    Dim i         As Integer
    
    If InStr(1, filename, ",") > 0 Then
        dirtemp = Split(filename, ",")
        
        For i = 1 To UBound(dirtemp)

            Up68POData (dirtemp(0) + "\" + dirtemp(i))

        Next
        
    Else
      
        Up68POData (filename)
        
    End If
    
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！", vbInformation, "提示"

    End If

End Sub

Private Sub Command28_Click()

    'GC
    On Error Resume Next

    Dim FName

    '帅选文件
    'CommonDialog7.Filter = "EXCEL文件(*.xls)|*.xls"
    CommonDialog7.Filter = "CSV文件(*.csv)|*.csv||EXCEL文件(*.xls)|*.xls "
    CommonDialog7.ShowOpen
    '得到文件名
    FName = CommonDialog7.filename

    If FName <> "" Then
        Text8.text = Replace(FName, Chr(0), ",")

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
        StrFileName = Mid(Trim(Text2.text), InStrRev(Trim(Text2.text), "\") + 1)
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

Private Sub Command32_Click()

    Dim beginTime   As String

    Dim endTime     As String

    Dim woTemp      As String

    Dim productTemp As String

    Dim sqlTemp     As String

    Dim cusPTTemp   As String

    beginTime = Format(DTPicker4.Value, "YYYY/MM/DD")
    endTime = Format(DTPicker5.Value, "YYYY/MM/DD")

    sqlTemp = "  select devicename as ""Wafer Type"",Assydevicename as ""Assy Part#"",batch as ""Fab Lot"",datecode as ""D/C"",wf as Qty, " & " sn as ""ICI Batch"" ,Bagno as ""Bag#"",'' as ""Comment"" ,Get_37_ICIHTPT(Assydevicename) as ""HT Part#"" " & "  from  MAPPINGDATA37 d  where Maptype='ICI' and d.Qtech_Created_Date >=to_date('" + beginTime + "','YYYY/MM/DD') and  d.Qtech_Created_Date <to_date('" + endTime + "','YYYY/MM/DD') " & " order by bagno "
 
    ExporToExcel (sqlTemp)

End Sub

Private Sub Command33_Click()
   MsgBox "此处无法上传,请勿在此上传", vbInformation, "提示"
    Exit Sub
    SumCount = 0
    ErrorInf = ""

    If Text9.text = "" Then
        MsgBox "先选择待上传的文件"

        Exit Sub
    
    End If
    
    Dim filename As String

    filename = Text9.text

    Dim dirtemp() As String

    Dim i         As Integer
    
    If InStr(1, filename, ",") > 0 Then
        dirtemp = Split(filename, ",")
        
        For i = 1 To UBound(dirtemp)
            Up37ExcelPO (dirtemp(0) + "\" + dirtemp(i))
         
            Sleep (1000)
        Next
        
    Else
        
        Up37ExcelPO (filename)

    End If
    
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

    End If
    
    If ErrorInf <> "" Then
        MsgBox "上传失败的有:" + ErrorInf + "数据库中已存在！"

    End If

End Sub

Private Sub Up37ExcelPO(dirtemp As String)

    Dim TPriceFlag           As Boolean

    Dim source_batch_id_Temp As String

    If Text9.text = "" Then
        MsgBox "先选择待上传的文件"

        Exit Sub

    End If

    Dim dirName  As String

    Dim filename As String

    TPriceFlag = False

    If InStrRev(Trim(dirtemp), "\") > 0 Then
        StrFileName = Mid(Trim(dirtemp), InStrRev(Trim(dirtemp), "\") + 1)
        dirName = Mid$(Trim(dirtemp), 1, InStrRev(Trim(dirtemp), "\"))

    End If

    Dim con As New ADODB.Connection

    Dim rs  As New ADODB.Recordset

    Set VBExcel = CreateObject("excel.application")

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(dirtemp)

    Set xlSheet = xlBook.Worksheets(1)

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 46 Then
   
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

    Dim waferStrTemp  As String

    Dim waferStrTemp1 As String

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count

        semPotemp.po1lot = ""
        TEMP = ""
        source_batch_id_Temp = ""
    
        For j = 1 To 46
      
            If j > 26 Then
                strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
            Else
                strChar = Chr(96 + j)

            End If
            
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
        
            If j = 1 Then
                semPotemp.DATE = tempVal

            End If
          
            If j = 4 Then
                semPotemp.MfgPlant = tempVal

            End If
          
            If j = 5 Then
                semPotemp.MfgPlant = semPotemp.MfgPlant & "-" & tempVal

            End If
  
            If j = 8 Then
                semPotemp.TypeService = tempVal

            End If
          
            If j = 9 Then
                semPotemp.PurchaseOrderNo = Trim(tempVal)

            End If
          
            If j = 10 Then
                semPotemp.ITEM = Trim(tempVal)

            End If
          
            If j = 11 Then
                semPotemp.MaterialDes = Trim(tempVal)

            End If
          
            If j = 12 Then
                semPotemp.YourMaterialNumber = Trim(tempVal)

            End If
          
            If j = 14 Then
                semPotemp.Quantity = Replace(tempVal, " ", "")

            End If
          
            If j = 16 Then
                semPotemp.LotNO = Trim(tempVal) ' 37二次的JOBID

            End If
          
            If j = 17 Then
                semPotemp.DelDate = Trim(tempVal)

            End If
          
            If j = 18 Then
                semPotemp.UnitPrice = Trim(tempVal)

            End If
          
            If j = 19 Then
                semPotemp.POPrice = "0.088"

            End If
          
            If j = 20 Then
                semPotemp.CURRENCY = Trim(tempVal)

            End If
          
            If j = 21 Then
                semPotemp.NetAmount = Trim(tempVal)

            End If
          
            If j = 23 Then
                semPotemp.TermsPayment = Trim(tempVal)

            End If
            
            If j = 27 Then
                semPotemp.ItemLineText = Trim(tempVal)

            End If
          
            If j = 34 Then
                semPotemp.Plant = Trim(tempVal)

            End If
             
            If j = 35 Then
                semPotemp.PartNumber = Trim(tempVal)

            End If
            
            If j = 36 Then
                semPotemp.Quantity = CLng(Trim(tempVal))
                  
            End If
                
            If j = 38 Then
                semPotemp.LotNumber = Trim(tempVal)
                  
            End If
                
            If j = 40 Then
                If Len(tempVal) < 2 Then
                    semPotemp.waferid = "0" & Trim(tempVal)
                Else
                    semPotemp.waferid = Trim(tempVal)
                  
                End If
                  
            End If
                
            If j = 41 Then
                If Len(tempVal) = 0 Then    ' 37的一次JOBID
                    semPotemp.WAFERLOT = semPotemp.LotNumber
                Else
                    semPotemp.WAFERLOT = Trim(tempVal)

                End If

            End If
                
            If j = 45 Then
                semPotemp.po1lot2 = Trim(tempVal)   ' HT的LOTID
                   
            End If
               
            If j = 46 Then
                semPotemp.ProductionOrderNo = Trim(tempVal)
                   
            End If

        Next j
    
        semPotemp.id = GetMaxID()

        semPotemp.QTECH_CREATED_BY = gUserName

        If Len(semPotemp.po1lot2) > 0 Then
            semPotemp.KeyStr = semPotemp.PurchaseOrderNo & "_" & semPotemp.LotNO & "_" & semPotemp.WAFERLOT & semPotemp.waferid & "+"
            semPotemp.po1lot = semPotemp.po1lot2
        Else

            semPotemp.KeyStr = semPotemp.PurchaseOrderNo & "_" & semPotemp.LotNO & "_" & semPotemp.WAFERLOT & semPotemp.waferid & "+"
            semPotemp.po1lot = GetLot(semPotemp.LotNumber)
        End If

        If (JudgeFlag37POHeader(semPotemp.KeyStr)) Then
            MsgBox "这笔：" & semPotemp.KeyStr & " 已存在，无需再次上传!"
            SumCount = SumCount - 1
            Exit Sub

        End If

        If Len(semPotemp.PurchaseOrderNo) < 0 Then

            MsgBox "PO为空不允许上传!"

            Exit Sub

        End If

        If Len(semPotemp.po1lot) > 0 Then

            waferStrTemp1 = Getcustpart(semPotemp.po1lot & semPotemp.waferid)
            Call Add37POHeaderICI(semPotemp, waferStrTemp1)

            
            If Left$(semPotemp.waferid, 1) = "0" Then
                waferStrTemp = Get_OracleStr("select max(substrateid) || '+' as substrateid from mappingdatatest where wafer_id in ('" & Right(semPotemp.waferid, 1) & "', '0' || '" & Right(semPotemp.waferid, 1) & "') and lotid = '" & semPotemp.po1lot & "'")
            Else
                 waferStrTemp = Get_OracleStr("select max(substrateid) || '+' as substrateid from mappingdatatest where wafer_id = '" & semPotemp.waferid & "' and lotid = '" & semPotemp.po1lot & "'")
            End If
            
            Call Add37POwaferDetail1(semPotemp.po1lot, CStr(waferStrTemp), semPotemp.waferid, semPotemp.PurchaseOrderNo, CStr(semPotemp.id), CStr(semPotemp.Quantity), waferStrTemp)
            Call Add37POwaferDetail(semPotemp.po1lot, CStr(waferStrTemp), semPotemp.waferid, semPotemp.PurchaseOrderNo)
        Else
  
            MsgBox "无LOT号"

            Exit Sub

        End If

    Next i
 
    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

End Sub

Private Sub Command34_Click()

    ExporToExcel (" select ID,PO_NUM as PurchaseOrderNo, MPN as ProductionOrderNo,  CREATED_DATE as PODate,  SHIPPING_MST_260 as Currency,  " & "  SHIP_SITE as ShippingAddress, COUNTRY_OF_ASSEMBLY as Termsofpayment,  PO_ITEM as Item, MPN_DESC as MaterialDescription,SOURCE_BATCH_ID as LotNo, SOURCE_MTRL_SLOC as WaferLot, " & "  DIE_QTY as Quantity, mtrl_num as BagNo,DATE_CODE as  DateCode, t_price as UnitPrice,  CURRENT_WAFER_QTY as NetAmount, " & "  SOURCE_MTRL_NUM as PartNumber, COUNTRY_OF_FAB as WaferFAB, IMAGER_CUSTOMER_REV as WaferREV " & "   from customeroitbl_test a  where  customershortname='37' and a.qtech_created_date>to_date('2016-03-26','YYYY-MM-DD') and a.flag='Y' order by id desc ")

End Sub

Private Sub Command35_Click()

Dim strSubstrateid As String
If txtSubstrateID.text = "" Then
    MsgBox "请输入需要删除的SubstrateID", vbInformation, "提示"
    Exit Sub
End If

Dim strsql As String
Dim strSql2 As String
strSubstrateid = Trim$(UCase(txtSubstrateID.text))

strsql = "insert into mappingdatatest_bak select * from mappingdatatest where substrateid = '" & strSubstrateid & "'"
If AddSql(strsql) = 0 Then
    MsgBox "没有该Substrateid:" & vbCrLf & strSubstrateid, vbInformation, "提示"
    Exit Sub
End If

MsgBox "备份成功", vbInformation, "提示"

strsql = "delete from mappingdatatest where substrateid = '" & strSubstrateid & "'"
strSql2 = "delete from erpbase..tblmappingData where substrateid = '" & strSubstrateid & "'"
If AddSql(strsql) > 0 And AddSql2(strSql2) > 0 Then
    MsgBox "删除成功", vbInformation, "提示"
    Exit Sub
End If

End Sub

Private Sub Command4_Click()

    Dim beginTime As String

    Dim endTime   As String

    beginTime = Format(DTPStart(2).Value, "YYYY/MM/DD")
    endTime = Format(DTPStop(2).Value, "YYYY/MM/DD")

    ExporToExcel ("select SUBSTRATEID, SUBSTRATETYPE, LOTID, PRODUCTID, CREATEDATE,MICRONLOTID, PASSBINCOUNT, FAILBINCOUNT, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE ,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE from mappingDataTest where customershortname ='AA' and qtech_created_date >= to_date('" + beginTime + "','YYYY/MM/DD') and qtech_created_date <=to_date('" + endTime + "','YYYY/MM/DD')  order by qtech_created_date desc , lotid, substrateid")

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
        StrFileName = Mid(Trim(Text3.text), InStrRev(Trim(Text3.text), "\") + 1)
        dirName = Mid$(Trim(Text3.text), 1, InStrRev(Trim(Text3.text), "\"))

    End If

    Dim con As New ADODB.Connection

    Dim rs  As New ADODB.Recordset

    con.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
    rs.Open "Select * From " & "[" & StrFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
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
            
        gcHeaderTemp.Lot_id = rs.Fields(9).Value
        gcDetailTemp.Lot_id = rs.Fields(9).Value
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
        
        If UCase(Left(Trim(Nextline), 4)) <> "ITEM" Then
             
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

    '读取CSV
    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim cusPTTemp            As String

    Dim gcVerTemp            As String

    Dim gcVerLastTemp        As String

    customerTemp = "GC"

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"

        Exit Sub

    End If

    'Dim dirName As String
    'Dim FileName As String

    ''获取文件名
    '    If InStrRev(Trim(Text3.Text), "\") > 0 Then
    '        strFileName = Mid(Trim(Text3.Text), InStrRev(Trim(Text3.Text), "\") + 1)
    '        dirName = Mid$(Trim(Text3.Text), 1, InStrRev(Trim(Text3.Text), "\"))
    '    End If

    '判断 GC类型，是不是
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
        
    SumCount = 0
        
    GCHeaderFlag = False

    Dim k        As Integer
        
    Dim FName    As String

    Dim Nextline As String

    FName = Trim(Text3.text)
    Open FName For Input As #1
        
    Do Until EOF(1)
        Line Input #1, Nextline
        cusPTTemp = ""
        gcVerTemp = ""
        gcVerLastTemp = ""
              
        If UCase(Left(Trim(Nextline), 4)) <> "ITEM" Then

            Dim bid

            bid = Split(Nextline, ",")
             
            id = 0
        
            '付值
            gcHeaderTemp.Created_By = gUserName
            gcDetailTemp.ITEM = Trim(bid(0))
            gcHeaderTemp.po_no = Trim(bid(1))
            gcHeaderTemp.SUPPLIER = Trim(bid(2))
            gcHeaderTemp.ShipTo = Trim(bid(3))
            gcHeaderTemp.Fab_Device = Trim(bid(4))
            
            gcHeaderTemp.Customer_Device = Trim(bid(5))
            cusPTTemp = Trim(gcHeaderTemp.Customer_Device)
            gcHeaderTemp.GC_Version = Trim(bid(6))
            gcVerTemp = Trim(UCase(gcHeaderTemp.GC_Version))
            
            '2015-04-27 jiayun add 第三位系统自动带
            
            '2015-11-17 jiayun add  GC2145
            
            If cusPTTemp = "GC2145-3" Then
                If Left(bid(9), 1) = "H" Then
                    gcVerLastTemp = "G"
               
                ElseIf Left(bid(9), 1) = "E" Then
                    gcVerLastTemp = "F"

                End If
            
            Else
        
                gcVerLastTemp = GetGCVerLastChar(cusPTTemp)

            End If
            
            If gcVerLastTemp <> "" Then
                gcHeaderTemp.GC_Version = gcVerTemp & gcVerLastTemp
                 
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
            
            gcHeaderTemp.Lot_id = Trim(bid(9))
            gcDetailTemp.Lot_id = Trim(bid(9))
            gcDetailTemp.wafer_id = Trim(bid(10))
            gcDetailTemp.Good_Die_Qty = CInt(Trim(bid(11)))
            gcHeaderTemp.WO_NO = Trim(bid(12))
            gcHeaderTemp.Ship_Out = Trim(bid(13))
            gcHeaderTemp.TradeType = Trim(bid(15))
            
            '2015-02-03 jiayunadd check shipOut
            '如果是COG的，则不可以为空
            
            '2016-01-28 jiayun modify Cog 根据客户机种来看

            '            If Left(gcHeaderTemp.Lot_ID, 3) = "GXS" Then
            '                If gcHeaderTemp.Ship_Out = "" Then
            '                  MsgBox "GC COG，最后一列发货地址不可以有空！"
            '                  Exit Sub
            '
            '                End If
            '
            '
            '            End If
            
            If Left(cusPTTemp, InStr(1, cusPTTemp, "-") - 1) = "GC9102" Then
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
            
            '2015-08-20 jiayun add 处理 GC0409-3
            
            If Trim(gcHeaderTemp.Customer_Device) = "GC0409-3" Then
            
                FabTem = Left(UCase(Trim(gcHeaderTemp.Fab_Device)), 5)
              
                If FabTem = "P6418" Then
                    gcDetailTemp.Good_Die_Qty = 5192
                     
                ElseIf FabTem = "P6820" Then
                 
                    gcDetailTemp.Good_Die_Qty = 11994
                     
                ElseIf FabTem = "P7238" Then
                 
                    gcDetailTemp.Good_Die_Qty = 5192 '5211

                End If
              
            ElseIf Trim(gcHeaderTemp.Customer_Device) = "GC2145-3" Then
            
                If Left(gcHeaderTemp.Lot_id, 1) = "H" Then
                    gcDetailTemp.Good_Die_Qty = 1676
               
                ElseIf Left(gcHeaderTemp.Lot_id, 1) = "E" Then
                    gcDetailTemp.Good_Die_Qty = 3920

                End If
                 
            End If
            
            '2013-12-27 jiayun add
            
            If gcDetailTemp.Good_Die_Qty <= 0 Then
                MsgBox "请确认客户机种对应的Die数是否有维护好！"

                Exit Sub

            End If
            
            '            '2015-10-29 jiayun add 客户机种后面加 C
            '
            '            If cusPTTemp = "GC030A-3" Then
            '                  gcHeaderTemp.Customer_Device = "GC030AC-3"
            '
            '            ElseIf cusPTTemp = "GC0406-3" Then
            '                  gcHeaderTemp.Customer_Device = "GC0406C-3"
            '
            '            ElseIf cusPTTemp = "GC2365-3" Then
            '                  gcHeaderTemp.Customer_Device = "GC2365C-3"
            '
            '            ElseIf cusPTTemp = "GC5005-3" Then
            '                  gcHeaderTemp.Customer_Device = "GC5005C-3"
            '
            '            ElseIf cusPTTemp = "GC8024-3" Then
            '                  gcHeaderTemp.Customer_Device = "GC8024C-3"
            '            ElseIf cusPTTemp = "GC6133-3" Then
            '                  gcHeaderTemp.Customer_Device = "GC6133C-3"
            '
            '            ElseIf cusPTTemp = "GC2003-3" Then
            '                  gcHeaderTemp.Customer_Device = "GC2003C-3"
            '
            '            ElseIf cusPTTemp = "GC1066-3" Then
            '                  gcHeaderTemp.Customer_Device = "GC1066C-3"
            '
            '            ElseIf cusPTTemp = "GC1064-3" Then
            '                  gcHeaderTemp.Customer_Device = "GC1064C-3"
            '
            '            ElseIf cusPTTemp = "GC1024-3" Then
            '                  gcHeaderTemp.Customer_Device = "GC1024C-3"
            '
            '            ElseIf cusPTTemp = "GC2375-3" Then
            '                  gcHeaderTemp.Customer_Device = "GC2375C-3"
            '
            '            ElseIf cusPTTemp = "GC2023-3" Then
            '                  gcHeaderTemp.Customer_Device = "GC2023C-3"
            '
            '            ElseIf cusPTTemp = "GC032A-3" Then
            '                  gcHeaderTemp.Customer_Device = "GC032AC-3"
            '
            '            End If
            
            '2016-03-07 jiayun modify add PT-C
            
            Set oiRS = GetGCPT_C(cusPTTemp)

            If (oiRS.RecordCount > 0) Then
                gcHeaderTemp.Customer_Device = oiRS.Fields("CUSTOMERPTNew").Value

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
 
        End If
        
    Loop

    Close #1
        
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"

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

    'Dim dirName As String
    'Dim FileName As String

    ''获取文件名
    '    If InStrRev(Trim(Text3.Text), "\") > 0 Then
    '        strFileName = Mid(Trim(Text3.Text), InStrRev(Trim(Text3.Text), "\") + 1)
    '        dirName = Mid$(Trim(Text3.Text), 1, InStrRev(Trim(Text3.Text), "\"))
    '    End If

    '判断 GC类型，是不是
    'If GetGCWLT(Trim(Text3.Text)) = "T" Then
    'UploadGCWLTNew
    '
    'Exit Sub
    'End If

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
    Open FName For Input As #1
        
    Do Until EOF(1)
        Line Input #1, Nextline
        cusPTTemp = ""
        gcVerTemp = ""
        gcVerLastTemp = ""
        waferIdTemp = ""
              
        If UCase(Left(Trim(Nextline), 2)) <> "NO" Then

            Dim bid

            bid = Split(Nextline, ",")
             
            id = 0
        
            '付值
            gcHeaderTemp.Created_By = gUserName
            gcDetailTemp.ITEM = Trim(bid(0))
            gcHeaderTemp.po_no = Trim(bid(6))
            gcHeaderTemp.SUPPLIER = Trim(bid(1))
            gcHeaderTemp.ShipTo = Trim(bid(2))
            gcHeaderTemp.Fab_Device = Trim(bid(3))
            
            gcHeaderTemp.Customer_Device = Trim(bid(4)) & "-3"
            cusPTTemp = Trim(gcHeaderTemp.Customer_Device)
            gcHeaderTemp.GC_Version = Trim(bid(5))
            gcVerTemp = Trim(UCase(gcHeaderTemp.GC_Version))
            
            '2015-04-27 jiayun add 第三位系统自动带
            '            gcVerLastTemp = GetGCVerLastChar(cusPTTemp)
            '
            '            If gcVerLastTemp <> "" Then
            '                 gcHeaderTemp.GC_Version = gcVerTemp & gcVerLastTemp
            '
            '            Else
            '
            '                If cusPTTemp = "GC1004-3" Then
            '
            '                      If Mid(gcVerTemp, 1, 1) = "A" Or Mid(gcVerTemp, 1, 1) = "B" Or Mid(gcVerTemp, 1, 1) = "C" Or Mid(gcVerTemp, 1, 1) = "D" Then
            '                       gcHeaderTemp.GC_Version = gcVerTemp & "A"
            '                      Else
            '                       gcHeaderTemp.GC_Version = gcVerTemp & "B"
            '                      End If
            '
            '
            '                ElseIf cusPTTemp = "GC0329-3" Then
            '                         If Len(gcVerTemp) = 2 Then
            '                            gcHeaderTemp.GC_Version = gcVerTemp & "D"
            '
            '                         ElseIf Len(gcVerTemp) = 3 Then
            '                             gcHeaderTemp.GC_Version = gcVerTemp
            '
            '                         Else
            '                            MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
            '                            Exit Sub
            '
            '                         End If
            '
            '
            '
            '                Else
            '                    '判断长度是否为3 ，如果是，则按市场部的来上传，否则提提示错误
            '                    If Len(gcVerTemp) = 3 Then
            '                         gcHeaderTemp.GC_Version = gcVerTemp
            '
            '                    Else
            '                            MsgBox "GC WO中，GCVersion列数据不对，请确认Wo!"
            '                            Exit Sub
            '
            '                    End If
            '
            '
            '
            '
            '                End If
            '
            '
            '
            '            End If
            
            waferIdTemp = Trim(bid(10)) & Right("0" & Trim(bid(11)), 2)
            
            gcDetailTemp.Marking_Lot_ID = GetGCWLDMaringCode(waferIdTemp)
   
            str01 = Trim(bid(9))
            
            If InStr(str01, "月") > 0 Then
            
                str03 = Replace(str01, "月", "-")
                str03 = Replace(str03, "日", "")
                str03 = Year(DATE) & "-" & str03
                gcHeaderTemp.GC_Date = str03
            
            Else
            
                gcHeaderTemp.GC_Date = bid(8)
            
            End If
            
            gcHeaderTemp.Lot_id = Trim(bid(10))
            gcDetailTemp.Lot_id = Trim(bid(10))
            gcDetailTemp.wafer_id = Trim(bid(11))
            gcDetailTemp.Good_Die_Qty = CInt(Trim(bid(12)))
            gcHeaderTemp.WO_NO = Trim(wo_HT_Temp)
            gcHeaderTemp.Ship_Out = Trim(bid(16))
            
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
            
            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & "-" & Right(("0" & gcDetailTemp.wafer_id), 2)
            
            If (JudgeGCDetailIdWLD(gcDetailTemp.Lot_id, gcDetailTemp.ITEM)) Then
                MsgBox "GC 这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.ITEM & "已存在，无需上传!"
               
            Else
                '上传到Detail表中
            
                '2012-11-05 jiayun 修改 GCT
                   
                'gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_ID), 2)
                   
                'gcDetailTemp.item = gcDetailTemp.Lot_ID & "-" & Right(("0" & gcDetailTemp.Wafer_ID), 2)
                   
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

    End If

End Sub

Private Sub UploadGCWLTNew()

    '读取CSV
    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim wo_HT_Temp           As String

    wo_HT_Temp = "WONO_" & Replace(Replace(Replace(Format(Now, "YYYY-MM-DD HH:MM:SS"), "-", ""), ":", ""), " ", "")

    customerTemp = "GC"

    '上传OI的CSV
    '处理文件名
    'If Text3.Text = "" Then
    '    MsgBox "先选择待上传的文件"
    '    Exit Sub
    'End If
    'Dim dirName As String
    'Dim FileName As String

    ''获取文件名
    '    If InStrRev(Trim(Text3.Text), "\") > 0 Then
    '        strFileName = Mid(Trim(Text3.Text), InStrRev(Trim(Text3.Text), "\") + 1)
    '        dirName = Mid$(Trim(Text3.Text), 1, InStrRev(Trim(Text3.Text), "\"))
    '    End If

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
        
        If UCase(Left(Trim(Nextline), 4)) <> "ITEM" Then

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
        StrFileName = Mid(Trim(Text3.text), InStrRev(Trim(Text3.text), "\") + 1)
        dirName = Mid$(Trim(Text3.text), 1, InStrRev(Trim(Text3.text), "\"))

    End If

    Dim con As New ADODB.Connection

    Dim rs  As New ADODB.Recordset

    con.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
    rs.Open "Select * From " & "[" & StrFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
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
        gcHeaderTemp.FAB_Device2 = IIf(IsNull(rs.Fields(4).Value), "", rs.Fields(4).Value)
            
        gcHeaderTemp.Fab_Device = IIf(IsNull(rs.Fields(15).Value), "", rs.Fields(15).Value)
            
        gcHeaderTemp.Customer_Device = IIf(IsNull(rs.Fields(5).Value), "", rs.Fields(5).Value)
        gcHeaderTemp.GC_Version = IIf(IsNull(rs.Fields(6).Value), "", rs.Fields(6).Value)
        'gcDetailTemp.Marking_Lot_ID = IIf(IsNull(Rs.fields(7).Value), "", Rs.fields(7).Value)
        gcHeaderTemp.GC_Date = rs.Fields(7).Value
            
        gcHeaderTemp.Lot_id = rs.Fields(8).Value
        gcDetailTemp.Lot_id = rs.Fields(8).Value
        gcDetailTemp.wafer_id = rs.Fields(9).Value
        gcDetailTemp.Good_Die_Qty = CInt(rs.Fields(10).Value)
        gcHeaderTemp.WO_NO = IIf(IsNull(rs.Fields(11).Value), "", rs.Fields(11).Value)
        gcHeaderTemp.remarkTemp = IIf(IsNull(rs.Fields(12).Value), "", rs.Fields(12).Value)
        gcHeaderTemp.DATE_CODE = IIf(IsNull(rs.Fields(13).Value), "", rs.Fields(13).Value)
        gcHeaderTemp.Marking_Lot_ID1 = IIf(IsNull(rs.Fields(14).Value), "", rs.Fields(14).Value)
        gcHeaderTemp.Marking_Lot_ID2 = IIf(IsNull(rs.Fields(15).Value), "", rs.Fields(15).Value)
        gcDetailTemp.Marking_Lot_ID = IIf(IsNull(rs.Fields(14).Value), "", rs.Fields(14).Value) & " " & IIf(IsNull(rs.Fields(15).Value), "", rs.Fields(15).Value)
            
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
            
        rs.MoveNext
        
    Next i
        
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
                eqISHeaderTemp.DEVICE = Trim(tempVal)
                
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
        StrFileName = Mid(Trim(Text3.text), InStrRev(Trim(Text3.text), "\") + 1)
        dirName = Mid$(Trim(Text3.text), 1, InStrRev(Trim(Text3.text), "\"))

    End If

    Dim con As New ADODB.Connection

    Dim rs  As New ADODB.Recordset

    con.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
    rs.Open "Select * From " & "[" & StrFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
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

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 21 Then

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

    Dim codeYY    As String

    Dim codeWW    As String

    Dim Bcode     As String

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

Private Sub UploadNormalCustomerNew(customerNameTemp As String)

    Dim source_batch_id_Temp As String

    Dim customerTemp         As String

    Dim SumCount             As Integer

    Dim tax                  As String

    Dim cusPTTemp            As String

    customerTemp = customerNameTemp

    '上传OI的CSV
    '处理文件名
    If Text3.text = "" Then
        MsgBox "先选择待上传的文件"

        Exit Sub

    End If

    If InStr(Text3.text, "-A") > 0 Then
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

    '2)判定最大列Excel中的和设定列是否相同
    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 21 Then
        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"

        Exit Sub

    End If
        
    ' 定义变量
    Dim i             As Integer

    Dim j             As Integer

    Dim id            As Long

    Dim TEMP          As String

    Dim temp2         As String

    Dim tempVal       As String

    Dim mCodetemp     As String

    Dim yTemp         As String

    Dim mTemp         As String

    Dim codeYY        As String

    Dim codeWW        As String

    Dim Bcode         As String

    Dim charTemp      As Long

    Dim FabTem        As String

    Dim gcVerTemp     As String

    Dim gcVerLastTemp As String

    Dim pjletter      As String

    Dim pjwafer       As String

    Dim waferseq      As String

    Dim monStr        As String

    Dim waferStr      As String

    Dim waferPos

    pjletter = "ABCDEFGHIJKLMNOPQRETUVWXY"
    waferseq = "123456789ABCDEFGHIJKLMNOP"

    SumCount = 0

    ' 遍历表格
    ' 第2行开始,循环更换行号
    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        TEMP = ""
        source_batch_id_Temp = ""
    
        ' 查询一行的值
        ' 第1列开始,循环增加列数
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
                cusPTTemp = gcHeaderTemp.Customer_Device
            
                ' 20180305,npi check
                If Get_OracleCnt("select * from tbltsvnpiproduct where customerptno1 = '" & cusPTTemp & "'") = 0 Then
                    MsgBox "NPI没有维护该客户机种:" & gcHeaderTemp.Fab_Device & ", 请联系NPI维护该机种", vbInformation

                    Exit Sub

                End If
            
            End If
            
            If j = 7 Then
                gcHeaderTemp.GC_Version = Trim(tempVal)

            End If
            
            If j = 8 Then

                '2016-03-30 add  2017-10-30 add JS140
                If customerTemp = "SX" Or customerTemp = "HJ" Or customerTemp = "TJ003" Or customerTemp = "JS140" Or customerTemp = "BJ153" Then
                    gcDetailTemp.Marking_Lot_ID = GetSXCodeID()
                ElseIf customerTemp = "81" And gcHeaderTemp.Customer_Device = "1103A_A" Then
                    codeYY = Year(Now)
                    codeWW = DatePart("ww", Now)
                    Bcode = "HS" & Mid(codeYY, 3, 1) & "A" & Mid(codeYY, 4, 1) & "S" & codeWW
                    gcDetailTemp.Marking_Lot_ID = Bcode
                ElseIf customerTemp = "81" And gcHeaderTemp.Customer_Device = "110F_A" Then
                    Bcode = "EHD" & "\\" & "510"
                    gcDetailTemp.Marking_Lot_ID = Bcode
                
                ElseIf customerTemp = "GT" And gcHeaderTemp.Customer_Device = "SIV121DU" Then
                    gcDetailTemp.Marking_Lot_ID = GetGTCodeID()
                    
                Else
                    gcDetailTemp.Marking_Lot_ID = Replace(Replace(Trim(tempVal), Chr(10), ""), Chr(13), "")
                
                    If InStr(gcDetailTemp.Marking_Lot_ID, "Z3A") > 0 Then
                        gcDetailTemp.Marking_Lot_ID = Left(gcDetailTemp.Marking_Lot_ID, 3) & "\\" & Right$(gcDetailTemp.Marking_Lot_ID, 3)

                    End If

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

                'If IsNumeric(Trim(tempVal)) = False Then
                If Trim(tempVal) = False Then  'css add 20160718
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
                gcHeaderTemp.remarkTemp = Trim(tempVal)

            End If
            
            If j = 16 Then
                gcHeaderTemp.TradeType = Trim(tempVal)
                gcHeaderTemp.Coo = Trim(tempVal)

            End If
            
            If j = 17 Then
                ' AT71, 增加DATECODE
                gcHeaderTemp.Attri01 = Trim(tempVal)
                gcHeaderTemp.Level = Trim(tempVal)

            End If
            
            If j = 18 Then
                ' AT71, 增加打标码part1
                gcHeaderTemp.Attri02 = Trim(tempVal)

            End If
            
            If j = 19 Then
                ' AT71, 增加打标码part2
                gcHeaderTemp.Attri03 = Trim(tempVal)

            End If
            
            If j = 20 Then
                gcHeaderTemp.Attri04 = Trim(tempVal)

            End If
            
            If j = 21 Then
                gcHeaderTemp.Attri05 = Trim(tempVal)

            End If

        Next j
    
        ' 读取一行后数据的处理
        gcHeaderTemp.taxTemp = tax

        If gcHeaderTemp.Customer_Device = "GW1N-LV1CS30C6/I5" Then
            gcDetailTemp.Marking_Lot_ID = gcHeaderTemp.Attri01 & "\\" & gcHeaderTemp.Attri02 & "\\" & gcHeaderTemp.Attri03 & "\\" & gcHeaderTemp.Lot_id
            
        End If
        
        If customerTemp = "69" Then
            If Mid(gcDetailTemp.wafer_id, 1, 1) = 0 Then
                pjwafer = Replace(gcDetailTemp.wafer_id, "0", "")
            Else
                pjwafer = gcDetailTemp.wafer_id

            End If
        
            gcDetailTemp.Marking_Lot_ID = Mid(gcHeaderTemp.Lot_id, 2, 6) & Mid(pjletter, pjwafer, 1)

        End If
    
        If customerTemp = "US026" Or customerTemp = "SG005" Or customerTemp = "TW079" Then
            ' 增加REMARK
            gcDetailTemp.Marking_Lot_ID = Mid$(gcHeaderTemp.Customer_Device, InStr(gcHeaderTemp.Customer_Device, "-") + 2, 1)
            gcDetailTemp.Marking_Lot_ID = gcDetailTemp.Marking_Lot_ID & Right(Year(Now), 1)
            gcDetailTemp.Marking_Lot_ID = gcDetailTemp.Marking_Lot_ID & Hex(Month(Now))
            gcDetailTemp.Marking_Lot_ID = gcDetailTemp.Marking_Lot_ID & Mid$(waferseq, gcDetailTemp.wafer_id, 1)

            Dim spos As String

            spos = InStr(gcHeaderTemp.Lot_id, ".")

            If spos > 0 Then
                gcDetailTemp.Marking_Lot_ID = gcDetailTemp.Marking_Lot_ID & Mid$(gcHeaderTemp.Lot_id, spos - 4, 4)
            Else
            
                gcDetailTemp.Marking_Lot_ID = gcDetailTemp.Marking_Lot_ID & Right$(gcHeaderTemp.Lot_id, 4)

            End If
              
        End If
    
        If customerTemp = "TW067" Then
            If gcHeaderTemp.Customer_Device = "PS5250LT" Or gcHeaderTemp.Customer_Device = "PS5250LT-AA" Then

                Dim waferStr1 As String

                waferStr1 = "123456789ABCDEFGHJKLMNPQRSTUVW"
                
                If Mid(gcDetailTemp.wafer_id, 1, 1) = 0 Then
                    waferPos = Replace(gcDetailTemp.wafer_id, "0", "")
                Else
                    waferPos = gcDetailTemp.wafer_id

                End If
            
                gcDetailTemp.Marking_Lot_ID = gcHeaderTemp.Attri01 & Mid$(waferStr1, waferPos, 1)
      
            End If

        End If
       
        'SH192打标码建立 2018-04-18
        If customerTemp = "SH192" And gcHeaderTemp.WO_NO = "XSH192002" Then
            If InStr(gcHeaderTemp.Lot_id, ".") > 0 Then
                gcDetailTemp.Marking_Lot_ID = "HTG6C" + "\\" + Mid(gcHeaderTemp.Lot_id, InStr(gcHeaderTemp.Lot_id, ".") - 4, 4) + "\\" + Trim(Right(Year(Now), 2)) + Right("0" & DatePart("WW", Now), 2)
            Else
                gcDetailTemp.Marking_Lot_ID = "HTG6C" + "\\" + Right(gcHeaderTemp.Lot_id, 4) + "\\" + Trim(Right(Year(Now), 2)) + Right("0" & DatePart("WW", Now), 2)

            End If

        End If
    
        ' Add by Tony , 20171206_01
        If (customerTemp = "SH115") Then
            gcDetailTemp.Marking_Lot_ID = Mid(cusPTTemp, 3, 4) + Trim(Right(Year(Now), 2)) + Right("0" & DatePart("WW", Now), 2)

        End If
    
        If customerTemp = "KR001" Then
            If cusPTTemp = "PC7090K" Then
                gcDetailTemp.Marking_Lot_ID = GetKRMark(gcHeaderTemp.Lot_id, gcDetailTemp.wafer_id)

            End If
        
            If cusPTTemp = "PS1130K" Or cusPTTemp = "PS4210K" Or cusPTTemp = "PC7080D" Or cusPTTemp = "PK2130K" Then
                gcDetailTemp.Marking_Lot_ID = GetKRMarkP(gcHeaderTemp.Lot_id, gcDetailTemp.wafer_id)

            End If

        End If
    
        If customerTemp = "KR002" Then

            Dim tempAB As String

            Dim tempCD As String
        
            tempAB = Right(gcHeaderTemp.Lot_id, 2)
            tempCD = IIf(Len(gcDetailTemp.wafer_id) = 1, "0" & gcDetailTemp.wafer_id, gcDetailTemp.wafer_id)
            gcDetailTemp.Marking_Lot_ID = tempAB & tempCD
        
        End If
    
        If customerTemp = "KR009" Then
            If cusPTTemp = "Hi-1A1" Then
                gcDetailTemp.Marking_Lot_ID = Right("0" & gcDetailTemp.wafer_id, 2) & "2" & Mid(gcDetailTemp.Lot_id, 5, 3)

            End If
        
        End If
    
        If customerTemp = "HY" Then
            If cusPTTemp = "Hi-258" Then
                gcDetailTemp.Marking_Lot_ID = Right("0" & gcDetailTemp.wafer_id, 2) & "2" & Mid(gcDetailTemp.Lot_id, 5, 3)

            End If
        
        End If
    
        ' Add by Tony 20171226, 客户说第一位空着
        If customerTemp = "US026" And gcHeaderTemp.Customer_Device = "TM2G1" Then
           
            monStr = "123456789ABC"
            waferStr = "123456789ABCDEFGHIJKLMNOP"
        
            If Mid(gcDetailTemp.wafer_id, 1, 1) = 0 Then
                waferPos = Replace(gcDetailTemp.wafer_id, "0", "")
            Else
                waferPos = gcDetailTemp.wafer_id

            End If
        
            gcDetailTemp.Marking_Lot_ID = " " & Right(Year(Now), 1) & Mid(monStr, Month(Now), 1) & Mid$(waferStr, waferPos, 1) & Right(Left(gcDetailTemp.Lot_id, InStr(gcDetailTemp.Lot_id, ".") - 1), 4)

        End If
    
        ' 20180621, by Tony
        If customerTemp = "AT71" Or customerTemp = "AH033" Or customerTemp = "SZ280" Then
            If cusPTTemp = "FP5510EE4" Or cusPTTemp = "FP5513E4" Then
                gcDetailTemp.Marking_Lot_ID = gcHeaderTemp.Attri02 & gcHeaderTemp.Attri03

            End If

        End If
    
        If customerTemp = "RD" Then
            If cusPTTemp = "RDA2216" Then
                gcDetailTemp.Marking_Lot_ID = "RDA" & "2216" & Mid(gcDetailTemp.Lot_id, 3, 4) & Right$("0" & gcDetailTemp.wafer_id, 2)

            End If

        End If
    
        '        If (JudgeSXHeaderId(gcHeaderTemp.lot_id, gcHeaderTemp.po_no, cusPTTemp) And customerTemp <> "US026" And customerTemp <> "SH07" And customerTemp <> "SG005" And customerTemp <> "TW079" And customerTemp <> "TW067" And (Check1.Value = 0)) Then
        '            id = GetSXLotIDPOId(gcHeaderTemp.lot_id, gcHeaderTemp.po_no, cusPTTemp)
        '        Else
        id = GetMaxID()

        If customerNameTemp = "GC" Then
            gcVerTemp = Trim(UCase(gcHeaderTemp.GC_Version))
    
            If cusPTTemp = "GC2145-3" Then
                If Left(gcHeaderTemp.Lot_id, 1) = "H" Then
                    gcVerLastTemp = "G"
                ElseIf Left(gcHeaderTemp.Lot_id, 1) = "E" Then
                    gcVerLastTemp = "F"

                End If

            Else
                gcVerLastTemp = GetGCVerLastChar(cusPTTemp)

            End If
                                
            If gcVerLastTemp <> "" Then
                gcHeaderTemp.GC_Version = gcVerTemp & gcVerLastTemp
                                     
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
            
            Set oiRS = GetGCPT_C(cusPTTemp)

            If (oiRS.RecordCount > 0) Then
                gcHeaderTemp.Customer_Device = oiRS.Fields("CUSTOMERPTNew").Value

            End If
                
            '2012-11-07 jiayun 修改Good_Die_Qty 根据市场部规则
            gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
            
            '2015-08-20 jiayun add 处理 GC0409-3
            If Trim(gcHeaderTemp.Customer_Device) = "GC0409-3" Then
                FabTem = Left(UCase(Trim(gcHeaderTemp.Fab_Device)), 5)
              
                If FabTem = "P6418" Then
                    gcDetailTemp.Good_Die_Qty = 5192
                     
                ElseIf FabTem = "P6820" Then
                 
                    gcDetailTemp.Good_Die_Qty = 11994
                     
                ElseIf FabTem = "P7238" Then
                 
                    gcDetailTemp.Good_Die_Qty = 5191 '5211 ccs alter 20160825

                End If
              
            ElseIf Trim(gcHeaderTemp.Customer_Device) = "GC2145-3" Then
            
                If Left(gcHeaderTemp.Lot_id, 1) = "H" Then
                    gcDetailTemp.Good_Die_Qty = 1676
                ElseIf Left(gcHeaderTemp.Lot_id, 1) = "E" Then
                    gcDetailTemp.Good_Die_Qty = 3920

                End If
            
            End If
            
            '2013-12-27 jiayun add
            If gcDetailTemp.Good_Die_Qty <= 0 Then
                MsgBox "请确认客户机种对应的Die数是否有维护好！"

                Exit Sub

            End If
        
        End If
                
        ' 上传至Header表
        Call AddNormalHeader(gcHeaderTemp, id, customerTemp)
        GCHeaderFlag = True
        
        '判断lotID在Detail表中是否已存在
    
        '2014-03-04 jiayun add  CN Wo  不用抛数据到Mapping表

        If customerNameTemp = "CN" Then
            SumCount = SumCount + 1
    
        ElseIf customerNameTemp = "GT" Then
            SumCount = SumCount + 1
      
        Else

            ' @ 判断是不是二次上传 20170814
            If (JudgeGCDetailId(gcDetailTemp.Lot_id, gcDetailTemp.wafer_id) And JudgeRepeId(customerNameTemp) And Check1.Value = 0) Then
                ' @ 不是二次上传的单
                MsgBox "这笔：" & gcDetailTemp.Lot_id & "; WaferId:" & gcDetailTemp.wafer_id & "已存在，无需上传!"
               
            Else
                '上传到Detail表中
                   
                '                If (JudgeRepeCount(gcDetailTemp.lot_id, gcDetailTemp.WAFER_ID)) Then
                '                    MsgBox "这笔二次上传：" & gcDetailTemp.lot_id & "; WaferId:" & gcDetailTemp.WAFER_ID & "已存在二次版本，无需上传!"
                '                Else

                If customerNameTemp = "CN" Then
                    gcDetailTemp.ITEM = gcDetailTemp.wafer_id
                   
                ElseIf customerNameTemp = "33" Or customerNameTemp = "SZ117" Or customerNameTemp = "HK005" Then  'ccs add 20160720

                    If Len(gcDetailTemp.wafer_id) < 2 Then
                        gcDetailTemp.ITEM = gcDetailTemp.Lot_id & "0" & (gcDetailTemp.wafer_id)
                    Else
                        gcDetailTemp.ITEM = gcDetailTemp.Lot_id & (gcDetailTemp.wafer_id)

                    End If
            
                ElseIf customerNameTemp = "JS113" Then  'ccs add 20160720
                    gcDetailTemp.ITEM = gcDetailTemp.Lot_id & (gcDetailTemp.wafer_id)
            
                ElseIf customerNameTemp = "XA" Then

                    If Len(gcDetailTemp.wafer_id) = 2 Then
                        gcDetailTemp.ITEM = gcDetailTemp.Lot_id & (gcDetailTemp.wafer_id)
                    Else
                        gcDetailTemp.ITEM = gcDetailTemp.Lot_id & "0" & (gcDetailTemp.wafer_id)

                    End If
          
                Else

                    If InStr(gcDetailTemp.wafer_id, "+") Then
                        gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 3)
                    Else
                        gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 2)

                    End If
          
                    If customerNameTemp = "AA(ON)" Then
                        If Len(gcDetailTemp.wafer_id) = 4 Then
                            gcDetailTemp.ITEM = gcDetailTemp.Lot_id & Right(("0" & gcDetailTemp.wafer_id), 4)

                        End If

                    End If
                   
                End If
                
                ' 判断SX  5位 8位      20171025
                '  If (customerTemp = "SX") And (check_sx_eight(cusPTTemp)) Then
                '      gcDetailTemp.Marking_Lot_ID = GetSX8CodeID(Trim(gcDetailTemp.lot_id), Trim(gcDetailTemp.wafer_id))
                '  End If
                
                '判断 是否特殊机种，要8位码的
                ' If (customerTemp = "SX" Or customerTemp = "HJ") And UCase(Trim(gcHeaderTemp.Customer_Device)) = "OV02A" Then
                If (customerTemp = "SX" Or customerTemp = "HJ") And (UCase(Trim(gcHeaderTemp.Customer_Device)) = "OV02A" Or UCase(Trim(gcHeaderTemp.Customer_Device)) = "OV02A-E" Or UCase(Trim(gcHeaderTemp.Customer_Device)) = "SP5506-M" Or UCase(Trim(gcHeaderTemp.Customer_Device)) = "SP5506" Or UCase(Trim(gcHeaderTemp.Customer_Device)) = "SP5506-E" Or UCase(Trim(gcHeaderTemp.Customer_Device)) = "SP5506-EM" Or UCase(Trim(gcHeaderTemp.Customer_Device)) = "SP8407-E" Or UCase(Trim(gcHeaderTemp.Customer_Device)) = "SP8407" Or UCase(Trim(gcHeaderTemp.Customer_Device)) = "SP5407-E" Or UCase(Trim(gcHeaderTemp.Customer_Device)) = "SP5407" Or UCase(Trim(gcHeaderTemp.Customer_Device)) = "SP2735") Then
                    gcDetailTemp.Marking_Lot_ID = GetSX8CodeID(Trim(gcDetailTemp.Lot_id), Trim(gcDetailTemp.wafer_id))

                End If
                 
                'ccs add 20161102 AB18客户把*号替换为年周
                If (customerTemp = "AB18") Then
                    gcDetailTemp.Marking_Lot_ID = Replace(gcDetailTemp.Marking_Lot_ID, "****", Trim(Right(Year(Now), 2)) + Trim(DatePart("WW", Now)))

                End If
            
                ' @ SUBSTRATEID '+ ' 20170814
                If (Not JudgeRepeId(customerNameTemp)) Then
                    gcDetailTemp.ITEM = gcDetailTemp.ITEM & "+"

                End If

                Dim sOra As String
                
                ' SX回货 2018/0302
                If Check1.Value = 1 Then

                    Dim strsql As String

                    Dim Rs2    As New ADODB.Recordset
                        
                    If Left(gcDetailTemp.wafer_id, 1) = "0" Then
                        gcDetailTemp.wafer_id = Replace$(gcDetailTemp.wafer_id, "0", "")

                    End If
                        
                    strsql = "select distinct b.filename,b.substrateid from customeroitbl_test a, mappingdatatest b where b.lotid = '" & gcDetailTemp.Lot_id & "' and b.wafer_id in  ('0'||'" & gcDetailTemp.wafer_id & "', '" & gcDetailTemp.wafer_id & "') and to_char(a.id) = b.filename and a.source_batch_id = b.lotid " & " and a.invflag = 0 and instr(b.substrateid, '+') > 0 and not exists (select 1 from ib_waferlist c where b.substrateid = c.waferid) "

                    If Rs2.State = adStateOpen Then Rs2.Close
                    Rs2.Open strsql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

                    If Rs2.RecordCount > 0 Then
                        MsgBox "第" & gcDetailTemp.wafer_id & "片已经生成了未开立工单的WaferID: " & Rs2("substrateid") & ",this is update case", vbCritical, "提醒"
                            
                        AddSql ("delete from mappingdatatest where filename = '" & Rs2(0) & "'")
                        AddSql ("delete from customeroitbl_test where id = '" & Rs2(0) & "'")
                        AddSql2 ("delete from [ERPBASE].[dbo].[tblCustomerOI]  where id = '" & Rs2(0) & "' ")
                        AddSql2 ("delete from [ERPBASE].[dbo].[tblmappingData]  where filename = '" & Rs2(0) & "'")
                            
                    End If
                    
                    sOra = "select a.productid  from mappingdatatest a WHERE a.substrateid IN ( '" & gcDetailTemp.ITEM & "')"
                    
                    gcDetailTemp.Marking_Lot_ID = Get_OracleStr(sOra)
                
                    gcDetailTemp.ITEM = Get_OracleStr("select max(substrateid) || '+' as substrateid from mappingdatatest where wafer_id in ('" & gcDetailTemp.wafer_id & "', '0' || '" & gcDetailTemp.wafer_id & "') and lotid = '" & gcDetailTemp.Lot_id & "' group by lotid")

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

    'VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！", vbInformation, "友情提示"

    End If
    
End Sub


Private Function check_sx_eight(Customer_Device As String) As Boolean
    check_sx_eight = False

    Dim cmd_ora As String

    cmd_ora = "select a.marking_code from tbltsvnpiproduct a where a.customerptno1  = '" & Customer_Device & "' or a.customerptno2 = '" & Customer_Device & "' or a.customerptno3 = '" & Customer_Device & "' or a.customerptno4 = '" & Customer_Device & "' "

    If GetSeqTW(cmd_ora) <> Null Then
    
        If GetSeqTW(cmd_ora) = 8 Then
            check_sx_eight = True

        End If

    End If

End Function

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
        StrFileName = Mid(Trim(Text3.text), InStrRev(Trim(Text3.text), "\") + 1)
        dirName = Mid$(Trim(Text3.text), 1, InStrRev(Trim(Text3.text), "\"))

    End If

    Dim con As New ADODB.Connection

    Dim rs  As New ADODB.Recordset

    con.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
    rs.Open "Select * From " & "[" & StrFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
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
        StrFileName = Mid(Trim(Text3.text), InStrRev(Trim(Text3.text), "\") + 1)
        dirName = Mid$(Trim(Text3.text), 1, InStrRev(Trim(Text3.text), "\"))

    End If

    Dim con As New ADODB.Connection

    Dim rs  As New ADODB.Recordset

    con.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
    rs.Open "Select * From " & "[" & StrFileName & "]", con, adOpenStatic, adLockReadOnly, adCmdText
        
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

MsgBox "老接口已关闭, 请使用新版上传通用版WO", vbInformation, "提示"
Exit Sub

    Dim customerStr As String

    If Trim(CmbCustomer.text) = "" Then
        MsgBox "请先选择客户！"

        Exit Sub

    End If

    customerStr = UCase(Trim(CmbCustomer.text))
    Call UploadNormalCustomerNew(customerStr)

End Sub

Private Function GetGCGoodDieQty(productNameTemp As String, dieQtyTemp As Long) As Integer
    '2013-12-26 jiayun add
    '根据Gc pt 查询数量

    GetGCGoodDieQty = 0

    Set updateRS = GetWO_GC_Die(productNameTemp)
    GetGCGoodDieQty = CInt(updateRS.Fields("dieqty").Value)

    'Dim productNameTemp2 As String
    '
    'If productNameTemp <> "" And dieQtyTemp > 0 Then
    '    productNameTemp2 = UCase(Left(Trim(productNameTemp), Len(Trim(productNameTemp)) - 2))
    '
    '    Select Case productNameTemp2
    '
    '    Case "GC6113"
    '        GetGCGoodDieQty = 6975
    '
    '    Case "GC0311"
    '        GetGCGoodDieQty = 5584
    '
    '    Case "GC0329"
    '        GetGCGoodDieQty = 4722
    '
    '    Case "GC0313"
    '        GetGCGoodDieQty = 5364
    '
    '    Case "GC2035"
    '        GetGCGoodDieQty = 1547
    '
    '    Case "GC6123"
    '        'GetGCGoodDieQty = 8688
    '        '2013-11-04 jiayun modify
    '
    '        GetGCGoodDieQty = 8706
    '
    '    Case "GC0328"
    '        GetGCGoodDieQty = 3382
    '
    '    Case "GC1004"
    '        GetGCGoodDieQty = 1302
    '
    '    Case Else
    '        GetGCGoodDieQty = 0
    '
    '    End Select
    '
    'Else
    '
    '    GetGCGoodDieQty = 0
    'End If

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

    ExporToExcel ("  select po_num as PO_NO, ship_site as Supplier,test_site as Ship_To, fab_conv_id as FAB_Device, mpn_desc as Customer_Device," & " imager_customer_rev as GC_Version,created_date as GC_Date,source_batch_id  as Lot_ID, mtrl_num   As WO_NO , probe_ship_part_type as 贸易类型 ,  RETICLE_LEVEL_71 as Attribute1,RETICLE_LEVEL_72 as Attribute2,RETICLE_LEVEL_73 as Attribute3,ASSEMBLY_FACILITY as Attribute4,BATCH_COMMENT_ASSY as Attribute5 " & " From CustomerOItbl_test  where CustomerShortName = '" & CmbCustomer.text & "'order by id ")

End Sub

Private Sub Command9_Click()

    If CmbCustomer.text = "" Then
        MsgBox "请先选择客户！"

        Exit Sub

    End If

    If CmbCustomer.text = "GC_WLD/T" Then

        ExporToExcel (" select substrateid as Item ,productid as Marking_Lot_ID ,lotid as Lot_ID ,wafer_id ,passbincount as Good_Die_Qty " & " from  mappingDataTest where  CustomerShortName = '" & CmbCustomer.text & "' and remark='WLT' order by id")
 
    Else

        ExporToExcel (" select substrateid as Item ,productid as Marking_Lot_ID ,lotid as Lot_ID ,wafer_id ,passbincount as Good_Die_Qty ,failbincount as NG_Die_Qty" & " from  mappingDataTest where  CustomerShortName = '" & CmbCustomer.text & "' order by id")
               
    End If

    ' ExporToExcel (" select substrateid as Item ,productid as Marking_Lot_ID ,lotid as Lot_ID ,wafer_id ,passbincount as Good_Die_Qty " & _
    '               " from  mappingDataTest where  CustomerShortName = '" & CmbCustomer.Text & "' order by id")
 
End Sub

Private Sub Form_Load()

    gCnt = 0
    Check1.Value = 0

    CommonDialog9.flags = cdlOFNAllowMultiselect + cdlOFNExplorer

    DTPicker4.Value = Now

    DTPicker5.Value = Now + 1

    Com.flags = &H80200

    ComSI.flags = &H80200

    CommonDialog7.flags = &H80200

    CmbCustomer37.AddItem ("37")
    CmbCustomer37.AddItem ("68")
    'CmbCustomer37.AddItem ("70")
    CmbCustomer37.AddItem ("HK006")
    CmbCustomer37.AddItem ("BJ128")

    IniCustomerName

    'IniCustomerName2

    'CmbCustomer.AddItem ("GC")
    'CmbCustomer.AddItem ("GC_WLD/T")
    'CmbCustomer.AddItem ("SX")
    'CmbCustomer.AddItem ("HJ")
    '
    'CmbCustomer.AddItem ("PT")
    'CmbCustomer.AddItem ("SY")
    'CmbCustomer.AddItem ("RD")
    'CmbCustomer.AddItem ("DN")
    'CmbCustomer.AddItem ("BD")
    'CmbCustomer.AddItem ("ZX")
    'CmbCustomer.AddItem ("HY")
    'CmbCustomer.AddItem ("HT")
    'CmbCustomer.AddItem ("OT")
    'CmbCustomer.AddItem ("MC")
    ''2014-09-17 jiayun modify si 改为GT
    'CmbCustomer.AddItem ("GT")
    '
    'CmbCustomer.AddItem ("CN")
    'CmbCustomer.AddItem ("KT")
    'CmbCustomer.AddItem ("HD")
    '
    'CmbCustomer.AddItem ("RS")
    'CmbCustomer.AddItem ("SD")
    '
    'CmbCustomer.AddItem ("QR")
    'CmbCustomer.AddItem ("QR2")
    '
    'CmbCustomer.AddItem ("MG")
    'CmbCustomer.AddItem ("LX")
    'CmbCustomer.AddItem ("GD")
    'CmbCustomer.AddItem ("AM")
    'CmbCustomer.AddItem ("EQ")
    'CmbCustomer.AddItem ("EQ_IS")
    'CmbCustomer.AddItem ("ZL")
    'CmbCustomer.AddItem ("YW")
    'CmbCustomer.AddItem ("RO")
    'CmbCustomer.AddItem ("MR")
    'CmbCustomer.AddItem ("CS")
    '
    'CmbCustomer.AddItem ("36")
    'CmbCustomer.AddItem ("34")
    'CmbCustomer.AddItem ("33")
    '
    'CmbCustomer.AddItem ("32")
    'CmbCustomer.AddItem ("45")
    'CmbCustomer.AddItem ("50")
    'CmbCustomer.AddItem ("60")
    '
    'CmbCustomer.AddItem ("30")
    'CmbCustomer.AddItem ("55")
    'CmbCustomer.AddItem ("54")
    'CmbCustomer.AddItem ("56")
    'CmbCustomer.AddItem ("57")
    'CmbCustomer.AddItem ("49")
    'CmbCustomer.AddItem ("59")
    'CmbCustomer.AddItem ("64")
    'CmbCustomer.AddItem ("61")
    '
    'CmbCustomer.AddItem ("68")
    'CmbCustomer.AddItem ("70")
    'CmbCustomer.AddItem ("69")
    'CmbCustomer.AddItem ("80")
    'CmbCustomer.AddItem ("81")
    'CmbCustomer.AddItem ("87")
    'CmbCustomer.AddItem ("88")
    'CmbCustomer.AddItem ("94")
    'CmbCustomer.AddItem ("93")
    'CmbCustomer.AddItem ("95")
    'CmbCustomer.AddItem ("B1")
    '
    '
    'CmbCustomer.AddItem ("XW")
    '
    '
    'CmbCustomer.AddItem ("YX")
    '
    'CmbCustomer.AddItem ("37")
    'CmbCustomer.AddItem ("77")
    'CmbCustomer.AddItem ("78")
    '
    '
    'CmbCustomer.AddItem ("XA")
    'CmbCustomer.AddItem ("HH")
    'CmbCustomer.AddItem ("SL")

    Combo1.AddItem ("AA")
    Combo1.AddItem ("自购")
    Combo1.AddItem ("CN")

    CmbPoType.AddItem ("Bumping")
    CmbPoType.AddItem ("ICI")

End Sub

Private Sub IniCustomerName()
    Set mainItemRS = GetJDCustomerName()
    Set CmbCustomer.RowSource = mainItemRS
    CmbCustomer.ListField = mainItemRS("productname").name
    CmbCustomer.BoundColumn = mainItemRS("PID").name

End Sub

'Private Sub IniCustomerName2()
'Set mainItemRS = GetJDCustomerName()
'Set DataCombo1.RowSource = mainItemRS
'DataCombo1.ListField = mainItemRS("productname").Name
'DataCombo1.BoundColumn = mainItemRS("PID").Name
'
'End Sub

Private Sub Label1_Click()

End Sub

