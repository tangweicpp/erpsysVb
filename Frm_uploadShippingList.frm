VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_uploadShippingList 
   Caption         =   "上传客户出货挑料信息"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20220
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
   ScaleHeight     =   10935
   ScaleWidth      =   20220
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   12495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   20415
      _ExtentX        =   36010
      _ExtentY        =   22040
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "DN上传"
      TabPicture(0)   =   "Frm_uploadShippingList.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "DN修改/删除"
      TabPicture(1)   =   "Frm_uploadShippingList.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "DN出货"
      TabPicture(2)   =   "Frm_uploadShippingList.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Toolbar3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Fra(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "SSTab2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin TabDlg.SSTab SSTab2 
         Height          =   9495
         Left            =   4080
         TabIndex        =   78
         Top             =   2760
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   16748
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   16777215
         TabCaption(0)   =   "出货申请"
         TabPicture(0)   =   "Frm_uploadShippingList.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Fra(2)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Fra(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "申请查看"
         TabPicture(1)   =   "Frm_uploadShippingList.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Fps_Ship_del(1)"
         Tab(1).Control(1)=   "Fps_Ship_del(0)"
         Tab(1).Control(2)=   "Opt1"
         Tab(1).Control(3)=   "Opt2"
         Tab(1).Control(4)=   "Opt3"
         Tab(1).ControlCount=   5
         Begin VB.OptionButton Opt3 
            Caption         =   "所有"
            Height          =   375
            Left            =   -72600
            TabIndex        =   95
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Opt2 
            Caption         =   "未审核"
            Height          =   375
            Left            =   -73680
            TabIndex        =   94
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton Opt1 
            Caption         =   "已审核"
            Height          =   375
            Left            =   -74760
            TabIndex        =   93
            Top             =   600
            Width           =   975
         End
         Begin VB.Frame Fra 
            Caption         =   "已申请"
            ForeColor       =   &H00FF0000&
            Height          =   2415
            Index           =   0
            Left            =   120
            TabIndex        =   89
            Top             =   6960
            Width           =   15735
            Begin FPSpreadADO.fpSpread Fps_Ship 
               Height          =   2055
               Index           =   2
               Left            =   120
               TabIndex        =   90
               Top             =   240
               Width           =   15495
               _Version        =   524288
               _ExtentX        =   27331
               _ExtentY        =   3625
               _StockProps     =   64
               EditEnterAction =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   0
               MaxRows         =   0
               SpreadDesigner  =   "Frm_uploadShippingList.frx":008C
               TextTip         =   2
               AppearanceStyle =   0
            End
         End
         Begin VB.Frame Fra 
            Caption         =   "待出货库存资料"
            ForeColor       =   &H00FF0000&
            Height          =   6495
            Index           =   2
            Left            =   120
            TabIndex        =   79
            Top             =   480
            Width           =   15700
            Begin VB.CheckBox Chk_All 
               Caption         =   "全选/全不选"
               Height          =   255
               Left            =   120
               TabIndex        =   96
               Top             =   240
               Width           =   1695
            End
            Begin FPSpreadADO.fpSpread Fps_Ship 
               Height          =   5775
               Index           =   0
               Left            =   120
               TabIndex        =   81
               Top             =   500
               Width           =   8415
               _Version        =   524288
               _ExtentX        =   14843
               _ExtentY        =   10186
               _StockProps     =   64
               EditEnterAction =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   0
               MaxRows         =   0
               SpreadDesigner  =   "Frm_uploadShippingList.frx":0525
               TextTip         =   2
               AppearanceStyle =   0
            End
            Begin FPSpreadADO.fpSpread Fps_Ship 
               Height          =   5775
               Index           =   1
               Left            =   8640
               TabIndex        =   82
               Top             =   500
               Width           =   6975
               _Version        =   524288
               _ExtentX        =   12303
               _ExtentY        =   10186
               _StockProps     =   64
               EditEnterAction =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   0
               MaxRows         =   0
               SpreadDesigner  =   "Frm_uploadShippingList.frx":09BE
               TextTip         =   2
               AppearanceStyle =   0
            End
            Begin VB.TextBox txtID_Text 
               Height          =   495
               Left            =   960
               TabIndex        =   80
               Top             =   5400
               Visible         =   0   'False
               Width           =   855
            End
         End
         Begin FPSpreadADO.fpSpread Fps_Ship_del 
            Height          =   7695
            Index           =   0
            Left            =   -74880
            TabIndex        =   91
            Top             =   1320
            Width           =   6615
            _Version        =   524288
            _ExtentX        =   11668
            _ExtentY        =   13573
            _StockProps     =   64
            EditEnterAction =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   0
            MaxRows         =   0
            SpreadDesigner  =   "Frm_uploadShippingList.frx":0E57
            TextTip         =   2
            AppearanceStyle =   0
         End
         Begin FPSpreadADO.fpSpread Fps_Ship_del 
            Height          =   7695
            Index           =   1
            Left            =   -68160
            TabIndex        =   92
            Top             =   1320
            Width           =   8895
            _Version        =   524288
            _ExtentX        =   15690
            _ExtentY        =   13573
            _StockProps     =   64
            EditEnterAction =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   0
            MaxRows         =   0
            SpreadDesigner  =   "Frm_uploadShippingList.frx":12F0
            TextTip         =   2
            AppearanceStyle =   0
         End
      End
      Begin VB.Frame Fra 
         Caption         =   "资料维护"
         ForeColor       =   &H00FF0000&
         Height          =   1455
         Index           =   1
         Left            =   4200
         TabIndex        =   51
         Top             =   1300
         Width           =   15735
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   9480
            TabIndex        =   64
            Top             =   180
            Width           =   1455
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   900
            TabIndex        =   63
            Top             =   180
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   3840
            TabIndex        =   62
            Top             =   180
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   6720
            TabIndex        =   61
            Top             =   180
            Width           =   1695
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   6
            ItemData        =   "Frm_uploadShippingList.frx":1789
            Left            =   900
            List            =   "Frm_uploadShippingList.frx":178B
            TabIndex        =   60
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   7
            Left            =   3840
            TabIndex        =   59
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   12000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   58
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   9480
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   57
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   900
            TabIndex        =   56
            Top             =   1020
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   3840
            TabIndex        =   55
            Top             =   1020
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   6720
            TabIndex        =   54
            Top             =   1020
            Width           =   1695
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   9
            Left            =   6720
            TabIndex        =   53
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton cmdCreate 
            Caption         =   "生成"
            Height          =   345
            Left            =   11040
            TabIndex        =   52
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "已选数量/SOLINE总量"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   24
            Left            =   11520
            TabIndex        =   102
            Top             =   1080
            Width           =   1710
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0/0"
            ForeColor       =   &H00008000&
            Height          =   195
            Index           =   23
            Left            =   13560
            TabIndex        =   101
            Top             =   1080
            Width           =   285
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单据编号"
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   6
            Left            =   8760
            TabIndex        =   77
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请人"
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   7
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请时间"
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   8
            Left            =   2900
            TabIndex        =   75
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请部门"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   5800
            TabIndex        =   74
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收货客户"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   2900
            TabIndex        =   73
            Top             =   675
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "运输方式"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   72
            Top             =   675
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收货地址"
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   12
            Left            =   5800
            TabIndex        =   71
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "备       注"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   8760
            TabIndex        =   70
            Top             =   720
            Width           =   675
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "良品数"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   120
            TabIndex        =   69
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "不良数"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   2900
            TabIndex        =   68
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "制程不良"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   5800
            TabIndex        =   67
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "已勾箱数/查询总箱数"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   18
            Left            =   8760
            TabIndex        =   66
            Top             =   1080
            Width           =   1680
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0/0"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   19
            Left            =   10560
            TabIndex        =   65
            Top             =   1080
            Width           =   45
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "查询条件"
         Height          =   10695
         Left            =   120
         TabIndex        =   33
         Top             =   1300
         Width           =   3855
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   11
            Left            =   1320
            TabIndex        =   99
            Top             =   3840
            Width           =   2415
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   10
            Left            =   1320
            TabIndex        =   98
            Top             =   3480
            Width           =   2415
         End
         Begin VB.Frame Frame9 
            Caption         =   "未OK"
            Height          =   2175
            Left            =   0
            TabIndex        =   86
            Top             =   8160
            Width           =   3735
            Begin VB.ListBox List_dn 
               Columns         =   2
               Height          =   1635
               Index           =   1
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   87
               Top             =   360
               Width           =   3375
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "待出货"
            Height          =   2895
            Left            =   0
            TabIndex        =   84
            Top             =   5160
            Width           =   3735
            Begin VB.ListBox List_dn 
               Columns         =   2
               Height          =   2310
               Index           =   0
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   85
               Top             =   240
               Width           =   3375
            End
         End
         Begin VB.CheckBox chkDZD 
            Caption         =   "免对账单勾选"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   83
            Top             =   4680
            Width           =   1575
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   8
            Left            =   1320
            TabIndex        =   41
            Top             =   3120
            Width           =   2415
         End
         Begin VB.CheckBox chk 
            Caption         =   "查询时保留已选择的LOT号或料号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   450
            Left            =   120
            TabIndex        =   40
            Top             =   4320
            Width           =   3500
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   5
            Left            =   1320
            Style           =   1  'Simple Combo
            TabIndex        =   39
            Top             =   2280
            Width           =   2415
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   4
            Left            =   1320
            Style           =   1  'Simple Combo
            TabIndex        =   38
            Top             =   1920
            Width           =   2415
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   3
            Left            =   1320
            TabIndex        =   37
            Top             =   1560
            Width           =   2415
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   2
            Left            =   1320
            TabIndex        =   36
            Top             =   1200
            Width           =   2415
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   35
            Top             =   765
            Width           =   2415
         End
         Begin VB.ComboBox Cob 
            Height          =   315
            Index           =   0
            ItemData        =   "Frm_uploadShippingList.frx":178D
            Left            =   1320
            List            =   "Frm_uploadShippingList.frx":1797
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   360
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker dtShipDate_Ship 
            Height          =   375
            Left            =   1320
            TabIndex        =   42
            Top             =   2640
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   393216
            Format          =   114098177
            CurrentDate     =   43879
         End
         Begin VB.TextBox TxtShipDate_Ship 
            Height          =   285
            Left            =   3120
            TabIndex        =   103
            Top             =   2640
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "客户机种"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   100
            Top             =   3600
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "      SO"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   97
            Top             =   3960
            Width           =   480
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出货日期"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   50
            Top             =   2760
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "      DN"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   20
            Left            =   240
            TabIndex        =   49
            Top             =   3240
            Width           =   480
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "客户PO_NUM"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   48
            Top             =   2355
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "料        号"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   47
            Top             =   1995
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "工单或Lot号"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   46
            Top             =   1635
            Width           =   945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "库房名称"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   45
            Top             =   1275
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "产线标记"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   44
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "客户代码"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   43
            Top             =   435
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "修改记录"
         ForeColor       =   &H00FF0000&
         Height          =   3975
         Left            =   -74760
         TabIndex        =   28
         Top             =   8280
         Width           =   20055
         Begin FPSpreadADO.fpSpread fpS_Mod 
            Height          =   3375
            Index           =   1
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   19575
            _Version        =   524288
            _ExtentX        =   34528
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
            MaxCols         =   0
            MaxRows         =   0
            SpreadDesigner  =   "Frm_uploadShippingList.frx":17A6
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "查询结果"
         ForeColor       =   &H00FF0000&
         Height          =   5175
         Left            =   -74760
         TabIndex        =   27
         Top             =   2880
         Width           =   20055
         Begin FPSpreadADO.fpSpread fpS_Mod 
            Height          =   4335
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   480
            Width           =   19575
            _Version        =   524288
            _ExtentX        =   34528
            _ExtentY        =   7646
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
            MaxCols         =   0
            MaxRows         =   0
            SpreadDesigner  =   "Frm_uploadShippingList.frx":1BC8
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "功能选项"
         ForeColor       =   &H00FF0000&
         Height          =   2295
         Left            =   -74760
         TabIndex        =   18
         Top             =   360
         Width           =   20055
         Begin VB.ComboBox CbItem_Mod 
            Height          =   315
            ItemData        =   "Frm_uploadShippingList.frx":1FEA
            Left            =   1320
            List            =   "Frm_uploadShippingList.frx":1FEC
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1920
            Width           =   2055
         End
         Begin VB.TextBox txtShipDate_Mod 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   9960
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1200
            Width           =   1455
         End
         Begin VB.ComboBox cbCustomerCode_Mod 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            ItemData        =   "Frm_uploadShippingList.frx":1FEE
            Left            =   1320
            List            =   "Frm_uploadShippingList.frx":1FF8
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox txtDN_Mod 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   19
            Top             =   1575
            Width           =   2055
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   11280
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":2007
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":2C59
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":38AB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":44FD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":514F
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   870
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   17265
            _ExtentX        =   30454
            _ExtentY        =   1535
            ButtonWidth     =   1455
            ButtonHeight    =   1482
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Appearance      =   1
            Style           =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   7
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "查询"
                  Key             =   "QUERY_MOD"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "确认修改"
                  Key             =   "SAVE_MOD"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "确认删除"
                  Key             =   "DEL_MOD"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "退出"
                  Key             =   "HOME_MOD"
                  ImageIndex      =   5
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin MSComCtl2.DTPicker dtShipDate_mod 
            Height          =   375
            Left            =   8400
            TabIndex        =   23
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   114098177
            CurrentDate     =   43271
         End
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   12000
            Top             =   360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
            Flags           =   524800
            MaxFileSize     =   9999
         End
         Begin VB.Label Label6 
            Caption         =   "修改项目"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "客户代码:"
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
            Left            =   240
            TabIndex        =   26
            Top             =   1230
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D.N:"
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
            Left            =   840
            TabIndex        =   25
            Top             =   1590
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择新的出货日期:"
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
            Left            =   6360
            TabIndex        =   24
            Top             =   1250
            Width           =   1875
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "挑料明细"
         ForeColor       =   &H00FF0000&
         Height          =   9255
         Left            =   -74760
         TabIndex        =   13
         Top             =   3000
         Width           =   20055
         Begin VB.TextBox txtDNCheck 
            BackColor       =   &H00FFC0FF&
            Height          =   4095
            Left            =   12960
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   480
            Width           =   5175
         End
         Begin FPSpreadADO.fpSpread Fps 
            Height          =   8775
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   12495
            _Version        =   524288
            _ExtentX        =   22040
            _ExtentY        =   15478
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
            MaxCols         =   80
            MaxRows         =   0
            SpreadDesigner  =   "Frm_uploadShippingList.frx":5DA1
            TextTip         =   2
         End
         Begin FPSpreadADO.fpSpread Fps 
            Height          =   3855
            Index           =   1
            Left            =   12960
            TabIndex        =   16
            Top             =   4920
            Width           =   6855
            _Version        =   524288
            _ExtentX        =   12091
            _ExtentY        =   6800
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
            SpreadDesigner  =   "Frm_uploadShippingList.frx":621F
            TextTip         =   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DateCode 检查"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   12960
            TabIndex        =   17
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "功能选项"
         ForeColor       =   &H00FF0000&
         Height          =   2055
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   20055
         Begin VB.TextBox txtDN 
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   4560
            TabIndex        =   5
            Top             =   1215
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.ComboBox cbCustomerCode 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            ItemData        =   "Frm_uploadShippingList.frx":6691
            Left            =   1320
            List            =   "Frm_uploadShippingList.frx":669B
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtFileName 
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1680
            Width           =   9015
         End
         Begin VB.TextBox txtShipDate 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   10320
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   1163
            Visible         =   0   'False
            Width           =   1455
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   11280
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":66AA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":72FC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":7F4E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":8BA0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_uploadShippingList.frx":97F2
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   600
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   10185
            _ExtentX        =   17965
            _ExtentY        =   1058
            ButtonWidth     =   1984
            ButtonHeight    =   1005
            AllowCustomize  =   0   'False
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   9
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "查询"
                  Key             =   "QUERY"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "打开"
                  Key             =   "OPEN"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Caption         =   "保存"
                  Key             =   "SAVE"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "删除"
                  Key             =   "DEL"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "退出"
                  Key             =   "HOME"
                  ImageIndex      =   5
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin MSComCtl2.DTPicker dtShipDate 
            Height          =   375
            Left            =   8880
            TabIndex        =   7
            Top             =   1163
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   114098177
            CurrentDate     =   43271
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   12000
            Top             =   360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
            Flags           =   524800
            MaxFileSize     =   9999
         End
         Begin VB.Label lblShipDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出货日期:"
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
            Left            =   7800
            TabIndex        =   12
            Top             =   1230
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblDN 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D.N:"
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
            Left            =   4080
            TabIndex        =   11
            Top             =   1230
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblCustomerCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "客户代码:"
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
            Left            =   240
            TabIndex        =   10
            Top             =   1230
            Width           =   975
         End
         Begin VB.Label lblFileName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "文件名(N):"
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
            Left            =   240
            TabIndex        =   9
            Top             =   1680
            Width           =   1020
         End
         Begin VB.Label lbl123 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DN已包含出货日期,上传请勿选择出货日期"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   11880
            TabIndex        =   8
            Top             =   1200
            Width           =   4485
         End
      End
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   870
         Left            =   120
         TabIndex        =   88
         Top             =   360
         Width           =   19665
         _ExtentX        =   34687
         _ExtentY        =   1535
         ButtonWidth     =   1455
         ButtonHeight    =   1482
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查询"
               Key             =   "QUERY_Ship"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "确认"
               Key             =   "SAVE_Ship"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "确认删除"
               Key             =   "DEL_Ship"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "HOME_Ship"
               ImageIndex      =   5
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
   End
End
Attribute VB_Name = "Frm_uploadShippingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public intFlag           As Integer '操作标记
Public strRealName As String
Public lsolineqty           As Long 'soline数量
Public strsono           As String 'sono
Public strsoline           As String 'soline
Public mailflag As Integer

Private Enum FpsM
    E_CHOOSE = 1        '选择
    e_ID                '主表ID
    E_cust              '客户代码
    E_DN                'DN
    e_GDH               '工单号或LOT号
    e_BigX              '大箱号
    e_LH                '料号
    e_NUM               '库存数
    E_GG                '规格
    E_XH                '型号
    E_UNIT              '单位
    e_KF                '出货库房

    e_MCol
End Enum
Private Enum FpsD
    e_BigX              '大箱号 隐藏字段
    E_XH                '箱号
    e_GDH               '工单号或LOT号
    e_LCK               '流程卡
    e_LH                '料号
    e_GNum              '合格数
    e_BLNum             '客户不良数
    e_ZCNum             '制程不良数
    e_ID                '明细ID
    e_KF                '出货库房
    E_DN                'DN
   ' E_Note               'DN
    E_Note2               'DN
    e_MCol
End Enum

Dim adoPrmReturn        As ADODB.Parameter
Dim adoprm1             As ADODB.Parameter
Dim adoprm2             As ADODB.Parameter
Dim adoPrm3             As ADODB.Parameter
Dim adoPrm4             As ADODB.Parameter
Dim adoPrm5             As ADODB.Parameter
Dim adoPrm6             As ADODB.Parameter
Dim adoPrm7             As ADODB.Parameter
Dim adoPrm8             As ADODB.Parameter
Dim adoPrm9             As ADODB.Parameter
Dim adoprmFG            As ADODB.Parameter
Dim adoprm10            As ADODB.Parameter
Dim adoPrm11            As ADODB.Parameter
Dim adoPrm12            As ADODB.Parameter
Dim adoPrm13            As ADODB.Parameter
Dim adoPrm14            As ADODB.Parameter
Dim adoPrm15            As ADODB.Parameter
Dim adoprm16            As ADODB.Parameter




Private Sub cbCustomerCode_Click()

Select Case cbCustomerCode.text

    Case "37"  '37
        lblDN.Visible = True
        txtDN.Visible = True
        lblShipDate.Visible = True
        dtShipDate.Visible = True
        txtShipDate.Visible = True
        Toolbar1.Buttons(7).Visible = True
        txtDNCheck.Visible = True
        fps(1).Visible = True
        lbl123.Visible = True
        Label1.Visible = True
    Case "SG005"
    
        lblDN.Visible = False
        txtDN.Visible = False
        lblShipDate.Visible = False
        dtShipDate.Visible = False
        txtShipDate.Visible = False
        Toolbar1.Buttons(7).Visible = False
        txtDNCheck.Visible = False
        fps(1).Visible = False
        lbl123.Visible = False
        Label1.Visible = False
    Case Else
        lblDN.Visible = False
        txtDN.Visible = False
        lblShipDate.Visible = False
        dtShipDate.Visible = False
        txtShipDate.Visible = False

End Select

End Sub

Private Function CheckDC() As Boolean
Dim i             As Integer
Dim strDN         As String
Dim strJobID      As String
Dim strDNDC       As String
Dim str37TESTDC   As String
Dim strSql        As String
Dim strWrongJobID As String
Dim bWrongJobID   As Boolean
txtDNCheck.text = ""
txtDNCheck.text = "DN                   JOBID                DNDC          正确DC" & vbCrLf
bWrongJobID = False
strWrongJobID = ""

With fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        strDN = Trim$("" & .text)
        .Col = 32
        strJobID = Trim$("" & .text)
        .Col = 52
        strDNDC = Trim$("" & .text)

        If Right$(strJobID, 1) = "M" Then
            str37TESTDC = Get_SqlStr("select DC from erptemp..tbl37testdc_m where JOBID = '" & strJobID & "'")
        Else
            str37TESTDC = Get_OracleStr("select dc from tbl37testdc where jobid = '" & strJobID & "'")

        End If

        If str37TESTDC = "" Then
            MsgBox "JOBID:" & strJobID & "查询不到本地对应的DC,无法检查客户DN是否正确,请联系IT", vbCritical, "警告"
            bWrongJobID = True
        Else

            If str37TESTDC <> "" And str37TESTDC <> strDNDC Then
                txtDNCheck.text = txtDNCheck.text & strDN & "       " & strJobID & "           " & strDNDC & "            " & str37TESTDC & vbCrLf
                bWrongJobID = True

            End If

        End If

    Next i

End With

If bWrongJobID = False Then
    CheckDC = True
Else
    CheckDC = False

End If

End Function





Private Sub CancerSelection(ROW_S, ROW_E)
Dim i As Integer
Dim strBigbox As String

     With Fps_Ship(0)
     For i = 1 To .MaxRows
         .Row = i
         .Col = FpsM.E_CHOOSE
         If .text = "1" Then
             Call fps_Ship_Click(0, 1, i)
         End If
     Next
     End With
     
End Sub


Private Sub cbCustomerCode_Mod_Click()

    If cbCustomerCode_Mod.text = "37" Then
        CbItem_Mod.Clear
        CbItem_Mod.AddItem ("修改出货日期")
        CbItem_Mod.AddItem ("修改出货数量")
        CbItem_Mod.AddItem ("删除")
        CbItem_Mod.AddItem ("查询所有修改记录")
        Toolbar2.Buttons(3).Enabled = True
        Label3.Caption = "D.N:"
    ElseIf cbCustomerCode_Mod.text = "SG005" Then
        CbItem_Mod.Clear
        CbItem_Mod.AddItem ("删除")
        Toolbar2.Buttons(3).Enabled = False
        Label3.Caption = "S.O:"
        Label2.Visible = False
        dtShipDate_mod.Visible = False
        txtShipDate_Mod.Visible = False
    End If
    
End Sub

Private Sub Chk_All_Click()

    Dim i As Integer
    
    If Chk_All.Value = 1 Then

         With Fps_Ship(0)
             For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                If Val(.Value) = 0 Then
                    Call fps_Ship_Click(0, 1, i)
                End If
             Next i
         End With

    
        
    ElseIf Chk_All.Value = 0 Then

         With Fps_Ship(0)
             For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                If Val(.Value) = 1 Then
                    Call fps_Ship_Click(0, 1, i)
                End If
             Next i
         End With
        
    End If
    
    
    
End Sub

Private Sub Cob_Change(Index As Integer)
Dim strSql As String
Dim rs     As New ADODB.Recordset
Dim i As Integer


 Select Case Index
 Case 0
    If Cob(0).text = "37" Then
        LoadDn
        lbl(20).Visible = True 'dn
        lbl(21).Visible = True '出货日期
        dtShipDate_Ship.Visible = True '出货日期
        Cob(8).Visible = True 'dn
        Frame7.Visible = True
        Frame9.Visible = True
        lbl(22).Visible = False '客户机种
        lbl(17).Visible = False 'so
        Cob(10).Visible = False '客户机种
        Cob(11).Visible = False 'so
        lbl(23).Visible = False
        lbl(24).Visible = False
        Cob(10).text = ""
        Cob(11).text = ""
        strsono = ""
        strsoline = ""
        lsolineqty = ""
    ElseIf Cob(0).text = "SG005" Then
        LoadSO
        
        lbl(20).Visible = False 'dn
        lbl(21).Visible = True '出货日期
        dtShipDate_Ship.Visible = False '出货日期
        Cob(8).Visible = False 'dn
        Frame7.Visible = False
        Frame9.Visible = False
        lbl(22).Visible = True '客户机种
        lbl(17).Visible = True 'so
        Cob(10).Visible = True '客户机种
        Cob(11).Visible = True 'so
        lbl(23).Visible = True
        lbl(24).Visible = True
              
        strSql = "select distinct DEVICE  from ERPBASE..tblCustomerShippingUp_So "
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        Cob(10).Clear
        If Not rs.EOF Then
            For i = 1 To rs.RecordCount
                Cob(10).AddItem Trim$("" & rs!device)
                rs.MoveNext
            Next
        End If
        rs.Close
       
    
    End If
 Case 10
 
    If Cob(0).text = "37" Then
        LoadDn
    ElseIf Cob(0).text = "SG005" Then
        LoadSO
    End If
 Case 11
    If Index = 11 Then
        If Trim(Cob(11).text) <> "" Then
            strsono = Split(Trim(Cob(11).text), "#")(0)
            strsoline = Split(Trim(Cob(11).text), "#")(1)
            lsolineqty = Val(Split(Trim(Cob(11).text), "#")(2))
        End If
    End If
  End Select

    

End Sub

Private Sub Cob_Click(Index As Integer)
Dim strSql As String
Dim rs     As New ADODB.Recordset
Dim i As Integer

 Select Case Index
 Case 0
    If Cob(0).text = "37" Then
        LoadDn
        lbl(20).Visible = True 'dn
        lbl(21).Visible = True '出货日期
        dtShipDate_Ship.Visible = True '出货日期
        Cob(8).Visible = True 'dn
        Frame7.Visible = True
        Frame9.Visible = True
        lbl(22).Visible = False '客户机种
        lbl(17).Visible = False 'so
        Cob(10).Visible = False '客户机种
        Cob(11).Visible = False 'so
        lbl(23).Visible = False
        lbl(24).Visible = False
        Cob(10).text = ""
        Cob(11).text = ""
        strsono = ""
        strsoline = ""
        lsolineqty = 0
    ElseIf Cob(0).text = "SG005" Then
        LoadSO
        
        lbl(20).Visible = False 'dn
        lbl(21).Visible = True '出货日期
        dtShipDate_Ship.Visible = True '出货日期
        Cob(8).Visible = False 'dn
        Frame7.Visible = False
        Frame9.Visible = False
        lbl(22).Visible = True '客户机种
        lbl(17).Visible = True 'so
        Cob(10).Visible = True '客户机种
        Cob(11).Visible = True 'so
        lbl(23).Visible = True
        lbl(24).Visible = True
              
        strSql = "select distinct DEVICE  from ERPBASE..tblCustomerShippingUp_So "
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        Cob(10).Clear
        If Not rs.EOF Then
            For i = 1 To rs.RecordCount
                Cob(10).AddItem Trim$("" & rs!device)
                rs.MoveNext
            Next
        End If
        rs.Close
    End If
 Case 10
 
    If Cob(0).text = "37" Then
        LoadDn
    ElseIf Cob(0).text = "SG005" Then
        LoadSO
    End If
 Case 11
    If Index = 11 Then
        If Trim(Cob(11).text) <> "" Then
            strsono = Split(Trim(Cob(11).text), "#")(0)
            strsoline = Split(Trim(Cob(11).text), "#")(1)
            lsolineqty = Val(Split(Trim(Cob(11).text), "#")(2))
        End If
    End If
  End Select

End Sub


Private Sub Cob_DblClick(Index As Integer)
Dim strSql As String
Dim rs     As New ADODB.Recordset
Dim i As Integer


 Select Case Index
 Case 0
    If Cob(0).text = "37" Then
        LoadDn
        lbl(20).Visible = True 'dn
        lbl(21).Visible = True '出货日期
        dtShipDate_Ship.Visible = True '出货日期
        Cob(8).Visible = True 'dn
        Frame7.Visible = True
        Frame9.Visible = True
        lbl(22).Visible = False '客户机种
        lbl(17).Visible = False 'so
        Cob(10).Visible = False '客户机种
        Cob(11).Visible = False 'so
        lbl(23).Visible = False
        lbl(24).Visible = False
        Cob(10).text = ""
        Cob(11).text = ""
        strsono = ""
        strsoline = ""
        lsolineqty = 0
    ElseIf Cob(0).text = "SG005" Then
        LoadSO
        
        lbl(20).Visible = False 'dn
        lbl(21).Visible = True '出货日期
        dtShipDate_Ship.Visible = True '出货日期
        Cob(8).Visible = False 'dn
        Frame7.Visible = False
        Frame9.Visible = False
        lbl(22).Visible = True '客户机种
        lbl(17).Visible = True 'so
        Cob(10).Visible = True '客户机种
        Cob(11).Visible = True 'so
        lbl(23).Visible = True
        lbl(24).Visible = True
              
        strSql = "select distinct DEVICE  from ERPBASE..tblCustomerShippingUp_So "
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        Cob(10).Clear
        If Not rs.EOF Then
            For i = 1 To rs.RecordCount
                Cob(10).AddItem Trim$("" & rs!device)
                rs.MoveNext
            Next
        End If
        rs.Close
       
    
    End If
 Case 10
 
    If Cob(0).text = "37" Then
        LoadDn
    ElseIf Cob(0).text = "SG005" Then
        LoadSO
    End If
 Case 11
    If Index = 11 Then
        If Trim(Cob(11).text) <> "" Then
            strsono = Split(Trim(Cob(11).text), "#")(0)
            strsoline = Split(Trim(Cob(11).text), "#")(1)
            lsolineqty = Val(Split(Trim(Cob(11).text), "#")(2))
        End If
    End If
  End Select

End Sub
Private Sub GetOptSelection()
    Dim i As Integer
    Dim j As Integer
    Dim strBigbox As String
    Dim strqty As String
    Dim strBigbox_temp As String
    Dim strBb_temp As String
    Dim lQty_temp As Long
    Dim lQty_select As Long


    strBigbox_temp = ""
    strBb_temp = ""
    strBigbox = ""
    lQty_select = 0
    lQty_temp = 0
        
    With Fps_Ship(0)
    '1.统计每个大箱的数量
        For i = 1 To .MaxRows
            .Row = i
            .Col = FpsM.e_BigX
            If InStr(strBigbox, Trim(.text)) = 0 Then
                strBigbox = strBigbox & "," & Trim(.text)
            End If
        Next
        '找出最优组合
        For i = 0 To UBound(Split(strBigbox, ","))
            strBigbox_temp = Split(strBigbox, ",")(i)
            If strBigbox_temp <> "" Then
                lQty_temp = 0
                For j = 1 To .MaxRows
                    .Row = j
                    .Col = FpsM.e_BigX
                    If Trim(.text) = strBigbox_temp Then
                        .Col = FpsM.e_NUM
                        lQty_temp = lQty_temp + Val(.text)
                    End If
                Next
                If lQty_select + lQty_temp <= lsolineqty Then
                   lQty_select = lQty_select + lQty_temp
                   For j = 1 To .MaxRows
                      .Row = j
                      .Col = FpsM.e_BigX
                      strBb_temp = Trim(.text)
                      .Col = 1
                      If strBb_temp = strBigbox_temp And Val(.Value) = 0 Then
                          Call fps_Ship_Click(0, 1, j)
                      End If
                   Next
                End If
            End If
        Next

     '   For i = 1 To .MaxRows
     '       .Row = i
     '       .Col = FpsM.E_CHOOSE
     '       If Val(.Value) = 0 Then
     '           .Col = FpsM.e_BigX
     '           If Trim(.text) <> strBigbox Then
     '               strBigbox=
     '               For j = i To .MaxCols
      '
                        
     '               Next
                    
                
                
                
            
          '  End If
            
        
        
        'Next
        
        
        
        
    End With
    

End Sub





Private Sub showdata_shiplist(cust As String, ordernolist As String, ShipDate As String)
Dim rs      As New ADODB.Recordset
Dim i As Integer
Dim strSql As String
Dim strDN As String
Dim j As Integer
Dim strTemp As String
If Trim(ordernolist) = "" Then
    Exit Sub
End If
If UCase(cust) = "37" Then
    
       strSql = "  SELECT CUST_DEVICE AS 客户机种,  Quality as   '数量(ea)', HT_DEVICE AS 厂内机种 ,dn as  'DN(#)',ShipOrder as 单据号 ,case BOND when 'A' THEN '保税' else '非保'end as  保税非保 ,SHIP_DATE as 出货日期 ,isnull(remark1,'') AS 快递,isnull(remark2,'') AS 出货地  " & _
                 "  from erpdata..tblShipOrder_Dn where shiporder in ('" & ordernolist & "') order by shiporder "
                
                
                
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        If rs.RecordCount > 0 Then
            With Fps_Ship(2)
               
                .MaxRows = 0
                 Set .DataSource = rs
                 
                strTemp = "<html><body><table  border=1 cellpadding=10>"
                For i = 0 To .MaxRows
                    
                    .Row = i
                    If i = 0 Then
                        strTemp = strTemp & "<tr bgcolor = #E6E6FA>"
                    Else
                        strTemp = strTemp & "<tr>"
                    End If
                    For j = 1 To .MaxCols
                        .Col = j
                        strTemp = strTemp & add_tr(Trim(.text))
                    Next
                    strTemp = strTemp & "</tr>"
                Next
                strTemp = strTemp & "</table></br></br>" '空行
                
               
            End With
            
            strSql = "  SELECT CUST_DEVICE AS 客户机种,  Quality as   '数量(ea)', HT_DEVICE AS 厂内机种 ,dn as  'DN(#)',ShipOrder as 单据号 ,case BOND when 'A' THEN '保税' else '非保'end as  保税非保 ,SHIP_DATE as 出货日期 ,isnull(remark1,'') AS 快递 ,isnull(remark2,'') AS 出货地 " & _
                     "  from erpdata..tblShipOrder_Dn where isnull(dn,'')<>'' and SHIP_DATE='" & ShipDate & "' order by shiporder "
            If rs.State = adStateOpen Then rs.Close
            rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
            If rs.RecordCount > 0 Then
                strTemp = strTemp & "以下为" & ShipDate & "出货汇总 <p></p><Table  border=1 cellpadding=10>"
                rs.MoveFirst
                '标题
                strTemp = strTemp & "<tr bgcolor = #E6E6FA >"
                For j = 0 To rs.Fields.count - 1
                    strTemp = strTemp & add_tr(rs.Fields(j).name)
                Next
                strTemp = strTemp & "</tr>"
                For i = 1 To rs.RecordCount
                    strTemp = strTemp & "<tr >"
                    For j = 0 To rs.Fields.count - 1
        
                        strTemp = strTemp & add_tr(rs.Fields(j))
  
                    Next
                    strTemp = strTemp & "</tr>"
                    rs.MoveNext

                Next
                
                strSql = "SELECT sum(Quality) as 合计数量 from erpdata..tblShipOrder_Dn where isnull(dn,'')<>'' and SHIP_DATE='" & ShipDate & "'"
                strTemp = strTemp & "<tr  border=1><th>合计</th><th> " & Get_SqlserverNo(strSql) & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
            End If
        
            strTemp = strTemp & "</table></body></html>"
        
        Call SentMesToStock(strTemp)
        
        End If
        
        
        
    ElseIf UCase(cust) = "SG005" Then
        strSql = "  SELECT a.HT_DEVICE AS 机种名 ,a.ShipOrder as 单据号,a.PCSNUM as 出货片数, a.Quality as  出货数量,case a.BOND when 'A' THEN '保税' else '非保'end as  保税非保 , a.SO_NO,a.SO_LINE ,a.remark1 as TERM,a.remark2 as 货代 ,case when b.库房名称 like '%不良%' then '不良' else '良品' end as '良品/不良'" & _
                 "  from erpdata..tblShipOrder_Dn a , erpdata..tblstock b where a.stockid=b.库房代码 and a.shiporder in ('" & ordernolist & "') order by shiporder"
   
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        If rs.RecordCount > 0 Then
            With Fps_Ship(2)
               
                .MaxRows = 0
                 Set .DataSource = rs
                strTemp = "<html><body><table  border=1 cellpadding=10>"
                For i = 0 To .MaxRows
                    If i = 0 Then
                        strTemp = strTemp & "<tr bgcolor = #E6E6FA >"
                    Else
                        strTemp = strTemp & "<tr>"
                    End If
                    .Row = i
                    For j = 1 To .MaxCols
                        .Col = j
                        strTemp = strTemp & add_tr(Trim(.text))
                    Next
                    strTemp = strTemp & "</tr>"
                Next
                strTemp = strTemp & "</table></body></html>"
               Call SentMesToStock(strTemp)
            End With
        End If
    End If

End Sub


Private Function add_tr(strTemp As String)
    add_tr = "<th>  " & strTemp & "  </th>"
End Function


Private Sub dtShipDate_Change()
txtShipDate.text = dtShipDate.Value

End Sub

Private Sub dtShipDate_Click()
txtShipDate.text = dtShipDate.Value

End Sub

Private Sub dtShipDate_CloseUp()
txtShipDate.text = dtShipDate.Value

End Sub


Private Sub dtShipDate_mod_CloseUp()
txtShipDate_Mod.text = dtShipDate_mod.Value

End Sub

Private Sub dtShipDate_mod_Change()
txtShipDate_Mod.text = dtShipDate_mod.Value

End Sub

Private Sub dtShipDate_mod_Click()
txtShipDate_Mod.text = dtShipDate_mod.Value
End Sub




Private Sub dtShipDate_Ship_Change()
TxtShipDate_Ship.text = dtShipDate_Ship.Value
LoadDn
End Sub



Private Sub dtShipDate_Ship_Click()
TxtShipDate_Ship.text = dtShipDate_Ship.Value
LoadDn
End Sub

Private Sub dtShipDate_Ship_CloseUp()
TxtShipDate_Ship.text = dtShipDate_Ship.Value
LoadDn
End Sub

Private Sub Form_Load()
InitCtrls
InitCtrl_Ship
End Sub

Private Sub InitCtrls()
Dim i As Integer

dtShipDate.Value = Format(Now(), "yyyy-MM-dd")
dtShipDate_mod.Value = Format(Now(), "yyyy-MM-dd")
dtShipDate_Ship.Value = Format(Now(), "yyyy-MM-dd")

cbCustomerCode.text = "37"
With fps(0)

    .Col = -1
    .Row = -1
    .Lock = True
    .TypeMaxEditLen = 500

End With

With fps(1)
    .TypeMaxEditLen = 500

    .Col = -1
    .Row = -1
    .Lock = True
    
    .SetText 1, 0, "DN"
    .SetText 2, 0, "QTY"
    .SetText 3, 0, "保税标记"
    .SetText 4, 0, "出货日期"
    
    .ColWidth(1) = 10
    .ColWidth(2) = 10
    .ColWidth(3) = 15

End With

    'Fps初始化
    With fpS_Mod(0)
        .ReDraw = False
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        .MaxRows = 0
        .ColsFrozen = 2
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        '设定列类型
        .Col = 1  '选择
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        '设定列宽
        .ColWidth(-1) = 10
        .ColWidth(1) = 4
        .RowHeight(-1) = 10
        '设定是否排序
        .UserColAction = UserColActionSort
        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
        .ZOrder
        .ReDraw = True
    End With
    

End Sub





Private Sub fpS_Mod_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i           As Long
Dim j           As Integer
Dim strTmp      As String

If Row < 1 Then Exit Sub
If Col <> 1 Then Exit Sub

Select Case CbItem_Mod.text


Case "修改出货日期"


    With fpS_Mod(0)

        .Col = 1
        .Row = Row
        .Value = Abs(Val(.Value) - 1)

        If Val(.Value) = 1 Then
            '将所有一样的DN号选择上
            .Col = 2
            .Row = Row
            strTmp = Trim$(.text)
            For i = 1 To .MaxRows
                .Row = i
                .Col = 2
                If Trim$(.text) = strTmp Then
                    .Col = 1
                    .Value = 1
                    .Col = -1
                    .ForeColor = &HFF8080
                End If
            Next
        Else
            '将所有一样的DN号选择上
            .Col = 2
            .Row = Row
            strTmp = Trim$(.text)
            For i = 1 To .MaxRows
                .Row = i
                .Col = 2
                If Trim$(.text) = strTmp Then
                    .Col = 1
                    .Value = 0
                    .Col = -1
                    .ForeColor = vbBlack
                End If
            Next
        End If
        
    End With
    
    
    



Case "修改出货数量", "删除"

    With fpS_Mod(0)

        .Col = 1
        .Row = Row
        .Value = Abs(Val(.Value) - 1)

        If Val(.Value) = 1 Then
            .Col = -1
            .ForeColor = &HFF8080

        Else
            .Col = -1
            .ForeColor = vbBlack

        End If
    End With

End Select




End Sub











Private Sub Fps_Ship_del_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i               As Long
Dim j               As Integer
Dim strorder      As String
Dim rs          As New ADODB.Recordset



    'Fps点击事件
    If Index <> 0 Then Exit Sub
    If Row < 1 Then Exit Sub
    If Col <> 1 Then Exit Sub
    
    With Fps_Ship_del(0)
        .Row = Row
        .Col = 1 '点击勾选列
        .Value = Abs(Val(.Value) - 1)
        For i = 1 To .MaxRows
             .Row = i
             .Col = 1
             If Val(.Value) = 1 Then
                 .Col = 4
                 If strorder = "" Then
                     strorder = Trim(.text)
                 Else
                     strorder = strorder & "," & Trim(.text)
                 End If
             End If
        Next
    End With
    Set rs = Get_SqlserveRs("select * from erpdata..tblstocksqfhsub where 单据编号 in ('" & Replace(strorder, ",", "','") & "') order by 单据编号 ,单据项次 ")

    With Fps_Ship_del(1)
        .MaxRows = 0
        Set .DataSource = rs
        
    End With
 

        
        
   
End Sub

Private Sub Opt1_Click()
ListOrderNumData
End Sub

Private Sub Opt2_Click()
ListOrderNumData
End Sub

Private Sub Opt3_Click()
ListOrderNumData
End Sub




Private Sub ListOrderNumData()

Dim strSql As String
Dim rs As ADODB.Recordset

     
  strSql = "select distinct  '0' as 选择 ,s.客户代码,s.客户名称,t.单据编号 , " & _
        " isnull(关联单据,'') as 关联单据 from erpdata..tblStockSQfh t inner join erpbase..tblXCustomer s on s.客户代码=t.客户代码 where t.接收标记=0  " & _
        " and t.领料员='" & txt(1).text & "' and t.单据类型=1 "
        
       
    If Opt1.Value = True Then '已审核
    
     strSql = strSql & " and isnull(t.财务审核,'')<>''"
            
    End If
           
    If Opt2.Value = True Then '未审核
    
     strSql = strSql & " and isnull(t.财务审核,'')=''"
     
    End If
    
       

    strSql = strSql & " order by s.客户代码,t.单据编号"
    Set rs = Get_SqlserveRs(strSql)
    
    With Fps_Ship_del(1)
        .MaxRows = 0
        
    End With
    
    With Fps_Ship_del(0)
        .MaxRows = 0
        Set .DataSource = rs
        
    End With

 

End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
   control_resize
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)


control_resize
End Sub






Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "QUERY"
        QueryData

    Case "OPEN"
        openData

    Case "SAVE"
        If cbCustomerCode.text = "37" Then
            SaveData
        ElseIf cbCustomerCode.text = "SG005" Then
            Call uploadSO(txtFileName.text)
        End If

    Case "DEL"
        delData

    Case "HOME"
        exitFrm

End Select

End Sub

Private Sub QueryData()

If cbCustomerCode.text = "" Then
    MsgBox "请选择客户代码", vbInformation, "提示"
    Exit Sub

End If

Select Case cbCustomerCode.ListIndex

    Case 0
        queryData_37

End Select

End Sub

Private Sub queryData_37()

    Dim strDN  As String
    Dim strSql As String
    Dim rs     As New ADODB.Recordset

    strDN = Trim$(txtDN.text)

    If strDN = "" Then
        If txtShipDate.text = "" Then
            strSql = "select a.lastupdatedate as 出货日期,a.* from CUSTOMERSHIPPINGUPTBL a order by a.id desc "
        Else
            strSql = "select a.lastupdatedate as 出货日期,a.* from CUSTOMERSHIPPINGUPTBL a where a.lastupdatedate = '" & Trim(txtShipDate.text) & "'  order by a.id desc "

        End If

    Else
        strSql = "select a.lastupdatedate as 出货日期,a.* from CUSTOMERSHIPPINGUPTBL a where delivery = '" & strDN & "' order by a.id desc "

    End If

    'Fps
'    Set rs = Get_OracleRs(strSql)
'
'    If Not rs.EOF Then
'
'        With Fps(0)
'            .MaxRows = 0
'            Set .DataSource = rs
'            Toolbar1.Buttons(5).Enabled = False
'
'        End With
'
'    Else
'        MsgBox "查询不到数据", vbInformation, "提示"
'        Exit Sub
'
'    End If

    'excel
   ExporToExcel (strSql)

End Sub

Private Sub openData()

If openFile Then
    If cbCustomerCode.text = "37" Then
    ShowData
 
    End If
       Toolbar1.Buttons(5).Enabled = True
End If

End Sub

Private Function openFile() As Boolean

On Error GoTo openFile_Err

openFile = False
CommonDialog1.ShowOpen

If CommonDialog1.filename = "" Then Exit Function
txtFileName.text = Replace(CommonDialog1.filename, Chr(0), ",")
CommonDialog1.filename = ""
openFile = True
Exit Function
openFile_Err:
MsgBox Err.DESCRIPTION & vbCrLf & "in 正式工程1.Frm_uploadShippingList.openFile ", vbExclamation + vbOKOnly, "Application Error"

Resume Next

End Function

Private Sub ShowData()
Dim strFileName() As String
Dim i             As Integer

fps(0).MaxRows = 0
fps(1).MaxRows = 0

If InStr(txtFileName.text, ",") > 0 Then
    strFileName = Split(Trim$(txtFileName.text), ",")

    For i = 1 To UBound(strFileName)
        Call ShowFps(strFileName(0) & "\" & strFileName(i))
    Next
Else
    ReDim strFileName(0)
    strFileName(0) = Trim$(txtFileName.text)
    Call ShowFps(strFileName(0))

End If

End Sub

Private Sub ShowFps(strFileName As String)

    On Error GoTo showFps_ErrON

    Dim VBExcel As Excel.Application
    Dim xlBook  As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim i       As Integer
    Dim j       As Integer
    Dim strDN   As String
    Dim strJob  As String
    Dim strShipDate As String
    Dim lQty    As Long
    Dim strChar As String
    
    Set VBExcel = CreateObject("excel.application")
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(strFileName)
    Set xlSheet = xlBook.Worksheets(1)
    
    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 58 Then
        MsgBox "DN列数不对", vbInformation, "警告"
        GoTo showFps_Err
                        
    End If

    With fps(0)

        For i = 1 To xlSheet.Range("A1").CurrentRegion.Rows.count

            If i <> 1 Then .MaxRows = .MaxRows + 1

            For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count

                If j > 26 Then
                    strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
                Else
                    strChar = Chr(96 + j)

                End If

                If i = 1 Then
                    If fps(0).MaxRows = 0 Then
                        .SetText j, .MaxRows, Trim$(Replace(Replace(Replace(xlSheet.Range(strChar & i).Value, ",", " "), "，", " "), "'", " "))

                    End If

                Else
                    .SetText j, .MaxRows, Trim$(Replace(Replace(Replace(xlSheet.Range(strChar & i).Value, ",", " "), "，", " "), "'", " "))

                End If

                If i > 1 Then
                    If j = 1 Then
                        strDN = Trim$(xlSheet.Range(strChar & i))

                        If strDN = "" Then
                            MsgBox "DN不可以为空", vbCritical, "警告"
                            GoTo showFps_Err
                        
                        End If
                    
                    ElseIf j = 10 Then
                        strShipDate = Trim$(xlSheet.Range(strChar & i))
                    
                    ElseIf j = 32 Then
                        strJob = Trim$(xlSheet.Range(strChar & i))
                    
                    ElseIf j = 33 Then
                        lQty = CLng(Trim$(xlSheet.Range(strChar & i)))

                    End If

                End If

            Next j

            If i > 1 Then
                If isNewDN(strDN) = True Then
                    If addDNQty(strDN, strJob, lQty, strShipDate) = False Then
                        GoTo showFps_Err

                    End If

                Else

                    If upDateDNQty(strDN, strJob, lQty, strShipDate) = False Then
                        GoTo showFps_Err

                    End If

                End If

            End If

        Next i

    End With

    If Not VBExcel Is Nothing Then
        VBExcel.Application.DisplayAlerts = False '关闭文档不弹出提示框
        xlBook.Close
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing

    End If

    Exit Sub
showFps_Err:
    fps(0).MaxRows = 0
    fps(1).MaxRows = 0

    If Not VBExcel Is Nothing Then
        VBExcel.Application.DisplayAlerts = False '关闭文档不弹出提示框
        xlBook.Close
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing

    End If

    Exit Sub
showFps_ErrON:
    GoTo showFps_Err
    MsgBox Err.DESCRIPTION & vbCrLf & "in 正式工程1.Frm_uploadShippingList.showFps", vbExclamation + vbOKOnly, "Application Error"

End Sub

Private Function isNewDN(strDN As String) As Boolean
Dim i As Integer

isNewDN = True

With fps(1)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1

        If .text = strDN Then
            isNewDN = False

        End If

    Next

End With

End Function

Private Function checkDNHistory(strDN As String) As Boolean
Dim strSql As String

checkDNHistory = False
strSql = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "'"

If Get_OracleCnt(strSql) > 0 Then
    MsgBox "之前已经上传过DN: " & strDN & vbCrLf & "本次禁止上传, 请重选打开正确的DN文件", vbExclamation, "警告"

    Exit Function
End If

checkDNHistory = True

End Function

Private Function upDateDNQty(strDN As String, strJob As String, lQty As Long, strShipDate As String) As Boolean
Dim i       As Integer
Dim strBand As String

upDateDNQty = False
strBand = getBandFlag(strJob)

If strBand = "" Then

    'MsgBox "JOBID: " & strJob & "在上传WO时没有维护保税标记. 请联系IT", vbInformation, "提示"
    ' Exit Function
End If

With fps(1)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1

        If .text = strDN Then
            .Col = 2
            .text = CLng(.text) + lQty
            
            .Col = 3
            If .text <> strBand Then
                MsgBox "DN:" & strDN & "存在保税非保税混合JOB, 请确认, 结束本次上传", vbCritical, "警告"
                Exit Function

            End If
            
            .Col = 4
            If .text <> strShipDate Then
                MsgBox "DN:" & strDN & "存在多笔出货日期, 请确认, 结束本次上传", vbCritical, "警告"
                Exit Function

            End If
            
        End If
        
    Next
    
End With

upDateDNQty = True

End Function

Private Function getBandFlag(strJob As String) As String
Dim strBand    As String
Dim strSql     As String
Dim strWaferID As String
Dim strWOID    As String

strJob = Replace$(strJob, "M", "")
'strSql = "select replace(aa.substrateid, '+','') as waferid from mappingdatatest aa  inner join customeroitbl_test bb on to_char(bb.id) = aa.filename and aa.lotid = bb.source_batch_id where bb.test_mtrl_desc = '" & strJob & "'  "
'strWaferID = Get_OracleStr(strSql)
'strSql = "select substratetype from mappingdatatest where substrateid = '" & strWaferID & "' and substratetype is not null "
'strBand = Get_OracleStr(strSql)
'
'If strBand = "" Then
'    strSql = "select distinct jobno from customeroitbl_test where test_mtrl_desc in (select source_mtrl_SLOC from customeroitbl_test where test_mtrl_desc = '" & strJob & "')"
'    If Get_OracleStr(strSql) = "A" Then
'        getBandFlag = "保税"
'    ElseIf Get_OracleStr(strSql) = "B" Then
'        getBandFlag = "非保税"
'    Else
'        getBandFlag = ""
'    End If
'Else
'
'    If strBand = "0" Then
'        getBandFlag = "非保税"
'    Else
'        getBandFlag = "保税"
'
'    End If
'
'End If

strSql = "select distinct cc.ordername from customeroitbl_test aa " & _
" inner join mappingdatatest bb on to_char(aa.id) = bb.filename and aa.source_batch_id = bb.lotid " & _
" inner join ib_waferlist cc on bb.substrateid = cc.waferid and bb.lotid = cc.waferlot " & _
" where aa.test_mtrl_desc = '" & strJob & "' "

strWOID = Trim("" & Get_OracleStr(strSql))

If Left$(strWOID, 1) = "A" Then
    getBandFlag = "保税"

ElseIf Left$(strWOID, 1) = "B" Then
    getBandFlag = "非保税"

Else
    'MsgBox "查询不到保税非保税工单", vbCritical, "警告"
getBandFlag = "保税"
End If



End Function

Private Function addDNQty(strDN As String, strJob As String, lQty As Long, strShipDate As String) As Boolean
Dim i       As Integer
Dim strBand As String

addDNQty = False
strBand = getBandFlag(strJob)

If strBand = "" Then

    'MsgBox "JOBID: " & strJob & "在上传WO时没有维护保税标记. 请联系IT", vbInformation, "提示"
    ' Exit Function
End If

With fps(1)
    .MaxRows = .MaxRows + 1
    i = .MaxRows
    .SetText 1, i, strDN
    .SetText 2, i, lQty
    .SetText 3, i, strBand
    .SetText 4, i, strShipDate

End With

addDNQty = True

End Function

Private Sub SaveData()
Dim i          As Integer
Dim strDN      As String
Dim strDNQuery As String

If CheckDC = False Then
    MsgBox "DateCode检查未通过, 请先修改AZ列的DateCode,再重新上传保存", vbInformation, "提示"
    Exit Sub
End If

strDNQuery = ""

If cbCustomerCode.text = "" Then
    MsgBox "请选择客户代码", vbInformation, "提示"
    Exit Sub

End If

If fps(0).MaxRows = 0 Then
    MsgBox "请打开需要上传的文件", vbInformation, "提示"
    Exit Sub

End If

Select Case cbCustomerCode.ListIndex

    Case 0  '37

        With fps(1)

            For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                strDN = Trim$(.text)

                If checkDNHistory(strDN) = False Then
                    fps(0).MaxRows = 0
                    fps(1).MaxRows = 0
                    Exit Sub
                Else

                    If getData(strDN) = True Then
                        strDNQuery = strDNQuery & strDN & "','"
                        MsgBox "DN:" & strDN & "上传成功", vbInformation, "提示"

                    End If

                End If

            Next

        End With

        If strDNQuery <> "" Then
            strDNQuery = Left$(strDNQuery, Len(strDNQuery) - 3)
            Call exportSuccessDN(strDNQuery)

        End If

End Select

Toolbar1.Buttons(5).Enabled = False

End Sub

Private Sub exportSuccessDN(strDN As String)
Dim strSql As String

strSql = "select a.lastupdatedate as 出货日期,a.delivery as DN, sum(a.quantity) as 上传总数 from CUSTOMERSHIPPINGUPTBL a where a.delivery in ('" & strDN & "') group by a.lastupdatedate,a.delivery order by a.lastupdatedate desc "
ExporToExcel (strSql)

End Sub

Private Function getData(strDN As String) As Boolean

On Error GoTo getData_ErrON

Dim tyDN As DN_DETAILS
Dim i    As Integer
Dim j    As Integer

getData = False
Cnn.BeginTrans
INIadoCon.BeginTrans

With fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        
        If .text = "" Then
            MsgBox "数据出错", vbCritical, "警告"
            GoTo getData_Err
        End If
        
        If .text = strDN Then
            .Col = 1
            tyDN.Delivery = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 2
            tyDN.ItemNo = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 3
            tyDN.DeliveryCreationDate = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 4
            tyDN.Plant = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 5
            tyDN.SalesDocument = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 6
            tyDN.SOItemNo = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 7
            tyDN.Material = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 8
            tyDN.MarketingPN = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 9
            tyDN.MaterialDescription = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 10
            tyDN.PlannedGIdate = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 11
            tyDN.CustomerPartnumber = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 12
            tyDN.ShiptoName = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 13
            tyDN.ShiptoCustomer = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 14
            tyDN.PurchasingDocNo = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 15
            tyDN.DateCodeRestrictions = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 16
            tyDN.LabelRequirement = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 17
            tyDN.ReLabelInstructions = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 18
            tyDN.ShipToStreet1 = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 19
            tyDN.ShipToStreet2 = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 20
            tyDN.ShipToStreet3 = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 21
            tyDN.City = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 22
            tyDN.State = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 23
            tyDN.PostalCode = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 24
            tyDN.CountryKey = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 25
            tyDN.ContactName = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 26
            tyDN.Phone = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 27
            tyDN.Fax = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 28
            tyDN.FreightForwarder = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 29
            tyDN.ShippingInstruction = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 30
            tyDN.AdditionalComments = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 31
            tyDN.StorageLocation = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 32
            tyDN.BatchNumber = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 33
            tyDN.Quantity = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 34
            tyDN.VolumeWeight = IIf((.text = ""), "0", Replace(Trim(.text), Chr(13) + Chr(10), ""))
            .Col = 35
            tyDN.GrossWeight = IIf((.text = ""), "0", Replace(Trim(.text), Chr(13) + Chr(10), ""))
            .Col = 36
            tyDN.netweight = IIf((.text = ""), "0", Replace(Trim(.text), Chr(13) + Chr(10), ""))
            .Col = 37
            tyDN.UoMForWeight = IIf((.text = ""), "0", Replace(Trim(.text), Chr(13) + Chr(10), ""))
            .Col = 38
            tyDN.NoOfCartons = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 39
            tyDN.VendorLotNumber = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 40
            tyDN.ShelfLocation = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 41
            tyDN.BOLOrAirwayBillNo = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 42
            tyDN.ActualShippingDate = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 43
            tyDN.PackagingDetails = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 44
            tyDN.PackingStatus = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 45
            tyDN.PickingStatus = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 46
            tyDN.CustomerCalendar = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 47
            tyDN.FatherBatch = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 48
            tyDN.MotherBatch = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 49
            tyDN.FatherBatchDateCode = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 50
            tyDN.MotherBatchDateCode = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 51
            tyDN.TransferOrderStatus = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 52
            tyDN.DATECODE = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 53
            tyDN.FatherBatchQty = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 54
            tyDN.ShippingPoint = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 55
            tyDN.ShipmentNumber = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 56
            tyDN.FabSite = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 57
            tyDN.AssemblySite = Replace(Trim(.text), Chr(13) + Chr(10), "")
            .Col = 58
            tyDN.TestSite = Replace(Trim(.text), Chr(13) + Chr(10), "")
            tyDN.id = GetshippingMaxID()

            If checkDNData(tyDN) = False Or saveDataToDB(tyDN) = False Then
                GoTo getData_Err

            End If

        End If

    Next

End With

Cnn.CommitTrans
INIadoCon.CommitTrans
getData = True
Exit Function

getData_Err:
Cnn.RollbackTrans
INIadoCon.RollbackTrans
Exit Function

getData_ErrON:
GoTo getData_Err
MsgBox Err.DESCRIPTION & ":插入数据库错误", vbExclamation, "警告"

End Function

Private Function checkDNData(tyDN As DN_DETAILS) As Boolean
checkDNData = False
Dim strSql As String
Dim strDC As String

If tyDN.DATECODE = "" Then
    MsgBox "DATECODE不可以为空,请联系客户,确保有值", vbCritical, "警告"
    Exit Function
Else
    If Right(tyDN.BatchNumber, 1) <> "M" And Right(tyDN.BatchNumber, 1) <> "R" Then
        
        strSql = "select dc from tbl37testdc where jobid = '" & tyDN.BatchNumber & "'   "
        
        strDC = Get_OracleStr(strSql)
        If strDC <> "" And strDC <> tyDN.DATECODE Then
           tyDN.DATECODE = strDC
           
            MsgBox "客户DC:" & tyDN.DATECODE & " 和厂内DC" & strDC & "不一致,请联系IT,确保一致", vbCritical, "警告"
            Exit Function

        End If

    End If

End If

If tyDN.FabSite <> "" Or tyDN.AssemblySite <> "" Or tyDN.TestSite <> "" Then
    If InStr(UCase(tyDN.LabelRequirement), "SAMSUNG") = 0 Then
        MsgBox "最后三列SITE有值,必须是37出SAMSUNG的标签", vbCritical, "警告"
        Exit Function

    End If

End If

If (InStr(UCase(tyDN.ShiptoName), "PREMIER") > 0) And (tyDN.CustomerPartnumber <> "") Then
    If InStr(UCase(tyDN.LabelRequirement), "SAMSUNG") = 0 Then
        MsgBox "label requirement必须是SAMSUNG的", vbCritical, "警告"
        Exit Function

    End If

End If

If tyDN.LabelRequirement = "" Then
    MsgBox "label requirement标签类型不可为空", vbCritical, "警告"
    Exit Function

End If

checkDNData = True

End Function

Private Function saveDataToDB(tyDN As DN_DETAILS) As Boolean
Dim strSql  As String
Dim strSql2 As String
Dim strSql3 As String

saveDataToDB = False
strSql = "select * from CUSTOMERSHIPPINGUPTBL where Delivery='" & tyDN.Delivery & "'  and batchnumber='" & tyDN.BatchNumber & "' and flag='Y'"

If Get_OracleCnt(strSql) > 0 Then
    strSql = "update CUSTOMERSHIPPINGUPTBL set quantity =(quantity + '" & tyDN.Quantity & "') where delivery = '" & tyDN.Delivery & "' and batchnumber = '" & tyDN.BatchNumber & "'"
'    strSql2 = "update [ERPBASE].[dbo].[tblCustomerShippingUp] set quantity =(quantity + '" & tyDN.Quantity & "')  where delivery = '" & tyDN.Delivery & "' and batchnumber = '" & tyDN.BatchNumber & "'"

    strSql2 = "insert into [ERPBASE].[dbo].[tblCustomerShippingUp](ID, Delivery, ItemNo, DeliveryCreationDate, Plant, SalesDocument, SOItemNo, Material, MarketingPN, MaterialDescription, PlannedGIDate, CustomerPartNumber, ShipToName, ShipToCustomer, " & _
       " PurchasingDocNo, DateCodeRestrictions, LabelRequirement, ReLabelInstructions, ShipToStreet1, ShipToStreet2, ShipToStreet3, City, State, PostalCode, CountryKey, ContactName, Phone, Fax, " & _
       " FreightForwarder, ShippingInstruction, AdditionalComments, StorageLocation, BatchNumber, Quantity, VolumeWeight, GrossWeight, Netweight, UoMForWeight, NoOfCartons, VendorLotNumber, " & _
       " ShelfLocation, BOLOrAirwayBillNo, ActualShippingDate, PackagingDetails, PackingStatus, PickingStatus, CustomerCalendar, customershortname, FLAG, CREATEDBY, CREATEDDATE, " & _
       " LASTUPDATEBY, LASTUPDATEDATE, FATHER_BATCH, MOTHER_BATCH, FATHER_BATCH_DATE_CODE, MOTHER_BATCH_DATE_CODE, TRANSFER_ORDER_STATUS, DATE_CODE, " & _
       " FATHER_BATCH_QTY, SHIPPING_POINT, SHIPMENT_NUMBER, FAB_SITE, ASSEMBLY_SITE, TEST_SITE)  " & _
       " values( '" & tyDN.id & "','" & tyDN.Delivery & "','" & tyDN.ItemNo & "','" & tyDN.DeliveryCreationDate & "','" & tyDN.Plant & "','" & tyDN.SalesDocument & "','" & tyDN.SOItemNo & "','" & tyDN.Material & "','" & tyDN.MarketingPN & "','" & tyDN.MaterialDescription & "','" & tyDN.PlannedGIdate & "','" & tyDN.CustomerPartnumber & "','" & tyDN.ShiptoName & "','" & tyDN.ShiptoCustomer & "', " & _
       " '" & tyDN.PurchasingDocNo & "','" & tyDN.DateCodeRestrictions & "','" & tyDN.LabelRequirement & "','" & tyDN.ReLabelInstructions & "','" & tyDN.ShipToStreet1 & "','" & tyDN.ShipToStreet2 & "','" & tyDN.ShipToStreet3 & "','" & tyDN.City & "','" & tyDN.State & "','" & tyDN.PostalCode & "','" & tyDN.CountryKey & "','" & tyDN.ContactName & "','" & tyDN.Phone & "','" & tyDN.Fax & "', " & _
       " '" & tyDN.FreightForwarder & "','" & tyDN.ShippingInstruction & "','" & tyDN.AdditionalComments & "','" & tyDN.StorageLocation & "','" & tyDN.BatchNumber & "','" & tyDN.Quantity & "','" & tyDN.VolumeWeight & "','" & tyDN.GrossWeight & "','" & tyDN.netweight & "','" & tyDN.UoMForWeight & "','" & tyDN.NoOfCartons & "','" & tyDN.VendorLotNumber & "', " & _
       " '" & tyDN.ShelfLocation & "','" & tyDN.BOLOrAirwayBillNo & "','" & tyDN.ActualShippingDate & "','" & tyDN.PackagingDetails & "','" & tyDN.PackingStatus & "','" & tyDN.PickingStatus & "','" & tyDN.CustomerCalendar & "','37','Y','" & gUserName & "',GETDATE(), " & _
       " '','" & tyDN.PlannedGIdate & "','" & tyDN.FatherBatch & "','" & tyDN.MotherBatch & "','" & tyDN.FatherBatchDateCode & "','" & tyDN.MotherBatchDateCode & "','" & tyDN.TransferOrderStatus & "','" & tyDN.DATECODE & "','" & tyDN.FatherBatchQty & "','" & tyDN.ShippingPoint & "','" & tyDN.ShipmentNumber & "','" & tyDN.FabSite & "','" & tyDN.AssemblySite & "','" & tyDN.TestSite & "' ) "

    strSql3 = "insert into DNSHIPMENTTBL(ID, Delivery, CUSTOMER_DEVICE,JOB_ID,HT_DEVICE,QUANTITY,SHIP_ORDER,BOND,SHIP_DATE,EXPRESS,REQUEST_FLAG,MAIL_FLAG) " & _
        " values( '" & tyDN.id & "','" & tyDN.Delivery & "','" & tyDN.Material & "','" & tyDN.BatchNumber & "','','" & tyDN.Quantity & "','','" & getBandFlag(tyDN.BatchNumber) & "','" & tyDN.PlannedGIdate & "','',0,0)"
    
    If AddSql(strSql) = 0 Or AddSql2(strSql2) = 0 Or AddSql(strSql3) = 0 Then
        MsgBox "DN: " & tyDN.Delivery & "没有上传成功, 请联系IT确定数据是否有问题", vbCritical, "警告"
        Exit Function

    End If
    
    

Else
    strSql = "insert into CUSTOMERSHIPPINGUPTBL(ID, Delivery, ItemNo, DeliveryCreationDate, Plant, SalesDocument, SOItemNo, Material, MarketingPN, MaterialDescription, PlannedGIDate, CustomerPartNumber, ShipToName, ShipToCustomer, " & _
       " PurchasingDocNo, DateCodeRestrictions, LabelRequirement, ReLabelInstructions, ShipToStreet1, ShipToStreet2, ShipToStreet3, City, State, PostalCode, CountryKey, ContactName, Phone, Fax, " & _
       " FreightForwarder, ShippingInstruction, AdditionalComments, StorageLocation, BatchNumber, Quantity, VolumeWeight, GrossWeight, Netweight, UoMForWeight, NoOfCartons, VendorLotNumber, " & _
       " ShelfLocation, BOLOrAirwayBillNo, ActualShippingDate, PackagingDetails, PackingStatus, PickingStatus, CustomerCalendar, customershortname, FLAG, CREATEDBY, CREATEDDATE, " & _
       " LASTUPDATEBY, LASTUPDATEDATE, FATHER_BATCH, MOTHER_BATCH, FATHER_BATCH_DATE_CODE, MOTHER_BATCH_DATE_CODE, TRANSFER_ORDER_STATUS, DATE_CODE, " & _
       " FATHER_BATCH_QTY, SHIPPING_POINT, SHIPMENT_NUMBER, FAB_SITE, ASSEMBLY_SITE, TEST_SITE)  " & _
       " values( '" & tyDN.id & "','" & tyDN.Delivery & "','" & tyDN.ItemNo & "','" & tyDN.DeliveryCreationDate & "','" & tyDN.Plant & "','" & tyDN.SalesDocument & "','" & tyDN.SOItemNo & "','" & tyDN.Material & "','" & tyDN.MarketingPN & "','" & tyDN.MaterialDescription & "','" & tyDN.PlannedGIdate & "','" & tyDN.CustomerPartnumber & "','" & tyDN.ShiptoName & "','" & tyDN.ShiptoCustomer & "', " & _
       " '" & tyDN.PurchasingDocNo & "','" & tyDN.DateCodeRestrictions & "','" & tyDN.LabelRequirement & "','" & tyDN.ReLabelInstructions & "','" & tyDN.ShipToStreet1 & "','" & tyDN.ShipToStreet2 & "','" & tyDN.ShipToStreet3 & "','" & tyDN.City & "','" & tyDN.State & "','" & tyDN.PostalCode & "','" & tyDN.CountryKey & "','" & tyDN.ContactName & "','" & tyDN.Phone & "','" & tyDN.Fax & "', " & _
       " '" & tyDN.FreightForwarder & "','" & tyDN.ShippingInstruction & "','" & tyDN.AdditionalComments & "','" & tyDN.StorageLocation & "','" & tyDN.BatchNumber & "','" & tyDN.Quantity & "','" & tyDN.VolumeWeight & "','" & tyDN.GrossWeight & "','" & tyDN.netweight & "','" & tyDN.UoMForWeight & "','" & tyDN.NoOfCartons & "','" & tyDN.VendorLotNumber & "', " & _
       " '" & tyDN.ShelfLocation & "','" & tyDN.BOLOrAirwayBillNo & "','" & tyDN.ActualShippingDate & "','" & tyDN.PackagingDetails & "','" & tyDN.PackingStatus & "','" & tyDN.PickingStatus & "','" & tyDN.CustomerCalendar & "','37','Y','" & gUserName & "',sysdate, " & _
       " '','" & tyDN.PlannedGIdate & "','" & tyDN.FatherBatch & "','" & tyDN.MotherBatch & "','" & tyDN.FatherBatchDateCode & "','" & tyDN.MotherBatchDateCode & "','" & tyDN.TransferOrderStatus & "','" & tyDN.DATECODE & "','" & tyDN.FatherBatchQty & "','" & tyDN.ShippingPoint & "','" & tyDN.ShipmentNumber & "','" & tyDN.FabSite & "','" & tyDN.AssemblySite & "','" & tyDN.TestSite & "' ) "
    
    strSql2 = "insert into [ERPBASE].[dbo].[tblCustomerShippingUp](ID, Delivery, ItemNo, DeliveryCreationDate, Plant, SalesDocument, SOItemNo, Material, MarketingPN, MaterialDescription, PlannedGIDate, CustomerPartNumber, ShipToName, ShipToCustomer, " & _
       " PurchasingDocNo, DateCodeRestrictions, LabelRequirement, ReLabelInstructions, ShipToStreet1, ShipToStreet2, ShipToStreet3, City, State, PostalCode, CountryKey, ContactName, Phone, Fax, " & _
       " FreightForwarder, ShippingInstruction, AdditionalComments, StorageLocation, BatchNumber, Quantity, VolumeWeight, GrossWeight, Netweight, UoMForWeight, NoOfCartons, VendorLotNumber, " & _
       " ShelfLocation, BOLOrAirwayBillNo, ActualShippingDate, PackagingDetails, PackingStatus, PickingStatus, CustomerCalendar, customershortname, FLAG, CREATEDBY, CREATEDDATE, " & _
       " LASTUPDATEBY, LASTUPDATEDATE, FATHER_BATCH, MOTHER_BATCH, FATHER_BATCH_DATE_CODE, MOTHER_BATCH_DATE_CODE, TRANSFER_ORDER_STATUS, DATE_CODE, " & _
       " FATHER_BATCH_QTY, SHIPPING_POINT, SHIPMENT_NUMBER, FAB_SITE, ASSEMBLY_SITE, TEST_SITE)  " & _
       " values( '" & tyDN.id & "','" & tyDN.Delivery & "','" & tyDN.ItemNo & "','" & tyDN.DeliveryCreationDate & "','" & tyDN.Plant & "','" & tyDN.SalesDocument & "','" & tyDN.SOItemNo & "','" & tyDN.Material & "','" & tyDN.MarketingPN & "','" & tyDN.MaterialDescription & "','" & tyDN.PlannedGIdate & "','" & tyDN.CustomerPartnumber & "','" & tyDN.ShiptoName & "','" & tyDN.ShiptoCustomer & "', " & _
       " '" & tyDN.PurchasingDocNo & "','" & tyDN.DateCodeRestrictions & "','" & tyDN.LabelRequirement & "','" & tyDN.ReLabelInstructions & "','" & tyDN.ShipToStreet1 & "','" & tyDN.ShipToStreet2 & "','" & tyDN.ShipToStreet3 & "','" & tyDN.City & "','" & tyDN.State & "','" & tyDN.PostalCode & "','" & tyDN.CountryKey & "','" & tyDN.ContactName & "','" & tyDN.Phone & "','" & tyDN.Fax & "', " & _
       " '" & tyDN.FreightForwarder & "','" & tyDN.ShippingInstruction & "','" & tyDN.AdditionalComments & "','" & tyDN.StorageLocation & "','" & tyDN.BatchNumber & "','" & tyDN.Quantity & "','" & tyDN.VolumeWeight & "','" & tyDN.GrossWeight & "','" & tyDN.netweight & "','" & tyDN.UoMForWeight & "','" & tyDN.NoOfCartons & "','" & tyDN.VendorLotNumber & "', " & _
       " '" & tyDN.ShelfLocation & "','" & tyDN.BOLOrAirwayBillNo & "','" & tyDN.ActualShippingDate & "','" & tyDN.PackagingDetails & "','" & tyDN.PackingStatus & "','" & tyDN.PickingStatus & "','" & tyDN.CustomerCalendar & "','37','Y','" & gUserName & "',GETDATE(), " & _
       " '','" & tyDN.PlannedGIdate & "','" & tyDN.FatherBatch & "','" & tyDN.MotherBatch & "','" & tyDN.FatherBatchDateCode & "','" & tyDN.MotherBatchDateCode & "','" & tyDN.TransferOrderStatus & "','" & tyDN.DATECODE & "','" & tyDN.FatherBatchQty & "','" & tyDN.ShippingPoint & "','" & tyDN.ShipmentNumber & "','" & tyDN.FabSite & "','" & tyDN.AssemblySite & "','" & tyDN.TestSite & "' ) "
    
    strSql3 = "insert into DNSHIPMENTTBL(ID, Delivery, CUSTOMER_DEVICE,JOB_ID,HT_DEVICE,QUANTITY,SHIP_ORDER,BOND,SHIP_DATE,EXPRESS,REQUEST_FLAG,MAIL_FLAG) " & _
        " values( '" & tyDN.id & "','" & tyDN.Delivery & "','" & tyDN.Material & "','" & tyDN.BatchNumber & "','','" & tyDN.Quantity & "','','" & getBandFlag(tyDN.BatchNumber) & "','" & tyDN.PlannedGIdate & "','',0,0)"
    
    If AddSql(strSql) = 0 Or AddSql2(strSql2) = 0 Or AddSql(strSql3) = 0 Then
        MsgBox "DN: " & tyDN.Delivery & "没有上传成功, 请联系IT确定数据是否有问题", vbCritical, "警告"
        Exit Function

    End If

End If

saveDataToDB = True

End Function

Private Sub delData()
Dim strDN As String

If txtDN.text = "" Then
    MsgBox "请输入要删除的DN", vbInformation, "提示"
    Exit Sub

End If

strDN = Trim$(txtDN.text)

If Get_OracleCnt("select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "'") = 0 Then
    MsgBox "您输入的DN不正确或没有上传记录,不可删除", vbInformation, "提示"
    Exit Sub

End If
If Get_OracleCnt("  select * from packing_detailed i WHERE i.DN_NUM= '" & strDN & "'") > 0 Then
    MsgBox "外包已开始作业，此DN不可删除,请通知外包", vbInformation, "提示"
    Exit Sub
End If
    
  
AddSql ("insert into CUSTOMERSHIPPINGUPTBL_BAK select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "' ")
MsgBox "DN备份成功", vbInformation, "提示"
AddSql ("delete from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "'")
AddSql ("delete from DNSHIPMENTTBL where delivery = '" & strDN & "'")
AddSql2 ("delete from  [ERPBASE].[dbo].[tblCustomerShippingUp] where Delivery = '" & strDN & "'")
MsgBox "已成功删除DN:" & strDN, vbInformation, "提示"
txtDN.text = ""

End Sub

Private Sub exitFrm()
Unload Me

End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key

    Case "QUERY_MOD"
        QueryData_Mod

    Case "SAVE_MOD"
       '确认修改
        SAVE_Modify
     
    Case "DEL_MOD"
          DEL_Single

    Case "HOME_MOD"
        exitFrm

End Select

End Sub
Private Sub DEL_Single()

Dim strid  As String
Dim strqty  As String
Dim i As Integer
Dim DelCnt As Integer
DelCnt = 0
With fpS_Mod(0)
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1

        If .text <> "" Then
            If .text = 1 Then
                DelCnt = DelCnt + 1
            End If

        End If
    Next i
    If DelCnt = 0 Then
        MsgBox "请选择要删除的数据", vbInformation, "提示"
        Exit Sub
    End If
    If MsgBox("你确认要删除选中的" & DelCnt & "笔数据吗?", vbOKCancel, "提示") = vbCancel Then
        Exit Sub
    End If
    If cbCustomerCode_Mod.text = "37" Then
         For i = 1 To .MaxRows
            .Row = i
            .Col = 1
    
            If .text <> "" Then
                If .text = 1 Then
        
                    .Col = 6 'ID
                    strid = Trim(.text) '
                   
                    AddSql ("insert into CUSTOMERSHIPPINGUPTBL_MODHIS select '" & gUserName & "' ,to_char(SYSDATE,'yyyy-MM-dd HH24:mi:ss'),'删除', a.* from CUSTOMERSHIPPINGUPTBL a where a.ID = '" & strid & "' ")
                    AddSql ("delete from  CUSTOMERSHIPPINGUPTBL where ID= '" & strid & "' ")
                    AddSql ("delete from  DNSHIPMENTTBL where ID= '" & strid & "' ")
                    AddSql2 ("delete from  [ERPBASE].[dbo].[tblCustomerShippingUp] where ID = '" & strid & "' ")
                End If
                
            End If
        Next
    ElseIf cbCustomerCode_Mod.text = "SG005" Then
         For i = 1 To .MaxRows
            .Row = i
            .Col = 1
    
            If .text <> "" Then
                If .text = 1 Then
        
                    .Col = 2 'ID
                    strid = Trim(.text) '
                   
                    AddSql ("insert into CUSTOMERSHIPPING_SO_MODHIS select '" & gUserName & "' ,to_char(SYSDATE,'yyyy-MM-dd HH24:mi:ss'),'删除', a.* from CUSTOMERSHIPPINGUPTBL_SO a where a.ID = '" & strid & "' ")
                    AddSql ("delete from  CUSTOMERSHIPPINGUPTBL_SO where ID= '" & strid & "' ")
                  '  AddSql ("delete from  DNSHIPMENTTBL where ID= '" & strid & "' ")
                    AddSql2 ("delete from  [ERPBASE].[dbo].[tblCustomerShippingUp_So] where ID = '" & strid & "' ")
                End If
                
            End If
        Next
    End If
    QueryData_Mod
    
End With

    
    

End Sub

Private Sub SAVE_Modify()

If CbItem_Mod.text = "修改出货数量" Then
    Quality_Modify
ElseIf CbItem_Mod.text = "修改出货日期" Then
    ShipDate_Modify
ElseIf CbItem_Mod.text = "删除" Then
    MsgBox "删除请点确认删除按钮", vbInformation, "提示"
 
Else
    MsgBox "请选择修改项目", vbInformation, "提示"
End If

End Sub


Private Sub ShipDate_Modify()
'出货日期修改
Dim strid  As String
Dim strqty  As String
Dim i As Integer
Dim rs     As New ADODB.Recordset
Dim strmsg As String
Dim ordernumlist As String
Dim strshipdate_old As String



If txtShipDate_Mod.text = "" Then

    MsgBox "请选择新的出货日期", vbInformation, "提示"
    Exit Sub

End If

strshipdate_old = GetSqlServerStr("select distinct ship_date from erpdata..tblShipOrder_Dn where dn='" & txtDN_Mod.text & "'")

With fpS_Mod(0)

     For i = 1 To .MaxRows
        .Row = i
        .Col = 1

        If .text <> "" Then
            If .text = 1 Then
    
                .Col = 6 'ID
                strid = Trim(.text) '
                
                AddSql ("insert into CUSTOMERSHIPPINGUPTBL_MODHIS select '" & gUserName & "' ,to_char(SYSDATE,'yyyy-MM-dd HH24:mi:ss'),'修改前日期', a.* from CUSTOMERSHIPPINGUPTBL a where a.ID = '" & strid & "' ")
                AddSql ("update  CUSTOMERSHIPPINGUPTBL set PLANNEDGIDATE='" & txtShipDate_Mod.text & "', LASTUPDATEDATE=to_date('" & txtShipDate_Mod.text & "','yyyy/mm/dd')   where ID= '" & strid & "' ")
                AddSql ("update  DNSHIPMENTTBL set SHIP_DATE='" & txtShipDate_Mod.text & "' where ID= '" & strid & "' ")
                AddSql2 ("update   [ERPBASE].[dbo].[tblCustomerShippingUp]  set PLANNEDGIDATE='" & txtShipDate_Mod.text & "', LASTUPDATEDATE='" & txtShipDate_Mod.text & "' where ID= '" & strid & "' ")
                AddSql ("insert into CUSTOMERSHIPPINGUPTBL_MODHIS select '" & gUserName & "' ,to_char(SYSDATE,'yyyy-MM-dd HH24:mi:ss'),'修改后日期', a.* from CUSTOMERSHIPPINGUPTBL a where a.ID = '" & strid & "' ")
                AddSql2 ("update  erpdata..tblShipOrder_Dn set ship_date='" & Format(dtShipDate_mod.Value, "yyyy/mm/dd") & "' where dn='" & txtDN_Mod.text & "'")
                      
            End If
            
        End If
    Next

    
    MsgBox "修改完成"
    QueryData_Mod
End With
 'merry20200422如果此dn已经生成出货单据 , 一并修改出货单据对应的出货日期

Set rs = Get_SqlserveRs("select distinct dn,shiporder,ship_date from  erpdata..tblShipOrder_Dn where dn='" & txtDN_Mod.text & "'")
If rs.RecordCount > 0 Then
    If Get_SqlserverCnt("select distinct ship_date from erpdata..tblShipOrder_Dn where dn='" & txtDN_Mod.text & "'") > 1 Then
        MsgBox "同一个DN对应不同出货时间，邮件发送失败！", vbInformation, "提示"
    Else
        rs.MoveFirst
        For i = 1 To rs.RecordCount
            If ordernumlist = "" Then
                ordernumlist = rs("shiporder")
            Else
                ordernumlist = ordernumlist & "," & rs("shiporder")
            End If
            rs.MoveNext
        Next
        Cob(0).text = "37"
        dtShipDate_Ship.Value = strshipdate_old
        mailflag = 2 '修改时间的格式
        Call showdata_shiplist(Cob(0).text, Replace(ordernumlist, ",", "','"), Format(dtShipDate_Ship.Value, "yyyy/mm/dd"))
        dtShipDate_Ship.Value = txtShipDate_Mod.text
        mailflag = 1 '正常出货的格式
        Call showdata_shiplist(Cob(0).text, Replace(ordernumlist, ",", "','"), Format(dtShipDate_Ship.Value, "yyyy/mm/dd"))
    End If
End If




End Sub



Private Sub Quality_Modify()
'出货数量修改
Dim strid  As String
Dim i As Integer
Dim strqty As Long

With fpS_Mod(0)



     For i = 1 To .MaxRows
        .Row = i
        .Col = 1

        If .text <> "" Then
            If .text = 1 Then
    
                .Col = 6 'ID
                strid = Trim(.text) '
                
                .Col = 5 '数量
                                        
                If IsNumeric(Trim(.text)) = False Then
                
                    MsgBox "出货数量" & strqty & "填写有误，请修正后再提交", vbInformation, "提示"
                    Exit Sub
                End If
                strqty = Trim(.text)
                If strqty <= 0 Then
                
                    MsgBox "出货数量" & strqty & "填写有误，请修正后再提交", vbInformation, "提示"
                    Exit Sub
                End If
                AddSql ("insert into CUSTOMERSHIPPINGUPTBL_MODHIS select '" & gUserName & "' ,to_char(SYSDATE,'yyyy-MM-dd HH24:mi:ss'),'修改前数量', a.* from CUSTOMERSHIPPINGUPTBL a where a.ID = '" & strid & "' ")
                AddSql ("update  CUSTOMERSHIPPINGUPTBL set QUANTITY=" & strqty & " where ID= '" & strid & "' ")
                AddSql ("update  DNSHIPMENTTBL set QUANTITY=" & strqty & " where ID= '" & strid & "' ")
                AddSql2 ("update  [ERPBASE].[dbo].[tblCustomerShippingUp]  set QUANTITY=" & strqty & " where ID= '" & strid & "' ")
                AddSql ("insert into CUSTOMERSHIPPINGUPTBL_MODHIS select '" & gUserName & "' ,to_char(SYSDATE,'yyyy-MM-dd HH24:mi:ss'),'修改后数量', a.* from CUSTOMERSHIPPINGUPTBL a where a.ID = '" & strid & "' ")
                
            End If
            
        End If
    Next
    MsgBox "修改完成"
    QueryData_Mod
End With




End Sub


Private Sub QueryData_Mod()
Dim strDN  As String
Dim strSql As String
Dim rs     As New ADODB.Recordset

fpS_Mod(0).MaxRows = 0
fpS_Mod(1).MaxRows = 0

If cbCustomerCode_Mod.text = "" Then
    MsgBox "请选择客户代码", vbInformation, "提示"
    Exit Sub

End If

If txtDN_Mod.text = "" Then
    MsgBox "请输入DN号", vbInformation, "提示"
    Exit Sub
End If
If CbItem_Mod.text = "" Then
    MsgBox "请选择修改项目", vbInformation, "提示"
    Exit Sub
End If


If cbCustomerCode_Mod.text = "37" Then

    strDN = Trim$(txtDN_Mod.text)
    strSql = "select 0 as 选择,a.delivery as DN, a.PLANNEDGIDATE as 出货日期,a.BATCHNUMBER AS JobID, a.QUANTITY as 出货数量 ,a.* from CUSTOMERSHIPPINGUPTBL a where delivery = '" & strDN & "' order by a.id desc "
    Set rs = Get_OracleRs(strSql)
    With fpS_Mod(0)
        Set .DataSource = rs
        .Row = -1
        .Col = 1  '选择
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
    End With
   
    
    If CbItem_Mod.text = "修改出货数量" Or CbItem_Mod.text = "删除" Then
         With fpS_Mod(0)
              .Row = -1
              .Col = -1
              .BackColor = &HFFFFFF
              .Col = 5
              .Lock = False
             .BackColor = vbGreen
             
             
             
             
         End With
    ElseIf CbItem_Mod.text = "修改出货日期" Then
         With fpS_Mod(0)
              .Row = -1
              .Col = -1
              .BackColor = &HFFFFFF
              .Col = 3
             .BackColor = vbGreen
  
         End With
    End If
    If Get_OracleCnt("  select * from packing_detailed i WHERE i.DN_NUM= '" & strDN & "'") > 0 Then
        If CbItem_Mod.text = "查询所有修改记录" Then
            Toolbar2.Buttons(3).Enabled = False
            Toolbar2.Buttons(5).Enabled = False
            
        ElseIf CbItem_Mod.text = "修改出货日期" Then
            Toolbar2.Buttons(3).Enabled = True

        Else
            MsgBox "外包已开始作业，此DN不可修改,请通知外包", vbInformation, "提示"
            Toolbar2.Buttons(3).Enabled = False
        End If
        Toolbar2.Buttons(5).Enabled = False
    Else
         If CbItem_Mod.text = "查询所有修改记录" Then
            Toolbar2.Buttons(3).Enabled = False
            Toolbar2.Buttons(5).Enabled = False
        Else
            Toolbar2.Buttons(3).Enabled = True
            Toolbar2.Buttons(5).Enabled = True
    End If
    End If
    If CbItem_Mod.text = "修改出货日期" Then
        strSql = "select distinct a.delivery as DN, a.MODIFYBY as 修改人员,a.MODIFYDATE as 修改时间,a.MODFIYITEM  as 修改项目, a.LASTUPDATEDATE as 出货时间   from CUSTOMERSHIPPINGUPTBL_MODHIS a where a.delivery = '" & strDN & "' and a.MODFIYITEM like '%日期%' order by a.MODIFYDATE desc "
    ElseIf CbItem_Mod.text = "修改出货数量" Then
        strSql = "select distinct  a.delivery as DN, a.MODIFYBY as 修改人员,a.MODIFYDATE as 修改时间,a.MODFIYITEM  as 修改项目, a.LASTUPDATEDATE as 出货时间,a.BatchNumber as JOBID, a.Quantity AS 数量   from CUSTOMERSHIPPINGUPTBL_MODHIS a where a.delivery = '" & strDN & "' and a.MODFIYITEM like '%数量%' order by a.BatchNumber,a.MODIFYDATE desc "
    ElseIf CbItem_Mod.text = "删除" Then
        strSql = "select distinct a.delivery as DN,  a.MODIFYBY as 修改人员,a.MODIFYDATE as 修改时间,a.MODFIYITEM  as 修改项目, a.LASTUPDATEDATE as 出货时间,a.BatchNumber as JOBID, a.Quantity AS 数量   from CUSTOMERSHIPPINGUPTBL_MODHIS a where a.delivery = '" & strDN & "' and a.MODFIYITEM like '%删除%' order by a.BatchNumber,a.MODIFYDATE desc "
    ElseIf CbItem_Mod.text = "查询所有修改记录" Then
        strSql = "select distinct  a.delivery as DN,  a.MODIFYBY as 修改人员,a.MODIFYDATE as 修改时间,a.MODFIYITEM  as 修改项目, a.LASTUPDATEDATE as 出货时间,a.BatchNumber as JOBID, a.Quantity AS 数量   from CUSTOMERSHIPPINGUPTBL_MODHIS a where a.delivery = '" & strDN & "' order by a.BatchNumber,a.MODIFYDATE desc "
        
     End If
    Set rs = Get_OracleRs(strSql)
    With fpS_Mod(1)
        Set .DataSource = rs
    End With
ElseIf cbCustomerCode_Mod.text = "SG005" Then

    strDN = Trim$(txtDN_Mod.text)
    strSql = "select 0 as 选择,a.* from ERPBASE..tblCustomerShippingUp_So a where SO_NO = '" & strDN & "' order by a.SO_LINE "
    Set rs = Get_SqlserveRs(strSql)
    With fpS_Mod(0)
        Set .DataSource = rs
        .Row = -1
        .Col = 1  '选择
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
    End With
   

End If

End Sub



'------------------------------------------------------
'此后为新代码

Private Sub Reconciliation(sDH)
' 免对账维护
Dim sFlag As String
Dim strSql As String
Dim rs          As New ADODB.Recordset
sFlag = "N"
If chkDZD.Value = 1 Then
    sFlag = "Y"
Else
    sFlag = "N"
End If

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = INIadoCon

rs.Source = "select * from [erpdata].[dbo].[MDZD_TBL] where SENT_ID = '" & sDH & "'"

rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount > 0 Then
    strSql = "update [erpdata].[dbo].[MDZD_TBL] set flag = '" & sFlag & "' where SENT_ID = '" & sDH & "'"
    If AddSql2(strSql) = 0 Then
        MsgBox "没有维护到免对账单", vbCritical, "提示"
        Exit Sub
    End If
Else
    strSql = "insert into  [erpdata].[dbo].[MDZD_TBL](SENT_ID, FLAG) values('" & sDH & "', '" & sFlag & "')"
    If AddSql2(strSql) = 0 Then
        MsgBox "没有维护到免对账单", vbCritical, "提示"
        Exit Sub
    End If
End If

End Sub


'保存资料
Private Sub SaveData_Ship()
Dim i           As Long
Dim j           As Long
Dim k           As Long
Dim strSql      As String
Dim rs          As New ADODB.Recordset
Dim strCust     As String   '客户代码
Dim strKF       As String   '库房
Dim strid       As String   'ID字符串
Dim strXH       As String   '箱号字符串
Dim strgdh      As String   '工单字符串
Dim strLCK      As String   '流程卡字符串
Dim strlps      As String   '合格数
Dim strbls      As String   '不良数
Dim strzcbls    As String   '制程不良数
Dim strDN    As String   'dn
Dim intCount    As Integer  '记录个数
Dim sDH As String
Dim sFlag As String
Dim rs1 As ADODB.Recordset
Dim flag_up As String
Dim qty_total As Long
Dim dn_temp  As String 'dn
Dim bigbox_temp  As String '大箱号
Dim XH_temp  As String
Dim KF_temp  As String
Dim ordernumlist  As String
Dim strKFList  As String

    mailflag = 1 '正常发货的邮件格式
    intFlag = 1
    intCount = 0
    strKF = ""
    strCust = ""
    strKFList = ""
    '资料检查
    If txt(1).text = "" Or txt(2).text = "" Or txt(3).text = "" Then
        MsgBox "请检查资料的完整性！", vbInformation, "提示"
        Exit Sub
    End If
    If intFlag <> 3 Then
        If Cob(6).text = "" Then
            MsgBox "请选择运输方式！", vbInformation, "提示"
            Exit Sub
        End If
        If Cob(7).text = "" Then
            MsgBox "请选择收货客户！", vbInformation, "提示"
            Exit Sub
        End If
        'If txt(4).Text = "" Then
        If Cob(9).text = "" Then
            MsgBox "请选择收货地址！", vbInformation, "提示"
            Exit Sub
        End If
        If Val(txt(6).text) + Val(txt(7).text) + Val(txt(8).text) <= 0 Then
            MsgBox "请先选择要出货的箱号！", vbInformation, "提示"
            Exit Sub
        End If
        '记录出货库房（校验出货的库房是不是一个）
        intCount = 0
        strKF = ""
        With Fps_Ship(0)
            If .MaxRows <= 0 Then
                MsgBox "请先查询资料！", vbInformation, "提示"
                Exit Sub
            End If
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsM.E_CHOOSE '选择
                If .Value = 1 Then
                    intCount = intCount + 1 '计数器
                    .Col = FpsM.E_cust '客户
                    If strCust = "" Then
                        strCust = Trim$(.text)
                    End If
                    '.Col = FpsM.e_KF '库房
                    'If strKF = "" Then
                    '    strKF = Trim$(.Text)
                    'End If
                    '校验下一个是否一样，不一样就提示错误
                    'If strKF <> Trim(.Text) Then
                    '    MsgBox "出货只能出同一个库房的资料！", vbInformation, "提示"
                    '    Exit Sub
                    'End If
                    .Col = FpsM.e_KF '库房
                    If InStr(strKFList, Split(Trim$(.text), " ")(0)) = 0 Then
                        If strKFList = "" Then
                            strKFList = Split(Trim$(.text), " ")(0)
                        Else
                            strKFList = strKFList & "," & Split(Trim$(.text), " ")(0)
                        End If
                    End If
                    
                End If
            Next
        End With
        If intCount <= 0 Then
            MsgBox "请先选择要出货的箱号！", vbInformation, "提示"
            Exit Sub
        End If
       ' strKF = Left(strKF, InStr(strKF, " ") - 1) '截取库房编号
        '记录明细资料
        With Fps_Ship(1)
            If .MaxRows <= 0 Then
                MsgBox "请先选择要出货的箱号资料！", vbInformation, "提示"
                Exit Sub
            End If
'新增

           If Cob(0).text = "37" Then
                
               dn_temp = ""
               For i = 1 To .MaxRows
                    .Row = i
                    .Col = FpsD.E_DN
                   If .text <> dn_temp Then '切换dn时切换申请单号
                        dn_temp = Trim(.text)
                        .Col = FpsD.E_XH
                        strXH = Trim(.text)
                        .Col = FpsD.e_BigX
                        bigbox_temp = Trim(.text)
                        .Col = FpsD.e_KF
                        KF_temp = Trim(.text)
                        .SetText FpsD.E_Note2, i, "切换申请单号"
                   Else
                        .Col = FpsD.e_KF
                        If KF_temp <> Trim(.text) Then
                            MsgBox dn_temp & "存于不同的库房，无法同一次出货，请确认！", vbInformation, "提示"
                            Exit Sub
                        End If
                   
                       .Col = FpsD.e_BigX
                       If Trim(.text) <> bigbox_temp Then '同dn不同大箱号，有可能需要变更申请单号
                          bigbox_temp = Trim(.text)
                          .Col = FpsD.E_XH    '箱号
                          XH_temp = Trim(.text)
                          For j = i + 1 To .MaxRows
                             .Row = j
                             .Col = FpsD.e_BigX
                             If Trim(.text) <> bigbox_temp Then Exit For
                             .Col = FpsD.E_XH    '箱号
                             XH_temp = XH_temp & Trim(.text) & "★"
                          Next
                          
                            If Len(strXH) + Len(XH_temp) < 6000 Then
                                .Col = FpsD.E_XH    '箱号
                                strXH = strXH & Trim(.text) & "★"
                            Else   '切换申请单号
                                .Col = FpsD.E_XH    '箱号
                                strXH = Trim(.text)
                                .SetText FpsD.E_Note2, i, "切换申请单号"
                            End If
                        Else
                            .Col = FpsD.E_XH    '箱号
                            strXH = strXH & Trim(.text) & "★"
                        End If
                   End If
               Next
                If CheckQtyByDn <> "" Then '核对DN数量
                   ' MsgBox CheckQtyByDn & "数量与库存不一致", vbInformation, "提示"
                    Exit Sub
                End If
            ElseIf UCase(Cob(0).text) = "SG005" Then
                If TxtShipDate_Ship.text = "" Then
                    MsgBox "请选择出货日期,今天出货也需要再选一次", vbInformation, "提示"
                    Exit Sub
                End If
                If CheckQtyBySO <> "" Then '核对所选数量是否超出SO数量
                    Exit Sub
                End If
            End If

           If MsgBox("确定要执行吗，点'是'继续，点'否'取消？", vbInformation + vbYesNo, "提示") = vbNo Then
                Exit Sub
           End If
        If UCase(Cob(0).text) = "37" Then
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsD.E_Note2
                If Trim(.text) <> "" Then
                    If i > 1 Then
                        AutoCode
                        Call Reconciliation(txt(0).text)   '免对账维护
                        If ExecProc(intFlag, txt(0).text, strid, strXH, strgdh, strLCK, strlps, strbls, strzcbls, strKF, strDN, strCust) = False Then
                            Call showdata_shiplist(Cob(0).text, Replace(ordernumlist, ",", "','"), Format(dtShipDate_Ship.Value, "yyyy/mm/dd"))
                            MsgBox "任务执行失败！", vbInformation, Me.Caption
                            Exit Sub
                        End If
                        If ordernumlist = "" Then
                            ordernumlist = txt(0).text
                        Else
                            ordernumlist = ordernumlist & "," & txt(0).text
                        End If
                    End If
                    strid = ""
                    strXH = ""
                    strgdh = ""
                    strLCK = ""
                    strlps = ""
                    strbls = ""
                    strzcbls = ""
                End If
                .Row = i
                .Col = FpsD.e_ID    'ID
                If InStr(strid, Trim$(.text)) <= 0 Then '不存在就累加
                    strid = strid & Trim$(.text) & "★"
                End If
                .Col = FpsD.E_XH    '箱号
                strXH = strXH & Trim(.text) & "★"
                .Col = FpsD.e_GDH   '工单号
                strgdh = strgdh & Trim(.text) & "★"
                .Col = FpsD.e_LCK   '流程卡
                strLCK = strLCK & Trim(.text) & "★"
                .Col = FpsD.e_GNum  '合格数
                strlps = strlps & Trim(.text) & "★"
                .Col = FpsD.e_BLNum '来料不良数
                strbls = strbls & Trim(.text) & "★"
                .Col = FpsD.e_ZCNum '制程不良数
                strzcbls = strzcbls & Trim(.text) & "★"
                .Col = FpsD.e_KF '库房
                strKF = Trim(.text)
                .Col = FpsD.E_DN 'dn
                strDN = Trim(.text)
                If i = .MaxRows Then
                    AutoCode
                    Call Reconciliation(txt(0).text)   '免对账维护
                    If ExecProc(intFlag, txt(0).text, strid, strXH, strgdh, strLCK, strlps, strbls, strzcbls, strKF, strDN, strCust) = False Then
                        MsgBox "任务执行失败！", vbInformation, Me.Caption
                        'Call showdata_shiplist(Replace(ordernumlist, ",", "','"))
                        Call showdata_shiplist(Cob(0).text, Replace(ordernumlist, ",", "','"), Format(dtShipDate_Ship.Value, "yyyy/mm/dd"))
                        Exit Sub
                    End If
                    If ordernumlist = "" Then
                        ordernumlist = txt(0).text
                    Else
                        ordernumlist = ordernumlist & "," & txt(0).text
                    End If
                End If
                
            Next
        ElseIf Cob(0).text = "SG005" Then
            For k = 0 To UBound(Split(strKFList, ","))
                strKF = Split(strKFList, ",")(k)
                strid = ""
                strXH = ""
                strgdh = ""
                strLCK = ""
                strlps = ""
                strbls = ""
                strzcbls = ""
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = FpsD.e_KF
                    If Trim(.text) = strKF Then

                        .Row = i
                        .Col = FpsD.e_ID    'ID
                        If InStr(strid, Trim$(.text)) <= 0 Then '不存在就累加
                            strid = strid & Trim$(.text) & "★"
                        End If
                        .Col = FpsD.E_XH    '箱号
                        strXH = strXH & Trim(.text) & "★"
                        .Col = FpsD.e_GDH   '工单号
                        strgdh = strgdh & Trim(.text) & "★"
                        .Col = FpsD.e_LCK   '流程卡
                        strLCK = strLCK & Trim(.text) & "★"
                        .Col = FpsD.e_GNum  '合格数
                        strlps = strlps & Trim(.text) & "★"
                        .Col = FpsD.e_BLNum '来料不良数
                        strbls = strbls & Trim(.text) & "★"
                        .Col = FpsD.e_ZCNum '制程不良数
                        strzcbls = strzcbls & Trim(.text) & "★"
                        .Col = FpsD.e_KF '库房
                        strKF = Trim(.text)
                        .Col = FpsD.E_DN 'dn
                        strDN = Trim(.text)
                    End If
                Next
                AutoCode
                Call Reconciliation(txt(0).text)   '免对账维护
                If ExecProc(intFlag, txt(0).text, strid, strXH, strgdh, strLCK, strlps, strbls, strzcbls, strKF, strDN, strCust) = False Then
                    MsgBox "任务执行失败！", vbInformation, Me.Caption
                   ' Call showdata_shiplist(Replace(ordernumlist, ",", "','"))
                    Call showdata_shiplist(Cob(0).text, Replace(ordernumlist, ",", "','"), Format(dtShipDate_Ship.Value, "yyyy/mm/dd"))
                    Exit Sub
                End If
                If ordernumlist = "" Then
                    ordernumlist = txt(0).text
                Else
                    ordernumlist = ordernumlist & "," & txt(0).text
                End If
            Next
        End If
        End With
    End If
    'Call showdata_shiplist(Replace(ordernumlist, ",", "','"))
    Call showdata_shiplist(Cob(0).text, Replace(ordernumlist, ",", "','"), Format(dtShipDate_Ship.Value, "yyyy/mm/dd"))
    If Cob(0).text = "37" Then
        LoadDn '刷新DNlist
    ElseIf Cob(0).text = "SG005" Then
        LoadSO '刷新SOlist
    End If
 End Sub
 
 
Private Function ExecProc(intFlag As Integer, strOrderNo As String, strid As String, strXH As String, strgdh As String, strLCK As String, strlps As String, strbls As String, strzcbls As String, strKF As String, strDN As String, strCust As String)
    Dim strSql As String
    Dim rs                  As New ADODB.Recordset
    Dim strShipDate As String
    Dim strTerm As String
    Dim strForwarder As String
    Dim strInstruction As String
    Dim strShipTo As String

On Error GoTo FError
    ExecProc = True
    If Len(strXH) >= 6000 Or Len(strLCK) >= 6000 Then
        MsgBox "选取的出货数量超出上限，请减少选取的数量", vbInformation, "提示"
        ExecProc = False
        Exit Function
    End If

    '保存资料
    Set adoCmd = New ADODB.Command
     Set adoCmd.ActiveConnection = INIadoCon2
     adoCmd.CommandText = "uspfp_fhsqXH1"
     adoCmd.Parameters.Refresh
     adoCmd.CommandType = adCmdStoredProc
  
        Set adoPrmReturn = New ADODB.Parameter         '返回执行成功标记
        adoPrmReturn.type = adInteger
        adoPrmReturn.Direction = adParamReturnValue
        adoCmd.Parameters.Append adoPrmReturn
        
        Set adoprm1 = New ADODB.Parameter               '单据编号
        adoprm1.type = adChar
        adoprm1.Size = 50
        adoprm1.Direction = adParamInput
        adoprm1.Value = Trim(strOrderNo)
        adoCmd.Parameters.Append adoprm1
        
        Set adoprm2 = New ADODB.Parameter                 '1ID号
        adoprm2.type = adChar
        adoprm2.Size = 8000
        adoprm2.Direction = adParamInput
        adoprm2.Value = Trim(strid)
        adoCmd.Parameters.Append adoprm2
        
        Set adoPrm3 = New ADODB.Parameter                 '2箱号
        adoPrm3.type = adChar
        adoPrm3.Size = 80000
        adoPrm3.Direction = adParamInput
        adoPrm3.Value = Trim(strXH)
        adoCmd.Parameters.Append adoPrm3
        
        Set adoPrm4 = New ADODB.Parameter                 '5lck
        adoPrm4.type = adChar
        adoPrm4.Size = 80000
        adoPrm4.Direction = adParamInput
        adoPrm4.Value = Trim(strLCK)
        adoCmd.Parameters.Append adoPrm4
        
        Set adoPrm5 = New ADODB.Parameter                 '6gdh
        adoPrm5.type = adChar
        adoPrm5.Size = 80000
        adoPrm5.Direction = adParamInput
        adoPrm5.Value = Trim(strgdh)
        adoCmd.Parameters.Append adoPrm5
        
        Set adoPrm6 = New ADODB.Parameter                 '7lps
        adoPrm6.type = adChar
        adoPrm6.Size = 80000
        adoPrm6.Direction = adParamInput
        adoPrm6.Value = Trim(strlps)
        adoCmd.Parameters.Append adoPrm6
        
        Set adoPrm7 = New ADODB.Parameter               '8bls
        adoPrm7.type = adChar
        adoPrm7.Size = 80000
        adoPrm7.Direction = adParamInput
        adoPrm7.Value = Trim(strbls)                 '
        adoCmd.Parameters.Append adoPrm7
        
        Set adoPrm8 = New ADODB.Parameter               '9zcbls
        adoPrm8.type = adChar
        adoPrm8.Size = 80000
        adoPrm8.Direction = adParamInput
        adoPrm8.Value = Trim(strzcbls)
        adoCmd.Parameters.Append adoPrm8
        
        
        Set adoPrm9 = New ADODB.Parameter               '申请人
        adoPrm9.type = adChar
        adoPrm9.Size = 50
        adoPrm9.Direction = adParamInput
        adoPrm9.Value = Trim(txt(1).text)
        adoCmd.Parameters.Append adoPrm9
        
        Set adoprm10 = New ADODB.Parameter               '申请部门
        adoprm10.type = adChar
        adoprm10.Size = 50
        adoprm10.Direction = adParamInput
        adoprm10.Value = Trim(txt(3).text)
        adoCmd.Parameters.Append adoprm10
        
        Set adoPrm11 = New ADODB.Parameter               '客户代码
        adoPrm11.type = adChar
        adoPrm11.Size = 50
        adoPrm11.Direction = adParamInput
        adoPrm11.Value = Trim(strCust)
        adoCmd.Parameters.Append adoPrm11
        
        Set adoPrm12 = New ADODB.Parameter             '发货地址
        adoPrm12.type = adChar
        adoPrm12.Size = 255
        adoPrm12.Direction = adParamInput
        'adoPrm12.Value = Trim(txt(4).Text)
        adoPrm12.Value = Trim(Cob(9).text)
        adoCmd.Parameters.Append adoPrm12
        
        Set adoPrm13 = New ADODB.Parameter               '运输方式
        adoPrm13.type = adChar
        adoPrm13.Size = 50
        adoPrm13.Direction = adParamInput
        adoPrm13.Value = Trim(Cob(6).text)
        adoCmd.Parameters.Append adoPrm13
        
        Set adoPrm14 = New ADODB.Parameter             '库房编号
        adoPrm14.type = adChar
        adoPrm14.Size = 50
        adoPrm14.Direction = adParamInput
        adoPrm14.Value = strKF
        adoCmd.Parameters.Append adoPrm14
        
        Set adoPrm15 = New ADODB.Parameter             '说明,备注
        adoPrm15.type = adChar
        adoPrm15.Size = 50
        adoPrm15.Direction = adParamInput
        adoPrm15.Value = Trim(txt(5).text)
        adoCmd.Parameters.Append adoPrm15
        
        Set adoprmFG = New ADODB.Parameter             '1、新增;3、删除;2、修改
        adoprmFG.type = adInteger
        adoprmFG.Direction = adParamInput
        adoprmFG.Value = intFlag
        adoCmd.Parameters.Append adoprmFG
        
        Set adoprm16 = New ADODB.Parameter             '收货客户
        adoprm16.type = adChar
        adoprm16.Size = 50
        adoprm16.Direction = adParamInput
        adoprm16.Value = Trim(Cob(7).text)
        adoCmd.Parameters.Append adoprm16
      
     adoCmd.Execute
     Screen.MousePointer = 0
     
     If adoPrmReturn.Value = 0 Then
        
         strSql = "SELECT * FROM erptemp..SHIP_ORDER_sql a WHERE a.ship_order='" & Trim(strOrderNo) & "'"
         If rs.State = adStateOpen Then rs.Close
         rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
         If rs.RecordCount < 1 Then
          
        '   MsgBox "保存失败！", vbInformation, Me.Caption
           ExecProc = False
           Exit Function
          
         End If
        If Cob(0).text = "37" And strDN <> "" Then
            If DateDiff("d", Format(dtShipDate_Ship.Value, "yyyy/mm/dd"), Format(Now(), "yyyy/mm/dd")) >= 1 Then
            '应该前一天出的货,当天没出，第二天出，出货日期带第二天的日期
                strShipDate = Format(Now(), "yyyy/mm/dd")
            Else
                strShipDate = Format(dtShipDate_Ship.Value, "yyyy/mm/dd")
            End If
            strInstruction = GetSqlServerStr(" select isnull(ShippingInstruction,'') FROM ERPBASE..tblCustomerShippingUp WHERE rtrim(delivery)='" & Trim(strDN) & "'")
            If InStr(UCase(strInstruction), "FED") > 0 Or InStr(UCase(strInstruction), "DHL") > 0 Then
                strInstruction = "FED"
            Else
                strInstruction = "自提"
            End If
            strShipTo = GetSqlServerStr(" select isnull(ShipToName,'') FROM ERPBASE..tblCustomerShippingUp WHERE rtrim(delivery)='" & Trim(strDN) & "'")
            'Shenzhen YH Global Logistics Co. Ltd=HW
            If Trim(UCase(strShipTo)) = "SHENZHEN YH GLOBAL LOGISTICS CO. LTD" Then
                strShipTo = "HW"
            Else
                strShipTo = ""
            End If
            
            strSql = " INSERT INTO erpdata..tblShipOrder_Dn(Dn, ShipOrder, Quality,BOND,CUST_DEVICE,HT_DEVICE,SHIP_DATE,Remark1,Remark2) " & _
            " SELECT '" & Trim(strDN) & "', a.单据编号,sum(a.数量) AS 数量 ,LEFT(a.大工单,1) AS 保税,c.MPN_DESC ,d.QTECHPTNO ,'" & strShipDate & "','" & strInstruction & "','" & strShipTo & "'" & _
            " FROM erpdata..tblstocksqfhSUB a , erpbase..tblmappingdata   b,erpbase..tblcustomeroi c,erptemp..tbltsvnpiproduct d " & _
            " WHERE a.流程卡编号=b.SUBSTRATEID AND a.工单号=b.LOTID AND b.LOTID =c.SOURCE_BATCH_ID AND b.FILENAME =convert(VARCHAR(30),convert(int,c.id)) AND a.料号=d.QTECHPTNO2  " & _
            " AND  a.单据编号 ='" & Trim(strOrderNo) & "'" & _
            " GROUP BY a.单据编号,LEFT(a.大工单,1),c.MPN_DESC ,d.QTECHPTNO "
            AddSql2 (strSql)

        End If
    
        If Cob(0).text = "SG005" And strsono <> "" Then
            strShipDate = Format(dtShipDate_Ship.Value, "yyyy/mm/dd")
            
            strTerm = GetSqlServerStr(" select isnull(TERM,'') FROM ERPBASE..tblCustomerShippingUp_So  WHERE rtrim(SO_NO)='" & Trim(strsono) & "' AND SO_LINE='" & strsoline & "'")
            
            strForwarder = GetSqlServerStr(" select isnull(Forwarder,'') FROM ERPBASE..tblCustomerShippingUp_So  WHERE rtrim(SO_NO)='" & Trim(strsono) & "' AND SO_LINE='" & strsoline & "'")
            
            
            strSql = " INSERT INTO erpdata..tblShipOrder_Dn(SO_NO, SO_LINE, ShipOrder, Quality,BOND,CUST_DEVICE,HT_DEVICE,PCSNUM,SHIPTO,STOCKID,SHIP_DATE,REMARK1,REMARK2) " & _
             " SELECT '" & strsono & "','" & strsoline & "', a.单据编号,sum(a.数量) AS 数量 ,LEFT(a.大工单,1) AS 保税,c.MPN_DESC ,d.QTECHPTNO,COUNT(*) AS 片数,'" & Trim(Cob(9).text) & "' ,'" & strKF & "','" & strShipDate & "','" & strTerm & "','" & strForwarder & "'" & _
             " FROM erpdata..tblstocksqfhSUB a , erpbase..tblmappingdata   b,erpbase..tblcustomeroi c,erptemp..tbltsvnpiproduct d " & _
             " WHERE a.流程卡编号=b.SUBSTRATEID AND a.工单号=b.LOTID AND b.LOTID =c.SOURCE_BATCH_ID AND b.FILENAME =convert(VARCHAR(30),convert(int,c.id)) AND a.料号=d.QTECHPTNO2  " & _
             " AND  a.单据编号 ='" & Trim(strOrderNo) & "'" & _
             " GROUP BY a.单据编号,LEFT(a.大工单,1),c.MPN_DESC ,d.QTECHPTNO "
            
            AddSql2 (strSql)
              
        End If
        
       ' MsgBox "已经成功执行您的任务！", vbInformation, Me.Caption

     Else
        ExecProc = False
       ' MsgBox "任务执行失败！", vbInformation, Me.Caption
        Exit Function
        
     End If
     Exit Function
FError:
    MsgBox "执行失败，原因：" & Err.DESCRIPTION, vbInformation, "Frm_uploadShippingList.ExecProc"
End Function




'自动获得单据编号，但是不一定是保存到资料库的单据编号
Private Sub AutoCode()
Dim strCompare          As String
Dim strDate             As String
Dim strSql              As String
Dim strSqlin              As String
Dim strSqlin_bak              As String
Dim maxid  As String
Dim rs                  As New ADODB.Recordset
Dim rs1                  As New ADODB.Recordset
    
    txt(0).text = ""
    strCompare = Trim("F" & Format(Now, "yymmdd"))
    strDate = Trim(Format(Now, "yymmdd"))
'    strSql = "SELECT MAX(LEFT(RTRIM(发货清单编号),LEN(RTRIM(发货清单编号))-4))+RIGHT('0000'+CAST(CAST(RIGHT(MAX(RTRIM(发货清单编号)),4) AS INT)+1 AS VARCHAR),4) 单号 " & _
'             " FROM tblCodeRule WHERE 发货清单编号 LIKE '" & strCompare & "%' "
             
    strSql = " select RIGHT(REPLICATE('0',4)+CAST((ISNULL(MAX(ID),0) + 1) AS varchar(10)),4) AS 单号,ISNULL(MAX(ID),0) + 1 as id from erptemp..SHIP_NUM_SEQ a  WHERE  CONVERT(VARCHAR(100),a.create_date ,12)  = '" & strDate & "'  "
    strSqlin_bak = " SELECT MAX(id) FROM erptemp..SHIP_NUM_SEQ_BAK "
     
     
     If rs1.State = adStateOpen Then rs1.Close
     rs1.Open strSqlin_bak, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     maxid = rs1.Fields(0).Value
     rs1.Close
     
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs!单号 = "0001" Then
    txt(0).text = strCompare & rs!单号
    txtID_Text.text = Val(maxid) + 1
    Else
    txt(0).text = strCompare & Format((Val(rs!单号) + 20000 - Val(maxid)), "0000")
    txtID_Text.text = rs.Fields(1).Value
    
    End If
    strSqlin = " INSERT INTO erptemp..SHIP_NUM_SEQ (CREATE_DATE,CREATE_BY,FLAG ) VALUES (GETDATE(),'" & txt(1).text & "','0') "
    
    
    
  If AddSql2(strSqlin) <> 0 Then
     MsgBox "获取单据号" & txt(0).text, vbInformation, "提示"
     End If
    
'    If IsNull(Rs!单号) Then
'        txt(0).Text = strCompare & "0001"
'    Else
'        txt(0).Text = Trim$("" & Rs!单号)
'    End If
    rs.Close
    
End Sub

Private Sub cmdCreate_Click()
AutoCode
End Sub


'界面布局

Private Sub Form_Resize()
    control_resize

End Sub
Private Sub control_resize()

    On Error Resume Next
    
    SSTab1.Move SSTab1.Left, SSTab1.Top, Me.Width - SSTab1.Left - 350, Me.Height - SSTab1.Top - 400
    Select Case SSTab1.Tab
    
    Case 0
    

    Case 1
    

    
    Case 2
    
    Toolbar3.Move Toolbar3.Left, Toolbar3.Top, SSTab1.Width - Toolbar3.Left - 250, Toolbar3.Height

    SSTab2.Move SSTab2.Left, SSTab2.Top, SSTab1.Width - SSTab2.Left - 350, SSTab1.Height - SSTab2.Top - 400

    
    Select Case SSTab2.Tab
    
    Case 0
    Fra(1).Move Fra(1).Left, Fra(1).Top, SSTab1.Width - Fra(1).Left - 250, Fra(1).Height
    Fra(0).Move Fra(0).Left, SSTab2.Height - 2800, SSTab2.Width - Fra(0).Left - 350, Fra(0).Height
    Fra(2).Move Fra(2).Left, Fra(2).Top, SSTab2.Width - Fra(2).Left - 350, Fra(0).Top - Fra(2).Top - 300
    Fps_Ship(0).Move Fps_Ship(0).Left, Fps_Ship(0).Top, Fra(2).Width - 6800, Fra(2).Height - Fps_Ship(0).Top - 100
    Fps_Ship(1).Move Fra(2).Width - 6700, Fps_Ship(1).Top, 6500, Fra(2).Height - Fps_Ship(1).Top - 100
    Fps_Ship(2).Move Fps_Ship(2).Left, Fps_Ship(2).Top, Fra(0).Width - Fps_Ship(2).Top - 350, Fra(0).Height - Fps_Ship(2).Top - 400
    
    Case 1
    
    Fps_Ship_del(0).Move Fps_Ship_del(0).Left, Fps_Ship_del(0).Top, Fps_Ship_del(0).Width, SSTab2.Height - Fps_Ship_del(0).Top - 100
    Fps_Ship_del(1).Move Fps_Ship_del(1).Left, Fps_Ship_del(1).Top, SSTab2.Width - Fps_Ship_del(1).Left - 350, SSTab2.Height - Fps_Ship_del(1).Top - 100
        
    End Select
    
    End Select
    
    
    
End Sub
'初始化控件

Private Sub InitCtrl_Ship()
Dim i                   As Integer
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
    
    'Fps初始化
    With Fps_Ship(0)
        .ReDraw = False
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        .MaxRows = 0
        .MaxCols = FpsM.e_MCol - 1
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        
        .Col = FpsM.E_CHOOSE
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        .SetText FpsM.E_CHOOSE, 0, "选择"
        .SetText FpsM.e_ID, 0, "项次"
        .SetText FpsM.E_cust, 0, "客户"
        .SetText FpsM.e_GDH, 0, "工单号/LOT号"
        .SetText FpsM.e_BigX, 0, "大箱号"
        .SetText FpsM.e_LH, 0, "料号"
        .SetText FpsM.e_NUM, 0, "库存数"
        .SetText FpsM.E_GG, 0, "规格"
        .SetText FpsM.E_XH, 0, "型号"
        .SetText FpsM.E_UNIT, 0, "单位"
        .SetText FpsM.e_KF, 0, "出货库房"
        .SetText FpsM.E_DN, 0, "DN"
        '设定是否排序
        .UserColAction = UserColActionSort
        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
        .ColWidth(-1) = 12
        .ColWidth(FpsM.E_CHOOSE) = 4
        .ColWidth(FpsM.e_ID) = 4
        .ColWidth(FpsM.E_cust) = 5
        .ColWidth(FpsM.E_UNIT) = 4
        .ZOrder
        .ReDraw = True
    End With
    With Fps_Ship(1)
        .ReDraw = False
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        .MaxRows = 0
        .MaxCols = FpsD.e_MCol - 1
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        
     '   .Col = FpsD.e_BigX
     '   .Row = -1
        '.ColHidden = True
     '   .Col = FpsD.E_ID
     '   .Row = -1
      '  .ColHidden = True
        .Col = FpsD.E_Note2
        .Row = -1
        .ColHidden = True
        
        .SetText FpsD.e_BigX, 0, "大箱号"
        .SetText FpsD.E_XH, 0, "箱号"
        .SetText FpsD.e_GDH, 0, "工单号/LOT号"
        .SetText FpsD.e_LCK, 0, "流程卡/WaferID"
        .SetText FpsD.e_LH, 0, "料号"
        .SetText FpsD.e_GNum, 0, "合格数"
        .SetText FpsD.e_BLNum, 0, "来料不良数"
        .SetText FpsD.e_ZCNum, 0, "制程不良数"
        .SetText FpsD.e_ID, 0, "项次"
        .SetText FpsD.e_KF, 0, "出货库房"
        .SetText FpsD.E_DN, 0, "DN"
        .SetText FpsD.E_Note2, 0, "Note"
'        '设定是否排序
'        .UserColAction = UserColActionSort
'        For i = 1 To .MaxCols
'            .Col = i
'            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
'        Next
'        .colwidth(-1) = 12
'        .ZOrder
        .ReDraw = True
    End With
    
     With Fps_Ship_del(0)
        .ReDraw = False
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        .MaxRows = 0
        .MaxCols = 6
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        
        .Col = 1
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        .SetText 1, 0, "选择"
        .SetText 2, 0, "客户代码"
        .SetText 3, 0, "客户名称"
        .SetText 4, 0, "单据编号"
        .SetText 5, 0, "关联单据"

        '设定是否排序
        .UserColAction = UserColActionSort
        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
        .ColWidth(-1) = 12
        .ColWidth(1) = 4
        .ColWidth(2) = 4
        .ColWidth(3) = 18
        .ColWidth(4) = 10
         
        .ZOrder
        .ReDraw = True
    End With
    
    
    
    strSql = "select EmpName from XTW..employee where empno = '" & gUserName & "'"
   strRealName = Get_SqlStr2(strSql)

    '申请人
    txt(1).text = Trim(gUserName) & " " & Trim(strRealName)
    '申请日期
    txt(2).text = Format(Now(), "YYYY-MM-DD")
    '申请部门
   ' txt(3).Text = Trim(strUserDepartNUM) & Space(1) & Trim(strUserDepart)
     txt(3).text = "06 销售部"
    '加载客户代码和收货客户
    strSql = "SELECT 客户代码 FROM dbo.tblXCustomer"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
  '  Cob(0).Clear
    Cob(7).Clear
    If Not rs.EOF Then
        For i = 1 To rs.RecordCount
      '      Cob(0).AddItem Trim$("" & rs!客户代码)
            Cob(7).AddItem Trim$("" & rs!客户代码)
            rs.MoveNext
        Next
    Else
        MsgBox "客户代码加载失败，请联系系统管理员！", vbInformation, "提示"
        Exit Sub
    End If
    rs.Close
    '加载产线标记
    strSql = "select rtrim(名称)+' '+rtrim(说明) 产线标记 from tblbase where 说明2='产线标记' order by 说明"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Cob(1).Clear
    If Not rs.EOF Then
        For i = 1 To rs.RecordCount
            Cob(1).AddItem Trim$("" & rs!产线标记)
            rs.MoveNext
        Next
    Else
        MsgBox "产线标记加载失败，请联系系统管理员！", vbInformation, "提示"
        Exit Sub
    End If
    rs.Close
    '加载库房名称
    strSql = "SELECT  库存名称 FROM ( SELECT rtrim(库房代码)+' '+rtrim(库房名称) 库存名称 from erpbase..tblstock  where 仓库属性='成品仓'  AND 库房名称  NOT LIKE '%委外%'  UNION   select rtrim(库房代码)+' '+rtrim(库房名称) 库存名称 from erpbase..tblstock  where  库房代码 = '85'  ) A ORDER BY 库存名称 "
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Cob(2).Clear
    If Not rs.EOF Then
        For i = 1 To rs.RecordCount
            Cob(2).AddItem Trim$("" & rs!库存名称)
            rs.MoveNext
        Next
    Else
        MsgBox "库房名称加载失败，请联系系统管理员！", vbInformation, "提示"
        Exit Sub
    End If
    rs.Close
    '加载运输方式
    strSql = "select RTRIM(运输方式代码)+' '+RTRIM(运输方式名称) 运输方式 from erpdata..tblXTransitMode"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Cob(6).Clear
    If Not rs.EOF Then
        For i = 1 To rs.RecordCount
            Cob(6).AddItem Trim$("" & rs!运输方式)
            rs.MoveNext
        Next
    Else
        MsgBox "运输方式加载失败，请联系系统管理员！", vbInformation, "提示"
        Exit Sub
    End If
    rs.Close
    '加载DN
    ' strSql = "SELECT DISTINCT Delivery FROM erpbase..tblCustomerShippingUp WHERE Flag='Y' Order By 1 Desc"
    ' If rs.State = adStateOpen Then rs.Close
    ' rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    ' Cob(8).Clear
    ' If Not rs.EOF Then
        ' For i = 1 To rs.RecordCount
            ' Cob(8).AddItem Trim$("" & rs!Delivery)
            ' rs.MoveNext
        ' Next
    ' Else
        ' MsgBox "DN加载失败，请联系系统管理员！", vbInformation, "提示"
        ' Exit Sub
    ' End If
    ' rs.Close
End Sub

Private Sub Cob_DropDown(Index As Integer)
Dim i                   As Integer
Dim strSql              As String
Dim rs                  As New ADODB.Recordset

    If Index = 3 Then
        '加载工单号
        strSql = "select distinct rtrim(a.工单号) 工单号 from erpdata..tblStockNum a inner join erpdata..tblbase b on a.产线标记=b.名称 and b.说明2='产线标记' " & _
                 " where a.库存数>0 and b.名称='" & Val(Trim(Cob(1).text)) & "'"
        If Trim$(Cob(2).text) <> "" Then
            If InStr(Cob(2).text, " ") > 0 Then
                strSql = strSql & " and a.库房编号='" & Left(Cob(2).text, InStr(Cob(2).text, " ") - 1) & "'"
            Else
                strSql = strSql & " and a.库房编号='" & Cob(2).text & "'"
            End If
        End If
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        Cob(3).Clear
        If Not rs.EOF Then
            For i = 1 To rs.RecordCount
                Cob(3).AddItem Trim$("" & rs!工单号)
                rs.MoveNext
            Next
        End If
        rs.Close
    End If
    If Index = 9 Then
        '加载料号
        strSql = " SELECT a.SHIP_TO 出货地 FROM erptemp..customer_information a WHERE a.CUSTOMER = '" & Cob(0).text & "'"
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        Cob(9).Clear
        If Not rs.EOF Then
            For i = 1 To rs.RecordCount
                Cob(9).AddItem Trim$("" & rs!出货地)
                rs.MoveNext
            Next
        Else
            MsgBox "出货地加载失败，请联系系统管理员！", vbInformation, "提示"
            Exit Sub
        End If
        rs.Close
    End If

End Sub



Private Sub LoadDn()
Dim rs      As New ADODB.Recordset
Dim i As Integer
Dim date1 As String
Dim date2 As String

Dim strSql As String

   If Cob(0).text = "37" Then
        With Fps_Ship(1)
            .MaxRows = 0
        End With
        With Fps_Ship(0)
            .MaxRows = 0
        End With

        date1 = Format(dtShipDate_Ship.Value, "YYYY/MM/DD")
        date2 = Format(dtShipDate_Ship.Value + 1, "YYYY/MM/DD")

        'merry20200224 加载可出货的DN
        strSql = "SELECT  DELIVERY ,sum(QUANTITY) AS QTY FROM ERPBASE..tblCustomerShippingUp  WHERE  LASTUPDATEDATE >='" & date1 & "' and LASTUPDATEDATE <'" & date2 & "'  GROUP BY DELIVERY ORDER BY DELIVERY"

        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        List_dn(0).Clear
        List_dn(1).Clear
        If Not rs.EOF Then
             For i = 1 To rs.RecordCount
                If GetDnFlag(Trim(rs("DELIVERY")), rs("QTY")) = "待出货" Then
                    List_dn(0).AddItem Trim$("" & rs!Delivery) & "#" & rs!QTY
                End If
                If GetDnFlag(Trim(rs("DELIVERY")), rs("QTY")) = "未完成" Then
                    List_dn(1).AddItem Trim$("" & rs!Delivery) & "#" & rs!QTY
                End If
                
                rs.MoveNext
             Next
         End If
         rs.Close
    Else
        List_dn(0).Clear
        List_dn(1).Clear
    
    End If
   
  End Sub
Private Sub LoadSO()
Dim rs      As New ADODB.Recordset
Dim i As Integer
Dim date1 As String
Dim date2 As String

Dim strSql As String

   If Cob(0).text <> "SG005" Then
       Exit Sub
   End If
   If Trim(Cob(10).text) = "" Then
       Exit Sub
   End If
   
    With Fps_Ship(1)
        .MaxRows = 0
    End With
    With Fps_Ship(0)
        .MaxRows = 0
    End With
    'merry20200224 加载可出货的SO
    strSql = "SELECT a.SO_NO,a.SO_LINE,a.SO_QTY  FROM ERPBASE..tblCustomerShippingUp_SO a WHERE NOT EXISTS (SELECT 1 FROM erpdata..tblShipOrder_Dn WHERE SO_NO =a.SO_NO AND SO_LINE=a.SO_LINE)  and  RTRIM(a.DEVICE) ='" & Trim(Cob(10).text) & "' order by a.PSD,a.SO_NO,a.SO_LINE"
  
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Cob(11).Clear
    If Not rs.EOF Then
         For i = 1 To rs.RecordCount
            Cob(11).AddItem Trim$("" & rs!SO_NO) & "#" & rs!SO_LINE & "#" & rs!SO_QTY
            rs.MoveNext
         Next
     End If
     rs.Close
    
  End Sub
Private Function GetDnFlag(dn As String, QTY As Long)
    Dim rs        As New ADODB.Recordset
    Dim i         As Integer
    Dim strSql    As String

    
    On Error Resume Next
    
    GetDnFlag = "待出货"
     '未完成挑料的不显示
     ' strSql = "SELECT  a.DELIVERY ,a.QUANTITY AS QTY_DN ,a.BATCHNUMBER ,sum(nvl(b.QTY,0))  as QTY_STOCK ,nvl(b.FLAG,0) as FLAG  FROM CUSTOMERSHIPPINGUPTBL  a  LEFT JOIN packing_detailed  b  ON  a.DELIVERY=b.DN_NUM AND a.BATCHNUMBER=b.JOB_ID " & _
     ' "WHERE a.DELIVERY='" & dn & "' GROUP BY  a.DELIVERY ,a.QUANTITY ,a.BATCHNUMBER ,b.FLAG  ORDER BY a.BATCHNUMBER"
    
     ' rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
     ' rs.MoveFirst
     ' If rs.RecordCount > 0 Then
          ' For i = 1 To rs.RecordCount
              ' If rs("QTY_DN") <> rs("QTY_STOCK") Or rs("FLAG") = 0 Then
                  ' GetDnFlag = "挑料中"
                  ' Exit Function
              ' End If
              ' rs.MoveNext
          ' Next
     ' Else
          ' GetDnFlag = "待挑料"
          ' Exit Function
     ' End If
     ' rs.Close

   strSql = "select isnull(sum(b.数量),0) as 库存数量 from erpdata..tblStockNumTree a,erpdata..tblStockNumSub b where  a.箱号=b.箱号 and b.库房编号 in ('07','16') and a.DN='" & dn & "'"
   rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
   If rs.RecordCount > 0 Then
      rs.MoveFirst
       If rs("库存数量") <> QTY Then
           GetDnFlag = "未完成"
           'Exit Function
       End If
   Else
       GetDnFlag = "未完成"
       'Exit Function
   End If
   rs.Close
    
    '已申请出货的，且数量=dn数量的不显示
    strSql = "select isnull(sum(quality),0) as 已申请数量 from [erpdata].[dbo].[tblShipOrder_Dn] where  dn='" & dn & "'"
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
       If rs("已申请数量") >= QTY Then
           GetDnFlag = "已完成"
           Exit Function
       End If
    End If
    rs.Close
    
    
    Set rs = Nothing
End Function

Private Sub fps_Ship_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i               As Long
Dim j               As Integer
Dim strBigbox       As String
Dim intID           As Long
Dim strDN       As String
Dim strCustPN       As String

    'Fps点击事件
    If Index <> 0 Then Exit Sub
    If Row < 1 Then Exit Sub
    If Col <> 1 Then Exit Sub
    
    With Fps_Ship(0)
        .Row = Row
        .Col = FpsM.E_CHOOSE '点击勾选列
        If .Value = 0 Then '选择列
            .Col = FpsM.e_ID        '主表ID
            intID = Val(Trim$(.text))
            .Col = FpsM.e_BigX      '大箱号
            strBigbox = Trim$(.text)
            .Col = FpsM.E_DN      'dn
            strDN = Trim$(.text)
             .Col = FpsM.E_GG      '规格
            strCustPN = Trim$(.text)
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsM.e_BigX
                If strBigbox = Trim$(.text) Then
                    .Col = FpsM.E_CHOOSE
                    .Value = 1
                    .Col = FpsM.E_GG      '规格
                    If strCustPN <> Trim$(.text) Then
                        MsgBox "不同机种不可装于同一个大箱", vbInformation, "提示"
                        Exit Sub
                    End If
                    '修改字体颜色
                    .Col = -1
                    .ForeColor = &HFF8080
                    
                End If
            Next
            '查询明细信息（调取大箱号信息）
            Call SearchDetail(strBigbox, strDN, 1)
            '查询是否有收货客户
            Call SerachSHKH(intID, 1)
            '查询出货地址 mwl 2017.12.25 add
            Call SearchSHDZ
        Else    '取消
            .Col = FpsM.e_ID        '主表ID
            intID = Val(Trim$(.text))
            .Col = FpsM.e_BigX '大箱号
            strBigbox = Trim$(.text)
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsM.e_BigX
                If strBigbox = Trim$(.text) Then
                    .Col = FpsM.E_CHOOSE
                    .Value = 0
                    '修改字体颜色
                    .Col = -1
                    .ForeColor = vbBlack
                End If
            Next
            '删除明细信息（删除取消的大箱号信息）
            Call SearchDetail(strBigbox, strDN, 2)
            '查询是否有收货客户
            Call SerachSHKH(intID, 2)
            '查询出货地址 mwl 2017.12.25 add
            Call SearchSHDZ
        End If
    End With
    '计算勾选的箱数和总箱数
    Call CalcBoxNum
    
End Sub
'查询收货地址
Private Sub SearchSHDZ()
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
Dim i                   As Long
Dim strCustTmp          As String
Dim strGDHTmp           As String

    With Fps_Ship(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = FpsM.E_CHOOSE
            If .Value = 1 Then
                .Col = FpsM.E_cust      '客户代码
                strCustTmp = Trim$(.text)
                .Col = FpsM.e_GDH       '工单号
                strGDHTmp = strGDHTmp + "," + Trim$(.text)
            End If
        Next
        If InStr(strGDHTmp, ",") > 0 Then
            strGDHTmp = Mid$(strGDHTmp, 2, Len(strGDHTmp) - 1)
        End If
    End With
    '查询资料库
    strSql = "SELECT ShipTo,COUNT(DISTINCT ShipTo) BS FROM erpdata..tblSale_Shipto WHERE CustCode='" + strCustTmp + "' AND LotID IN('" + Replace$(strGDHTmp, ",", "','") + "')" & _
             " GROUP BY ShipTo"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
        If Val(Trim$("" & rs!BS)) > 1 Then
            'txt(4).Text = "地址不同"
        Cob(9).text = "地址不同"
        Else
          '  txt(4).Text = Trim("" & Rs!ShipTo)
          Cob(9).text = Trim("" & rs!ShipTo)
        End If
    Else
        'txt(4).Text = ""
    End If
    rs.Close
End Sub
'计算勾选的箱数和总箱数
Private Sub CalcBoxNum()
Dim i               As Long
Dim strBox          As String
Dim strBoxDetail    As String
Dim intBoxNum       As Integer
Dim intTotalBox     As Integer
    
    strBox = ""
    strBoxDetail = ""
    intBoxNum = 0
    intTotalBox = 0
    With Fps_Ship(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = FpsM.e_BigX '大箱号
            If InStr(strBox, Trim$(.text)) <= 0 Then
                strBox = strBox & Trim$(.text) & ","
                intTotalBox = intTotalBox + 1
            End If
            .Col = FpsM.E_CHOOSE '选择
            If .Value = 1 Then
                .Col = FpsM.e_BigX '大箱号
                If InStr(strBoxDetail, Trim$(.text)) <= 0 Then
                    strBoxDetail = strBoxDetail & Trim$(.text) & ","
                    intBoxNum = intBoxNum + 1
                End If
            End If
        Next
    End With
    lbl(19).Caption = Trim$(intBoxNum) & "/" & Trim$(intTotalBox)
    
End Sub

Private Sub SerachSHKH(intID As Long, intBJ As Integer)
'查询收货客户
Dim i               As Long
Dim strSql          As String
Dim rs              As New ADODB.Recordset
    
    '查询资料
    If intBJ = 1 And Trim(Cob(7).text) = "" Then
        strSql = "SELECT * FROM erpdata..Vw_AutoSearchSHKH Where ID=" & intID & ""
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If rs.RecordCount > 0 Then
            Cob(7).text = Trim$("" & rs!收货客户)
        End If
        rs.Close
    End If
    
    If intBJ = 2 And Trim(Cob(7).text) = "" Then '删除明细时，查询选择的那些项目收货客户
        With Fps_Ship(0)
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsM.E_CHOOSE    '选择
                If Val(.text) = 1 Then
                    .Col = FpsM.e_ID    '主表ID
                    strSql = "SELECT * FROM erpdata..Vw_AutoSearchSHKH Where ID=" & Val(Trim$(.text)) & ""
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
                    If rs.RecordCount > 0 Then
                        Cob(7).text = Trim$("" & rs!收货客户)
                        Exit For
                    End If
                    rs.Close
                End If
            Next
        End With
    End If
    
End Sub
Private Sub SearchDetail(strBigbox As String, strDN As String, intBJ As Integer)
'查询箱号信息
Dim i               As Long
Dim strSql          As String
Dim rs              As New ADODB.Recordset
    
    If intBJ = 1 Then '查询明细
        With Fps_Ship(1)
            '先判定传过来的大箱号在这边有没有加载过了
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsD.e_BigX '大箱号
                If strBigbox = Trim$(.text) Then
                    Exit Sub
                End If
            Next
        End With
        '查询资料
        strSql = "SELECT '" & strBigbox & "' 大箱号,a.箱号,b.工单号,b.流程卡编号,b.料号" & _
                 ",CASE WHEN b.合格标记=0 THEN b.数量 ELSE 0 END 合格数 " & _
                 ",CASE WHEN b.合格标记=2 THEN b.数量 ELSE 0 END 来料不良数 " & _
                 ",CASE WHEN b.合格标记=1 THEN b.数量 ELSE 0 END 制程不良数 " & _
                 ",b.id,b.库房编号 " & _
                 " FROM dbo.f_getChild_1('" & strBigbox & "') a " & _
                 " INNER JOIN erpdata..tblStockNumSub b ON a.箱号 = b.箱号"
        If rs.State = adStateOpen Then rs.Close
        rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If rs.RecordCount > 0 Then
            With Fps_Ship(1)
                '向Fps中加载资料
                For i = 1 To rs.RecordCount
                    .MaxRows = .MaxRows + 1
                    .SetText FpsD.e_BigX, .MaxRows, Trim$("" & rs!大箱号)
                    .SetText FpsD.E_XH, .MaxRows, Trim$("" & rs!箱号)
                    .SetText FpsD.e_GDH, .MaxRows, Trim$("" & rs!工单号)
                    .SetText FpsD.e_LCK, .MaxRows, Trim$("" & rs!流程卡编号)
                    .SetText FpsD.e_LH, .MaxRows, Trim$("" & rs!料号)
                    .SetText FpsD.e_GNum, .MaxRows, Trim$("" & rs!合格数)
                    .SetText FpsD.e_BLNum, .MaxRows, Trim$("" & rs!来料不良数)
                    .SetText FpsD.e_ZCNum, .MaxRows, Trim$("" & rs!制程不良数)
                    .SetText FpsD.e_ID, .MaxRows, Trim$("" & rs!id)
                    .SetText FpsD.e_KF, .MaxRows, Trim$("" & rs!库房编号)
                    .SetText FpsD.E_DN, .MaxRows, strDN
    
                    rs.MoveNext
                Next
            End With
        End If
        rs.Close
    End If
    
    If intBJ = 2 Then '删除明细
        With Fps_Ship(1)
            Set .DataSource = Nothing
            For i = .MaxRows To 1 Step -1
                .Row = i
                .Col = FpsD.e_BigX '大箱号
                If strBigbox = Trim$(.text) Then    '箱号相等删除
                    .DeleteRows i, 1
                    .MaxRows = .MaxRows - 1
                End If
            Next
        End With
    End If
    '计算数量
    Call CalcNum
End Sub
'计算数量
Private Sub CalcNum()
Dim i               As Long
Dim j               As Integer
Dim lngGNum         As Long
Dim lngBLNum        As Long
Dim lngZCNum        As Long
    
    lngGNum = 0
    lngBLNum = 0
    lngZCNum = 0
    With Fps_Ship(1)
        For i = 1 To .MaxRows
            .Row = i
            .Col = FpsD.e_GNum '合格数
            lngGNum = lngGNum + Val(Trim$(.text))
            .Col = FpsD.e_BLNum '来料不良数
            lngBLNum = lngBLNum + Val(Trim$(.text))
            .Col = FpsD.e_ZCNum '制程不良数
            lngZCNum = lngZCNum + Val(Trim$(.text))
        Next
    End With
    '赋值到界面上
    txt(6).text = lngGNum
    txt(7).text = lngBLNum
    txt(8).text = lngZCNum
    
    If Cob(11).text <> "" Then '显示SO数量，已选择数量，及剩余数量
        lbl(23).Caption = lngGNum + lngBLNum + lngZCNum & "/" & lsolineqty
        If lngGNum + lngBLNum + lngZCNum > lsolineqty Then
            lbl(23).ForeColor = &HFF&
        Else
            lbl(23).ForeColor = &H8000&
            lbl(23).Caption = lbl(23).Caption & ",余量:" & lsolineqty - (lngGNum + lngBLNum + lngZCNum)
     
        End If
    End If
End Sub

'计算数量
Private Function CheckQtyByDn()
Dim i               As Long
Dim j               As Integer
Dim lngGNum         As Long
Dim lngBLNum        As Long
Dim lngZCNum        As Long
Dim lngqty_dn       As Long
Dim strDN           As String
    
    CheckQtyByDn = ""
    
    
    With List_dn(0)
        For i = 0 To .ListCount - 1
            strDN = ""
            lngqty_dn = 0
            If .Selected(i) = True Then
                strDN = Split(Trim$("" & .List(i)), "#")(0)
                lngqty_dn = Val(Split(Trim$("" & .List(i)), "#")(1))
                lngGNum = 0
                lngBLNum = 0
                lngZCNum = 0
                With Fps_Ship(1)
                    For j = 1 To .MaxRows
                        .Row = j
                        .Col = FpsD.E_DN '合格数
                        If Trim(.text) = strDN Then
                            .Col = FpsD.e_GNum '合格数
                            lngGNum = lngGNum + Val(Trim$(.text))
                            .Col = FpsD.e_BLNum '来料不良数
                            lngBLNum = lngBLNum + Val(Trim$(.text))
                            .Col = FpsD.e_ZCNum '制程不良数
                            lngZCNum = lngZCNum + Val(Trim$(.text))
                         End If
                    Next
                    
                End With
                
                If lngGNum + lngBLNum + lngZCNum <> lngqty_dn Then
                    CheckQtyByDn = strDN
                    MsgBox strDN & "  数量有误,请核对 ", vbInformation, "提示"
                    Exit Function
                End If
            End If
        Next
    End With

    


    
End Function
Private Function CheckQtyBySO()
Dim i               As Long
Dim j               As Integer
Dim lngGNum         As Long
Dim lngBLNum        As Long
Dim lngZCNum        As Long

    CheckQtyBySO = ""
    
    lngGNum = 0
    lngBLNum = 0
    lngZCNum = 0
    With Fps_Ship(1)
        For j = 1 To .MaxRows
            .Row = j
            .Col = FpsD.e_GNum '合格数
            lngGNum = lngGNum + Val(Trim$(.text))
            .Col = FpsD.e_BLNum '来料不良数
            lngBLNum = lngBLNum + Val(Trim$(.text))
            .Col = FpsD.e_ZCNum '制程不良数
            lngZCNum = lngZCNum + Val(Trim$(.text))

        Next
        
    End With
    
    If lngGNum + lngBLNum + lngZCNum > lsolineqty Then
        CheckQtyBySO = "选择数量超出SOline数量"
        MsgBox " 选择数量超出SOline数量,请核对 ", vbInformation, "提示"
        Exit Function
    End If


    


    
End Function


'删除或修改赋值到Fps
Public Sub GiveFps(strdjbh As String, intBJ As Integer)
Dim i           As Long
Dim strSql      As String
Dim rs          As New ADODB.Recordset
    
    strSql = "select d.工单号,b.客户代码,dbo.f_getparent(f.箱号) 大箱号,A.发货地址,a.备注,a.运输方式,a.领料部门,b.客户名称,a.单据编号,a.序号,a.物料编号,c.料号,isnull(c.规格型号,'') as 规格,c.型号," & _
            " c.计量单位名称 单位,a.领料员,a.单据日期,isnull(a.审核,'') as 审核, isnull(a.审核部门,'') as 审核部门,SUM(f.数量) 数量,c.物料名称,dbo.usp_date(isnull(审核日期,'')) as 审核日期 ,e.库房代码+' '+e.库房名称 库房,a.id,a.收货客户 " & _
            " FROM erpdata..tblStockSQfh AS a " & _
            " inner join erpdata..tblStockSQfhsub f on a.单据编号=f.单据编号 and a.序号=f.单据项次" & _
            " INNER JOIN dbo.tblXCustomer AS b ON a.客户代码 = b.客户代码 " & _
            " INNER JOIN dbo.tblSmainM2 AS c ON a.物料编号 = c.物料编号  " & _
            " inner join erpdata..tblstock e on a.仓库编号=e.库房代码" & _
            " INNER JOIN erpdata..tblstocknum d on a.id=d.id where a.单据编号='" & strdjbh & "' and a.接收标记=0 and a.单据类型=1" & _
            " group by d.工单号,b.客户代码,dbo.f_getparent(f.箱号),A.发货地址,a.备注,a.运输方式,a.领料部门,b.客户名称,a.单据编号,a.序号,a.物料编号,c.料号,isnull(c.规格型号,''),c.型号," & _
            " c.计量单位名称,a.领料员,a.单据日期,isnull(a.审核,''), isnull(a.审核部门,''),c.物料名称,dbo.usp_date(isnull(审核日期,'')) ,e.库房代码,e.库房名称,a.id,a.收货客户" & _
            " order by 3"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount = 0 Then
      Screen.MousePointer = 0
      MsgBox "您选择的入库单已被库房接收，请刷新当前界面重新操作！", vbInformation, Me.Caption
      Exit Sub
   End If
   If rs.RecordCount > 0 Then
        Me.Tag = intBJ
        Me.Caption = "发货清单" & IIf(intBJ = 1, "(新增)", IIf(intBJ = 3, "(删除)", "(修改)"))
        txt(0).text = strdjbh
        txt(1).text = Trim("" & rs("领料员"))
        txt(2).text = Format(Trim("" & rs("单据日期")), "YYYY-MM-DD")
        txt(3).text = Trim("" & rs("领料部门"))
        Cob(0).text = Trim("" & rs("客户代码"))
        Cob(6).text = Trim("" & rs("运输方式"))
        'txt(4).Text = Trim("" & Rs("发货地址"))
        Cob(9).text = Trim("" & rs("发货地址"))
        txt(5).text = Trim("" & rs("备注"))
        Cob(2).text = Trim("" & rs("库房"))
        Cob(7).text = Trim$("" & rs("收货客户"))
        rs.MoveFirst
        With Fps_Ship(0)
            For i = 1 To rs.RecordCount
                .MaxRows = .MaxRows + 1
                .SetText FpsM.e_ID, .MaxRows, Trim$("" & rs!id)
                .SetText FpsM.E_cust, .MaxRows, Trim$("" & rs!客户代码)
                .SetText FpsM.e_GDH, .MaxRows, Trim$("" & rs!工单号)
                .SetText FpsM.e_BigX, .MaxRows, Trim$("" & rs!大箱号)
                .SetText FpsM.e_LH, .MaxRows, Trim$("" & rs!料号)
                .SetText FpsM.e_NUM, .MaxRows, Trim$("" & rs!数量)
                .SetText FpsM.E_GG, .MaxRows, Trim$("" & rs!规格)
                .SetText FpsM.E_XH, .MaxRows, Trim$("" & rs!型号)
                .SetText FpsM.E_UNIT, .MaxRows, Trim$("" & rs!单位)
                .SetText FpsM.e_KF, .MaxRows, Trim$("" & rs!库房)
                rs.MoveNext
            Next
        End With
        '显示窗体
        Screen.MousePointer = 0
        Me.Show vbModal
    End If
    rs.Close
End Sub





Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key

    Case "QUERY_Ship"
        QueryData_Ship

    Case "SAVE_Ship"
        SaveData_Ship
        
    Case "DEL_Ship"
        DelData_Ship
      
    Case "HOME_Ship"
        exitFrm

End Select
End Sub


Private Sub DelData_Ship()
On Error GoTo ErrHandle
Dim strdjbh              As String
Dim i                   As Long
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
Dim strsql1              As String
Dim strSql2              As String
Dim strSql3              As String
Dim strSql4              As String
Dim strSql5              As String
Dim DelCnt               As Integer



   ' MsgBox "待更新", vbInformation, "提示"
   ' Exit Sub
   DelCnt = 0
    With Fps_Ship_del(0)
    '校验
    For i = 1 To .MaxRows
        
        .Row = i
        .Col = 1
        If .text = "1" Then
            DelCnt = DelCnt + 1
            .Col = 4
            strdjbh = Trim(.text)
            If Get_SqlserverCnt(" select 1 from erpdata..tblStockSQfh where 单据编号='" & strdjbh & "' and  isnull(财务审核,'')<>''") > 0 Then
                MsgBox strdjbh & "单据已经被财务审核，不能进行删除！", vbInformation, "提示"
                Exit Sub
            End If
            If Get_SqlserverCnt(" select 1 from erpdata..tblStockSQfhsub where 单据编号='" & strdjbh & "' and 开票标记=1 ") > 0 Then
                MsgBox strdjbh & "单据已经开票，不能进行删除或修改！", vbInformation, "提示"
                Exit Sub
            End If
        End If
    Next
    If DelCnt = 0 Then
        MsgBox "请选择需要删除的单据！", vbInformation, "提示"
        Exit Sub
    
    End If
    '删除
     For i = 1 To .MaxRows
        
        .Row = i
        .Col = 1
        If .text = "1" Then
            .Col = 4
            strdjbh = Trim(.text)
            If MsgBox("你确认要删除" & strdjbh & "吗?", vbOKCancel, "提示") = vbCancel Then
            
            Else
                AddSql2 ("insert into erpdata..tblStockSQfh_bak select GETDATE(),'删除','" & gUserName & "', a.* from erpdata..tblStockSQfh a where a.单据编号='" & strdjbh & "'")
                AddSql2 ("insert into erpdata..tblStockSQfhsub_bak select GETDATE(),'删除','" & gUserName & "', a.* from erpdata..tblStockSQfhsub a where a.单据编号='" & strdjbh & "'")

                AddSql2 ("DELETE FROM erpdata..tblStockSQfh where 单据编号='" & strdjbh & "'")
                AddSql2 ("DELETE FROM erpdata..tblStockSQfhsub where 单据编号='" & strdjbh & "'")
                AddSql2 ("DELETE FROM erpdata..tblStockMoveRec WHERE 单据编号='" & strdjbh & "'")
                AddSql2 ("DELETE FROM erptemp..tblshipreport_new WHERE ship_order='" & strdjbh & "'")
                AddSql2 ("DELETE FROM erpdata..tblshiporder_dn WHERE shiporder='" & strdjbh & "'")
                
                MsgBox "已成功删除单据编号:" & strdjbh, vbInformation, "提示"
            End If
            
        End If
    Next
    End With
    ListOrderNumData

Exit Sub
ErrHandle:
    MsgBox "执行失败！" + Chr(13) + "原因:" + Err.DESCRIPTION, vbExclamation, Me.Caption
End Sub
Private Sub QueryData_Ship()
On Error GoTo ErrHandle
Dim i                   As Long
Dim j                   As Long
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
Dim strTmpID            As String
Dim strTmpDXH           As String
Dim IsHave              As Boolean
Dim strDNList           As String
Dim strSOList           As String

    If Cob(0).text = "" Then
        MsgBox "请选择客户代码！", vbInformation, "提示"
        Exit Sub
    End If
    strDNList = ""
    strSOList = ""
    With List_dn(0)
    If Cob(0).text = "37" Then
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
                If strDNList = "" Then
                    strDNList = Split(Trim$("" & .List(i)), "#")(0)
                Else
                    strDNList = strDNList & "','" & Split(Trim$("" & .List(i)), "#")(0)
                End If
            End If
        Next
    ElseIf Cob(0).text = "SG005" Then
    End If
    End With

        

    
    
    If Cob(0).text = "37" And (strDNList <> "" Or Cob(8).text <> "") Then
           Cob(6).text = "01 航空运输"
           Cob(7).text = "37"
           Cob(9).text = "37"
           
    ElseIf Cob(0).text = "SG005" And strsono <> "" Then
           Cob(6).text = "10 货代"
           Cob(7).text = "SG005"
           Cob(9).text = GetSqlServerStr("Select End_Customer_Label_cust_name from ERPBASE..tblCustomerShippingUp_SO where so_no='" & strsono & "' and so_line='" & strsoline & "'")
    Else
       If Cob(2).text = "" Then
           MsgBox "请选择库房名称！", vbInformation, "提示"
           Exit Sub
       End If
     End If
    strSql = "SELECT 1 AS 选择,a.ID,a.客户代码,a.工单号,dbo.f_getparent(b.箱号) 大箱号,a.料号,d.规格,d.型号,d.计量单位名称 单位" & _
             ",SUM(b.数量) 数量,c.库房代码+' '+c.库房名称 库房 "
    If strDNList <> "" Or Cob(8).text <> "" Then

        strSql = strSql + ", f.dn as DN "
    Else
        strSql = strSql + ", '' as DN "
    End If
         
    strSql = strSql + "  FROM erpdata..tblStockNum a INNER JOIN erpdata..tblStockNumSub b ON a.id=b.ID " & _
     " INNER JOIN tblstock c ON a.库房编号=c.库房代码 " & _
     " INNER JOIN tblSmainM2 d ON a.物料编号=d.物料编号 " & _
     " LEFT JOIN erpdata..tblWithWork e ON a.订单编号=e.订单编号 AND a.订单项次=e.订单项次  " & _
     " INNER JOIN erpbase..tblmappingData g ON b.流程卡编号=g.SUBSTRATEID AND b.工单号=g.lotid  " & _
     " INNER JOIN ERPBASE..TBLCUSTOMEROI h ON g.lotid=h.SOURCE_BATCH_ID AND g.filename=convert(varchar(20),h.id) "
     
     
    '根据市场部需求增加37客户DN查询条件 2017.2.14 add mwl------
    
    If strDNList <> "" Then
        strSql = strSql + " INNER JOIN erpdata..tblStockNumTree f ON b.箱号=f.箱号 AND f.DN in('" & strDNList & "')"
    Else
        If Cob(8).text <> "" Then
           strSql = strSql + " INNER JOIN erpdata..tblStockNumTree f ON b.箱号=f.箱号 AND f.DN in('" & Cob(8).text & "')"
        End If
    End If

    '----------------------------------------------------------
    strSql = strSql + " Where a.库存数 > 0 "
    If Trim$(Cob(0).text) <> "" Then
        strSql = strSql & " And a.客户代码='" & Trim(Cob(0).text) & "'"
    End If
    If Val(Trim$(Cob(1).text)) > 0 Then
        strSql = strSql & " And a.产线标记='" & Val(Trim(Cob(1).text)) & "'"
    End If
    If Trim$(Cob(2).text) <> "" Then
        strSql = strSql & " And a.库房编号='" & Left(Trim(Cob(2).text), InStr(Trim(Cob(2).text), " ") - 1) & "'"
    Else
        If Cob(0).text = "37" And (strDNList <> "" Or Cob(8).text <> "") Then
           strSql = strSql & " And a.库房编号 in ('07','16')"
        End If
        If Cob(0).text = "SG005" And strsono <> "" Then
           If UCase(Right(Trim$(Cob(10).text), 2)) = "-D" Then 'merry20200408以-D结尾代表不良品，在30仓查找库存
               strSql = strSql & " And a.库房编号 in ('30')"
           Else
               strSql = strSql & " And a.库房编号 in ('07','20')"
           End If
        End If
    End If
    If Trim$(Cob(3).text) <> "" Then
        strSql = strSql & " And a.工单号 Like '" & Trim(Cob(3).text) & "%'"
    End If
    If Trim$(Cob(4).text) <> "" Then
        strSql = strSql & " And a.料号 Like '" & Trim(Cob(4).text) & "%'"
    End If
    If Trim$(Cob(5).text) <> "" Then
        strSql = strSql & " And e.po_num Like '" & Trim(Cob(5).text) & "%'"
    End If
    If Trim$(Cob(10).text) <> "" Then
        If UCase(Right(Trim$(Cob(10).text), 2)) = "-D" Then 'merry20200408以-D结尾代表不良品，把-D去除匹配客户机种
            strSql = strSql & " And rtrim(h.MPN_DESC) = '" & Left(Trim(Cob(10).text), Len(Trim(Cob(10).text)) - 2) & "'"
        Else
            strSql = strSql & " And rtrim(h.MPN_DESC) = '" & Trim(Cob(10).text) & "'"
        End If
    End If
    
    strSql = strSql & " GROUP BY a.ID,a.客户代码,a.工单号,dbo.f_getparent(b.箱号),a.料号,d.规格,d.型号,d.计量单位名称,c.库房代码,c.库房名称"
    
    If strDNList <> "" Or Cob(8).text <> "" Then

        strSql = strSql + " , f.dn order by 12, 5"
    Else
        strSql = strSql + " order by  5"
    End If
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If chk.Value = 0 Then '没有勾选以选择的信息
        Fps_Ship(0).MaxRows = 0
        Fps_Ship(1).MaxRows = 0
        txt(6).text = ""
        txt(7).text = ""
        txt(8).text = ""
    Else
        With Fps_Ship(0)
            Set .DataSource = Nothing
            '先删除没有勾选的信息
            For i = .MaxRows To 1 Step -1
                .Row = i
                .Col = FpsM.E_CHOOSE
                If .Value = 0 Then
                    .DeleteRows i, 1 '删除行
                    .MaxRows = .MaxRows - 1
                End If
            Next
        End With
    End If
    '赋值数据到Fps控件
    
    If rs.RecordCount > 0 Then
        With Fps_Ship(0)
            rs.MoveFirst
            For i = 1 To rs.RecordCount
                IsHave = False
                '先判定已勾选的资料是否已经存在 （ID和大箱号）
                If chk.Value = 1 Then
                    For j = 1 To .MaxRows
                        .Row = j
                        .Col = FpsM.e_ID 'ID
                        strTmpID = Trim$(.text)
                        .Col = FpsM.e_BigX  '大箱号
                        strTmpDXH = Trim$(.text)
                        If Trim$("" & rs!id) = strTmpID And Trim$("" & rs!大箱号) = strTmpDXH Then
                            IsHave = True
                            Exit For
                        End If
                    Next
                End If
                If IsHave = False Then  '不存在已经勾选的信息，新增以后进去
                    .MaxRows = .MaxRows + 1
                    .SetText FpsM.e_ID, .MaxRows, Trim$("" & rs!id)
                    .SetText FpsM.E_cust, .MaxRows, Trim$("" & rs!客户代码)
                    .SetText FpsM.e_GDH, .MaxRows, Trim$("" & rs!工单号)
                    .SetText FpsM.e_BigX, .MaxRows, Trim$("" & rs!大箱号)
                    .SetText FpsM.e_LH, .MaxRows, Trim$("" & rs!料号)
                    .SetText FpsM.e_NUM, .MaxRows, Trim$("" & rs!数量)
                    .SetText FpsM.E_GG, .MaxRows, Trim$("" & rs!规格)
                    .SetText FpsM.E_XH, .MaxRows, Trim$("" & rs!型号)
                    .SetText FpsM.E_UNIT, .MaxRows, Trim$("" & rs!单位)
                    .SetText FpsM.e_KF, .MaxRows, Trim$("" & rs!库房)
                    .SetText FpsM.E_DN, .MaxRows, Trim$("" & rs!dn)
                End If
                rs.MoveNext
            Next
        End With
    End If
    rs.Close
    '计算勾选的箱数和总箱数
    Call CalcBoxNum
    If Cob(0).text = "SG005" Then
        GetOptSelection '按line数量自动勾选
    End If
Exit Sub
ErrHandle:
    MsgBox "执行失败！" + Chr(13) + "原因:" + Err.DESCRIPTION, vbExclamation, Me.Caption
End Sub





Private Sub SentMesToStock(strText As String)
'发送邮件给仓库
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
Dim strRealName As String
Dim strSql As String
Dim strFileName As String





If gUserName = "07885" Then
  '   Exit Sub
End If
strSql = "select EmpName from XTW..employee where empno = '" & gUserName & "'"
strRealName = Get_SqlStr2(strSql)

i = 0


If Cob(0).text = "37" Then
    dirtemp = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\邮件接收\SentTo_ShipByDn.cfg"
    strSentTitle = Cob(0).text & "客户出货(安排" & Format(dtShipDate_Ship.Value, "YYYY/MM/DD") & "出货）---请审核"
    If mailflag = 1 Then '出货自动邮件
        strSentText = "内勤部" & strRealName & ",工号:" & gUserName & "," & Format(dtShipDate_Ship.Value, "YYYY/MM/DD") & "新增出货" & strText
    ElseIf mailflag = 2 Then '修改出货时间自动邮件
        strSentText = "内勤部" & strRealName & ",工号:" & gUserName & ", 以下DN修改出货时间为" & txtShipDate_Mod.text & strText
    End If
ElseIf Cob(0).text = "SG005" Then
    dirtemp = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\邮件接收\SentTo_ShipBySO.cfg"
    strSentTitle = Format(dtShipDate_Ship.Value, "YYYY/MM/DD") & " " & Cob(0).text & "出货" & Cob(9).text & "   贴关务发票号标签 "
    strSentText = "内勤部" & strRealName & ",工号:" & gUserName & "," & strText
Else
    dirtemp = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\邮件接收\SentTo_ShipByDn.cfg"
    strSentTitle = Format(Now(), "YYYY/MM/DD") & " " & Cob(0).text & "出货" & Cob(9).text
    strSentText = "内勤部" & strRealName & ",工号:" & gUserName & "," & strText
End If
If gUserName = "07885" Then
    dirtemp = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\邮件接收\SentTo_ShipByDn_test.cfg"
End If

Open dirtemp For Input As #1

Do While Not EOF(1)
    Line Input #1, strTemp
    strSentTo(i) = Trim$(strTemp)
    i = i + 1
Loop
Close #1

strSentCC(0) = ""
strSentCC(1) = ""
If SentMess(strSentTitle, strSentText, strSentTo, strFileName, strSentCC) = True Then
    MsgBox "邮件已发送", vbInformation, Me.Caption
Else
    MsgBox "邮件发送失败", vbCritical, Me.Caption

End If

End Sub


Public Function SentMess(Subject As String, SentText As String, Recipient() As String, Attachment As String, RecipientCC() As String) As Boolean

If gUserName = "07885" Then
    SentMess = True
    'Exit Function
End If

    Dim JM As Object

    Set JM = CreateObject("JMAIL.Message")
    
    Dim Recipients()   As String

    Dim RecipientCCs() As String

    Dim strBodyinfo    As String

    Dim i              As Integer

    Dim strSql         As String

    Dim j              As Integer

    Dim rs             As New ADODB.Recordset

    Dim RsD            As New ADODB.Recordset

    On Error GoTo ErrHandler

    SentMess = False

    JM.Charset = "gb2312"
    JM.Silent = False
    JM.Priority = 1
    JM.MailServerUserName = "sqladmin" '帐号
    JM.MailServerPassWord = "ksitadmin" '密码
    JM.From = "sqladmin@ht-tech.com"    '名字
    JM.FromName = "sqladmin"  '发件人姓名
    
    '收件人
    For i = 0 To UBound(Recipient) - 1
        If Recipient(i) <> "" Then
            JM.AddRecipient Recipient(i)
        End If
        
    Next
 
    '抄送人
    For i = 0 To UBound(RecipientCC) - 1
        If RecipientCC(i) <> "" Then
            JM.AddRecipientCC RecipientCC(i)
        End If
        
    Next
    
    '附件
    If Attachment <> "" Then
        If Dir(Attachment, vbNormal Or vbArchive) = "" Then
            Exit Function
        Else
            JM.AddAttachment Attachment

        End If

    End If
    
    JM.Subject = Subject
    'JM.AppendText SentText
    JM.HTMLBody = SentText
    SentMess = JM.Send("mail.ht-tech.com")
    
ErrHandler:
    Set JM = Nothing
    Exit Function

End Function

'-----------------------------------------------------------


Private Sub uploadSO(strFileName As String)

    On Error GoTo uploadSO_ErrON

    Dim VBExcel As Excel.Application
    Dim xlBook  As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim i       As Integer
    Dim j       As Integer
    Dim strso   As String
    Dim fSoLine  As Double
    Dim strShipDate As String
    Dim strCustDevice As String
    Dim lQty    As Long
    Dim strChar As String
    Dim strSOid    As Long
    Dim struploadid    As Long
    Dim insertsqlstr As String
    Dim rs    As New ADODB.Recordset
    Dim strTemp As String
    Dim strSql As String
    Dim strSql2 As String
    
    
    Set VBExcel = CreateObject("excel.application")
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(strFileName)
    Set xlSheet = xlBook.Worksheets(1)
    
    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 43 Then
        MsgBox "SO列数不对", vbInformation, "警告"
        GoTo uploadSO_Err
                        
    End If


    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
       
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count

            If j > 26 Then
                strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
            Else
                strChar = Chr(96 + j)

            End If

            ' If i = 1 Then
                ' If Fps(0).MaxRows = 0 Then
                    ' .SetText j, .MaxRows, Trim$(Replace(Replace(Replace(xlSheet.Range(strChar & i).Value, ",", " "), "，", " "), "'", " "))

                ' End If

            ' Else
                ' .SetText j, .MaxRows, Trim$(Replace(Replace(Replace(xlSheet.Range(strChar & i).Value, ",", " "), "，", " "), "'", " "))

            ' End If

            If i > 1 Then
                If j = 4 Then
                    If Trim$(xlSheet.Range(strChar & i)) = "" Then
                        MsgBox "SO#不可以为空", vbCritical, "警告"
                        GoTo uploadSO_Err
                    
                    End If
                
                    strso = Trim$(xlSheet.Range(strChar & i))

                ElseIf j = 5 Then
                    If Trim$(xlSheet.Range(strChar & i)) = "" Then
                        MsgBox "SO_LINE不可以为空", vbCritical, "警告"
                        GoTo uploadSO_Err
                    
                    End If
                    fSoLine = Trim$(xlSheet.Range(strChar & i))

                ElseIf j = 3 Then
                    If Trim$(xlSheet.Range(strChar & i)) = "" Then
                        MsgBox "PSD不可以为空", vbCritical, "警告"
                        GoTo uploadSO_Err
                    
                    End If
                
                    strShipDate = Trim$(xlSheet.Range(strChar & i))
                
                ElseIf j = 7 Then
                    If Trim$(xlSheet.Range(strChar & i)) = "" Then
                        MsgBox "DEVICE不可以为空", vbCritical, "警告"
                        GoTo uploadSO_Err
                    
                    End If
                    strCustDevice = Trim$(xlSheet.Range(strChar & i))
                
                ElseIf j = 6 Then
                    If Trim$(xlSheet.Range(strChar & i)) = "" Then
                        MsgBox "SO_QTY不可以为空", vbCritical, "警告"
                        GoTo uploadSO_Err
                    
                    End If
                    lQty = CLng(Trim$(xlSheet.Range(strChar & i)))

                End If

            End If

        Next j
        strSql = "select * from erpbase..tblCustomerShippingUp_So where SO_NO='" & strso & "'  and SO_LINE='" & fSoLine & "'"
        If Get_SqlserverCnt(strSql) > 0 Then
        '直接以新的覆盖旧的
            strSql = "delete  from erpbase..tblCustomerShippingUp_So where SO_NO='" & strso & "'  and SO_LINE='" & fSoLine & "'"
            AddSql2 (strSql)
            strSql = "delete  from CUSTOMERSHIPPINGUPTBL_SO where SO_NO='" & strso & "'  and SO_LINE='" & fSoLine & "'"
            AddSql (strSql)
           ' MsgBox strso & " " & fSoLine & "已上传过，请不要重复上传", vbInformation, "提示"
          '  GoTo uploadSO_Err

        End If
    Next i
    
    '开始上传
  '  strSOid = Get_SqlserverNo("select SHIPPING_SO.nextval ID from dual")
  
     struploadid = Get_SqlserverNo("select max(UPLOADID) as maxid from erpbase..tblCustomerShippingUp_So") + 1
    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        strSOid = Get_SqlserverNo("select max(ID) as maxid from erpbase..tblCustomerShippingUp_So") + 1
        insertsqlstr = "'" & struploadid & "','" & strSOid & "'"
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count

            If j > 26 Then
                strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
            Else
                strChar = Chr(96 + j)
            End If
            strTemp = "'" & Trim$(Replace(Replace(Replace(xlSheet.Range(strChar & i).Value, ",", " "), "，", " "), "'", " ")) & "'"
            insertsqlstr = insertsqlstr & " ," & strTemp
            
        Next j
        strSql = " insert into erpbase..tblCustomerShippingUp_So(UPLOADID,ID, ISSUE_DATE, SUBCON, PSD, SO_NO, SO_LINE, SO_QTY, DEVICE, Sales_Part_ID, Customer_PN, " & _
        " Customer_PO, packing, Description, Forwarder, Ship_to_Address, Country, ZIP, End_Customer_Label_cust_name, End_Customer_SHIP_Name,  " & _
        " End_Customer_attention, End_Customer_TEL, End_Customer_FAX, REMARK_1, REMARK_2, Express_Account, TERM, Ship_To_Location, Label_Version," & _
        " Revision_Index, Additional_Part_Information, Supplier_ID, Customer_Label_Code, Manufacturer_Number, Ship_To_City, Ship_To_Postal_Code, " & _
        " Customer_Part_Description, Customer_Dock_Code, Material_Handling_Code, Ship_to_Address1, Ship_to_Address2, Ship_to_Address3, OVT_Place_of_Origin, Sales_Region, Ship_To_Customer) values("
        strSql = strSql & insertsqlstr & ")"
       
         strSql2 = " insert into CUSTOMERSHIPPINGUPTBL_SO(UPLOADID,ID, ISSUE_DATE, SUBCON, PSD, SO_NO, SO_LINE, SO_QTY, DEVICE, Sales_Part_ID, Customer_PN, " & _
        " Customer_PO, packing, Description, Forwarder, Ship_to_Address, Country, ZIP, End_Customer_Label_cust_name, End_Customer_SHIP_Name,  " & _
        " End_Customer_attention, End_Customer_TEL, End_Customer_FAX, REMARK_1, REMARK_2, Express_Account, TERM, Ship_To_Location, Label_Version," & _
        " Revision_Index, Additional_Part_Information, Supplier_ID, Customer_Label_Code, Manufacturer_Number, Ship_To_City, Ship_To_Postal_Code, " & _
        " Customer_Part_Description, Customer_Dock_Code, Material_Handling_Code, Ship_to_Address1, Ship_to_Address2, Ship_to_Address3, OVT_Place_of_Origin, Sales_Region, Ship_To_Customer) values("
        strSql2 = strSql2 & insertsqlstr & ")"
       
             
       
       If AddSql2(strSql) = 0 Then
           GoTo uploadSO_Err
       End If
       If AddSql(strSql2) = 0 Then
           GoTo uploadSO_Err
        End If
    Next i
'显示在界面上
    Set rs = Get_SqlserveRs("select * from erpbase..tblCustomerShippingUp_So where UPLOADid='" & struploadid & "'")
  fps(0).MaxRows = 0
    If rs.RecordCount > 0 Then
        With fps(0)
            Set .DataSource = rs
        End With
    End If
    If Not VBExcel Is Nothing Then
        VBExcel.Application.DisplayAlerts = False '关闭文档不弹出提示框
        xlBook.Close
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing
        Set xlBook = Nothing

   End If
   MsgBox "上传成功", vbInformation, "提示"

    Exit Sub
uploadSO_Err:
    MsgBox Err.DESCRIPTION & vbCrLf & "in 正式工程1.Frm_uploadShippingList.showFps_SG005", vbExclamation + vbOKOnly, "Application Error"
    If Not VBExcel Is Nothing Then
        VBExcel.Application.DisplayAlerts = False '关闭文档不弹出提示框
        xlBook.Close
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing

    End If

    Exit Sub
uploadSO_ErrON:
    GoTo uploadSO_Err
   ' MsgBox Err.DESCRIPTION & vbCrLf & "in 正式工程1.Frm_uploadShippingList.showFps_SG005", vbExclamation + vbOKOnly, "Application Error"

End Sub






















