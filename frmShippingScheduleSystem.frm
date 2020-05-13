VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form frmShippingScheduleSystem 
   BackColor       =   &H00E0E0E0&
   Caption         =   "出货计划系统"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   17205
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "菜单选项(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   0
      TabIndex        =   11
      Top             =   840
      Width           =   20535
      Begin VB.ComboBox cbPO 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   3840
         TabIndex        =   64
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "每日入库量自动出货"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   63
         Top             =   1680
         Width           =   2775
      End
      Begin VB.ComboBox cbCXBJ 
         BackColor       =   &H00FFC0FF&
         Height          =   315
         ItemData        =   "frmShippingScheduleSystem.frx":0000
         Left            =   6960
         List            =   "frmShippingScheduleSystem.frx":000D
         TabIndex        =   62
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtWOID 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3840
         TabIndex        =   58
         Top             =   188
         Width           =   1935
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   17280
         Top             =   480
      End
      Begin VB.ComboBox cbProductNO 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1080
         TabIndex        =   45
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox cbCustPN 
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1080
         TabIndex        =   44
         Top             =   525
         Width           =   1935
      End
      Begin VB.ComboBox cbHTPN 
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1080
         TabIndex        =   43
         Top             =   855
         Width           =   1935
      End
      Begin VB.TextBox txtShipDate 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1253
         Width           =   1935
      End
      Begin VB.TextBox txtAdd 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   9000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   18
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox cbShipGoodOrNot 
         BackColor       =   &H00E0E0E0&
         Caption         =   "是否出不良品"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   17
         Top             =   1590
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.ComboBox cbShipBy 
         BackColor       =   &H00FFC0FF&
         Height          =   315
         Left            =   6960
         TabIndex        =   16
         Top             =   893
         Width           =   1935
      End
      Begin VB.ComboBox cbShipAddr 
         BackColor       =   &H00FFC0FF&
         Height          =   315
         Left            =   6960
         TabIndex        =   15
         Top             =   533
         Width           =   1935
      End
      Begin VB.ComboBox cbShipTo 
         BackColor       =   &H00FFC0FF&
         Height          =   315
         Left            =   6960
         TabIndex        =   14
         Top             =   188
         Width           =   1935
      End
      Begin VB.TextBox txtCustLot 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3840
         TabIndex        =   13
         Top             =   533
         Width           =   1935
      End
      Begin VB.ComboBox cbCustCode 
         BackColor       =   &H00FFC0FF&
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
         Left            =   1080
         TabIndex        =   12
         Top             =   188
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "产线标记"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6000
         TabIndex        =   61
         Top             =   1612
         Width           =   840
      End
      Begin VB.Label lblSJ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "双击->"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5400
         TabIndex        =   59
         Top             =   1298
         Width           =   540
      End
      Begin VB.Label lblWOID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3240
         TabIndex        =   57
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间"
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
         Left            =   16200
         TabIndex        =   54
         Top             =   1905
         Width           =   420
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日期"
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
         Left            =   16200
         TabIndex        =   53
         Top             =   1545
         Width           =   420
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "员工"
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
         Left            =   16080
         TabIndex        =   52
         Top             =   1185
         Width           =   420
      End
      Begin VB.Label lblSysDate 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "当前时间:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   15120
         TabIndex        =   51
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Line Line2 
         X1              =   11160
         X2              =   14520
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Shape Shape1 
         Height          =   1815
         Left            =   11160
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lblQtyPecs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   12720
         TabIndex        =   50
         Top             =   1680
         Width           =   165
      End
      Begin VB.Label LabPecs 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "累计片数(Wafer &PCS):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   11280
         TabIndex        =   49
         Top             =   1320
         Width           =   2100
      End
      Begin VB.Label lblProductNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "成品料号"
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
         Left            =   120
         TabIndex        =   46
         Top             =   1245
         Width           =   900
      End
      Begin VB.Label lblShipDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出货日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   6000
         TabIndex        =   42
         Top             =   1290
         Width           =   840
      End
      Begin VB.Label lblCreater 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "制单日期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   1
         Left            =   15120
         TabIndex        =   40
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblCreater 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "制单员:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   0
         Left            =   15120
         TabIndex        =   39
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   12720
         TabIndex        =   38
         Top             =   720
         Width           =   165
      End
      Begin VB.Label lblShippingQty 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "当前计划累计DIE数(DIE &PCS):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   11280
         TabIndex        =   37
         Top             =   360
         Width           =   2835
      End
      Begin VB.Label lblHTPN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "厂内机种"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   27
         Top             =   915
         Width           =   840
      End
      Begin VB.Label lblShipBy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "其他备注"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   9000
         TabIndex        =   26
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lblShipBy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "货运方式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   6000
         TabIndex        =   25
         Top             =   945
         Width           =   840
      End
      Begin VB.Label lblShipAddr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发货地址"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6000
         TabIndex        =   24
         Top             =   585
         Width           =   840
      End
      Begin VB.Label lblShipTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收货客户"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   6000
         TabIndex        =   23
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lblCustPO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P.O"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3360
         TabIndex        =   22
         Top             =   942
         Width           =   315
      End
      Begin VB.Label lblCustLot 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOTID"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3240
         TabIndex        =   21
         Top             =   585
         Width           =   525
      End
      Begin VB.Label lblCustPN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户机种"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   585
         Width           =   840
      End
      Begin VB.Label lblCustCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10335
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   20565
      _ExtentX        =   36274
      _ExtentY        =   18230
      _Version        =   393216
      Style           =   1
      MousePointer    =   1
      TabHeight       =   617
      BackColor       =   14737632
      ForeColor       =   192
      TabCaption(0)   =   "出货计划制定[销售部]"
      TabPicture(0)   =   "frmShippingScheduleSystem.frx":0030
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "出货计划预览[销售部]"
      TabPicture(1)   =   "frmShippingScheduleSystem.frx":004C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_Schedule"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "SOD修改/审核[PMC+销售部]"
      TabPicture(2)   =   "frmShippingScheduleSystem.frx":0068
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame5 
         Caption         =   "SO&D修改"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   9975
         Left            =   -74880
         TabIndex        =   28
         Top             =   465
         Width           =   20175
         Begin VB.CheckBox CheckSOD_MOD 
            Caption         =   "全选/反选"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker DT_MOD 
            Height          =   300
            Left            =   3120
            TabIndex        =   34
            Top             =   330
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   210960385
            CurrentDate     =   43594
         End
         Begin FPSpreadADO.fpSpread fpS_SOD_MOD 
            Height          =   8175
            Left            =   120
            TabIndex        =   60
            Top             =   720
            Width           =   19815
            _Version        =   524288
            _ExtentX        =   34951
            _ExtentY        =   14420
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
            SpreadDesigner  =   "frmShippingScheduleSystem.frx":0084
            TextTip         =   2
            AppearanceStyle =   0
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "修改SOD"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   2400
            TabIndex        =   35
            Top             =   390
            Width           =   630
         End
      End
      Begin VB.Frame Fra_Schedule 
         Caption         =   "出货计划明细(&SHIP)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   9015
         Left            =   -74880
         TabIndex        =   6
         Top             =   465
         Width           =   20295
         Begin VB.TextBox txtShipID 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2760
            TabIndex        =   32
            Top             =   330
            Width           =   1935
         End
         Begin VB.CheckBox Check1 
            Caption         =   "全选/反选"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   10
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chkSchedule_COMPLETED 
            BackColor       =   &H00C0C0C0&
            Caption         =   "已完成"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   9
            Top             =   360
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chkSchedule_ON 
            BackColor       =   &H00C0C0FF&
            Caption         =   "进行中"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6240
            TabIndex        =   8
            Top             =   360
            Value           =   2  'Grayed
            Width           =   975
         End
         Begin VB.CheckBox chkSchedule_OTHER 
            BackColor       =   &H00FFC0FF&
            Caption         =   "明细"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9000
            TabIndex        =   7
            Top             =   360
            Width           =   855
         End
         Begin FPSpreadADO.fpSpread fpS_ShipSchedule 
            Height          =   8055
            Left            =   0
            TabIndex        =   33
            Top             =   720
            Width           =   20175
            _Version        =   524288
            _ExtentX        =   35586
            _ExtentY        =   14208
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
            MaxCols         =   20
            MaxRows         =   0
            SpreadDesigner  =   "frmShippingScheduleSystem.frx":0502
            TextTip         =   2
            AppearanceStyle =   0
         End
         Begin VB.Label lblShipID 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0FF&
            Caption         =   "出货计划ID:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1560
            TabIndex        =   31
            Top             =   382
            Width           =   1155
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "片明细(&WAFERS)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   9015
         Left            =   5640
         TabIndex        =   4
         Top             =   345
         Width           =   14775
         Begin VB.CheckBox Check_Wafers 
            Caption         =   "全选/反选"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1095
         End
         Begin FPSpreadADO.fpSpread fpS_Wafer 
            Height          =   8175
            Left            =   120
            TabIndex        =   56
            Top             =   600
            Width           =   14535
            _Version        =   524288
            _ExtentX        =   25638
            _ExtentY        =   14420
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
            SpreadDesigner  =   "frmShippingScheduleSystem.frx":098A
            TextTip         =   2
            AppearanceStyle =   0
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "工单明细(&WORKORDERS)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   9015
         Left            =   0
         TabIndex        =   2
         Top             =   345
         Width           =   5535
         Begin MSComCtl2.DTPicker DT_Begin 
            Height          =   300
            Left            =   2040
            TabIndex        =   47
            Top             =   210
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   210960385
            CurrentDate     =   43594
         End
         Begin VB.CheckBox chkONSOD 
            BackColor       =   &H00E0E0E0&
            Caption         =   "SOD"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1320
            TabIndex        =   36
            Top             =   262
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox Check_WO 
            Caption         =   "全选/反选"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker DT_End 
            Height          =   300
            Left            =   3600
            TabIndex        =   30
            Top             =   210
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   210960385
            CurrentDate     =   43594
         End
         Begin FPSpreadADO.fpSpread fps_WO 
            Height          =   8175
            Left            =   120
            TabIndex        =   55
            Top             =   600
            Width           =   5295
            _Version        =   524288
            _ExtentX        =   9340
            _ExtentY        =   14420
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
            SpreadDesigner  =   "frmShippingScheduleSystem.frx":0E08
            TextTip         =   2
            AppearanceStyle =   0
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "至"
            Height          =   195
            Left            =   3360
            TabIndex        =   48
            Top             =   270
            Width           =   180
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17205
      _ExtentX        =   30348
      _ExtentY        =   1535
      ButtonWidth     =   2143
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  查  询"
            Key             =   "QUERY"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  生 成"
            Key             =   "CONFIRM"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "A004"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "删除"
            Key             =   "DELETE"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "修改"
            Key             =   "MODIFY"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "批准"
            Key             =   "PASS"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退回"
            Key             =   "CANCEL_PASS"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SOD待审核"
            Key             =   "WAIT_PASS"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SOD已审核"
            Key             =   "SOD_PASS"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "清空"
            Key             =   "CLEAR"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "导出SOD记录"
            Key             =   "EXPORT_SOD"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  退  出"
            Key             =   "EXIT"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   12360
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
               Picture         =   "frmShippingScheduleSystem.frx":1286
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmShippingScheduleSystem.frx":33C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmShippingScheduleSystem.frx":624A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmShippingScheduleSystem.frx":89FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmShippingScheduleSystem.frx":AB36
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmShippingScheduleSystem.frx":D2E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmShippingScheduleSystem.frx":FA9A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmShippingScheduleSystem.frx":12B1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmShippingScheduleSystem.frx":152CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmShippingScheduleSystem.frx":155E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmShippingScheduleSystem.frx":162C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmShippingScheduleSystem.frx":19344
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmShippingScheduleSystem.frx":1BAF6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmShippingScheduleSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim glColorInStock   As Long
Dim glColorInProcess As Long
Dim glColorShip      As Long

Enum E_WO

    E_CHOOSE = 1
    E_WOID
    E_WOCREATEDATE
    E_SOD
    E_QTY
    E_END

End Enum

Enum E_Wafer

    E_CHOOSE = 1
    E_CUSTPN
    E_WOID
    E_BAND
    E_LOTID
    E_WAFERID
    E_WO_QTY
    E_PLAN_QTY
    E_STOCK_QTY
    E_SENT_QTY
    E_CARTON_NO
    E_SHIP
    e_PO
    E_ShipTo
    E_END

End Enum

Enum E_WOSOD

    E_CHOOSE = 1
    E_WOID
    E_CUSTCODE
    E_CUSTPN
    E_CUSTPRODUCT
    E_WOQTY
    E_STOCKQTY
    E_NOSTOCKQTY
    E_OLDSOD
    E_NEWSOD
    E_FLAG
    E_REASON
    E_END

End Enum

Private Sub cbCustCode_Change()
Dim rs          As New ADODB.Recordset
Dim strSql      As String
Dim strCustCode As String

strCustCode = UCase(Trim$(cbCustCode.text))
strSql = "select distinct CUSTOMERPN  from ib_wohistory where customer = '" & strCustCode & "' and CUSTOMERPN is not null"
Set rs = Get_OracleRs(strSql)
cbCustPN.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbCustPN.AddItem Trim(rs("CUSTOMERPN"))
        rs.MoveNext
    Loop

End If

strSql = "select distinct SALESORDER from ib_wohistory where customer = '" & strCustCode & "' and SALESORDER is not null"
Set rs = Get_OracleRs(strSql)
cbHTPN.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbHTPN.AddItem Trim(rs("SALESORDER"))
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub

Private Sub cbCustCode_Click()
Dim rs          As New ADODB.Recordset
Dim strSql      As String
Dim strCustCode As String

strCustCode = UCase(Trim$(cbCustCode.text))
strSql = "select distinct customerptno1  from tbltsvnpiproduct where customershortname = '" & strCustCode & "' and customerptno1 is not null"
Set rs = Get_OracleRs(strSql)
cbCustPN.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbCustPN.AddItem Trim(rs("customerptno1"))
        rs.MoveNext
    Loop

End If

strSql = "select distinct qtechptno  from tbltsvnpiproduct where customershortname = '" & strCustCode & "' and qtechptno is not null"
Set rs = Get_OracleRs(strSql)
cbHTPN.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbHTPN.AddItem Trim(rs("qtechptno"))
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub

Private Sub cbHTPN_Change()
Dim strSql  As String
Dim strHTPN As String
Dim rs      As New ADODB.Recordset

strHTPN = UCase(Trim$(cbHTPN.text))
If "" = strHTPN Then Exit Sub
strSql = "select distinct customerptno1  from tbltsvnpiproduct where qtechptno = '" & strHTPN & "' and customerptno1 is not null"
cbCustPN.text = Get_OracleStr(strSql)
strSql = "select distinct qtechptno2  from tbltsvnpiproduct where qtechptno = '" & strHTPN & "' and qtechptno2 is not null"
Set rs = Get_OracleRs(strSql)
cbProductNO.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbProductNO.AddItem Trim(rs("qtechptno2"))
        cbProductNO.text = Trim(rs("qtechptno2"))
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub

Private Sub cbHTPN_Click()
Dim strSql  As String
Dim strHTPN As String
Dim rs      As New ADODB.Recordset

strHTPN = UCase(Trim$(cbHTPN.text))
strSql = "select distinct customerptno1  from tbltsvnpiproduct where qtechptno = '" & strHTPN & "' and customerptno1 is not null"
cbCustPN.text = Get_OracleStr(strSql)
strSql = "select distinct qtechptno2  from tbltsvnpiproduct where qtechptno = '" & strHTPN & "' and qtechptno2 is not null"
Set rs = Get_OracleRs(strSql)
cbProductNO.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbProductNO.AddItem Trim(rs("qtechptno2"))
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub

Private Sub cbPO_Click()
fpS_wafer.MaxRows = 0
cbShipAddr.text = ""
If InStr(cbPO.text, " ") > 0 Then
    cbShipAddr.text = Split(cbPO.text, " ")(1)
End If
End Sub
Private Sub cbShipAddr_DropDown()
Dim i      As Integer
Dim strSql As String
Dim rs     As New ADODB.Recordset

If cbCustCode.text = "" Then
    MsgBox "请先选择客户代码", vbInformation, "提示"
    Exit Sub

End If

strSql = " SELECT a.SHIP_TO 出货地 FROM erptemp..customer_information a WHERE a.CUSTOMER = '" & cbCustCode.text & "'"
If rs.State = adStateOpen Then rs.Close
rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
cbShipAddr.Clear
If Not rs.EOF Then

    For i = 1 To rs.RecordCount
        cbShipAddr.AddItem Trim$("" & rs!出货地)
        rs.MoveNext
    Next
Else
    MsgBox "出货地加载失败，请联系系统管理员！", vbInformation, "提示"
    Exit Sub

End If

rs.Close

End Sub

Private Sub Check_Wafers_Click()
Dim i As Integer

With fpS_wafer
    If Check_Wafers.Value = 1 Then

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_Wafer.E_CHOOSE
            .text = 1
        Next i

    ElseIf Check_Wafers.Value = 0 Then

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_Wafer.E_CHOOSE
            .text = 0
        Next i

    End If

End With

reflashQty

End Sub

Private Sub Check_WO_Click()
Dim i       As Integer
Dim strWOID As String

With fps_WO
    If Check_WO.Value = 1 Then

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_WO.E_CHOOSE
            .Value = 1
            .Col = -1
            .ForeColor = &HFF8080
            .Col = E_WO.E_WOID
            strWOID = Trim$(.text)
            Call SearchDetail_ByWOID(strWOID, 1)
        Next i

    ElseIf Check_WO.Value = 0 Then

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_WO.E_CHOOSE
            .Value = 0
            .Col = -1
            .ForeColor = vbBlack
            .Col = E_WO.E_WOID
            strWOID = Trim$(.text)
            Call SearchDetail_ByWOID(strWOID, 2)
        Next i

    End If

End With

reflashQty

End Sub

Private Sub CheckSOD_MOD_Click()
Dim i As Integer

With fpS_SOD_MOD
    If CheckSOD_MOD.Value = 1 Then

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_WOSOD.E_CHOOSE
            .text = 1
        Next i

    ElseIf CheckSOD_MOD.Value = 0 Then

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_WOSOD.E_CHOOSE
            .text = 0
        Next i

    End If

End With

End Sub

Private Sub chkONSOD_Click()
If chkONSOD.Value = 1 Then
    DT_End.Visible = True
Else
    DT_End.Visible = False

End If

End Sub

Private Sub chkSchedule_COMPLETED_Click()
showShipScheduleHistory

End Sub

Private Sub chkSchedule_OTHER_Click()
showShipScheduleHistory

End Sub

Private Sub Form_Load()
InitData
InitCtrls

End Sub

Private Sub InitData()
glColorInStock = &HFF00&
glColorInProcess = &H80FFFF
glColorShip = &HFF80FF

End Sub

Private Sub fpS_ShipSchedule_DblClick(ByVal Col As Long, ByVal Row As Long)

With fpS_ShipSchedule
    .Row = Row
    .Col = 2
    txtShipID.text = Trim$(.text)

End With

End Sub

Private Sub fpS_Wafer_Change(ByVal Col As Long, ByVal Row As Long)
reflashQty

End Sub

Private Sub fpS_Wafer_Click(ByVal Col As Long, ByVal Row As Long)
Dim strWaferID As String
Dim strQboxNO  As String
Dim i          As Integer

If Row < 1 Then Exit Sub
If Col <> E_Wafer.E_CHOOSE Then Exit Sub

With fpS_wafer
    .Row = Row
    .Col = E_Wafer.E_CHOOSE
    If .Value = 1 Then
        .Value = 0
        .Col = E_Wafer.E_WAFERID
        strWaferID = Trim$(.text)
        .Col = E_Wafer.E_CARTON_NO
        strQboxNO = Trim$(.text)

        For i = 1 To .MaxRows
            .Row = i
            '            .Col = E_Wafer.e_WaferID
            '
            '            If strWaferID = Trim$(.Text) Then
            '                .Col = E_Wafer.e_Choose
            '                .Value = 0
            '
            '            End If
            .Col = E_Wafer.E_CARTON_NO
            If strQboxNO <> "" And strQboxNO <> "WIP" Then
                If strQboxNO = Trim$(.text) Then
                    .Col = E_Wafer.E_CHOOSE
                    .Value = 0

                End If

            End If

        Next
    Else
        .Value = 1
        .Col = E_Wafer.E_WAFERID
        strWaferID = Trim$(.text)
        .Col = E_Wafer.E_CARTON_NO
        strQboxNO = Trim$(.text)

        For i = 1 To .MaxRows
            .Row = i
            '            .Col = E_Wafer.e_WaferID
            '
            '            If strWaferID = Trim$(.Text) Then
            '                .Col = E_Wafer.e_Choose
            '                .Value = 1
            '
            '            End If
            .Col = E_Wafer.E_CARTON_NO
            If strQboxNO <> "" And strQboxNO <> "WIP" Then
                If strQboxNO = Trim$(.text) Then
                    .Col = E_Wafer.E_CHOOSE
                    .Value = 1

                End If

            End If

        Next

    End If

End With

reflashQty

End Sub

Private Sub fpS_Wafer_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim strWaferID As String
Dim i          As Integer

If Row < 1 Then Exit Sub

With fpS_wafer
    .Row = Row
    .Col = E_Wafer.E_CHOOSE
    If .Value = 1 Then
        .Col = E_Wafer.E_WAFERID
        strWaferID = Trim$(.text)

        For i = 1 To .MaxRows
            Call MoveSingerWafer(strWaferID)
        Next

    End If

End With

reflashQty

End Sub

Private Function MoveSingerWafer(strWaferID As String)
Dim i As Integer

With fpS_wafer

    For i = 0 To .MaxRows
        .Col = E_Wafer.E_WAFERID
        .Row = i
        If Trim$(.text) = strWaferID Then
            .DeleteRows i, 1
            .MaxRows = .MaxRows - 1

        End If

    Next

End With

End Function

Private Sub fpS_WO_Change(ByVal Col As Long, ByVal Row As Long)
Dim i       As Long
Dim j       As Integer
Dim strWOID As String

If Row < 1 Then Exit Sub

With fps_WO
    .Row = Row
    .Col = E_WO.E_CHOOSE
    If .Value = 0 Then
        .Value = 1
        .Col = -1
        .ForeColor = &HFF8080
        .Col = E_WO.E_WOID
        strWOID = Trim$(.text)
        Call SearchDetail_ByWOID(strWOID, 1)
    Else
        .Value = 0
        .Col = -1
        .ForeColor = vbBlack
        .Col = E_WO.E_WOID
        strWOID = Trim$(.text)
        Call SearchDetail_ByWOID(strWOID, 2)

    End If

End With

End Sub

Private Sub fpS_WO_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim i       As Long
Dim j       As Integer
Dim strWOID As String

If Row < 1 Then Exit Sub

With fps_WO
    .Row = Row
    .Col = E_WO.E_CHOOSE
    If .Value = 0 Then
        .Value = 1
        .Col = -1
        .ForeColor = &HFF8080
        .Col = E_WO.E_WOID
        strWOID = Trim$(.text)
        Call SearchDetail_ByWOID(strWOID, 1)
    Else
        .Value = 0
        .Col = -1
        .ForeColor = vbBlack
        .Col = E_WO.E_WOID
        strWOID = Trim$(.text)
        Call SearchDetail_ByWOID(strWOID, 2)

    End If

End With

End Sub

Private Sub fpS_WO_Click(ByVal Col As Long, ByVal Row As Long)
Dim i       As Long
Dim j       As Integer
Dim strWOID As String

If Row < 1 Then Exit Sub
If Col <> 1 Then Exit Sub

With fps_WO
    .Row = Row
    .Col = E_WO.E_CHOOSE
    If .Value = 0 Then
        .Value = 1
        .Col = -1
        .ForeColor = &HFF8080
        .Col = E_WO.E_WOID
        strWOID = Trim$(.text)
        Call SearchDetail_ByWOID(strWOID, 1)
    Else
        .Value = 0
        .Col = -1
        .ForeColor = vbBlack
        .Col = E_WO.E_WOID
        strWOID = Trim$(.text)
        Call SearchDetail_ByWOID(strWOID, 2)

    End If

End With

End Sub

Private Sub SearchDetail_ByWOID(strWOID As String, intBJ As Integer)
Dim i      As Long
Dim strSql As String
Dim rs     As New ADODB.Recordset

If intBJ = 1 Then

    With fpS_wafer

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_Wafer.E_WOID
            If strWOID = Trim$(.text) Then
                Exit Sub

            End If

        Next

    End With

    '查询资料
    strSql = " select aa.ORDERNAME,ss.CUSTOMERPN,aa.WAFERLOT, aa.waferid,SUM(convert(int,aa.DIEQTY)) as 工单数量, SUM(convert(int,aa.DIEQTY)) - (isnull(SUM(cc.数量),0) - isnull(SUM(dd.数量),0)) as 可发货量, isnull(sum(bb.数量),0) as 库存量,(isnull(SUM(cc.数量),0) - isnull(SUM(dd.数量),0))  as 已发货量, isnull(bb.箱号,'WIP') as 入库箱号, case isnull(bb.箱号,'WIP')  when 'WIP' then '入满即发' else '只发库存'  end as 出货选项" & _
       " from  [erpdata].[dbo].[tblTSVwaferlist] aa inner join [erpdata].[dbo].[tblTSVworkorder] ss on aa.ORDERNAME = ss.ORDERNAME left join erpdata..tblStockNumSub bb on aa.WAFERID = bb.流程卡编号 and aa.WAFERLOT = bb.工单号 and aa.ORDERNAME = bb.大工单 " & " left join erpdata..tblStocksqfhsub cc on aa.WAFERID = cc.流程卡编号 and aa.WAFERLOT = cc.工单号 and aa.ORDERNAME = cc.大工单 and CHARINDEX('F',cc.单据编号) > 0 " & " left join erpdata..tblStocksqfhsub dd on aa.WAFERID = dd.流程卡编号 and aa.WAFERLOT = dd.工单号 and aa.ORDERNAME = dd.大工单 and CHARINDEX('T',dd.单据编号) > 0 " & _
       " where aa.ORDERNAME = '" & strWOID & "' and not exists (select 1 from erptemp..SHIP_PLAN_DETAILED ee where ee.WAFER_ID = aa.WAFERID) group by ss.CUSTOMERPN,aa.WAFERID, bb.箱号, aa.ORDERNAME,aa.WAFERLOT " & " order by aa.WAFERID "
    Set rs = Get_SqlserveRs(strSql)
    If rs.RecordCount > 0 Then

        With fpS_wafer

            For i = 1 To rs.RecordCount
                .MaxRows = .MaxRows + 1
                .SetText E_Wafer.E_CHOOSE, .MaxRows, 1
                .SetText E_Wafer.E_CUSTPN, .MaxRows, Trim$("" & rs!CUSTOMERPN)
                .SetText E_Wafer.E_WOID, .MaxRows, Trim$("" & rs!ORDERNAME)
                If Left$(Trim$("" & rs!ORDERNAME), 1) = "A" Then
                    .SetText E_Wafer.E_BAND, .MaxRows, "保税"
                Else
                    .SetText E_Wafer.E_BAND, .MaxRows, "非保税"

                End If

                .SetText E_Wafer.E_LOTID, .MaxRows, Trim$("" & rs!WAFERLOT)
                .SetText E_Wafer.E_WAFERID, .MaxRows, Trim$("" & rs!waferid)
                .SetText E_Wafer.E_WO_QTY, .MaxRows, Trim$("" & rs!工单数量)
                .SetText E_Wafer.E_PLAN_QTY, .MaxRows, Trim$("" & rs!可发货量)
                .SetText E_Wafer.E_STOCK_QTY, .MaxRows, Trim$("" & rs!库存量)
                .SetText E_Wafer.E_SENT_QTY, .MaxRows, Trim$("" & rs!已发货量)
                .SetText E_Wafer.E_CARTON_NO, .MaxRows, Trim$("" & rs!入库箱号)
                .SetText E_Wafer.E_SHIP, .MaxRows, Trim$("" & rs!出货选项)
                rs.MoveNext
            Next

        End With

    End If

    rs.Close
    Set rs = Nothing

End If

If intBJ = 2 Then

    With fpS_wafer
        Set .DataSource = Nothing

        For i = .MaxRows To 1 Step -1
            .Row = i
            .Col = E_Wafer.E_WOID
            If strWOID = Trim$(.text) Then
                .DeleteRows i, 1
                .MaxRows = .MaxRows - 1

            End If

        Next

    End With

End If

'计算数量
reflashQty

End Sub

Private Sub SearchDetail_ByLotID(strLotID As String)
Dim i      As Long
Dim strSql As String
Dim strpo As String
Dim strShipTo As String
Dim rs     As New ADODB.Recordset
If strLotID = "*" Then
    strLotID = ""
    fpS_wafer.MaxRows = 0
End If

With fpS_wafer

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_Wafer.E_LOTID
        If strLotID = Trim$(.text) Then
            Exit Sub

        End If

    Next

End With

If strLotID <> "" Then
cbCustCode.text = Get_OracleStr("select distinct b.customer from ib_waferlist a, ib_wohistory b where a.ordername = b.ordername and a.waferlot = '" & strLotID & "'")


'查询资料
    strSql = " select aa.ORDERNAME,ss.CUSTOMERPN ,aa.WAFERLOT, aa.waferid,SUM(convert(int,aa.DIEQTY)) as 工单数量, SUM(convert(int,aa.DIEQTY)) - (isnull(SUM(cc.数量),0) - isnull(SUM(dd.数量),0)) as 可发货量, isnull(sum(bb.数量),0) as 库存量,(isnull(SUM(cc.数量),0) - isnull(SUM(dd.数量),0))  as 已发货量, isnull(bb.箱号,'WIP') as 入库箱号, case isnull(bb.箱号,'WIP')  when 'WIP' then '入满即发' else '只发库存'  end as 出货选项,ff.PO_NUM AS PO号,ff.comp_code AS 出货地址 " & _
   " from  [erpdata].[dbo].[tblTSVwaferlist] aa inner join [erpdata].[dbo].[tblTSVworkorder] ss on aa.ORDERNAME = ss.ORDERNAME left join erpdata..tblStockNumSub bb on aa.WAFERID = bb.流程卡编号 and aa.WAFERLOT = bb.工单号 and aa.ORDERNAME = bb.大工单 " & " left join erpdata..tblStocksqfhsub cc on aa.WAFERID = cc.流程卡编号 and aa.WAFERLOT = cc.工单号 and aa.ORDERNAME = cc.大工单 and CHARINDEX('F',cc.单据编号) > 0 " & " left join erpdata..tblStocksqfhsub dd on aa.WAFERID = dd.流程卡编号 and aa.WAFERLOT = dd.工单号 and aa.ORDERNAME = dd.大工单 and CHARINDEX('T',dd.单据编号) > 0 " & _
    " INNER JOIN erpbase..tblmappingData gg ON gg.SUBSTRATEID=aa.WAFERID AND gg.LOTID=aa.WAFERLOT " & _
    " INNER JOIN erpbase..tblCustomerOI ff ON ff.SOURCE_BATCH_ID=gg.LOTID AND convert(VARCHAR(20),ff.ID)=gg.FILENAME  " & _
    " where aa.WAFERLOT = '" & strLotID & "' and not exists (select 1 from erptemp..SHIP_PLAN_DETAILED ee where ee.WAFER_ID = aa.WAFERID)"
Else
    If Trim(cbCustCode.text) = "" Then
        Exit Sub
    End If
    '查询资料
    strSql = " select aa.ORDERNAME,ss.CUSTOMERPN ,aa.WAFERLOT, aa.waferid,SUM(convert(int,aa.DIEQTY)) as 工单数量, SUM(convert(int,aa.DIEQTY)) - (isnull(SUM(cc.数量),0) - isnull(SUM(dd.数量),0)) as 可发货量, isnull(sum(bb.数量),0) as 库存量,(isnull(SUM(cc.数量),0) - isnull(SUM(dd.数量),0))  as 已发货量, isnull(bb.箱号,'WIP') as 入库箱号, case isnull(bb.箱号,'WIP')  when 'WIP' then '入满即发' else '只发库存'  end as 出货选项,ff.PO_NUM AS PO号,ff.comp_code AS 出货地址 " & _
    " from  [erpdata].[dbo].[tblTSVwaferlist] aa inner join [erpdata].[dbo].[tblTSVworkorder] ss on aa.ORDERNAME = ss.ORDERNAME left join erpdata..tblStockNumSub bb on aa.WAFERID = bb.流程卡编号 and aa.WAFERLOT = bb.工单号 and aa.ORDERNAME = bb.大工单 " & " left join erpdata..tblStocksqfhsub cc on aa.WAFERID = cc.流程卡编号 and aa.WAFERLOT = cc.工单号 and aa.ORDERNAME = cc.大工单 and CHARINDEX('F',cc.单据编号) > 0 " & " left join erpdata..tblStocksqfhsub dd on aa.WAFERID = dd.流程卡编号 and aa.WAFERLOT = dd.工单号 and aa.ORDERNAME = dd.大工单 and CHARINDEX('T',dd.单据编号) > 0 " & _
    " INNER JOIN erpbase..tblmappingData gg ON gg.SUBSTRATEID=aa.WAFERID AND gg.LOTID=aa.WAFERLOT " & _
    " INNER JOIN erpbase..tblCustomerOI ff ON ff.SOURCE_BATCH_ID=gg.LOTID AND convert(VARCHAR(20),ff.ID)=gg.FILENAME  " & _
    " where aa.WAFERLOT in ( SELECT distinct b.工单号 FROM erpdata..tblstocknum a , erpdata..tblstocknumsub b WHERE a.客户代码='" & Trim(cbCustCode.text) & "' AND a.id=b.ID  ) and not exists (select 1 from erptemp..SHIP_PLAN_DETAILED ee where ee.WAFER_ID = aa.WAFERID)"
End If
strpo = ""
strShipTo = ""
If Trim(cbPO.text) <> "" And Trim(cbPO.text) <> "所有" Then
    strpo = Split(Trim(cbPO.text), " ")(0)
    strSql = strSql & " and ff.po_num='" & strpo & "'"
    If InStr(Trim(cbPO.text), " ") > 0 Then
        strShipTo = Split(Trim(cbPO.text), " ")(1)
    strSql = strSql & "  and ff.comp_code='" & strShipTo & "'"
    End If
End If
strSql = strSql & " group by ss.CUSTOMERPN,aa.WAFERID, bb.箱号, aa.ORDERNAME,aa.WAFERLOT,ff.PO_NUM,ff.comp_code " & " order by aa.WAFERID "
Set rs = Get_SqlserveRs(strSql)
If rs.RecordCount > 0 Then

    With fpS_wafer

        For i = 1 To rs.RecordCount
            .MaxRows = .MaxRows + 1
            .SetText E_Wafer.E_CHOOSE, .MaxRows, 1
            .SetText E_Wafer.E_CUSTPN, .MaxRows, Trim$("" & rs!CUSTOMERPN)
            .SetText E_Wafer.E_WOID, .MaxRows, Trim$("" & rs!ORDERNAME)
            .SetText E_Wafer.E_LOTID, .MaxRows, Trim$("" & rs!WAFERLOT)
            .SetText E_Wafer.E_WAFERID, .MaxRows, Trim$("" & rs!waferid)
            .SetText E_Wafer.E_WO_QTY, .MaxRows, Trim$("" & rs!工单数量)
            .SetText E_Wafer.E_PLAN_QTY, .MaxRows, Trim$("" & rs!可发货量)
            .SetText E_Wafer.E_STOCK_QTY, .MaxRows, Trim$("" & rs!库存量)
            .SetText E_Wafer.E_SENT_QTY, .MaxRows, Trim$("" & rs!已发货量)
            .SetText E_Wafer.E_CARTON_NO, .MaxRows, Trim$("" & rs!入库箱号)
            .SetText E_Wafer.E_SHIP, .MaxRows, Trim$("" & rs!出货选项)
            .SetText E_Wafer.e_PO, .MaxRows, Trim$("" & rs!PO号)
            .SetText E_Wafer.E_ShipTo, .MaxRows, Trim$("" & rs!出货地址)
            rs.MoveNext
        Next

    End With

Else
    MsgBox "查询不到该LOT可以出货的工单,请确认是否输入错误", vbCritical, "警告"

End If

rs.Close
Set rs = Nothing
'计算数量
reflashQty

End Sub

Private Sub reflashQty()
Dim lQty       As Long
Dim lQty2      As Long
Dim strWaferID As String
Dim i          As Integer

lQty = 0
lQty2 = 0
strWaferID = ""

With fpS_wafer

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_Wafer.E_CHOOSE
        If .Value = 1 Then
            .Col = E_Wafer.E_WAFERID
            If strWaferID = "" Then
                strWaferID = Trim$(.text)
                lQty2 = 1
            Else
                If strWaferID <> Trim$(.text) Then
                    strWaferID = Trim$(.text)
                    lQty2 = lQty2 + 1

                End If

            End If

            .Col = E_Wafer.E_SHIP
            If Trim$(.text) = "只发库存" Then
                .Col = E_Wafer.E_STOCK_QTY
                lQty = lQty + CLng(.text)
            Else
                .Col = E_Wafer.E_PLAN_QTY
                lQty = lQty + CLng(.text)

            End If

        End If

    Next

End With

lblQTY.Caption = lQty
lblQtyPecs.Caption = lQty2

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

Select Case SSTab1.Tab

    Case 0
        Toolbar1.Buttons("CONFIRM").Enabled = True
        Toolbar1.Buttons("QUERY").Enabled = True
        Toolbar1.Buttons("DELETE").Enabled = False
        Toolbar1.Buttons("MODIFY").Enabled = False
        Toolbar1.Buttons("PASS").Enabled = False
        Toolbar1.Buttons("CANCEL_PASS").Enabled = False
        Toolbar1.Buttons("WAIT_PASS").Enabled = False
        Toolbar1.Buttons("SOD_PASS").Enabled = False

    Case 1
        Toolbar1.Buttons("CONFIRM").Enabled = False
        Toolbar1.Buttons("QUERY").Enabled = True
        Toolbar1.Buttons("DELETE").Enabled = True
        Toolbar1.Buttons("MODIFY").Enabled = False
        Toolbar1.Buttons("PASS").Enabled = True
        Toolbar1.Buttons("CANCEL_PASS").Enabled = True
        Toolbar1.Buttons("WAIT_PASS").Enabled = False
        Toolbar1.Buttons("SOD_PASS").Enabled = False
        showShipScheduleHistory

    Case 2
        Toolbar1.Buttons("CONFIRM").Enabled = False
        Toolbar1.Buttons("QUERY").Enabled = True
        Toolbar1.Buttons("DELETE").Enabled = False
        Toolbar1.Buttons("MODIFY").Enabled = True
        Toolbar1.Buttons("MODIFY").Caption = "修改SOD"
        Toolbar1.Buttons("PASS").Enabled = True
        Toolbar1.Buttons("CANCEL_PASS").Enabled = True
        Toolbar1.Buttons("WAIT_PASS").Enabled = True
        Toolbar1.Buttons("SOD_PASS").Enabled = True

End Select

End Sub

Private Sub Timer1_Timer()
lbltime.Caption = Format(Now(), "HH:mm:ss")

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "QUERY"
        Toolbar1.Buttons("QUERY").Enabled = False

        Select Case SSTab1.Tab

            Case 0
                queryDataCreate

            Case 1
                queryDataHistory

            Case 2
                queryDataSOD

        End Select

        Toolbar1.Buttons("QUERY").Enabled = True

    Case "CONFIRM"
        SaveData

    Case "CLEAR"
        ClearData

    Case "WAIT_PASS"
        SSTab1.Tab = 2
        queryDataSOD_WAITPASS
        
    Case "SOD_PASS"
        SSTab1.Tab = 2
        queryDataSOD_PASS

    Case "PASS"

        Select Case SSTab1.Tab

            Case 1
                If gUserName <> "12725" And gUserName <> "07885" And gUserName <> "16642" Then
                    MsgBox "您的工号没有审核的权限", vbInformation, "提示"
                    Exit Sub

                End If

                passData_ShipPlan

            Case 2
                If gUserName <> "12725" And gUserName <> "07885" And gUserName <> "16642" Then
                    MsgBox "您的工号没有审核的权限", vbInformation, "提示"
                    Exit Sub

                End If

                passData_SOD

        End Select

    Case "CANCEL_PASS"
        If gUserName <> "12725" And gUserName <> "07885" And gUserName <> "16642" Then
            MsgBox "您的工号没有反审核/退出的权限", vbInformation, "提示"
            Exit Sub

        End If

        If SSTab1.Tab = 1 Then
            Call nopassData_ShipPlan
        ElseIf SSTab1.Tab = 2 Then
            Call nopassData_SOD
        End If

    Case "MODIFY"
        modData_SOD

    Case "DELETE"

        Select Case SSTab1.Tab

            Case 0

            Case 1
                delShipPlan

        End Select

    Case "EXPORT_SOD"
        
         ExporToExcel ("select a.work_order_id 工单号,b.planenddate 原工单SOD,a.new_sod_date 修改后SOD,a.remark1 修改原因,a.update_by 修改人,a.update_time 修改日期 from PMC_SOD_UPDATE_TBL a  " & _
"inner join ib_workorder b on a.work_order_id = b.ordername " & _
"order by a.update_time desc ")
        
    Case "EXIT"
        Unload Me

End Select

End Sub

Private Sub passData_ShipPlan()
Dim i         As Integer
Dim bSel      As Boolean
Dim strShipID As String
Dim strSql    As String

bSel = False

With fpS_ShipSchedule

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .Value = 1 Then
            bSel = True

        End If

    Next

End With

If bSel = False Then
    MsgBox "请选择需要审核的出货计划", vbInformation, "提示"
    Exit Sub

End If

With fpS_ShipSchedule

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .Value = 1 Then
            .Row = i
            .Col = 2
            strShipID = Trim$(.text)
            .Col = 3
            If .text = "" Then
                strSql = "update erptemp..SHIP_PLAN set APPROVER = '" & gUserName & "',APPROV_TIME = GetDate() where PLAN_ID = '" & strShipID & "'"
                AddSql2 (strSql)
            Else
                MsgBox strShipID & "已经审核过， 请勿重复审核", vbInformation, "提示"

            End If

        End If

    Next

End With

MsgBox "操作已执行完成", vbInformation, "提示"
showShipScheduleHistory

End Sub

Private Sub passData_SOD()
Dim i          As Integer
Dim bSel       As Boolean
Dim strDateMod As String
Dim strSql     As String
Dim strSql2    As String
Dim strWOID    As String

bSel = False

With fpS_SOD_MOD

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_WOSOD.E_CHOOSE
        If .Value = 1 Then
            bSel = True

        End If

    Next

End With

If bSel = False Then
    MsgBox "请选择需要审核的SOD", vbInformation, "提示"
    Exit Sub

End If

With fpS_SOD_MOD

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_WOSOD.E_CHOOSE
        If .Value = 1 Then
            .Col = E_WOSOD.E_WOID
            strWOID = Trim$(.text)
            .Col = E_WOSOD.E_NEWSOD
            strDateMod = Trim$(.text)
            If strDateMod <> "" Then
                strSql2 = "update erptemp..TBL_SHIPPLAN_SODWAITPASS set FLAG = '1',PASS_BY = '" & gUserName & "',PASS_DATE = GETdate()  where ordername = '" & strWOID & "'"
                AddSql2 (strSql2)
                strSql = "update PMC_SOD_UPDATE_TBL set flag = '1' where WORK_ORDER_ID = '" & strWOID & "'"
                AddSql (strSql)
                strSql2 = "update erpdata..tblTSVworkorder set PLANENDDATE = '" & strDateMod & "' where ordername = '" & strWOID & "'"
                AddSql2 (strSql2)
                strSql = "update ib_wohistory  set PLANENDDATE = '" & strDateMod & "' where ordername = '" & strWOID & "'"
                AddSql (strSql)
                ' SOD修改, 抛MES接口
                strSql = "INSERT INTO mes_reference (IDENTIFIER,KEY1,KEY2,KEY3,propertyname,Propertyvalue,Flag )" & "VALUES ('ORDER_SOD','" & strWOID & "',to_char(sysdate,'YYMMDDHH24miss'),'NULL','SOD','" & strDateMod & "',0  )"
                AddSql (strSql)

            End If

        End If

    Next

End With

Call showSOD_WAIT
MsgBox "计划出货日期已经修改并审核通过,请确认", vbInformation, "提示"

End Sub

Private Sub nopassData_ShipPlan()
Dim i         As Integer
Dim bSel      As Boolean
Dim strShipID As String
Dim strSql    As String

bSel = False

With fpS_ShipSchedule

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .Value = 1 Then
            bSel = True

        End If

    Next

End With

If bSel = False Then
    MsgBox "请选择需要反审核的出货计划", vbInformation, "提示"
    Exit Sub

End If

With fpS_ShipSchedule

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .Value = 1 Then
            .Row = i
            .Col = 2
            strShipID = Trim$(.text)
            .Col = 3
            If .text <> "" Then
                strSql = "update erptemp..SHIP_PLAN set APPROVER = '',APPROV_TIME = null where PLAN_ID = '" & strShipID & "'"
                AddSql2 (strSql)
            Else
                MsgBox strShipID & "没有审核过， 无法反审", vbInformation, "提示"

            End If

        End If

    Next

End With

MsgBox "操作已执行完成", vbInformation, "提示"
showShipScheduleHistory

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       nopassData_SOD
' Description:       退出SOD
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/18-17:29:49
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub nopassData_SOD()
Dim bChoose As Boolean
Dim i As Integer
Dim strWOID As String
Dim strSql As String


If SSTab1.Tab <> 2 Then
    Exit Sub
End If

bChoose = False

With fpS_SOD_MOD
    For i = 1 To .MaxRows
        .Row = i
        .Col = E_WOSOD.E_CHOOSE
        If .Value = 1 Then
            .Col = E_WOSOD.E_WOID
            strWOID = Trim$("" & .text)
            strSql = "update erptemp..TBL_SHIPPLAN_SODWAITPASS set flag = '2' where ORDERNAME = '" & strWOID & "'"
            AddSql2 (strSql)
            
            bChoose = True
        End If
    Next
End With

If Not bChoose Then
    MsgBox "请选择需要退回的SOD", vbInformation, "提示"
    Exit Sub
End If

MsgBox "已退回", vbInformation, "提示"

Call showSOD_WAIT

End Sub

Private Sub modData_SOD()
Dim i          As Integer
Dim bSel       As Boolean
Dim strDateMod As String
Dim strSql     As String
Dim strSql2    As String
Dim strSql3    As String
Dim strWOID    As String
bSel = False

With fpS_SOD_MOD

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_WOSOD.E_CHOOSE

        If .Value = 1 Then
            bSel = True

        End If

    Next

End With

If bSel = False Then
    MsgBox "请选择需要更改的SOD" & vbCrLf & "支持批量修改", vbInformation, "提示"
    Exit Sub

End If

'Dialog_SOD_UPDATE.Show 1
strDateMod = Format(DT_MOD.Value, "yyyy-MM-dd")

With fpS_SOD_MOD

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_WOSOD.E_CHOOSE

        If .Value = 1 Then
            .Col = E_WOSOD.E_WOID
            strWOID = Trim$(.text)

            strSql = "update ib_wohistory  set PLANENDDATE = '" & strDateMod & "' where ordername = '" & strWOID & "'"
            AddSql (strSql)
            
            strSql = "update shop_order set plan_end_date =  '" & strDateMod & "' where shop_order = '" & strWOID & "'  "
            AddSql (strSql)
            
            strSql2 = "update erpdata..tblTSVworkorder set PLANENDDATE = '" & strDateMod & "' where ordername = '" & strWOID & "'"
            AddSql2 (strSql2)
            ' SOD修改, 抛MES接口
            strSql = "INSERT INTO mes_reference (IDENTIFIER,KEY1,KEY2,KEY3,propertyname,Propertyvalue,Flag )" & "VALUES ('ORDER_SOD','" & strWOID & "',to_char(sysdate,'YYMMDDHH24miss'),'NULL','SOD','" & strDateMod & "',0  )"
            AddSql (strSql)

        End If

    Next

End With

showSOD
MsgBox "计划出货日期修改已完成", vbInformation, "提示"

End Sub

Private Sub InitCtrls()
SSTab1.Tab = 0
Dim rs     As New ADODB.Recordset, i As Integer
Dim strSql As String

Set rs.ActiveConnection = SqlConnect
rs.Source = "select distinct 客户代码 from tblxcustomer"
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
cbCustCode.Clear
Cbshipto.Clear
If rs.RecordCount > 0 Then
    rs.MoveFirst

    For i = 1 To rs.RecordCount
        cbCustCode.AddItem Trim(rs("客户代码"))
        Cbshipto.AddItem Trim$(rs("客户代码"))
        rs.MoveNext
    Next i

End If

rs.Close
'lblCreater
lblUserName.Caption = gUserName
lblDate.Caption = Format(Now(), "yyyy-MM-dd")
lbltime.Caption = Format(Now(), "HH:mm:ss")
'DT
DT_Begin.Value = Format(Now() - 60, "yyyy-MM-dd")
DT_End.Value = Format(Now() + 60, "yyyy-MM-dd")
DT_MOD.Value = Format(Year(Now()) & "-" & Month(Now()) & "-" & "28", "yyyy-MM-dd")

'fps
With fps_WO
    .MaxCols = E_WO.E_END - 1
    .Col = -1
    .Row = -1
    .Lock = True
    .Col = E_WO.E_CHOOSE
    .CellType = CellTypeCheckBox
    .TypeHAlign = TypeVAlignCenter
    .TypeVAlign = TypeVAlignCenter
    .SetText 0, 0, "序号"
    .ColWidth(0) = 4
    .SetText E_WO.E_CHOOSE, 0, "√"
    .ColWidth(E_WO.E_CHOOSE) = 4
    .SetText E_WO.E_WOID, 0, "工单号"
    .ColWidth(E_WO.E_WOID) = 10
    .SetText E_WO.E_WOCREATEDATE, 0, "开工单日期"
    .ColWidth(E_WO.E_WOCREATEDATE) = 10
    .SetText E_WO.E_SOD, 0, "SOD"
    .ColWidth(E_WO.E_SOD) = 6
    .SetText E_WO.E_QTY, 0, "可出货量"
    .ColWidth(E_WO.E_QTY) = 8
    .Col = E_WO.E_QTY
    .BackColor = glColorInProcess

End With

With fpS_wafer
    .MaxCols = E_Wafer.E_END - 1
    .Col = -1
    .Row = -1
    .Lock = True
    .Col = E_Wafer.E_CHOOSE
    .CellType = CellTypeCheckBox
    .TypeHAlign = TypeVAlignCenter
    .TypeVAlign = TypeVAlignCenter
    .SetText 0, 0, "序号"
    .ColWidth(0) = 4
    .SetText E_Wafer.E_CHOOSE, 0, "√"
    .ColWidth(E_Wafer.E_CHOOSE) = 4
    .SetText E_Wafer.E_CUSTPN, 0, "客户机种"
    .ColWidth(E_Wafer.E_CUSTPN) = 10
    .SetText E_Wafer.E_WOID, 0, "工单号"
    .ColWidth(E_Wafer.E_WOID) = 10
    .SetText E_Wafer.E_BAND, 0, "保税/非保"
    .ColWidth(E_Wafer.E_BAND) = 6
    .SetText E_Wafer.E_LOTID, 0, "LOTID"
    .ColWidth(E_Wafer.E_LOTID) = 10
    .SetText E_Wafer.E_WAFERID, 0, "WAFERID"
    .ColWidth(E_Wafer.E_WAFERID) = 12
    .SetText E_Wafer.E_WO_QTY, 0, "工单数量"
    .ColWidth(E_Wafer.E_WO_QTY) = 8
    .SetText E_Wafer.E_PLAN_QTY, 0, "最大发货量"
    .ColWidth(E_Wafer.E_PLAN_QTY) = 8
    .SetText E_Wafer.E_STOCK_QTY, 0, "库存量"
    .ColWidth(E_Wafer.E_STOCK_QTY) = 8
    .SetText E_Wafer.E_SENT_QTY, 0, "已发货量"
    .ColWidth(E_Wafer.E_SENT_QTY) = 8
    .SetText E_Wafer.E_CARTON_NO, 0, "库存箱号/WIP"
    .ColWidth(E_Wafer.E_CARTON_NO) = 10
    .SetText E_Wafer.E_SHIP, 0, "出货条件"
    .ColWidth(E_Wafer.E_SHIP) = 12
    .SetText E_Wafer.e_PO, 0, "PO"
    .ColWidth(E_Wafer.e_PO) = 12
    .SetText E_Wafer.E_ShipTo, 0, "出货地址"
    .ColWidth(E_Wafer.E_ShipTo) = 12
    .Col = E_Wafer.E_SHIP
    .Lock = False
    .CellType = CellTypeComboBox
    .TypeComboBoxList = .TypeComboBoxList & "入满即发"
    .TypeComboBoxList = .TypeComboBoxList & "只发库存"
    .Col = E_Wafer.E_SHIP
    .BackColor = glColorShip
    .Col = E_Wafer.E_PLAN_QTY
    .BackColor = glColorInProcess
    .Col = E_Wafer.E_STOCK_QTY
    .BackColor = glColorInStock

End With

With fpS_ShipSchedule
    .DAutoCellTypes = False
    .Col = -1
    .Row = -1
    .Lock = True
    .Col = 1
    .Lock = False
    .CellType = CellTypeCheckBox
    .TypeHAlign = TypeVAlignCenter
    .TypeVAlign = TypeVAlignCenter
    .ColWidth(4) = 4
    .ColWidth(5) = 4

End With

With fpS_SOD_MOD
    .MaxCols = E_WOSOD.E_END - 4
    .Col = -1
    .Row = -1
    .Lock = True
    .Col = E_WOSOD.E_CHOOSE
    .CellType = CellTypeCheckBox
    .TypeHAlign = TypeVAlignCenter
    .TypeVAlign = TypeVAlignCenter
    .Lock = False
    .SetText 0, 0, "序号"
    .ColWidth(0) = 4
    .SetText E_WOSOD.E_CHOOSE, 0, "√"
    .ColWidth(E_WOSOD.E_CHOOSE) = 4
    .SetText E_WOSOD.E_WOID, 0, "工单号"
    .ColWidth(E_WOSOD.E_WOID) = 10
    .SetText E_WOSOD.E_CUSTCODE, 0, "客户代码"
    .ColWidth(E_WOSOD.E_CUSTCODE) = 6
    .SetText E_WOSOD.E_CUSTPN, 0, "客户机种"
    .ColWidth(E_WOSOD.E_CUSTPN) = 12
    .SetText E_WOSOD.E_CUSTPRODUCT, 0, "成品料号"
    .ColWidth(E_WOSOD.E_CUSTPRODUCT) = 12
    .SetText E_WOSOD.E_WOQTY, 0, "工单数量"
    .ColWidth(E_WOSOD.E_WOQTY) = 6
    .SetText E_WOSOD.E_STOCKQTY, 0, "入库数量"
    .ColWidth(E_WOSOD.E_STOCKQTY) = 6
    .SetText E_WOSOD.E_NOSTOCKQTY, 0, "未入库数量"
    .ColWidth(E_WOSOD.E_NOSTOCKQTY) = 8
    .SetText E_WOSOD.E_OLDSOD, 0, "当前工单SOD"
    .ColWidth(E_WOSOD.E_OLDSOD) = 12
'    .SetText E_WOSOD.E_NEWSOD, 0, "修改后SOD"
'    .ColWidth(E_WOSOD.E_NEWSOD) = 12
'    .SetText E_WOSOD.E_FLAG, 0, "审核状态"
'    .ColWidth(E_WOSOD.E_FLAG) = 8
'    .SetText E_WOSOD.E_REASON, 0, "修改原因"
'    .ColWidth(E_WOSOD.E_REASON) = 50
    .Col = E_WOSOD.E_OLDSOD
    .BackColor = glColorShip
'    .Col = E_WOSOD.E_NEWSOD
'    .BackColor = glColorInProcess

End With

' 运输方式初始化
strSql = "select RTRIM(运输方式代码)+' '+RTRIM(运输方式名称) 运输方式 from erpdata..tblXTransitMode"
If rs.State = adStateOpen Then rs.Close
rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
cbShipBy.Clear
If Not rs.EOF Then

    For i = 1 To rs.RecordCount
        cbShipBy.AddItem Trim$("" & rs!运输方式)
        rs.MoveNext
    Next
Else
    MsgBox "运输方式加载失败，请联系系统管理员！", vbInformation, "提示"
    Exit Sub

End If

rs.Close

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       queryDataCreate2
' Description:       查询可出货工单
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-16:13:06
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub queryDataCreate()
Dim strCustCode  As String
Dim strCustPN    As String
Dim strHTPN      As String
Dim strProductNO As String
Dim strWOID      As String
Dim strLotID     As String
Dim lWOQty       As Long
Dim lSentQty     As Long
Dim lSentQty2    As Long
Dim lPlanQty     As Long
Dim strSql       As String
Dim rs           As New ADODB.Recordset
Dim strBeginDate As String
Dim strEndDate   As String

strCustCode = Trim$(cbCustCode.text)
strCustPN = Trim$(cbCustPN.text)
strHTPN = Trim$(cbHTPN.text)
strProductNO = Trim$(cbProductNO.text)
strBeginDate = Format(DT_Begin.Value, "yyyy-MM-dd")
strEndDate = Format(DT_End.Value, "yyyy-MM-dd")
'lblQty.Caption = "0"
'lblQtyPecs.Caption = "0"
'fpS_WO.MaxRows = 0
'fpS_Wafer.MaxRows = 0
If strCustCode = "" Then
    MsgBox "请输入客户代码", vbInformation, "提示"
    Exit Sub

End If

If txtCustLot.text <> "" Then
    strLotID = Trim$(txtCustLot.text)
    Call SearchDetail_ByLotID(strLotID)
    Exit Sub

End If

If strCustPN = "" And strHTPN = "" Then
    MsgBox "请输入厂内机种/客户机种" & vbCrLf & "或者直接输入客户批号", vbInformation, "提示"
    Exit Sub

End If

If strHTPN <> "" Then
    If strProductNO <> "" Then
        strSql = "select distinct ORDERNAME,planenddate,erpcreatedate from ib_wohistory where CUSTOMER = '" & strCustCode & "' and salesorder = '" & strHTPN & "' and product = '" & strProductNO & "' and instr(ORDERNAME,'BSM-') = 0 "
    Else
        strSql = "select distinct ORDERNAME,planenddate,erpcreatedate from ib_wohistory where CUSTOMER = '" & strCustCode & "' and salesorder = '" & strHTPN & "' and instr(ORDERNAME,'BSM-') = 0 and instr(ORDERNAME,'BSG-') = 0 "

    End If

Else
    If strProductNO <> "" Then
        strSql = "select distinct ORDERNAME,planenddate,erpcreatedate from ib_wohistory where CUSTOMER = '" & strCustCode & "' and CUSTOMERPN = '" & strCustPN & "' and product = '" & strProductNO & "' and instr(ORDERNAME,'BSM-') = 0 "
    Else
        strSql = "select distinct ORDERNAME,planenddate,erpcreatedate from ib_wohistory where CUSTOMER = '" & strCustCode & "' and CUSTOMERPN = '" & strCustPN & "' and instr(ORDERNAME,'BSM-') = 0 and instr(ORDERNAME,'BSG-') = 0"

    End If

End If

If chkONSOD.Value = 1 Then
    strSql = strSql & " and planenddate >= '" & strBeginDate & "' and planenddate <=  '" & strEndDate & "' order by planenddate "
Else
    strSql = strSql & " and planenddate >= '" & strBeginDate & "' order by planenddate "

End If

Set rs = Get_OracleRs(strSql)
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        strWOID = Trim(rs(0).Value)
        strSql = "SELECT qty FROM erpdata..tblTSVworkorder where ORDERNAME = '" & strWOID & "'"
        lWOQty = Get_SqlserverNo(strSql)
        strSql = "SELECT SUM(数量) FROM erpdata..tblStocksqfhsub where 大工单 = '" & strWOID & "' and 单据编号 like 'F%' "
        lSentQty = Get_SqlserverNo(strSql)
        strSql = "SELECT SUM(数量) FROM erpdata..tblStocksqfhsub where 大工单 = '" & strWOID & "' and 单据编号 like 'T%' "
        lSentQty = lSentQty - Get_SqlserverNo(strSql)
        strSql = "select SUM(GOOD_DIE) from erptemp..SHIP_PLAN_DETAILED where SHOP_ORDER = '" & strWOID & "'  "
        lSentQty2 = Get_SqlserverNo(strSql)
        lPlanQty = lWOQty - lSentQty - lSentQty2
        If lPlanQty > 0 Then

            With fps_WO
                .MaxRows = .MaxRows + 1
                .SetText E_WO.E_CHOOSE, .MaxRows, 0
                .SetText E_WO.E_WOID, .MaxRows, strWOID
                .SetText E_WO.E_WOCREATEDATE, .MaxRows, Trim$("" & rs(2).Value)
                .SetText E_WO.E_SOD, .MaxRows, Trim("" & rs(1).Value)
                .SetText E_WO.E_QTY, .MaxRows, lPlanQty

            End With

        End If

        rs.MoveNext
    Loop
Else
    MsgBox "该机种没有工单开立记录", vbInformation
    Exit Sub

End If

If fps_WO.MaxRows > 0 Then
    MsgBox "工单可出货记录已查询,请展开明细", vbInformation, "提示"
Else
    MsgBox "该机种在该日期段暂无可以出货的工单" & vbCrLf & "请确认是否扩大查询日期段", vbInformation, "提示"

End If

End Sub

Private Sub queryDataHistory()
showShipScheduleHistory

End Sub

Private Sub queryDataSOD()
showSOD

End Sub

Private Sub queryDataSOD_WAITPASS()
showSOD_WAIT

End Sub

Private Sub queryDataSOD_PASS()
showSOD_PASS

End Sub

Private Sub showSOD()
Dim strSql       As String
Dim rs           As New ADODB.Recordset
Dim strCustCode  As String
Dim strCustPN    As String
Dim strCustLotID As String
Dim strWOID      As String
Dim lWOQty       As Long
Dim lStockQty    As Long

If txtWOID.text = "" And txtCustLot.text = "" And cbCustCode.text = "" And cbCustPN.text = "" Then
    MsgBox "请输入工单号或客户批号(LOTID)或客户代码+客户机种", vbInformation, "提示"
    Exit Sub

End If

fpS_SOD_MOD.MaxRows = 0

If txtWOID.text <> "" Then
    strWOID = Trim$(txtWOID.text)

    If Get_OracleCnt("select * from ib_Wohistory where ordername = '" & strWOID & "' ") > 0 Then
        lWOQty = Get_OracleNo("select qty from ib_Wohistory where ordername = '" & strWOID & "'")
        lStockQty = Get_SqlserverNo("select sum(入库数) as QTY from erpdata..tblPackToHouse where 大工单 = '" & strWOID & "'")

        If lStockQty < lWOQty Then
            Call showSODFps(strWOID, lWOQty, lStockQty)

        End If

    Else
        MsgBox "您输入的工单号不存在", vbInformation, "提示"
        Exit Sub

    End If

Else

    If txtCustLot.text <> "" Then
        strCustLotID = Trim$(txtCustLot.text)
        strSql = "select distinct ordername from ib_waferlist where waferlot = '" & strCustLotID & "' "
        Set rs = Get_OracleRs(strSql)

        If Not rs.EOF Then

            Do While Not rs.EOF
                strWOID = Trim$("" & rs.Fields("ordername").Value)
                lWOQty = Get_OracleNo("select qty from ib_Wohistory where ordername = '" & strWOID & "'")
                lStockQty = Get_SqlserverNo("select sum(入库数) as QTY from erpdata..tblPackToHouse where 大工单 = '" & strWOID & "'")

                If lStockQty < lWOQty Then
                    Call showSODFps(strWOID, lWOQty, lStockQty)

                End If

                rs.MoveNext
            Loop
        Else
            MsgBox "您输入的LOTID查询不到开工单记录", vbInformation, "提示"
            Exit Sub

        End If

    Else

        If cbCustCode.text <> "" Then
            If cbHTPN.text <> "" Then
                strSql = "select distinct ordername from ib_wohistory where salesorder = '" & Trim$(cbHTPN.text) & "' and instr(ordername,'BSM-') = 0 and instr(ordername,'BSG-') = 0 "
                Set rs = Get_OracleRs(strSql)

                If Not rs.EOF Then

                    Do While Not rs.EOF
                        strWOID = Trim$("" & rs.Fields("ordername").Value)
                        lWOQty = Get_OracleNo("select qty from ib_Wohistory where ordername = '" & strWOID & "'")
                        lStockQty = Get_SqlserverNo("select sum(入库数) as QTY from erpdata..tblPackToHouse where 大工单 = '" & strWOID & "'")

                        If lStockQty < lWOQty Then
                            Call showSODFps(strWOID, lWOQty, lStockQty)

                        End If

                        rs.MoveNext
                    Loop
                Else
                    MsgBox "您输入的厂内机种查询不到开工单记录", vbInformation, "提示"
                    Exit Sub

                End If

            Else
                strCustCode = Trim$(cbCustCode.text)
                strSql = "select distinct ordername from ib_wohistory where customer = '" & strCustCode & "' and instr(ordername,'BSM-') = 0 and instr(ordername,'BSG-') = 0 "
                Set rs = Get_OracleRs(strSql)

                If Not rs.EOF Then

                    Do While Not rs.EOF
                        strWOID = Trim$("" & rs.Fields("ordername").Value)
                        lWOQty = Get_OracleNo("select qty from ib_Wohistory where ordername = '" & strWOID & "'")
                        lStockQty = Get_SqlserverNo("select sum(入库数) as QTY from erpdata..tblPackToHouse where 大工单 = '" & strWOID & "'")

                        If lStockQty < lWOQty Then
                            Call showSODFps(strWOID, lWOQty, lStockQty)

                        End If

                        rs.MoveNext
                    Loop
                Else
                    MsgBox "您输入的客户代码查询不到开工单记录", vbInformation, "提示"
                    Exit Sub

                End If

            End If

        End If

        If cbHTPN.text <> "" Then
            strSql = "select distinct ordername from ib_wohistory where salesorder = '" & Trim$(cbHTPN.text) & "' and instr(ordername,'BSM-') = 0 and instr(ordername,'BSG-') = 0 "
            Set rs = Get_OracleRs(strSql)

            If Not rs.EOF Then

                Do While Not rs.EOF
                    strWOID = Trim$("" & rs.Fields("ordername").Value)
                    lWOQty = Get_OracleNo("select qty from ib_Wohistory where ordername = '" & strWOID & "'")
                    lStockQty = Get_SqlserverNo("select sum(入库数) as QTY from erpdata..tblPackToHouse where 大工单 = '" & strWOID & "'")

                    If lStockQty < lWOQty Then
                        Call showSODFps(strWOID, lWOQty, lStockQty)

                    End If

                    rs.MoveNext
                Loop
            Else
                MsgBox "您输入的厂内机种查询不到开工单记录", vbInformation, "提示"
                Exit Sub

            End If

        End If

    End If

End If

End Sub

Private Sub showSODFps(strWOID As String, lWOQty As Long, lStockQty As Long)
Dim strSql    As String
Dim strSql2   As String
Dim rs        As New ADODB.Recordset
Dim strReason As String

strSql = " select distinct aa.shop_order,aa.cust_id,'',aa.prd_ID,aa.plan_end_date from shop_order aa,tbltsvnpiproduct bb where aa.prd_ID = bb.qtechptno2 and aa.HT_DEVICE = bb.qtechptno and aa.shop_order = '" & strWOID & "' "
Set rs = Get_OracleRs(strSql)
If Not rs.EOF Then

    Do While Not rs.EOF

        With fpS_SOD_MOD
            .MaxRows = .MaxRows + 1
            .SetText E_WOSOD.E_CHOOSE, .MaxRows, 0
            .SetText E_WOSOD.E_WOID, .MaxRows, Trim$("" & rs(0).Value)
            .SetText E_WOSOD.E_CUSTCODE, .MaxRows, Trim$("" & rs(1).Value)
           ' .SetText E_WOSOD.E_CUSTPN, .MaxRows, Trim$("" & rs(2).Value)
            .SetText E_WOSOD.E_CUSTPRODUCT, .MaxRows, Trim$("" & rs(3).Value)
            .SetText E_WOSOD.E_WOQTY, .MaxRows, lWOQty
            .SetText E_WOSOD.E_STOCKQTY, .MaxRows, lStockQty
            .SetText E_WOSOD.E_NOSTOCKQTY, .MaxRows, lWOQty - lStockQty
            .SetText E_WOSOD.E_OLDSOD, .MaxRows, Trim$("" & rs(4).Value)

            strSql2 = "select remark1 from PMC_SOD_UPDATE_TBL where WORK_ORDER_ID = '" & strWOID & "' "
            strReason = Get_OracleStr(strSql2)
            .SetText E_WOSOD.E_REASON, .MaxRows, strReason

        End With

        rs.MoveNext
    Loop

End If

End Sub

Private Sub showSODFps_Pass(strWOID As String, lWOQty As Long, lStockQty As Long)
Dim strSql    As String
Dim strSql2   As String
Dim rs        As New ADODB.Recordset
Dim strReason As String

strSql = "select distinct aa.ORDERNAME as 工单号,aa.customer as 客户代码,aa.customerpn as 客户机种,aa.product as 成品料号 ,CONVERT(varchar(100), aa.PLANENDDATE, 23) as 原SOD日期 ,dd.SOD as 修改SOD日期,'已审核' as 审核状态 from erpdata..tblTSVworkorder aa " & " inner join erpdata..tblTSVwaferlist bb on aa.ORDERNAME = bb.ORDERNAME " & " left join erptemp..TBL_SHIPPLAN_SODWAITPASS dd on  dd.ORDERNAME = aa.ordername where aa.ORDERNAME = '" & strWOID & "' "
Set rs = Get_SqlserveRs(strSql)
If Not rs.EOF Then

    Do While Not rs.EOF

        With fpS_SOD_MOD
            .MaxRows = .MaxRows + 1
            .SetText E_WOSOD.E_CHOOSE, .MaxRows, 0
            .SetText E_WOSOD.E_WOID, .MaxRows, Trim$("" & rs("工单号").Value)
            .SetText E_WOSOD.E_CUSTCODE, .MaxRows, Trim$("" & rs("客户代码").Value)
            .SetText E_WOSOD.E_CUSTPN, .MaxRows, Trim$("" & rs("客户机种").Value)
            .SetText E_WOSOD.E_CUSTPRODUCT, .MaxRows, Trim$("" & rs("成品料号").Value)
            .SetText E_WOSOD.E_WOQTY, .MaxRows, lWOQty
            .SetText E_WOSOD.E_STOCKQTY, .MaxRows, lStockQty
            .SetText E_WOSOD.E_NOSTOCKQTY, .MaxRows, lWOQty - lStockQty
            .SetText E_WOSOD.E_OLDSOD, .MaxRows, Trim$("" & rs("原SOD日期").Value)
            .SetText E_WOSOD.E_NEWSOD, .MaxRows, Trim$("" & rs("修改SOD日期").Value)
            .SetText E_WOSOD.E_FLAG, .MaxRows, Trim$("" & rs("审核状态").Value)
            strSql2 = "select remark1 from PMC_SOD_UPDATE_TBL where WORK_ORDER_ID = '" & strWOID & "' "
            strReason = Get_OracleStr(strSql2)
            .SetText E_WOSOD.E_REASON, .MaxRows, strReason

        End With

        rs.MoveNext
    Loop

End If

End Sub

Private Sub showSOD_WAIT()
Dim strSql      As String
Dim rs          As New ADODB.Recordset
Dim strWOID     As String
Dim lWOQty      As Long
Dim lStockQty   As Long
Dim strCustCode As String

strCustCode = Trim$(cbCustCode.text)
If strCustCode = "" Then
    strSql = "select aa.ORDERNAME from erptemp .. TBL_SHIPPLAN_SODWAITPASS  aa where aa.FLAG <> 1 "
Else
    strSql = "select aa.ORDERNAME,bb.CUSTOMER from erptemp .. TBL_SHIPPLAN_SODWAITPASS  aa" & " inner join  [erpdata].[dbo].[tblTSVworkorder]  bb on aa.ORDERNAME = bb.ORDERNAME " & " where aa.FLAG <> 1 and bb.CUSTOMER = '" & strCustCode & "' "
    
End If

Set rs = Get_SqlserveRs(strSql)
fpS_SOD_MOD.MaxRows = 0
If Not rs.EOF Then

    Do While Not rs.EOF
        strWOID = Trim$("" & rs.Fields("ordername").Value)
        lWOQty = Get_OracleNo("select qty from ib_Wohistory where ordername = '" & strWOID & "'")
        lStockQty = Get_SqlserverNo("select sum(入库数) as QTY from erpdata..tblPackToHouse where 大工单 = '" & strWOID & "'")
        If lStockQty < lWOQty Then
            Call showSODFps(strWOID, lWOQty, lStockQty)

        End If

        rs.MoveNext
    Loop

End If

If fpS_SOD_MOD.MaxRows = 0 Then
    MsgBox "没有待审核的SOD", vbInformation, "提示"

End If

End Sub

Private Sub showSOD_PASS()
Dim strSql      As String
Dim rs          As New ADODB.Recordset
Dim strWOID     As String
Dim lWOQty      As Long
Dim lStockQty   As Long
Dim strCustCode As String

strCustCode = Trim$(cbCustCode.text)
If strCustCode = "" Then
    strSql = "select aa.ORDERNAME from erptemp .. TBL_SHIPPLAN_SODWAITPASS  aa where aa.FLAG = 1 "
Else
    strSql = "select aa.ORDERNAME,bb.CUSTOMER from erptemp .. TBL_SHIPPLAN_SODWAITPASS  aa" & " inner join  [erpdata].[dbo].[tblTSVworkorder]  bb on aa.ORDERNAME = bb.ORDERNAME " & " where aa.FLAG = 1 and bb.CUSTOMER = '" & strCustCode & "' "
    
End If

Set rs = Get_SqlserveRs(strSql)
fpS_SOD_MOD.MaxRows = 0
If Not rs.EOF Then

    Do While Not rs.EOF
        strWOID = Trim$("" & rs.Fields("ordername").Value)
        lWOQty = Get_OracleNo("select qty from ib_Wohistory where ordername = '" & strWOID & "'")
        lStockQty = Get_SqlserverNo("select sum(入库数) as QTY from erpdata..tblPackToHouse where 大工单 = '" & strWOID & "'")
        If lStockQty < lWOQty Then
            Call showSODFps_Pass(strWOID, lWOQty, lStockQty)

        End If

        rs.MoveNext
    Loop

End If

If fpS_SOD_MOD.MaxRows = 0 Then
    MsgBox "没有待审核的SOD", vbInformation, "提示"

End If

End Sub

Private Sub delShipPlan()
Dim strSql As String

If txtShipID.text = "" Then
    MsgBox "请输入要删除的出货计划ID", vbInformation, "提示"
    Exit Sub

End If

If Get_SqlStr("select approver from erptemp..SHIP_PLAN where  PLAN_ID = '" & Trim(txtShipID.text) & "' and ship_flag <> '0'  ") <> "" Then
    MsgBox "该计划已经审核通过,请联系相关人员反审核,否则无法删除", vbInformation, "提示"
    Exit Sub

End If

strSql = "delete from  erptemp..SHIP_PLAN where PLAN_ID = '" & Trim(txtShipID.text) & "' "
AddSql2 (strSql)
strSql = "delete from erptemp..SHIP_PLAN_DETAILED where PLAN_ID = '" & Trim(txtShipID.text) & "'"
AddSql2 (strSql)
MsgBox "已经删除出货计划" & vbCrLf & Trim(txtShipID.text), vbInformation, "提示"
txtShipID.text = ""
showShipScheduleHistory

End Sub

Private Sub showShipScheduleHistory()
Dim strSql    As String, strsql1 As String, strSql2 As String, strSql3 As String, strSql4 As String, strSql5 As String
Dim rs        As New ADODB.Recordset
Dim strShipID As String
Dim i         As Integer

If chkSchedule_OTHER.Value = 0 Then
    strSql = " SELECT distinct '' as 选择,a.PLAN_ID as 出货计划ID,a.APPROVER as 审核人员,  a.APPROV_TIME as 审核时间,case a.SHIP_FLAG when '0' then '未执行' when '1' then '已执行' else '未知状态' end AS 执行状态,   " & " a.PLAN_DATE as 出货计划日期,a.REMARK2 as 出货单号, a.CUSTOMER AS 客户代码,  a.PLAN_ITEM as ITEM, a.CUST_PART as 客户机种, a.PRODUCT_NAME as 料号, " & " a.PRODUCT_ID as 物料编号, SUM(b.GOOD_DIE) as 出货数量, a.SHIP_AD as 出货地址, a.SHIP_CUST as 出货客户, a.SHIP_TYPE as 货运方式, b.CREATE_BY as 创建人员,a.REMARK1 as 备注  " & " FROM erptemp .. SHIP_PLAN a  inner join erptemp .. SHIP_PLAN_DETAILED b on a.PLAN_ID = b.PLAN_ID and a.PLAN_ITEM = b.PLAN_ITEM  "
Else
    strSql = " SELECT distinct '' as 选择,a.PLAN_ID as 出货计划ID,a.APPROVER as 审核人员,  a.APPROV_TIME as 审核时间,case a.SHIP_FLAG when '0' then '未执行' when '1' then '已执行' else '未知状态' end AS 执行状态,  a.PLAN_DATE as 出货计划日期,a.REMARK2 as 出货单号, a.CUSTOMER AS 客户代码,  a.PLAN_ITEM as ITEM, a.CUST_PART as 客户机种, a.PRODUCT_NAME as 料号, a.PRODUCT_ID as 物料编号, " & " a.SHIP_AD as 出货地址, a.SHIP_CUST as 出货客户, a.SHIP_TYPE as 货运方式, b.LOT_ID,b.WAFER_ID ,b.TOTAL_DIE,b.GOOD_DIE, b.SHOP_ORDER as 工单号,b.CREATE_BY as 创建人员,a.REMARK1 as 备注 " & " FROM erptemp .. SHIP_PLAN a " & " inner join erptemp .. SHIP_PLAN_DETAILED b on a.PLAN_ID = b.PLAN_ID and a.PLAN_ITEM = b.PLAN_ITEM "

End If

If txtShipID.text <> "" Then
    strsql1 = " where a.PLAN_ID = '" & Trim(txtShipID.text) & "' "
Else
    strsql1 = ""

End If

If chkSchedule_COMPLETED.Value = 0 Then
    strSql2 = " and a.SHIP_FLAG = '0' "
Else
    strSql2 = ""

End If

strSql4 = " and CONVERT(varchar(100), a.PLAN_DATE, 23)  > CONVERT(varchar(100), GETDATE()-7, 23) "


If chkSchedule_OTHER.Value = 0 Then
    strSql3 = "group  by a.PLAN_ID, a.SHIP_FLAG,a.PLAN_DATE,a.CUSTOMER,a.PLAN_ITEM, a.CUST_PART,a.PRODUCT_NAME,a.PRODUCT_ID,a.SHIP_AD,a.SHIP_CUST ,a.SHIP_TYPE,b.CREATE_BY,a.APPROVER ,a.APPROV_TIME ,a.REMARK1,a.REMARK2 order by a.PLAN_ID,a.PLAN_ITEM"
Else
    strSql3 = " order by a.PLAN_ID,a.PLAN_ITEM "

End If

strSql = strSql & strsql1 & strSql2 & strSql4 & strSql3
Set rs = Get_SqlserveRs(strSql)

With fpS_ShipSchedule
    .MaxRows = 0
    If rs.RecordCount > 0 Then
        Set .DataSource = rs

    End If

End With

With fpS_ShipSchedule

    For i = 1 To .MaxRows
        .Row = i
        .Col = 3
        If .text <> "" Then
            .BackColor = vbGreen

        End If

    Next

End With

End Sub

Private Sub ClearData()
If txtCustLot.text <> "*" Then
    cbCustCode.text = ""
    txtCustLot.text = ""
Else
    txtCustLot.text = ""
    txtCustLot.text = "*"
End If
'cbCustCode.text = ""
cbHTPN.text = ""
cbCustPN.text = ""
cbProductNO.text = ""
'txtCustPO.text = ""
cbPO.text = ""
'txtCustLot.text = ""
Cbshipto.text = ""
cbShipAddr.text = ""
cbShipBy.text = ""
txtAdd.text = ""
txtShipDate.text = ""
fps_WO.MaxRows = 0
fpS_wafer.MaxRows = 0
lblQTY.Caption = "0"
lblQtyPecs.Caption = "0"

End Sub

Private Sub SaveData()
Dim tyData         As SHIP_PLAN
Dim i              As Integer
Dim strProductName As String
Dim strWOShipTo    As String

On Error GoTo hErr

If Cbshipto.text = "" Then
    MsgBox "请选择收货客户", vbInformation, "提示"
    Exit Sub

End If

If cbShipAddr.text = "" Then
    MsgBox "请选择发货地址", vbInformation, "提示"
    Exit Sub

End If

If cbShipBy.text = "" Then
    MsgBox "请选择货运方式", vbInformation, "提示"
    Exit Sub

End If

If lblQTY.Caption = "" Or lblQTY.Caption = "0" Then
    MsgBox "请勾选需要出货的WAFER明细", vbExclamation, "提示"
    Exit Sub

End If

If txtShipDate.text = "" Then
    MsgBox "请选择出货日期", vbInformation, "提示"
    Exit Sub

End If

If cbCXBJ.text = "" Then
    MsgBox "请选择产线标记", vbInformation, "提示"
    Exit Sub

End If

tyData.PLAN_ID = getPlanID
tyData.PLAN_DATE = txtShipDate.text
tyData.ALARM_TIME = txtShipDate.text
tyData.CUSTOMER = Trim$(cbCustCode.text)
If cbShipGoodOrNot.Value = 1 Then
    tyData.BAD_FLAG = "Y"
Else
    tyData.BAD_FLAG = "N"

End If

tyData.SHIP_AD = Trim(cbShipAddr.text)
tyData.SHIP_TYPE = Trim$(cbShipBy.text)
tyData.SHIP_CUST = Trim(Cbshipto.text)
tyData.REMARK1 = Trim(cbCXBJ.text)     '产线标记
tyData.REMARK4 = Trim(txtAdd.text)     '附加信息

'检查发货地址
With fpS_wafer

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_Wafer.E_CHOOSE
        If .Value = 1 Then
            .Col = E_Wafer.E_WAFERID
            strWOShipTo = Get_OracleStr(" select b.comp_code from mappingdatatest a " & " inner join customeroitbl_test b on a.filename = to_char(b.id)  " & " where a.substrateid = '" & Trim$(.text) & "' ")
            If strWOShipTo <> "" And strWOShipTo <> tyData.SHIP_AD Then
                MsgBox "您选择的出货地址: " & tyData.SHIP_AD & "和WO上的:" & strWOShipTo & "不一致", vbCritical, "错误"
                Exit Sub

            End If

        End If

    Next

End With

INIadoCon.BeginTrans

With fpS_wafer

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_Wafer.E_CHOOSE
        If .Value = 1 Then
            .Col = E_Wafer.E_WOID
            tyData.SHOP_ORDER = Trim$(.text)
            If Left$(tyData.SHOP_ORDER, 1) = "A" Then
                tyData.REMARK5 = "保税"
            Else
                tyData.REMARK5 = "非保税"

            End If

            .Col = E_Wafer.E_CARTON_NO
            If Trim(.text) <> "WIP" Then
                tyData.QBOXNO = Trim$(.text)
            Else
                tyData.QBOXNO = ""

            End If

            .Col = E_Wafer.E_CUSTPN
            tyData.CUST_PART = Trim(.text)
            .Col = E_Wafer.E_LOTID
            tyData.Lot_id = Trim$(.text)
            .Col = E_Wafer.E_WAFERID
            tyData.wafer_id = Trim$(.text)
            .Col = E_Wafer.E_WO_QTY
            tyData.GROSS_DIE = Trim$(.text)
            tyData.TOTAL_DIES = Trim$(.text)
            .Col = E_Wafer.E_SHIP
            If .text = "入满即发" Then
                .Col = E_Wafer.E_PLAN_QTY
            Else
                .Col = E_Wafer.E_STOCK_QTY

            End If

            tyData.GOOD_DIES = Trim$(.text)
            strProductName = Get_OracleStr("select distinct product from ib_wohistory where ordername = '" & tyData.SHOP_ORDER & "'")
            tyData.PRODUCT_ID = Get_SqlStr("select FNumber from AIS20141114094336..t_ICItem where F_101 = '" & strProductName & "'")
            tyData.PRODUCT_NAME = strProductName
            If Get_SqlserverCnt("select * from erptemp..SHIP_PLAN where PLAN_ID = '" & tyData.PLAN_ID & "' and PRODUCT_NAME = '" & strProductName & "'") = 0 Then
                tyData.PLAN_ITEM = tyData.PLAN_ITEM + 1
                If saveDataToDB_Header(tyData) = False Then
                    GoTo hErr

                End If

            End If

            'Detail: group by wafer
            tyData.PLAN_ITEM = Get_SqlStr("select distinct PLAN_ITEM from erptemp..SHIP_PLAN where PLAN_ID = '" & tyData.PLAN_ID & "' and PRODUCT_NAME = '" & strProductName & "'")
            If saveDataToDB_Datails(tyData) = False Then
                GoTo hErr

            End If

        End If

    Next

End With

INIadoCon.CommitTrans
lblQTY.Caption = "0"
MsgBox "出货计划已生成" & vbCrLf & tyData.PLAN_ID, vbInformation, "提示"
ClearData
Exit Sub
hErr:
MsgBox "产生错误: " & Err.DESCRIPTION, vbExclamation, "警告"
INIadoCon.RollbackTrans

End Sub

Private Function getPlanID() As String
getPlanID = "SP" & Right(Year(Now), 2) & Right$("00" & Month(Now), 2) & Right$("00" & Day(Now), 2) & Right("000" & Get_OracleStr("select SHIP_PLAN_ID_SEQ.Nextval from dual"), 3)

End Function

Private Function saveDataToDB_Header(tD As SHIP_PLAN) As Boolean
Dim strSql As String

saveDataToDB_Header = False
strSql = "insert into erptemp..SHIP_PLAN(PLAN_ID,CUSTOMER,CUST_PART,PRODUCT_NAME,PRODUCT_ID,GROSS_DIE,BAD_FLAG,SHIP_AD,SHIP_TYPE,SHIP_CUST,PLAN_DATE,ALARM_TIME,SHIP_FLAG,REMARK1,REMARK4,PLAN_ITEM) " & " values('" & tD.PLAN_ID & "','" & tD.CUSTOMER & "','" & tD.CUST_PART & "','" & tD.PRODUCT_NAME & "', '" & tD.PRODUCT_ID & "','" & tD.GROSS_DIE & "','" & tD.BAD_FLAG & "','" & tD.SHIP_AD & "','" & tD.SHIP_TYPE & "','" & tD.SHIP_CUST & "','" & tD.PLAN_DATE & "','" & tD.ALARM_TIME & "','0','" & tD.REMARK1 & "','" & tD.REMARK4 & "','" & tD.PLAN_ITEM & "')  "
If AddSql2(strSql) = 0 Then
    MsgBox "头表数据未插入", vbExclamation, "警告"
    Exit Function

End If

saveDataToDB_Header = True

End Function

Private Function updateDataToDB_Header(tD As SHIP_PLAN) As Boolean
Dim strSql As String

updateDataToDB_Header = False
strSql = "update erptemp..SHIP_PLAN  set GROSS_DIE = GROSS_DIE + " & tD.GROSS_DIE & " where PLAN_ID = '" & tD.PLAN_ID & "' and  PRODUCT_NAME = '" & tD.PRODUCT_NAME & "'        "
If AddSql2(strSql) = 0 Then
    MsgBox "头表数据未更新", vbExclamation, "警告"
    Exit Function

End If

updateDataToDB_Header = True

End Function

Private Function saveDataToDB_Datails(tD As SHIP_PLAN) As Boolean
Dim strSql As String

saveDataToDB_Datails = False
strSql = "insert into erptemp..SHIP_PLAN_DETAILED(PLAN_ID,WAFER_ID,LOT_ID,PO_NUM,SHOP_ORDER,CREATE_BY,FLAG,PLAN_ITEM,PRODUCT_NAME,TOTAL_DIE,GOOD_DIE,QBOX,REMARK5) " & " values('" & tD.PLAN_ID & "','" & tD.wafer_id & "','" & tD.Lot_id & "','" & tD.PO_NUM & "','" & tD.SHOP_ORDER & "','" & gUserName & "','0','" & tD.PLAN_ITEM & "','" & tD.PRODUCT_NAME & "','" & tD.TOTAL_DIES & "','" & tD.GOOD_DIES & "','" & tD.QBOXNO & "','" & tD.REMARK5 & "')  "
If AddSql2(strSql) = 0 Then
    MsgBox "子表数据未插入", vbExclamation, "警告"
    Exit Function

End If

saveDataToDB_Datails = True

End Function

Private Sub txtCustLot_Change()
If Trim(txtCustLot.text) = "*" Then
    Call updatepolist("")
End If
End Sub
Private Sub txtShipDate_DblClick()
Dialog_ShipPlan.Show 1

End Sub

Private Sub updatepolist(strLotID As String)

Dim i      As Integer
Dim strSql As String
Dim rs     As New ADODB.Recordset
If strLotID = "" Then
    strSql = " select distinct  rtrim(ff.PO_NUM) + ' ' + rtrim(ff.comp_code) AS PO出货地址 from  erpbase..tblCustomerOI ff inner join erpbase..tblmappingData gg  ON ff.SOURCE_BATCH_ID=gg.LOTID AND convert(VARCHAR(20),ff.ID)=gg.FILENAME  where  gg.SUBSTRATEID in ( SELECT distinct b.流程卡编号 FROM erpdata..tblstocknum a , erpdata..tblstocknumsub b WHERE a.客户代码='" & Trim(cbCustCode.text) & "' AND a.id=b.ID and not exists (select 1 from erptemp..SHIP_PLAN_DETAILED ee where ee.WAFER_ID =  b.流程卡编号 ) )"
Else
    strSql = " select  distinct rtrim(ff.PO_NUM) + ' ' + rtrim(ff.comp_code) AS PO出货地址 from  erpbase..tblCustomerOI ff inner join erpbase..tblmappingData egg  ON ff.SOURCE_BATCH_ID=gg.LOTID AND convert(VARCHAR(20),ff.ID)=gg.FILENAME  where  gg.SUBSTRATEID in ( SELECT distinct b.流程卡编号 FROM erpdata..tblstocknum a , erpdata..tblstocknumsub b WHERE a.客户代码='" & Trim(cbCustCode.text) & "' AND a.id=b.ID and b.工单号='" & strLotID & "' and not exists (select 1 from erptemp..SHIP_PLAN_DETAILED ee where ee.WAFER_ID =  b.流程卡编号 )) "
End If
If rs.State = adStateOpen Then rs.Close
rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
cbPO.Clear
cbPO.AddItem "所有"
If Not rs.EOF Then
    For i = 1 To rs.RecordCount
        cbPO.AddItem Trim$("" & rs!PO出货地址)
        rs.MoveNext
    Next
End If
rs.Close
End Sub


