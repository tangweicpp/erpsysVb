VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_ProductionPlanDetailNew 
   Caption         =   "工单明细"
   ClientHeight    =   9495
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
   LockControls    =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin FPSpreadADO.fpSpread fpS1 
      Height          =   4695
      Left            =   120
      TabIndex        =   69
      Top             =   3960
      Width           =   11895
      _Version        =   524288
      _ExtentX        =   20981
      _ExtentY        =   8281
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
      SpreadDesigner  =   "Frm_ProductionPlanDetailNew.frx":0000
      Appearance      =   1
      AppearanceStyle =   0
   End
   Begin VB.Frame Frame2 
      Height          =   8535
      Left            =   12120
      TabIndex        =   67
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtLog 
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
         Height          =   8415
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   68
         Text            =   "Frm_ProductionPlanDetailNew.frx":0414
         Top             =   120
         Width           =   6735
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "全选/反选"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10920
      TabIndex        =   62
      Top             =   3720
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   600
      TabIndex        =   60
      Top             =   3240
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "工单信息"
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   11895
      Begin VB.TextBox txtMapping 
         Height          =   285
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox txtCusDieQty 
         Height          =   285
         Left            =   5520
         TabIndex        =   64
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtWOType 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   203
         Width           =   2175
      End
      Begin VB.TextBox txtPO 
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtCloseDate 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtOpenDate 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtLotType 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   1883
         Width           =   2175
      End
      Begin VB.TextBox txt37Pri 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   1643
         Width           =   2175
      End
      Begin VB.TextBox txtWODept 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1403
         Width           =   2175
      End
      Begin VB.TextBox txtPN 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1163
         Width           =   2175
      End
      Begin VB.TextBox txtCusPN 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   683
         Width           =   2175
      End
      Begin VB.TextBox txtHTPN 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   923
         Width           =   2175
      End
      Begin VB.TextBox txtCusCode 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   443
         Width           =   2175
      End
      Begin VB.TextBox txtMPN 
         Height          =   285
         Left            =   9240
         TabIndex        =   46
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox txtType 
         Height          =   285
         Left            =   9240
         TabIndex        =   44
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtDateCode 
         Height          =   285
         Left            =   5520
         TabIndex        =   42
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtShipSite 
         Height          =   285
         Left            =   9240
         TabIndex        =   40
         Top             =   1883
         Width           =   2535
      End
      Begin VB.TextBox txtCusRequest 
         Height          =   285
         Left            =   9240
         TabIndex        =   38
         Top             =   1643
         Width           =   2535
      End
      Begin VB.TextBox txtPFA 
         Height          =   285
         Left            =   9240
         TabIndex        =   36
         Top             =   1403
         Width           =   2535
      End
      Begin VB.TextBox txtLotStatus 
         Height          =   285
         Left            =   9240
         TabIndex        =   33
         Top             =   1163
         Width           =   2535
      End
      Begin VB.TextBox txtPOItem 
         Height          =   285
         Left            =   9240
         TabIndex        =   31
         Top             =   923
         Width           =   2535
      End
      Begin VB.TextBox txtMM 
         Height          =   285
         Left            =   9240
         TabIndex        =   29
         Top             =   683
         Width           =   2535
      End
      Begin VB.TextBox txtCtyFab 
         Height          =   285
         Left            =   9240
         TabIndex        =   27
         Top             =   443
         Width           =   2535
      End
      Begin VB.TextBox txtPercent 
         Height          =   285
         Left            =   9240
         TabIndex        =   25
         Text            =   "25"
         Top             =   203
         Width           =   2535
      End
      Begin VB.TextBox txtMarkingCode 
         Height          =   285
         Left            =   5520
         TabIndex        =   23
         Top             =   1650
         Width           =   1815
      End
      Begin VB.TextBox txtNGFlag 
         Height          =   285
         Left            =   5520
         TabIndex        =   21
         Text            =   "Y"
         Top             =   1410
         Width           =   1815
      End
      Begin VB.TextBox txt260 
         Height          =   285
         Left            =   5520
         TabIndex        =   19
         Top             =   1170
         Width           =   1815
      End
      Begin VB.TextBox txt235 
         Height          =   285
         Left            =   5520
         TabIndex        =   17
         Top             =   930
         Width           =   1815
      End
      Begin VB.TextBox txtDesignedID 
         Height          =   285
         Left            =   5520
         TabIndex        =   15
         Top             =   690
         Width           =   1815
      End
      Begin VB.TextBox txtICR 
         Height          =   285
         Left            =   5520
         TabIndex        =   13
         Top             =   450
         Width           =   1815
      End
      Begin VB.TextBox txtFABDevice 
         Height          =   285
         Left            =   5520
         TabIndex        =   11
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "是否有MAPPING"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   31
         Left            =   7800
         TabIndex        =   65
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户设计DIE数量"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   30
         Left            =   3960
         TabIndex        =   63
         Top             =   2400
         Width           =   1365
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单类型: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   29
         Left            =   240
         TabIndex        =   58
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   28
         Left            =   240
         TabIndex        =   49
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MPN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   8760
         TabIndex        =   45
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "贸易类型"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   27
         Left            =   8400
         TabIndex        =   43
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DateCode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   26
         Left            =   4560
         TabIndex        =   41
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ShipSite"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   24
         Left            =   8520
         TabIndex        =   39
         Top             =   1920
         Width           =   660
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户需求日"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   23
         Left            =   8280
         TabIndex        =   37
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "订单单号"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   22
         Left            =   4680
         TabIndex        =   35
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ProtectiveFilmApld"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   21
         Left            =   7680
         TabIndex        =   34
         Top             =   1440
         Width           =   1500
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LotStatus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   20
         Left            =   8400
         TabIndex        =   32
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POItem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   19
         Left            =   8520
         TabIndex        =   30
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MicronMaterial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   18
         Left            =   8040
         TabIndex        =   28
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CountryFab"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   17
         Left            =   8250
         TabIndex        =   26
         Top             =   480
         Width           =   930
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "比率(*)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   16
         Left            =   8565
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MarkingCode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   15
         Left            =   4320
         TabIndex        =   22
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NG标志"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   4800
         TabIndex        =   20
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level260加工地"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   13
         Left            =   4125
         TabIndex        =   18
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level235(CS)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   4365
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DesignedID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   4455
         TabIndex        =   14
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ImagerCustomerRev"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   3810
         TabIndex        =   12
         Top             =   480
         Width           =   1650
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FAB设备"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   4740
         TabIndex        =   10
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户机种: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "厂内机种: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   195
         TabIndex        =   8
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "成品料号: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单部门: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   195
         TabIndex        =   6
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "37PRI: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   360
         TabIndex        =   5
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOTTYPE: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   915
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开工日期: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   195
         TabIndex        =   3
         Top             =   2131
         Width           =   840
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "完工日期: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   195
         TabIndex        =   2
         Top             =   2384
         Width           =   840
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   8625
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   1535
      ButtonWidth     =   1984
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "     "
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  保存工单   "
            Key             =   "SAVE"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "EXIT"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   240
         Top             =   5160
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_ProductionPlanDetailNew.frx":0429
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_ProductionPlanDetailNew.frx":0D03
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "进度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   61
      Top             =   3315
      Width           =   420
   End
End
Attribute VB_Name = "Frm_ProductionPlanDetailNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim lPart As Integer

Private Sub Check1_Click()

   Dim i As Integer

    If Check1.Value = 1 Then

        For i = 1 To fpS1.MaxRows

            With fpS1
                .Row = i
                .Col = 7
                .Text = 1

            End With

        Next i
        
    ElseIf Check1.Value = 0 Then

        For i = 1 To fpS1.MaxRows

            With fpS1
                .Row = i
                .Col = 7
                .Text = 0

            End With

        Next i
        
    End If

End Sub

Private Sub Form_Load()
    InitFps
    ListWOHead
    ListWODetail

End Sub

Private Sub InitFps()

    With fpS1
        .Col = -1
        .Row = -1
        
        .Lock = True

        .MaxCols = 7
        .MaxRows = 0
        .FontBold = True
    
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .SetText 1, 0, "工单_ID"
        .SetText 2, 0, "LOT_ID"
        .SetText 3, 0, "WAFER_ID"
        .SetText 4, 0, "GROSS_DIES"
        .SetText 5, 0, "GOOD_DIES"
        .SetText 6, 0, "打标码"
        .SetText 7, 0, "选择"
          
        .ColWidth(1) = 14
        .ColWidth(2) = 15
        .ColWidth(3) = 20
        .ColWidth(4) = 12
        .ColWidth(5) = 12
        .ColWidth(6) = 12
        .ColWidth(7) = 4
        
        If Frm_ProductionPlanNew.cbWOType = "FO_CSP工单" Then
            .Col = 4
            .Lock = False
        
            .Col = 5
            .Lock = False
        
        End If
       
        .Col = 7
        .CellType = CellTypeCheckBox
        .Lock = False
        
        .MaxRows = 0
    End With

End Sub

Private Sub ListWOHead()

    Dim Rs1        As New ADODB.Recordset

    Dim strCusCode As String

    Dim strLotID   As String

    Dim oiRS       As New ADODB.Recordset
    
    Dim rs         As New ADODB.Recordset
    
    With Frm_ProductionPlanNew.List1

        For i = 0 To .ListCount - 1

            If .Selected(i) = True Then
                strLotID = Trim$(.List(i))
                
            End If

        Next

    End With
    
    strCusCode = UCase(Trim(Frm_ProductionPlanNew.cbCusCode.Text))
    
    Set oiRS = GetOIData(strCusCode, strLotID)
    
    If (oiRS.RecordCount > 0) Then
        If JudgeCustomerPTNum(strLotID) Then
            MsgBox "此LotID：" & strLotID & " 客户WO上有多笔客户机种号，请注意确认目前带出信息是否正确！", vbInformation, "提示"

        End If
        
        txtFABDevice.Text = "" & Trim(oiRS("fabrication_facility"))
        txtICR.Text = "" & Trim(oiRS("imager_customer_rev"))
        txtDesignedID.Text = "" & Trim(oiRS("design_id"))
        txt260.Text = "" & Trim(oiRS("shipping_mst_260"))
        txt235.Text = "" & Trim(oiRS("shipping_mst_level"))
        txtMarkingCode.Text = "" & Trim(oiRS("encoded_mark_id"))
        txtCtyFab.Text = "" & Trim(oiRS("country_of_fab"))
        txtMM.Text = "" & Trim(oiRS("micron_material"))
        txtPOItem.Text = "" & Trim(oiRS("po_item"))
        txtLotStatus.Text = "" & Trim(oiRS("lot_status"))
        txtType.Text = "" & Trim(oiRS("PROBE_SHIP_PART_TYPE"))
        txtPFA.Text = IIf("" & Trim(oiRS("protective_film_apld")) = "YES", "PF", "" & Trim(oiRS("protective_film_apld")))
        txtCusRequest.Text = "" & Trim(oiRS("lot_priority"))
        txtShipSite.Text = "" & Trim(oiRS("ship_site"))
        txtDateCode.Text = "" & GetONDateCode()

    End If
    
    txtWOType.Text = Trim(Frm_ProductionPlanNew.cbWOType.Text)
    txtCusCode.Text = Trim(Frm_ProductionPlanNew.cbCusCode.Text)
    txtCusPN.Text = Trim(Frm_ProductionPlanNew.txtCusPN.Text)
    txtPO.Text = Trim$(Frm_ProductionPlanNew.cbHTPN.Text)

    If (txtCusCode.Text = "AA" Or txtCusCode.Text = "AA(ON)") And (txtWOType.Text <> "玻璃工单" And txtWOType.Text <> "DUMMY工单") Then
        txtMPN.Text = txtCusPN.Text
        txtDateCode.Text = GetONDateCode()
        txtMarkingCode.Text = GetONWoMarkingCode(Trim(txtMPN.Text))
        txtCusPN.Text = GetONOPN_WSG(Trim(txtCusPN.Text))
    Else
        txtMPN.Text = Trim(Frm_ProductionPlanNew.txtCusPN.Text)
        txtDateCode.Text = GetONDateCode()
        txtMarkingCode.Text = GetONWoMarkingCode(Trim(txtMPN.Text))

    End If
    
    txtPN.Text = Trim(Frm_ProductionPlanNew.cbPN.Text)
    txtHTPN.Text = Trim(Frm_ProductionPlanNew.cbHTPN.Text)
    txtWODept.Text = Trim(Frm_ProductionPlanNew.txtWODept.Text)
    txt37Pri.Text = Trim(Frm_ProductionPlanNew.cb37Pri.Text)
    
    Select Case Trim(Frm_ProductionPlanNew.cbLotType.Text)
    
        Case "量产批(M)"
            txtLotType.Text = "M"

        Case "工程批(E)"
            txtLotType.Text = "E"

        Case "客户实验(Q)"
            txtLotType.Text = "Q"

        Case "DC片(D)"
            txtLotType.Text = "D"

    End Select
    
    ' 查询CUSDIE
    Set Rs1.ActiveConnection = OraConnect
    
    Rs1.Source = "select customerdieqty,mapping from tbltsvnpiproduct where qtechptno2 = '" & Trim(txtPN.Text) & "' and customershortname = '" & txtCusCode.Text & "' and customerptno1 = '" & Trim(Frm_ProductionPlanNew.txtCusPN.Text) & "' and qtechptno = '" & txtHTPN.Text & "' "
 
    Rs1.Open , , adOpenStatic, adLockReadOnly, adCmdText
    
    txtCusDieQty.Text = Trim(Rs1("customerdieqty") & "")
    txtMapping.Text = Trim(Rs1("MAPPING") & "")
    
    txtOpenDate.Text = Trim(Frm_ProductionPlanNew.DTPicker1(0).Value)
    txtCloseDate.Text = Trim(Frm_ProductionPlanNew.DTPicker1(1).Value)
    
End Sub

Private Sub ListWODetail()
    
    Dim strLotID As String
    
    Dim strWOID  As String

    With Frm_ProductionPlanNew.List1

        If Frm_ProductionPlanNew.Check2.Value = 1 Then

            ' 批量工单
            For i = 0 To .ListCount - 1

                If .Selected(i) = True Then
                
                    strLotID = Trim$(.List(i))
               
                    strWOID = Frm_ProductionPlanNew.GetWOID()
 
                    Call ListData(strLotID, strWOID)

                End If

            Next
            
        Else
            ' 单工单
            strWOID = Frm_ProductionPlanNew.GetWOID()
            
            For i = 0 To .ListCount - 1

                If .Selected(i) = True Then
                
                    strLotID = Trim$(.List(i))
                    
                    Call ListData(strLotID, strWOID)

                End If

            Next

        End If
        
    End With
    
End Sub
   
Private Sub ListData(strLotID As String, strWOID As String)

    Dim adorst1  As New ADODB.Recordset

    Dim listRS2  As New ADODB.Recordset
    
    Dim listRS3  As New ADODB.Recordset
    
    Dim strCusPN As String
    
    Dim i        As Long
    
    strCusPN = Trim$(txtCusPN.Text)

    On Error GoTo there
    
    Set adorst1.ActiveConnection = OraConnect

    If Trim(txtCusCode.Text) = "AA" Or Trim(txtCusCode.Text) = "AA(ON)" Then
        If Right(Trim(UCase(Frm_ProductionPlanNew.cbWO)), 2) = "ST" Or Right(Trim(UCase(Frm_ProductionPlanNew.cbWO)), 2) = "ET" Then
        
            adorst1.Source = "select distinct '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, ONSTMarkingCodeSeq.QTSeq(b.substrateid,b.lotid) as productid from mappingdatatest b where b.lotid = '" & strLotID & "' and  not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"
        Else
            adorst1.Source = "select distinct  '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, ONMarkingCodeSeq.QTSeq(b.substrateid) as productid from WAFERDETAILTMP a, mappingdatatest b where  b.lotid = '" & strLotID & "' and  not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"

        End If
        
    Else

        If Frm_ProductionPlanNew.cbWOType = "重工工单" Or Frm_ProductionPlanNew.cbWOType = "FO_CSP工单" Then
            adorst1.Source = "select distinct  '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, replace(b.productid,'_','') as productid from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and b.filename = to_char(a.id) and a.mpn_desc = '" & strCusPN & "' and b.substrateid like '%+%' and  not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"
        
        ElseIf Frm_ProductionPlanNew.cbWOType = "硅基工单" Then
            adorst1.Source = "select distinct '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, replace(b.productid,'_','') as productid from customeroitbl_test a,mappingdatatest b where b.lotid = '" & strLotID & "' and b.filename = to_char(a.id) and a.mpn_desc = '" & strCusPN & "' and instr(b.substrateid, '+') = 0 and not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"
        
        Else
            adorst1.Source = "select distinct '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, replace(b.productid,'_','') as productid from customeroitbl_test a,mappingdatatest b where b.lotid = '" & strLotID & "' and b.filename = to_char(a.id) and a.mpn_desc = '" & strCusPN & "' and not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"

        End If
        
    End If
    
    adorst1.Open , , adOpenStatic, adLockReadOnly, adCmdText
    
    Dim lQty As Long
    
    If Frm_ProductionPlanNew.cbWOType = "FO_CSP工单" Then
        lQty = Get_OracleStr("select distinct customerdieqty from tbltsvnpiproduct where qtechptno2 = '" & txtPN.Text & "'")
    End If
    
    With fpS1

        If adorst1.RecordCount > 0 Then

            adorst1.MoveFirst

            For i = 1 To adorst1.RecordCount
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                .Col = 1
                .Text = Trim("" & adorst1(0))
                    
                .Col = 2
                .Text = Trim$("" & adorst1(1))
            
                .Col = 3
                .Text = Trim$("" & adorst1(2))
             
                .Col = 4
                
                If Frm_ProductionPlanNew.cbWOType = "FO_CSP工单" Then
                    .Text = lQty
                Else
                    .Text = Trim$("" & adorst1(3))
                End If
    
                .Col = 5
                If Frm_ProductionPlanNew.cbWOType = "FO_CSP工单" Then
                    .Text = lQty
                Else
                    .Text = Trim$("" & adorst1(4))
                End If
                
                .Col = 6
                .Text = Trim$("" & adorst1(5))
                
                .Col = 7
                .Text = "1"
                
                adorst1.MoveNext
            Next
        Else
            MsgBox "没有查询到明细", vbInformation, "提示"
        
        End If

    End With
    
    Screen.MousePointer = 0

    adorst1.Close
    Set adorst1 = Nothing

    Exit Sub
    
there:
    Screen.MousePointer = 0
    MsgBox "查询失败！" + Chr(13) + "原因:" + Err.DESCRIPTION, vbInformation, Me.Caption
  
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key

        Case "SAVE"
            ForSave
        
        Case "EXIT"
            Unload Me

    End Select

End Sub

Private Sub ForSave()

    Toolbar1.Buttons("SAVE").Enabled = False
    Frm_ProductionPlanNew.tb1.Buttons("SEARCH").Enabled = True
    Frm_ProductionPlanNew.tb1.Buttons("PREVIEW").Enabled = False
    
    Dim strWOID As String

    Dim bRtn    As Boolean
    
    Dim i       As Integer
    
    Dim iPart   As Integer
    
    i = 0
    iPart = 0
    bRtn = False
    ProgressBar1.Value = 1
    strWOID = ""
    
    With fpS1

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If
    
        For i = 1 To .MaxRows
            .Row = i
            .Col = 7
        
            If .Text = "1" Then
                bRtn = True

            End If
            
        Next

    End With
    
    If bRtn = False Then
        MsgBox "请勾选要开立的WAFER ID", vbCritical, "警告"
        Exit Sub

    End If
    
    If CheckData = False Then
        Exit Sub

    End If
    
    With fpS1

        For i = 1 To .MaxRows
            .Row = 1
            .Col = 7
            
            If .Text = "1" Then
                .Col = 1

                If .Text <> strWOID Then
                    strWOID = .Text
                    
                    iPart = iPart + 1

                End If
          
            End If
           
        Next
    
    End With

    lPart = 20 * (1 / iPart)
    strWOID = ""
    
    With fpS1

        For i = 1 To .MaxRows
            .Row = i
            .Col = 7
            
            If .Text = "1" Then
                .Col = 1

                If .Text <> strWOID Then
                    strWOID = .Text
                    Call SaveWOID(strWOID)
                End If
                
                
    
            End If
           
        Next

    End With

    ProgressBar1.Value = 100
 
    SaveAsExcel
    
End Sub
    
Private Function CheckData() As Boolean

    Dim i              As Integer

    Dim strWaferID     As String

    Dim strLotID       As String
    
    Dim strGrossDies   As String
    
    Dim strMarkingCode As String
    
    Dim listRS2        As New ADODB.Recordset

    Dim listRS3        As New ADODB.Recordset

    CheckData = False
        
    With fpS1

        '37PO, JOBID
        If txtCusCode.Text = "37" And cbWOType <> "Dummy工单" And Frm_ProductionPlanNew.cbWOType <> "FO_CSP工单" And Frm_ProductionPlanNew.cbWOType <> "硅基工单" Then

            For i = 1 To fpS1.MaxRows
                .Row = i
                .Col = 7

                If .Text = "1" Then
                    .Col = 3
                    strWaferID = UCase(Trim$(.Text))
                    
                    If Check37PO(strWaferID) = False Then
                        MsgBox "37客户, 市场部未维护订单PO信息, 禁止开立工单, 请联系市场部确认", vbCritical, "警告"

                        Exit Function

                    End If

                End If

            Next i

        End If
    
        '37DateCode确认
        If txtCusCode.Text = "37" Then
            If Get_OracleCnt("select * from tbltsvnpiproduct a where instr(a.struckstr1,'TR') >0  and a.customerptno1 = '" & txtCusPN.Text & "' ") > 0 Then

                For i = 1 To fpS1.MaxRows
                    .Row = i
                    .Col = 7

                    If .Text = "1" Then
                        .Col = 3
                        strWaferID = UCase(Trim$(.Text))

                        If Check37DATECODE(Replace$(strWaferID, "+", "")) = False Then
                            MsgBox "37客户, 市场部未维护DATECODE信息, 禁止开立工单, 请联系市场部确认", vbCritical, "警告"

                            Exit Function

                        End If

                    End If

                Next i
            
            End If
        
        End If

        'Mapping确认
        If txtMapping.Text = "Y" And (Frm_ProductionPlanNew.cbWOType <> "重工工单" And Frm_ProductionPlanNew.cbWOType <> "玻璃工单" And Frm_ProductionPlanNew.cbWOType <> "Dummy工单" And Frm_ProductionPlanNew.cbWOType <> "FO_CSP工单" And Frm_ProductionPlanNew.cbWOType <> "硅基工单") Then

            Dim checkLot As String
            
            checkLot = ""
            
            For i = 1 To fpS1.MaxRows
                .Row = i
                .Col = 7
            
                If .Text = "1" Then
                    .Col = 2
                    strLotID = UCase(Trim$(.Text))
                    .Col = 3
                    strWaferID = UCase(Trim$(.Text))
                
                    If strLotID <> checkLot Then
                        If InStr(strWaferID, "+") = 0 Then
                            If CheckMapping(strLotID) = True Then
                                MsgBox "市场部未维护订单MAPPING信息, 禁止开立工单, 请联系市场部确认", vbCritical, "警告"

                                Exit Function

                            End If
                    
                            checkLot = strLotID

                        End If

                    End If
                   
                End If

            Next i
        
        End If
    
        'GrossDies确认
        If (Frm_ProductionPlanNew.cbWOType <> "重工工单" And Frm_ProductionPlanNew.cbWOType <> "FO_CSP工单" And Frm_ProductionPlanNew.cbWOType <> "玻璃工单" And Frm_ProductionPlanNew.cbCusCode.Text <> "37" And Frm_ProductionPlanNew.cbWOType <> "Dummy工单" And Frm_ProductionPlanNew.cbWOType <> "硅基工单") Then

            For i = 1 To fpS1.MaxRows
            
                .Row = i
                .Col = 7

                If .Text = "1" Then

                    .Col = 3
                    strWaferID = UCase(Trim$(.Text))
                    .Col = 4
                    strGrossDies = UCase(Trim$(.Text))
                
                    If Frm_ProductionPlanNew.cbCusCode.Text = "KR001" And InStr(strWaferID, "+") > 0 Then
                        Exit For

                    End If

                    If strGrossDies <> txtCusDieQty.Text Then
                        MsgBox "GROSSDIE数量错误", vbCritical, "警告"

                        Exit Function

                    End If

                End If

            Next i

        End If
    
        ' 打标码校验
        If Frm_ProductionPlanNew.cbWOType <> "Dummy工单" And Frm_ProductionPlanNew.cbWOType <> "FO_CSP工单" And Frm_ProductionPlanNew.cbWOType <> "硅基工单" And Frm_ProductionPlanNew.cbWOType <> "玻璃工单" Then

            Select Case txtCusCode.Text
    
                Case "SX", "HJ", "TJ003", "JS140", "BJ153", "GC"
                
                    For i = 1 To fpS1.MaxRows
            
                        .Row = i
                        .Col = 7

                        If .Text = "1" Then

                            .Col = 3
                            strWaferID = UCase(Trim$(.Text))
                            .Col = 6
                            strMarkingCode = UCase(Trim$(.Text))
                                
                            Set listRS2 = GetMARK2(strWaferID, txtPN.Text)
    
                            If listRS2.RecordCount > 0 Then
                                MsgBox "打标码长度异常"
    
                                Exit Function
    
                            End If
    
                            If Get_OracleCnt("select * from tbltsvnpiproduct where customerptno1 = '" & txtCusPN.Text & "' and qtechptno2 = '" & txtPN.Text & "'  and  marking_code > 0 ") > 0 Then
                                Set listRS3 = GetMARK3(strMarkingCode)
    
                                If listRS3.RecordCount > 0 Then
                                    MsgBox "存在重复打标码"
    
                                    Exit Function
    
                                End If
    
                            End If
                    
                        End If

                    Next i
        
            End Select
    
        End If
    
    End With
   
    CheckData = True

End Function


Private Sub SaveWOID(strWOID As String)
    AddLog strWOID & ": 工单正在开立...."

    If ProgressBar1.Value < 100 Then
    
        If ProgressBar1.Value + lPart > 100 Then
            ProgressBar1.Value = 100
        Else
            ProgressBar1.Value = ProgressBar1.Value + lPart
        End If
    End If
    
    If SaveOldData(strWOID) = False Then
      Exit Sub
    End If
            
    If ProgressBar1.Value < 100 Then
    
        If ProgressBar1.Value + lPart > 100 Then
            ProgressBar1.Value = 100
        Else
            ProgressBar1.Value = ProgressBar1.Value + lPart
        End If
    End If
            
    If SaveNewData(strWOID) = False Then
       Exit Sub
    End If
            
    If ProgressBar1.Value < 100 Then
    
        If ProgressBar1.Value + lPart > 100 Then
            ProgressBar1.Value = 100
        Else
            ProgressBar1.Value = ProgressBar1.Value + lPart
        End If
    End If
            
    Sleep (500)
            
    If ProgressBar1.Value < 100 Then
    
        If ProgressBar1.Value + lPart > 100 Then
            ProgressBar1.Value = 100
        Else
            ProgressBar1.Value = ProgressBar1.Value + lPart
        End If
    End If
    
    AddLog strWOID & ": 工单开立成功!!!"
    
End Sub

Private Function SaveOldData(strWOID As String) As Boolean

    SaveOldData = False
    
    Dim lWaferQty As Long

    Dim lDieQty   As Long

    On Error GoTo DealError
 
    Cnn.BeginTrans
    INIadoCon.BeginTrans
    
    Call SaveWOPri(strWOID)
    
    Call SaveWODetail(strWOID, lWaferQty, lDieQty)
    
    Call SaveWOHead(strWOID, lDieQty)
    
    Call SaveWOBom(strWOID, lWaferQty, lDieQty)
  
    Cnn.CommitTrans
    INIadoCon.CommitTrans
    
    SaveOldData = True
    
    Exit Function

DealError:
    Cnn.RollbackTrans
    INIadoCon.RollbackTrans
    
    MsgBox "保存失败！" + Chr(13) + "原因:" + Err.DESCRIPTION, vbInformation, Me.Caption
    AddLog strWOID & ": 旧工单接口插入失败"

End Function

Private Function SaveNewData(strWOID As String) As Boolean

    SaveNewData = False
    
    Dim strOra As String

    Dim strSql As String
    
    Dim lRtn   As Long
    
    Dim i      As Integer
    
    On Error GoTo DealError
 
    Cnn.BeginTrans
    INIadoCon.BeginTrans
 
    ' 1. 表shop_order_detail
    strOra = "insert into shop_order_detail(SHOP_ORDER,CUST_LOT_ID,WAFER_ID,GROSS_DIE_QTY,GOOD_DIE_QTY,MARK_CODE)" & "select a.ordername as SHOP_ORDER,a.waferlot as CUST_LOT_ID,a.waferid as WAFER_ID,a.dieqty      as GROSS_DIE_QTY,a.fgdieqty    as GOOD_DIE_QTY, " & "decode(b.customer,'AA',d.markingcodefirst || a.markingcode,a.markingcode) as MARK_CODE from ib_waferlist a inner join ib_workorder b on b.ordername = a.ordername " & "left join  CUSTOMERMPNAttributes d on d.part = b.mpn where  b.ordername  in ('" & strWOID & "')"
    
    If AddSql(strOra) = 0 Then
        AddLog strWOID & ": shop_order_detail插入失败"
        GoTo DealError
    End If
        
    ' 2. 表SHOP_ORDER_PROPERTY
    strOra = "select SHOP_ORDER_PROPERTY_PKG.SHOP_ORDER_PROPERTY('" & strWOID & "')  from dual"
    
    If AddSql(strOra) = 0 Then
        AddLog strWOID & ": SHOP_ORDER_PROPERTY插入失败"
        GoTo DealError
    End If

    ' 3. 表shop_order
    strOra = "insert into shop_order(SHOP_ORDER,PRD_ID,PRD_VER,ERP_ROUTING,ORDER_QTY,CUST_LOT_QTY,PLAN_STAR_DATE,PLAN_END_DATE,MANF_DEPT,MANF_DEPT_DESC,LOT_TYPE,PRIORITY,PKG,CUST_ID,ERP_CREATE_DATE,CREATOR,flag,ht_device,RELEASE_TYPE) " & _
       " select trim(a.ordername) as SHOP_ORDER, trim(b.product) as PRD_ID, 'A' as PRD_VER, '' as ERP_ROUTING,trim(COUNT(distinct A.WAFERID)) as ORDER_QTY, trim(COUNT(distinct A.WAFERLOT)) as CUST_LOT_QTY, trim(B.PLANSTARTDATE) AS PLAN_STAR_DATE, " & _
       " trim(B.PLANENDDATE) AS PLAN_END_DATE,  trim(B.PARA8) AS MANF_DEPT, trim(g.manf_dept_desc) AS MANF_DEPT_DESC,trim(e.lot_type) as LOT_TYPE,trim(decode(e.pri, 'Hot Lot', 1, 'Super Hot Lot', 1, 4)) as PRIORITY,trim(f.pkg_type) as PKG, " & _
       " trim(shop_order_property_pkg.SHOP_ORDER_CUSTOMER(b.customer)) AS CUST_ID, trim(b.erpcreatedate) as ERP_CREATE_DATE,'" & gUserName & "' as CREATOR, '0' as flag, trim(f.qtechptno) as ht_device, '1' as RELEASE_TYPE " & _
       " from ib_waferlist a, ib_workorder b, MAPPINGDATATEST  C, PJ_WO_PRI e,tbltsvnpiproduct f, MES_DEPT g " & _
       " where b.ordername = a.ordername and f.customerptno1 = b.mpn AND B.CUSTOMER = F.CUSTOMERSHORTNAME and b.ordername = '" & strWOID & "'  AND C.SUBSTRATEID = a.waferid  and e.wo = b.ordername and f.qtechptno2 = b.product and  rownum = 1  and g.manf_dept = substr(b.para8, 1, instr(b.para8, '_') - 1) " & _
       " group by a.ordername, b.product, B.PLANSTARTDATE, B.PLANENDDATE, B.PARA8,e.lot_type, e.pri,f.pkg_type,shop_order_property_pkg.SHOP_ORDER_CUSTOMER(b.customer), b.erpcreatedate, e.creat_by, f.qtechptno, g.manf_dept_desc"
    
    If AddSql(strOra) = 0 Then
        AddLog strWOID & ": shop_order插入失败"
        GoTo DealError
    End If

    ' 4. 表ERPBASE..TblERPFLToME
    strSql = "insert into ERPBASE..TblERPFLToME (STOCK_TYPE,STOCK_ID,PRD_ID,PRD_VER,QTY,PRD_DATE,EFF_DATE,SHOP_ORDER,SupSN,Flag,Memo,CreateDate,FStauts,HeaderID) " & "select 'W',b.ORDERNAME + c.WAFERLOT, b.PRODUCT,'A',COUNT(*),GETDATE() - 1,GETDATE() + 300,b.ORDERNAME,c.WAFERLOT,0,'',GETDATE(),'','' from erpdata .. tblTSVworkorder b,erpdata .. tblTSVwaferlist c where c.ORDERNAME = b.ORDERNAME and b.ORDERNAME in ( '" & strWOID & "' ) group by b.PRODUCT, b.ORDERNAME, c.WAFERLOT"

    If AddSql2(strSql) = 0 Then
        AddLog strWOID & ": ERPBASE..TblERPFLToME插入失败"
        GoTo DealError
    End If
   
    Cnn.CommitTrans
    INIadoCon.CommitTrans
    
    SaveNewData = True
    
    Exit Function

DealError:
    Cnn.RollbackTrans
    INIadoCon.RollbackTrans
    
    MsgBox "新工单接口保存失败！" + Chr(13) + "原因:" + Err.DESCRIPTION, vbInformation, Me.Caption
    AddLog strWOID & ": 新工单接口保存失败插入失败"
    
End Function

Private Sub SaveWOPri(ByVal strWOID As String)

    Dim strOra     As String

    Dim strPri     As String
    
    Dim strPri2    As String
    
    Dim strLotType As String

    Dim strUser    As String

    strPri = Trim(txt37Pri.Text)
    
    strLotType = Trim$(txtLotType.Text)
    
    strPri2 = Trim$(Frm_ProductionPlanNew.cb37Pri2.Text)
    
    strUser = gUserName
    
    strOra = "insert into PJ_WO_PRI(wo,pri,great_date,lot_type,Creat_by,return) values('" & strWOID & "','" & strPri & "',to_char(sysdate,'YYYY-MM-DD'),'" & strLotType & "', '" & strUser & "','" & strPri2 & "')"
    
    If AddSql(strOra) = 0 Then
        AddLog strWOID & ": Pri插入失败"
    End If

End Sub

Private Sub SaveWOHead(strWOID As String, lDieQty As Long)

    Dim stHeader As BillHeader

    stHeader.ID = GetSeqID()
    stHeader.qty = lDieQty
    stHeader.ORDERNAME = strWOID
    stHeader.ORDERTYPE = 1
    stHeader.EVENTTYPE = "CREATED"
    stHeader.ERPUSER = "Auto"
    stHeader.product = Trim(txtPN.Text)
    stHeader.RequestDate = Now
    stHeader.PLANSTARTDATE = Trim(txtOpenDate.Text)
    stHeader.PLANENDDATE = Trim(txtCloseDate.Text)
    stHeader.CUSTOMER = Trim(txtCusCode.Text)
    stHeader.SALESORDER = Trim(txtPO.Text)
    stHeader.MODIFYFLAG = 0
    stHeader.CustomerERPN = Trim(txtCusPN.Text)

    If stHeader.CUSTOMER = "GC" Then
        stHeader.CustomerERPN = Mid(stHeader.CustomerERPN, 1, InStr(stHeader.CustomerERPN, "-") - 1)

    End If
    
    stHeader.FABFACILITY = Trim(txtFABDevice.Text)
    stHeader.IMAGERREV = Trim(txtICR.Text)
    stHeader.DESIGNID = Trim(txtDesignedID.Text)
    stHeader.MLEVEL235 = Trim(txt235.Text)
    stHeader.MLEVEL260 = Trim(txt260.Text)
    stHeader.NGFLAG = Val(txtNGFlag.Text)
    stHeader.PARA1 = Trim(txtMarkingCode.Text)
    stHeader.PARA2 = Trim(txtPercent.Text)
    stHeader.PARA3 = Trim(txtCtyFab.Text)
    stHeader.PARA4 = Trim(txtType.Text)
    stHeader.PARA5 = Trim(txtPOItem.Text)
    stHeader.PARA6 = Trim(txtShipSite.Text)
    stHeader.PARA8 = Trim(txtWODept.Text)
    stHeader.PARA10 = Trim(txtDateCode.Text)
    stHeader.PROTECTIVE_FILM_APLD = Trim(txtPFA.Text)
    stHeader.Lot_Stauts = Trim(txtLotStatus.Text)
    stHeader.MPN = Trim(txtMPN.Text)
    
    Call AddWOHead(stHeader)

End Sub

Private Sub AddWOHead(stHeader As BillHeader)

    Dim strOra       As String

    Dim strSql       As String

    Dim strOrderDept As String
    
    strOrderDept = Right(stHeader.PARA8, Len(stHeader.PARA8) - InStr(stHeader.PARA8, "_"))

    If Len(strOrderDept) < 3 Then
        strOrderDept = "ERROR"

    End If
    
    strOra = "insert into ib_workorder (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANstARTDATE ,PLANENDDATE ," & _
       " CUstOMER ,SALESORDER,CUstOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
       "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PARA10,PROTECTIVE_FILM_APLD ,LOT_stATUS ,MPN)" & _
       " Values (" & stHeader.ID & ",'" & stHeader.ORDERNAME & "','" & stHeader.ORDERTYPE & "' ,'CREATED','" & stHeader.ERPUSER & "','" & stHeader.product & "'," & stHeader.qty & ",sysdate,to_date('" & stHeader.PLANSTARTDATE & "','yyyy-mm-dd'),to_date('" & stHeader.PLANENDDATE & "','yyyy-mm-dd')," & _
       " '" & stHeader.CUSTOMER & "','" & stHeader.SALESORDER & "','" & stHeader.CustomerERPN & "','" & stHeader.FABFACILITY & "','" & stHeader.IMAGERREV & "','" & stHeader.DESIGNID & "','" & stHeader.MLEVEL235 & "','" & stHeader.MLEVEL260 & "','" & stHeader.NGFLAG & "','" & stHeader.PARA1 & "'," & _
       "  '" & stHeader.PARA2 & "','" & stHeader.PARA3 & "','" & stHeader.PARA4 & "','" & stHeader.PARA5 & "','" & stHeader.PARA6 & "','" & stHeader.RequestDate & "','" & stHeader.PARA8 & "','" & stHeader.PARA10 & "','" & stHeader.PROTECTIVE_FILM_APLD & "','" & stHeader.Lot_Stauts & "'," & _
       " '" & stHeader.MPN & "')"
    
    If AddSql(strOra) = 0 Then
        AddLog strWOID & ": ib_workorder插入失败"
    End If
       
    strOra = "insert into ib_wohistory (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANstARTDATE ,PLANENDDATE ," & _
       " CUstOMER ,SALESORDER,CUstOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
       "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PARA10,PROTECTIVE_FILM_APLD ,LOT_stATUS ,MPN)" & _
       " Values (" & stHeader.ID & ",'" & stHeader.ORDERNAME & "','" & stHeader.ORDERTYPE & "' ,'CREATED','" & stHeader.ERPUSER & "','" & stHeader.product & "'," & stHeader.qty & ",sysdate,to_date('" & stHeader.PLANSTARTDATE & "','yyyy-mm-dd'),to_date('" & stHeader.PLANENDDATE & "','yyyy-mm-dd')," & _
       " '" & stHeader.CUSTOMER & "','" & stHeader.SALESORDER & "','" & stHeader.CustomerERPN & "','" & stHeader.FABFACILITY & "','" & stHeader.IMAGERREV & "','" & stHeader.DESIGNID & "','" & stHeader.MLEVEL235 & "','" & stHeader.MLEVEL260 & "','" & stHeader.NGFLAG & "','" & stHeader.PARA1 & "'," & _
       "  '" & stHeader.PARA2 & "','" & stHeader.PARA3 & "','" & stHeader.PARA4 & "','" & stHeader.PARA5 & "','" & stHeader.PARA6 & "','" & stHeader.RequestDate & "','" & stHeader.PARA8 & "','" & stHeader.PARA10 & "','" & stHeader.PROTECTIVE_FILM_APLD & "','" & stHeader.Lot_Stauts & "'," & _
       " '" & stHeader.MPN & "')"
    
    If AddSql(strOra) = 0 Then
        AddLog strWOID & ": ib_wohistory插入失败"
    End If
    
 
    strSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANstARTDATE ,PLANENDDATE ," & _
       " CUstOMER ,SALESORDER,CUstOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
       "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PARA10,PROTECTIVE_FILM_APLD ,LOT_stATUS ,MPN)" & _
       " Values (" & stHeader.ID & ",'" & stHeader.ORDERNAME & "','" & stHeader.ORDERTYPE & "' ,'CREATED','" & stHeader.ERPUSER & "','" & stHeader.product & "'," & stHeader.qty & ",GetDate(),convert(datetime,'" & stHeader.PLANSTARTDATE & "'),convert(datetime,'" & stHeader.PLANENDDATE & "')," & _
       " '" & stHeader.CUSTOMER & "','" & stHeader.SALESORDER & "','" & stHeader.CustomerERPN & "','" & stHeader.FABFACILITY & "','" & stHeader.IMAGERREV & "','" & stHeader.DESIGNID & "','" & stHeader.MLEVEL235 & "','" & stHeader.MLEVEL260 & "','" & stHeader.NGFLAG & "','" & stHeader.PARA1 & "'," & _
       "  '" & stHeader.PARA2 & "','" & stHeader.PARA3 & "','" & stHeader.PARA4 & "','" & stHeader.PARA5 & "','" & stHeader.PARA6 & "','" & stHeader.RequestDate & "','" & strOrderDept & "', '" & stHeader.PARA10 & "','" & stHeader.PROTECTIVE_FILM_APLD & "','" & stHeader.Lot_Stauts & "'," & _
       " '" & stHeader.MPN & "')"
    
    If AddSql2(strSql) = 0 Then
        AddLog strWOID & ": tblTSVworkorder插入失败"
    End If
    
End Sub

Private Sub AddWODetail(stDetail As BillDetail)

    Dim strOra As String

    Dim strSql As String

    strOra = "insert into ib_waferlist(ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & stDetail.ORDERNAME & "'," & " '" & stDetail.WAFERID & "'," & stDetail.DIEQTY & "," & stDetail.FGDIEQTY & ",'" & stDetail.WAFERLOT & "',100,'" & stDetail.MARKINGCODE & "')"

    strSql = "insert into [erpdata].[dbo].[tblTSVwaferlist](ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & stDetail.ORDERNAME & "'," & " '" & stDetail.WAFERID & "'," & stDetail.DIEQTY & "," & stDetail.FGDIEQTY & ",'" & stDetail.WAFERLOT & "',100,'" & stDetail.MARKINGCODE & "')"

    If AddSql(strOra) = 0 Then
        AddLog strWOID & ": ib_waferlist插入失败"
    End If
     
    If AddSql2(strSql) = 0 Then
        AddLog strWOID & ": tblTSVwaferlist插入失败"
    End If

    Call addLogTxt(stDetail.ORDERNAME, " 插入表:ib_waferlist: " & stDetail.WAFERID)
End Sub

Private Sub SaveWOLot(strLotID As String)

    Dim strCusCode As String

    Dim strpo      As String

    strCusCode = Replace(Trim(txtCusCode.Text), "(ICI)", "")
    strpo = Trim(txtPO.Text)

    If strCusCode = "AA" Or strCusCode = "AA(ON)" Then
            
        Call ONLotIDClose(strLotID)
    Else
        Call updateHeaderDateForGC(strLotID, "Y", 0, strpo)
            
    End If

End Sub

Private Sub SaveWODetail(ByVal strWOID As String, _
                         ByRef lWaferQty As Long, _
                         ByRef lDieQty As Long)

    Dim stDetail As BillDetail

    Dim strLotID As String

    Dim strWO    As String

    With fpS1

        For i = 1 To .MaxRows
            .Row = i
            .Col = 7
            
            If .Text = "1" Then
                .Col = 1
                strWO = UCase(Trim$(.Text))

                If strWO = strWOID Then
                    .Col = 1
                    stDetail.ORDERNAME = UCase(Trim$(.Text))
                    
                    .Col = 2
                    stDetail.WAFERLOT = UCase(Trim$(.Text))
                    
                    .Col = 3
                    stDetail.WAFERID = UCase(Trim$(.Text))
                    
                    .Col = 4
                    stDetail.DIEQTY = UCase(Trim$(.Text))
                    
                    .Col = 5
                    stDetail.FGDIEQTY = UCase(Trim$(.Text))
                    
                    .Col = 6
                    stDetail.MARKINGCODE = UCase(Trim$(.Text))
                    
                    Call AddWODetail(stDetail)
            
                    If stDetail.WAFERLOT <> strLotID Then
                        strLotID = stDetail.WAFERLOT
                
                        Call SaveWOLot(strLotID)

                    End If
   
                    lWaferQty = lWaferQty + 1
                    lDieQty = lDieQty + stDetail.DIEQTY
    
                End If

            End If

        Next i
    
    End With
   
End Sub

Private Sub SaveWOBom(strWOID As String, lWaferQty As Long, lDieQty As Long)

    Dim strSql As String

    Dim strPN  As String
    
    strPN = Trim$(txtPN.Text)
  
    If Frm_ProductionPlanNew.cbWOType = "重工工单" Or Frm_ProductionPlanNew.cbWOType = "Dummy工单" Or Frm_ProductionPlanNew.cbWOType = "玻璃工单" Then
        strSql = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记)" & "select b.ORDERNAME ,a.物料编号,'1','主选材料',SUM(convert(int, c.DIEQTY)),'1' " & "from erpdata..tblSmainM2 a ,erpdata..tblTSVworkorder b, erpdata .. tblTSVwaferlist c " & "where b.ORDERNAME = '" & strWOID & "'  and  b.PRODUCT = a.料号 and c.ORDERNAME = b.ORDERNAME " & "group by b.ORDERNAME ,a.物料编号"
        
    Else
        strSql = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) " & _
        " SELECT distinct  '" & strWOID & "',X.物料编号,'1','主选材料', " & " CAST( (CAST(X.用量 AS DECIMAL(18,8)) * '" & lWaferQty & "' ) AS  DECIMAL(18,3))  ,1 " & _
        " FROM ( SELECT b.材料规范编号, b.物料编号,sum(b.每只用量) as 用量 " & _
        " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
        " WHERE a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & strPN & "'  AND b.物料编号 NOT LIKE '%03.06%' " & _
        " GROUP by b.材料规范编号, b.物料编号 )  X  " & " UNION SELECT distinct  '" & strWOID & "',y.物料编号,'1','主选材料',  y.用量  ,1 FROM ( SELECT b.材料规范编号, b.物料编号,SUM(CONVERT(INT,d.DIEQTY)) as 用量 " & _
        " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b ,erpdata .. tblTSVworkorder c, erpdata .. tblTSVwaferlist d  " & _
        " WHERE a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & strPN & "'  " & " AND b.物料编号 LIKE '%03.06%' AND c.PRODUCT = a.物料编号 " & "  " & _
        " AND c.ORDERNAME = '" & strWOID & "' AND d.ORDERNAME = c.ORDERNAME " & " GROUP by b.材料规范编号, b.物料编号 ) y "

    End If

    If AddSql2(strSql) = 0 Then
        AddLog strWOID & ": Bom表插入失败"
    
    End If

End Sub

Private Sub SaveAsExcel()

    Dim xlsApp      As Excel.Application

    Dim xlsBook     As Excel.Workbook

    Dim xlsSheet    As Excel.Worksheet

    Dim i           As Long

    Dim J           As Long

    Dim StrFileName As String

    Dim strPartName As String

    On Error GoTo Ert

    Set xlsApp = CreateObject("Excel.Application")
    Set xlsBook = xlsApp.Workbooks.Add
    Set xlsSheet = xlsBook.Worksheets(1)

    With xlsApp
        .Rows(1).Font.Bold = True

    End With

    With fpS1

        For i = 0 To fpS1.MaxRows
            For J = 1 To fpS1.MaxCols + 3
                .Col = J
                .Row = i
            
                If .Col = 7 And .Text = "1" Then
                    If .Col = 8 Then
                        If .Row = 0 Then
                            xlsSheet.Cells(i + 1, J) = "厂内机种"
                        Else
                            xlsSheet.Cells(i + 1, J) = txtHTPN.Text

                        End If
        
                    ElseIf .Col = 9 Then

                        If .Row = 0 Then
                            xlsSheet.Cells(i + 1, J) = "成品料号"
                        Else
                            xlsSheet.Cells(i + 1, J) = txtPN.Text

                        End If
            
                    ElseIf .Col = 10 Then

                        If .Row = 0 Then
                            xlsSheet.Cells(i + 1, J) = "客户机种"
                        Else
                            xlsSheet.Cells(i + 1, J) = txtCusPN.Text

                        End If
            
                    Else
                        xlsSheet.Cells(i + 1, J) = .Text

                    End If

                End If
        
            Next J
       
        Next i

    End With

    xlsApp.Visible = True

    StrFileName = "C:\others\" & Format(Now(), "MMDDHHMMSS") & ".clsx"
    
    xlsBook.SaveAs StrFileName

    Set xlsApp = Nothing

    Exit Sub

Ert:

    If Not (xlsApp Is Nothing) Then
        
        Set xlsApp = Nothing

    End If

End Sub

Private Sub AddLog(strLog As String)
txtLog = txtLog & Time() & "->" & strLog & vbCrLf

End Sub
