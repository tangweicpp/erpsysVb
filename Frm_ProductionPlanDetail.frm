VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form Frm_ProductionPlanDetail 
   Caption         =   "工单明细"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13860
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
   ScaleHeight     =   11010
   ScaleWidth      =   13860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Height          =   11415
      Left            =   13800
      TabIndex        =   70
      Top             =   0
      Width           =   5655
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
         Height          =   11175
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   71
         Text            =   "Frm_ProductionPlanDetail.frx":0000
         Top             =   120
         Width           =   6255
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "全选/反选"
      Height          =   255
      Left            =   120
      TabIndex        =   65
      Top             =   3600
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   720
      TabIndex        =   63
      Top             =   3240
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "工单信息"
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13815
      Begin VB.TextBox txtNpiOwner 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txtMapping 
         Height          =   285
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox txtCusDieQty 
         Height          =   285
         Left            =   5520
         TabIndex        =   67
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtWOType 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   203
         Width           =   2175
      End
      Begin VB.TextBox txtPO 
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtCloseDate 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtOpenDate 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtLotType 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   1883
         Width           =   2175
      End
      Begin VB.TextBox txt37Pri 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1643
         Width           =   2175
      End
      Begin VB.TextBox txtWODept 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1403
         Width           =   2175
      End
      Begin VB.TextBox txtPN 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   1163
         Width           =   2175
      End
      Begin VB.TextBox txtCusPN 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   683
         Width           =   2175
      End
      Begin VB.TextBox txtHTPN 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   923
         Width           =   2175
      End
      Begin VB.TextBox txtCusCode 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   443
         Width           =   2175
      End
      Begin VB.TextBox txtMPN 
         Height          =   285
         Left            =   9240
         TabIndex        =   49
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox txtType 
         Height          =   285
         Left            =   9240
         TabIndex        =   47
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtDateCode 
         Height          =   285
         Left            =   5520
         TabIndex        =   45
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ComboBox cbCus 
         Height          =   315
         Left            =   5520
         TabIndex        =   43
         Top             =   173
         Width           =   1815
      End
      Begin VB.TextBox txtShipSite 
         Height          =   285
         Left            =   9240
         TabIndex        =   41
         Top             =   1883
         Width           =   2535
      End
      Begin VB.TextBox txtCusRequest 
         Height          =   285
         Left            =   9240
         TabIndex        =   39
         Top             =   1643
         Width           =   2535
      End
      Begin VB.TextBox txtPFA 
         Height          =   285
         Left            =   9240
         TabIndex        =   37
         Top             =   1403
         Width           =   2535
      End
      Begin VB.TextBox txtLotStatus 
         Height          =   285
         Left            =   9240
         TabIndex        =   34
         Top             =   1163
         Width           =   2535
      End
      Begin VB.TextBox txtPOItem 
         Height          =   285
         Left            =   9240
         TabIndex        =   32
         Top             =   923
         Width           =   2535
      End
      Begin VB.TextBox txtMM 
         Height          =   285
         Left            =   9240
         TabIndex        =   30
         Top             =   683
         Width           =   2535
      End
      Begin VB.TextBox txtCtyFab 
         Height          =   285
         Left            =   9240
         TabIndex        =   28
         Top             =   443
         Width           =   2535
      End
      Begin VB.TextBox txtPercent 
         Height          =   285
         Left            =   9240
         TabIndex        =   26
         Text            =   "25"
         Top             =   203
         Width           =   2535
      End
      Begin VB.TextBox txtMarkingCode 
         Height          =   285
         Left            =   5520
         TabIndex        =   24
         Top             =   1883
         Width           =   1815
      End
      Begin VB.TextBox txtNGFlag 
         Height          =   285
         Left            =   5520
         TabIndex        =   22
         Text            =   "Y"
         Top             =   1643
         Width           =   1815
      End
      Begin VB.TextBox txt260 
         Height          =   285
         Left            =   5520
         TabIndex        =   20
         Top             =   1403
         Width           =   1815
      End
      Begin VB.TextBox txt235 
         Height          =   285
         Left            =   5520
         TabIndex        =   18
         Top             =   1163
         Width           =   1815
      End
      Begin VB.TextBox txtDesignedID 
         Height          =   285
         Left            =   5520
         TabIndex        =   16
         Top             =   923
         Width           =   1815
      End
      Begin VB.TextBox txtICR 
         Height          =   285
         Left            =   5520
         TabIndex        =   14
         Top             =   683
         Width           =   1815
      End
      Begin VB.TextBox txtFABDevice 
         Height          =   285
         Left            =   5520
         TabIndex        =   12
         Top             =   443
         Width           =   1815
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "负责人: "
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
         Index           =   32
         Left            =   360
         TabIndex        =   73
         Top             =   2760
         Width           =   840
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
         TabIndex        =   68
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
         TabIndex        =   66
         Top             =   2640
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
         TabIndex        =   61
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
         TabIndex        =   52
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
         TabIndex        =   48
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
         TabIndex        =   46
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
         TabIndex        =   44
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "接口中的客户"
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
         Index           =   25
         Left            =   4320
         TabIndex        =   42
         Top             =   240
         Width           =   1140
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
         TabIndex        =   40
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
         TabIndex        =   38
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
         TabIndex        =   36
         Top             =   2400
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
         TabIndex        =   35
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
         TabIndex        =   33
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
         TabIndex        =   31
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
         TabIndex        =   29
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
         TabIndex        =   27
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
         TabIndex        =   25
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
         TabIndex        =   23
         Top             =   1920
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
         TabIndex        =   21
         Top             =   1680
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
         TabIndex        =   19
         Top             =   1440
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
         TabIndex        =   17
         Top             =   1200
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
         TabIndex        =   15
         Top             =   960
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
         TabIndex        =   13
         Top             =   720
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
         TabIndex        =   11
         Top             =   480
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   2384
         Width           =   840
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   10140
      Width           =   13860
      _ExtentX        =   24448
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
               Picture         =   "Frm_ProductionPlanDetail.frx":0015
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_ProductionPlanDetail.frx":08EF
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   3840
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   13361
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "序号"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "工单号"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "LotID"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "WaferID"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "TotalDie数量"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "GoodDie数量"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "打标码"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "OVT_JOB"
         Object.Width           =   3528
      EndProperty
   End
   Begin FPSpreadADO.fpSpread Fps 
      Height          =   7575
      Index           =   0
      Left            =   120
      TabIndex        =   72
      Top             =   3840
      Width           =   11895
      _Version        =   524288
      _ExtentX        =   20981
      _ExtentY        =   13361
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
      SpreadDesigner  =   "Frm_ProductionPlanDetail.frx":11C9
      Appearance      =   1
      TextTip         =   2
      AppearanceStyle =   0
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
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   240
      TabIndex        =   64
      Top             =   3315
      Width           =   420
   End
End
Attribute VB_Name = "Frm_ProductionPlanDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim lPart As Integer

Private Sub Check1_Click()
Dim i As Integer

If Check1.Value = 1 Then

    With ListView1

        For i = 1 To .ListItems.count
            .ListItems(i).Checked = True
        Next

    End With

    With Fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            .text = 1
        Next

    End With

ElseIf Check1.Value = 0 Then

    With ListView1

        For i = 1 To .ListItems.count
            .ListItems(i).Checked = False
        Next

    End With

    With Fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            .text = 0
        Next

    End With

End If

End Sub

Private Sub Form_Load()
Call Init

End Sub

Private Sub Init()
If Frm_ProductionPlan.cbWOType.text = "FO_CSP工单" Then
    Fps(0).Visible = True
    ListView1.Visible = False
    MsgBox "FO_CSP工单需要手动维护GoodDies数量, 请选择对应Wafer,编辑合理的Dies数量", vbInformation, "提示"
Else
    Fps(0).Visible = False
    ListView1.Visible = True

End If

InitFps
InitWOHead
InitWODetail

End Sub

Private Sub InitFps()

With Fps(0)
    .MaxRows = 0
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsBest
    .Col = -1
    .Row = -1
    .Lock = True
    .Col = 1
    .CellType = CellTypeCheckBox
    .TypeHAlign = TypeHAlignCenter
    .TypeVAlign = TypeVAlignCenter
    .SetText 1, 0, "选择"
    .SetText 2, 0, "工单号"
    .SetText 3, 0, "LotID"
    .SetText 4, 0, "WaferID"
    .SetText 5, 0, "GoodDies"
    .SetText 6, 0, "打标码"
    .Col = 1
    .Lock = False
    .Col = 5
    .Lock = False
    .BackColor = vbGreen
    .ColWidth(1) = 4

End With

End Sub

Private Sub InitWOHead()
Dim rs1 As New ADODB.Recordset

cbCus.AddItem ("AA")
cbCus.AddItem ("WLC")
cbCus.AddItem ("GC")
cbCus.AddItem ("SX")
cbCus.AddItem ("SY")
cbCus.AddItem ("BD")
cbCus.AddItem ("HY")
cbCus.AddItem ("MG")
InitWOInfo
txtWOType.text = Trim(Frm_ProductionPlan.cbWOType.text)
txtCusCode.text = Trim(Frm_ProductionPlan.cbCusCode.text)
txtCusPN.text = Trim(Frm_ProductionPlan.txtCusPN.text)
txtPO.text = Trim$(Frm_ProductionPlan.cbHTPN.text)
If (txtCusCode.text = "AA" Or txtCusCode.text = "AA(ON)") And (txtWOType.text <> "玻璃工单" And txtWOType.text <> "DUMMY工单") Then
    txtMPN.text = txtCusPN.text
    txtDateCode.text = GetONDateCode()
    txtMarkingCode.text = GetONWoMarkingCode(Trim(txtMPN.text))
    txtCusPN.text = GetONOPN_WSG(Trim(txtCusPN.text))
Else
    txtMPN.text = Trim(Frm_ProductionPlan.txtCusPN.text)
    txtDateCode.text = GetONDateCode()
    txtMarkingCode.text = GetONWoMarkingCode(Trim(txtMPN.text))

End If

TxtPN.text = Trim(Frm_ProductionPlan.cbPn.text)
txtHTPN.text = Trim(Frm_ProductionPlan.cbHTPN.text)
txtWODept.text = Trim(Frm_ProductionPlan.txtWODept.text)
txt37Pri.text = Trim(Frm_ProductionPlan.cb37Pri.text)

Select Case Frm_ProductionPlan.cbLotType.ListIndex

    Case 0
        txtLotType.text = "M"

    Case 1
        txtLotType.text = "E"

    Case 2
        txtLotType.text = "Q"

    Case 3
        txtLotType.text = "D"

    Case Else
        txtLotType.text = "M"

End Select

' 查询CUSDIE
Set rs1.ActiveConnection = OraConnect
rs1.Source = "select customerdieqty,mapping from tbltsvnpiproduct where qtechptno2 = '" & Trim(TxtPN.text) & "' and customershortname = '" & txtCusCode.text & "' and customerptno1 = '" & Trim(Frm_ProductionPlan.txtCusPN.text) & "' and qtechptno = '" & txtHTPN.text & "' "
rs1.Open , , adOpenStatic, adLockReadOnly, adCmdText
txtCusDieQty.text = Trim(rs1("customerdieqty")) & ""
txtMapping.text = Trim(rs1("MAPPING")) & ""
txtOpenDate.text = Trim(Frm_ProductionPlan.DTPicker1(0).Value)
txtCloseDate.text = Trim(Frm_ProductionPlan.DTPicker1(1).Value)
txtNPIOwner.text = Trim(Frm_ProductionPlan.txtNPIOwner.text)

End Sub

Private Sub InitWOInfo()
Dim strCusCode As String
Dim strLotID   As String
Dim oiRS       As New ADODB.Recordset
Dim rs         As New ADODB.Recordset

With Frm_ProductionPlan.List1

    For i = 0 To .ListCount - 1
        If .Selected(i) = True Then
            strLotID = Trim$(.List(i))

        End If

    Next

End With

strCusCode = UCase(Trim(Frm_ProductionPlan.cbCusCode.text))
Set oiRS = GetOIData(strCusCode, strLotID)
If (oiRS.RecordCount > 0) Then
    txtFABDevice.text = "" & Trim(oiRS("fabrication_facility"))
    txtICR.text = "" & Trim(oiRS("imager_customer_rev"))
    txtDesignedID.text = "" & Trim(oiRS("design_id"))
    txt260.text = "" & Trim(oiRS("shipping_mst_260"))
    txt235.text = "" & Trim(oiRS("shipping_mst_level"))
    txtMarkingCode.text = "" & Trim(oiRS("encoded_mark_id"))
    txtCtyFab.text = "" & Trim(oiRS("country_of_fab"))
    txtMM.text = "" & Trim(oiRS("micron_material"))
    txtPOItem.text = "" & Trim(oiRS("po_item"))
    txtLotStatus.text = "" & Trim(oiRS("lot_status"))
    txtType.text = "" & Trim(oiRS("PROBE_SHIP_PART_TYPE"))
    txtPFA.text = IIf("" & Trim(oiRS("protective_film_apld")) = "YES", "PF", "" & Trim(oiRS("protective_film_apld")))
    txtCusRequest.text = "" & Trim(oiRS("lot_priority"))
    txtShipSite.text = "" & Trim(oiRS("ship_site"))
    txtDateCode.text = "" & GetONDateCode()
    If txtShipSite.text = "Qtech" And strCusCode = "AA" Then
        cbCus.text = "WLC"
    ElseIf txtShipSite.text = "SG" And strCusCode = "AA" Then
        cbCus.text = "AA"
    ElseIf strCusCode = "GC" Then
        cbCus.text = "GC"

    End If

End If

End Sub

Private Sub InitWODetail()
Dim strLotID   As String
Dim strWOID    As String
Dim strLotList As String

With Frm_ProductionPlan.List1
    If Frm_ProductionPlan.Check2.Value = 1 Then

        ' 批量工单
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
                strLotID = Trim$(.List(i))
                strWOID = Frm_ProductionPlan.GetWOID()
                Call ListData(strLotID, strWOID)

            End If

        Next
    Else
        ' 单工单
        strWOID = Frm_ProductionPlan.GetWOID()

        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
                strLotID = Trim$(.List(i))
                Call ListData(strLotID, strWOID)

            End If

        Next

    End If

End With

With Frm_ProductionPlan.ListView1
    If Frm_ProductionPlan.Check2.Value = 1 Then

        ' 批量工单
        For i = 1 To .ListItems.count
            If .ListItems(i).Checked Then
                strLotID = Trim$(.ListItems(i).SubItems(1))
                If InStr(strLotList, strLotID) = 0 Then
                    strLotList = strLotList & strLotID & ","
                    strWOID = Frm_ProductionPlan.GetWOID()
                    Call ListData(strLotID, strWOID)

                End If

            End If

        Next
    Else
        ' 单工单
        strWOID = Frm_ProductionPlan.GetWOID()

        For i = 1 To .ListItems.count
            If .ListItems(i).Checked Then
                strLotID = Trim$(.ListItems(i).SubItems(1))
                If InStr(strLotList, strLotID) = 0 Then
                    strLotList = strLotList & strLotID & ","
                    strLotID = Trim$(.ListItems(i).SubItems(1))
                    Call ListData(strLotID, strWOID)

                End If

            End If

        Next

    End If

End With

With ListView1

    For i = 1 To .ListItems.count
        .ListItems(i).Checked = True
    Next

End With

End Sub
   
Private Sub updateMarkingCode(strLotID As String, strCusPN As String)
Dim rs              As New ADODB.Recordset
Dim strSql          As String
Dim strWaferID      As String
Dim strid           As String
Dim strMarkingCode  As String
Dim strMarkingCode2 As String
Dim strHTPN         As String
Dim strYear         As String
Dim strWeek         As String
Dim strWeekSeq      As String

strWeekSeq = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"

Select Case txtCusCode.text

    Case "GD108", "HK080"
        strSql = "select b.substrateid,b.productid, a.mtrl_num from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists " & " (select 1 from ib_waferlist c where c.waferid = b.substrateid) order by b.substrateid "
        Set rs = Get_OracleRs(strSql)
        If Not rs.EOF Then
            rs.MoveFirst

            Do While Not rs.EOF
                strWaferID = Trim("" & rs(0))
                strMarkingCode = Trim$("" & rs(1))
                strHTPN = Trim$("" & rs(2))
                strid = Get_OracleStr("select filename from  mappingdatatest where substrateid = '" & strWaferID & "'")
                strMarkingCode2 = Trim(Get_OracleStr("select RETICLE_LEVEL_71 || '\\' || RETICLE_LEVEL_72 || '\\' || to_char(sysdate + 2,'YYWW') || '\\' || source_batch_id from customeroitbl_test where id = '" & strid & "'"))
                If strHTPN = "YGD108B1" Then
                    AddSql ("update mappingdatatest set productid = '" & strMarkingCode2 & "' where substrateid = '" & strWaferID & "' ")
                    AddSql2 ("update ERPBASE..tblmappingData set productid = '" & strMarkingCode2 & "' where substrateid = '" & strWaferID & "' ")

                End If

                rs.MoveNext
            Loop

        End If

        If strCusPN = "GW1N-LV4CS72" Then
            strSql = "select b.substrateid,b.productid, a.mtrl_num from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists " & " (select 1 from ib_waferlist c where c.waferid = b.substrateid) order by b.substrateid "
            Set rs = Get_OracleRs(strSql)
            If Not rs.EOF Then
                rs.MoveFirst

                Do While Not rs.EOF
                    strWaferID = Trim("" & rs(0))
                    If InStr(strWaferID, "+") > 0 Then
                        strMarkingCode = Trim$("" & rs(1))
                        strHTPN = Trim$("" & rs(2))
                        strid = Get_OracleStr("select filename from  mappingdatatest where substrateid = '" & strWaferID & "'")
                        ' strMarkingCode2 = Trim(Get_OracleStr("select RETICLE_LEVEL_71 || '\\' || RETICLE_LEVEL_72 || '\\' || to_char(sysdate + 2,'YYWW') || 'B' || '\\' || source_batch_id ||  '@@' || RETICLE_LEVEL_71 || '\\' || RETICLE_LEVEL_73 || '\\' || to_char(sysdate + 1,'YYWW') || 'B' || '\\' || source_batch_id " & " from customeroitbl_test where id = '" & strID & "'"))
                        strMarkingCode2 = Trim(Get_OracleStr("select 'GW1N-LV4\\CS72C6/I5' || '\\' || to_char(sysdate + 2,'YYWW') || 'B' || '\\' || source_batch_id ||  '@@' || 'GW1N-LV4\\CS72C5/I4' || '\\' || to_char(sysdate + 1,'YYWW') || 'B' || '\\' || source_batch_id " & " from customeroitbl_test where id = '" & strid & "'"))
                        If Left$(strMarkingCode2, 4) = "CS72" Then
                            MsgBox "YGD108B4打标码获取错误,请联系IT处理,本次工单无效", vbCritical, "警告"
                            Exit Sub

                        End If

                        AddSql ("update mappingdatatest set productid = '" & strMarkingCode2 & "' where substrateid = '" & strWaferID & "' ")
                        AddSql2 ("update ERPBASE..tblmappingData set productid = '" & strMarkingCode2 & "' where substrateid = '" & strWaferID & "' ")

                    End If

                    rs.MoveNext
                Loop

            End If

        End If

        If Trim(txtHTPN.text) = "YGD10806" Then
            strSql = "select b.substrateid,b.productid, a.mtrl_num from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists " & " (select 1 from ib_waferlist c where c.waferid = b.substrateid) order by b.substrateid "
            Set rs = Get_OracleRs(strSql)
            If Not rs.EOF Then
                rs.MoveFirst

                Do While Not rs.EOF
                    strWaferID = Trim("" & rs(0))
                    strMarkingCode = Trim$("" & rs(1))
                    strHTPN = Trim$("" & rs(2))
                    strid = Get_OracleStr("select filename from  mappingdatatest where substrateid = '" & strWaferID & "'")
                    strMarkingCode2 = Trim(Get_OracleStr("select 'GW1NZ-LV1\\CS16C6/I5' || '\\' || to_char(sysdate + 2,'YYWW') || '\\' || source_batch_id ||  '@@' || 'GW1NZ-LV1\\CS16C5/I4' || '\\' || to_char(sysdate + 1,'YYWW') || '\\' || source_batch_id " & " from customeroitbl_test where id = '" & strid & "'"))
                    If Left$(strMarkingCode2, 4) = "CS72" Then
                        MsgBox "YGD10806打标码获取错误,请联系IT处理,本次工单无效", vbCritical, "警告"
                        Exit Sub

                    End If

                    AddSql ("update mappingdatatest set productid = '" & strMarkingCode2 & "' where substrateid = '" & strWaferID & "' ")
                    AddSql2 ("update ERPBASE..tblmappingData set productid = '" & strMarkingCode2 & "' where substrateid = '" & strWaferID & "' ")
                    rs.MoveNext
                Loop

            End If

        End If

        If Trim(txtHTPN.text) = "YGD10811" Then
            strSql = "select b.substrateid,b.productid, a.mtrl_num from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists " & " (select 1 from ib_waferlist c where c.waferid = b.substrateid) order by b.substrateid "
            Set rs = Get_OracleRs(strSql)
            If Not rs.EOF Then
                rs.MoveFirst

                Do While Not rs.EOF
                    strWaferID = Trim("" & rs(0))
                    strMarkingCode = Trim$("" & rs(1))
                    strHTPN = Trim$("" & rs(2))
                    strid = Get_OracleStr("select filename from  mappingdatatest where substrateid = '" & strWaferID & "'")
                    strMarkingCode2 = Trim(Get_OracleStr("select 'GW1NZ-ZV1\\CS16I3' || '\\' || to_char(sysdate + 2,'YYWW') || '\\' || source_batch_id ||  '@@' || 'GW1NZ-ZV1\\CS16I2' || '\\' || to_char(sysdate + 1,'YYWW') || '\\' || source_batch_id " & " from customeroitbl_test where id = '" & strid & "'"))
                    If Left$(strMarkingCode2, 4) = "CS72" Then
                        MsgBox "YGD10811打标码获取错误,请联系IT处理,本次工单无效", vbCritical, "警告"
                        Exit Sub

                    End If

                    AddSql ("update mappingdatatest set productid = '" & strMarkingCode2 & "' where substrateid = '" & strWaferID & "' ")
                    AddSql2 ("update ERPBASE..tblmappingData set productid = '" & strMarkingCode2 & "' where substrateid = '" & strWaferID & "' ")
                    rs.MoveNext
                Loop

            End If

        End If
        
          '2020-2-20 ZYF
       If Trim(txtHTPN.text) = "YGD10810B" Then
            strSql = "select b.substrateid,b.productid, a.mtrl_num from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists " & " (select 1 from ib_waferlist c where c.waferid = b.substrateid) order by b.substrateid "
            Set rs = Get_OracleRs(strSql)
            If Not rs.EOF Then
                rs.MoveFirst

                Do While Not rs.EOF
                    strWaferID = Trim("" & rs(0))
                    strMarkingCode = Trim$("" & rs(1))
                    strHTPN = Trim$("" & rs(2))
                    strid = Get_OracleStr("select filename from  mappingdatatest where substrateid = '" & strWaferID & "'")
                    strMarkingCode2 = Trim(Get_OracleStr("select 'GW1NS-LV4\\CS49C6/I5' || '\\' || to_char(sysdate + 2,'YYWW') || '\\' || source_batch_id ||  '@@' || 'GW1NS-LV4\\CS49C5/I4' || '\\' || to_char(sysdate + 2,'YYWW') || '\\' || source_batch_id " & " from customeroitbl_test where id = '" & strid & "'"))
                    If Left$(strMarkingCode2, 4) = "CS72" Then
                        MsgBox "YGD10810B打标码获取错误,请联系IT处理,本次工单无效", vbCritical, "警告"
                        Exit Sub

                    End If

                    AddSql ("update mappingdatatest set productid = '" & strMarkingCode2 & "' where substrateid = '" & strWaferID & "' ")
                    AddSql2 ("update ERPBASE..tblmappingData set productid = '" & strMarkingCode2 & "' where substrateid = '" & strWaferID & "' ")
                    rs.MoveNext
                Loop

            End If

        End If
    
     
        '2020-4-10 ZYF
        If Trim(txtHTPN.text) = "YGD10812" Then
            strSql = "select b.substrateid,b.productid, a.mtrl_num from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists " & " (select 1 from ib_waferlist c where c.waferid = b.substrateid) order by b.substrateid "
            Set rs = Get_OracleRs(strSql)
            If Not rs.EOF Then
                rs.MoveFirst

                Do While Not rs.EOF
                    strWaferID = Trim("" & rs(0))
                    strMarkingCode = Trim$("" & rs(1))
                    strHTPN = Trim$("" & rs(2))
                    strid = Get_OracleStr("select filename from  mappingdatatest where substrateid = '" & strWaferID & "'")
                    strMarkingCode2 = Trim(Get_OracleStr("select 'GW1NS-UX2\\CS36UC6/I5' || '\\' || to_char(sysdate,'YYWW') || '\\' || source_batch_id ||  '@@' || 'GW1NS-UX2\\CS36UC5/I4' || '\\' || to_char(sysdate,'YYWW') || '\\' || source_batch_id " & " from customeroitbl_test where id = '" & strid & "'"))
                    If Left$(strMarkingCode2, 4) = "CS72" Then
                        MsgBox "YGD10812打标码获取错误,请联系IT处理,本次工单无效", vbCritical, "警告"
                        Exit Sub

                    End If

                    AddSql ("update mappingdatatest set productid = '" & strMarkingCode2 & "' where substrateid = '" & strWaferID & "' ")
                    AddSql2 ("update ERPBASE..tblmappingData set productid = '" & strMarkingCode2 & "' where substrateid = '" & strWaferID & "' ")
                    rs.MoveNext
                Loop

            End If

        End If
    
        '2020-3-22 ZYF
        If Trim(txtHTPN.text) = "YGD10813" Then
            strSql = "select b.substrateid,b.productid, a.mtrl_num from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists " & " (select 1 from ib_waferlist c where c.waferid = b.substrateid) order by b.substrateid "
            Set rs = Get_OracleRs(strSql)
            If Not rs.EOF Then
                rs.MoveFirst

                Do While Not rs.EOF
                    strWaferID = Trim("" & rs(0))
                    strMarkingCode = Trim$("" & rs(1))
                    strHTPN = Trim$("" & rs(2))
                    strid = Get_OracleStr("select filename from  mappingdatatest where substrateid = '" & strWaferID & "'")
                    strMarkingCode2 = Trim(Get_OracleStr("select 'GW1N-LV9\\CS81MC6/I5' || '\\' || to_char(sysdate + 2,'YYWW') || '\\' || source_batch_id ||  '@@' || 'GW1N-LV9\\CS49C5/I4' || '\\' || to_char(sysdate + 2,'YYWW') || '\\' || source_batch_id " & " from customeroitbl_test where id = '" & strid & "'"))
                    If Left$(strMarkingCode2, 4) = "CS72" Then
                        MsgBox "YGD10810B打标码获取错误,请联系IT处理,本次工单无效", vbCritical, "警告"
                        Exit Sub

                    End If

                    AddSql ("update mappingdatatest set productid = '" & strMarkingCode2 & "' where substrateid = '" & strWaferID & "' ")
                    AddSql2 ("update ERPBASE..tblmappingData set productid = '" & strMarkingCode2 & "' where substrateid = '" & strWaferID & "' ")
                    rs.MoveNext
                Loop

            End If

        End If
    
    

    Case "SH48"

        Select Case strCusPN

            Case "BL24SA128B", "BL24SA128B-CT", "BL24SA128C-CS", "BL24SA128C-CT", "BL24SA128B-CS"
                strSql = "select b.substrateid from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists " & " (select 1 from ib_waferlist c where c.waferid = b.substrateid) order by b.substrateid "
                strYear = Right(Year(Now), 1)
                strWeek = Right("00" & DatePart("WW", Now), 2)
                strWeek = Mid$(strWeekSeq, CLng(strWeek), 1)
                strMarkingCode = "7" & strYear & strWeek
                AddSql ("update mappingdatatest set productid = '" & strMarkingCode & "' where substrateid in (select b.substrateid from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists (select 1 from ib_waferlist c where c.waferid = b.substrateid) )    ")

            Case "BL24SA64B-CS"
                strSql = "select b.substrateid from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists " & " (select 1 from ib_waferlist c where c.waferid = b.substrateid) order by b.substrateid "
                strYear = Right(Year(Now), 1)
                strWeek = Right("00" & DatePart("WW", Now), 2)
                strWeek = Mid$(strWeekSeq, CLng(strWeek), 1)
                strMarkingCode = "6" & strYear & strWeek
                AddSql ("update mappingdatatest set productid = '" & strMarkingCode & "' where substrateid in (select b.substrateid from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists (select 1 from ib_waferlist c where c.waferid = b.substrateid) )    ")

        End Select

    Case "SH105"

        Select Case Trim(txtHTPN.text)

            Case "XSH10501B"     ' Changed by: Project Administrator at: 2019/7/19-13:33:40 NPI:黄和鸣
                strSql = "select b.substrateid from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists " & " (select 1 from ib_waferlist c where c.waferid = b.substrateid) order by b.substrateid "
                strYear = Right(Year(Now), 2)
                strWeek = Right("00" & DatePart("WW", Now), 2)
                AddSql ("update mappingdatatest set productid =  replace(productid,'\\','\\' || '" & strWeek & "' || '" & strYear & "' || 'T' || '\\') where length(productid) = 14 and substrateid in (select b.substrateid from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists (select 1 from ib_waferlist c where c.waferid = b.substrateid) ) ")

        End Select

    Case "33"

        Select Case Trim(txtHTPN.text)

            Case "X33005B"     ' Changed by: Project Administrator at: 2019/7/19-13:33:40 NPI:黄和鸣
                strSql = "select b.substrateid from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists " & " (select 1 from ib_waferlist c where c.waferid = b.substrateid) order by b.substrateid "
                strYear = Right(Year(Now), 2)
                strWeek = Right("00" & DatePart("WW", Now), 2)
                AddSql ("update mappingdatatest set productid = 'CVSMicro' || '\\' || 'CV8035D' || '\\' || '" & strYear & "' || '" & strWeek & "' where substrateid in (select b.substrateid from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists (select 1 from ib_waferlist c where c.waferid = b.substrateid) ) ")

        End Select

    Case Else

End Select

Select Case Trim(txtHTPN.text)

    Case "Y68559B"
        strSql = "select a.FAB_CONV_ID from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists " & " (select 1 from ib_waferlist c where c.waferid = b.substrateid) order by b.substrateid "
        Dim strFabDevice As String

        strFabDevice = Get_OracleStr(strSql)
        Dim strMarkCode As String

        strMarkCode = "BNA" & "\\" & Mid$("KMNPRSTVWXYZ", Year(Now) - 2018, 1) & Right("00" & DatePart("WW", Now), 2) & "\\" & Right(strFabDevice, 3)
        AddSql ("update mappingdatatest set productid = '" & strMarkCode & "' where substrateid in (select b.substrateid from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and a.mpn_desc = '" & strCusPN & "' and b.filename = to_char(a.id) and not exists (select 1 from ib_waferlist c where c.waferid = b.substrateid) ) ")

End Select

End Sub

Private Function CheckMarkingCode(strMarkingCode As String) As Boolean
CheckMarkingCode = False
Dim strHTPN   As String
Dim strCustPN As String

strCustPN = Trim(txtCusPN.text)
strHTPN = Trim(txtHTPN.text)

Select Case strHTPN

    Case "XAH01701B"
        If Left$(strMarkingCode, 4) <> "6001" Or Right(strMarkingCode, 3) <> "WFF" Then
            MsgBox "打标码错误,请联系IT确认", vbCritical, "警告"
            Exit Function

        End If

    Case "XAH01702B"
        If Left$(strMarkingCode, 4) <> "1619" Or Right(strMarkingCode, 3) <> "WFF" Then
            MsgBox "打标码错误,请联系IT确认", vbCritical, "警告"
            Exit Function

        End If

End Select

CheckMarkingCode = True

End Function

Private Sub ListData(strLotID As String, strWOID As String)
Dim rs                     As New ADODB.Recordset
Dim listRS2                As New ADODB.Recordset
Dim listRS3                As New ADODB.Recordset
Dim strCusPN               As String
Dim strMarkCode            As String
Dim strCheckMarkingCodeRes As String
Dim strMarkingCode         As String
Dim strWaferID             As String
Dim strLotWaferID          As String
Dim strCustPN              As String
Dim strHTPN                As String
Dim strCustCode            As String


Dim strdevice            As String
Dim rsdevice             As New ADODB.Recordset
Dim strpo                As String
Dim rspo                 As New ADODB.Recordset


strCustCode = txtCusCode.text
strCustPN = txtCusPN.text
strHTPN = txtHTPN.text
strCusPN = Trim$(txtCusPN.text)


If InStr(strWOID, "SM") = 0 And InStr(strWOID, "SG") = 0 And InStr(strWOID, "SS") = 0 Then

strdevice = " SELECT * FROM erptemp..ht_price_control a,erptemp..ht_price_config b   WHERE a.cust_id = '" & strCustCode & "' AND a.cust_device = '" & strCusPN & "' AND a.flag = 0 AND b.cust_id =  a.cust_id AND B.FLAG = 0 AND (b.PO_PRICE = 'Y' OR B.OPENPO = 'Y') "

If rsdevice.State = adStateOpen Then rsdevice.Close
rsdevice.Open strdevice, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If Not rsdevice.EOF Then

 strpo = " select hh.* from customeroitbl_test h  ,TSV_MD_POPrice HH  " & _
         " where h.source_batch_id =  '" & strLotID & "'  AND HH.CUSTOMERSHORTNAME = H.CUSTOMERSHORTNAME and hh.pt = h.mpn_desc and hh.po_num = h.po_num "
  
If rspo.State = adStateOpen Then rspo.Close
rspo.Open strpo, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If rspo.EOF Then

MsgBox "LOT" & strLotID & "无PO信息", vbInformation, "提示"

Exit Sub
End If

End If
End If



On Error GoTo there

Call updateMarkingCode(strLotID, strCusPN)
Set rs.ActiveConnection = OraConnect
If Trim(txtCusCode.text) = "AA" Or Trim(txtCusCode.text) = "AA(ON)" Then
    If (Right(Trim(UCase(Frm_ProductionPlan.cbWO)), 2) = "ST" Or Right(Trim(UCase(Frm_ProductionPlan.cbWO)), 2) = "ET") And txtHTPN.text <> "GX251AM" Then
        rs.Source = "select distinct '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, ONSTMarkingCodeSeq.QTSeq(b.substrateid,b.lotid) as productid, '' from mappingdatatest b where b.lotid = '" & strLotID & "' and  not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"
    Else
        rs.Source = "select distinct  '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, ONMarkingCodeSeq.QTSeq(b.substrateid) as productid,''  from WAFERDETAILTMP a, mappingdatatest b where  b.lotid = '" & strLotID & "' and  not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"

    End If

    If Trim(txtHTPN.text) = "X76001B" Then
        rs.Source = "select distinct  '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, ONMarkingCodeSeq.QTSeq2(b.substrateid) as productid,''  from WAFERDETAILTMP a, mappingdatatest b where  b.lotid = '" & strLotID & "' and  not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"

    End If

ElseIf Trim(txtCusCode.text) = "SG005" Or Trim(txtCusCode.text) = "US026" Then
    If Frm_ProductionPlan.cbWOType = "重工工单" Or Frm_ProductionPlan.cbWOType = "FO_CSP工单" Then
        rs.Source = "select distinct  '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, replace(b.productid,'_','') as productid, a.reticle_level_71 from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and b.filename = to_char(a.id) and a.mpn_desc = '" & strCusPN & "' and b.substrateid like '%+%' and  not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"
    ElseIf Frm_ProductionPlan.cbWOType = "硅基工单" Then
        rs.Source = "select distinct '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, replace(b.productid,'_','') as productid,a.reticle_level_71 from customeroitbl_test a,mappingdatatest b where b.lotid = '" & strLotID & "' and b.filename = to_char(a.id) and a.mpn_desc = '" & strCusPN & "' and instr(b.substrateid, '+') = 0 and not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"
    Else
        rs.Source = "select distinct '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, replace(b.productid,'_','') as productid,a.reticle_level_71 from customeroitbl_test a,mappingdatatest b where b.lotid = '" & strLotID & "' and b.filename = to_char(a.id) and a.mpn_desc = '" & strCusPN & "' and not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"

    End If

ElseIf Trim(txtCusCode.text) = "57" Then
    If Frm_ProductionPlan.cbWOType = "重工工单" Or Frm_ProductionPlan.cbWOType = "FO_CSP工单" Then
        rs.Source = "select distinct  '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, b.productid as productid, a.reticle_level_71 from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and b.filename = to_char(a.id) and a.mpn_desc = '" & strCusPN & "' and b.substrateid like '%+%' and  not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"
    ElseIf Frm_ProductionPlan.cbWOType = "硅基工单" Then
        rs.Source = "select distinct '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, b.productid as productid,a.reticle_level_71 from customeroitbl_test a,mappingdatatest b where b.lotid = '" & strLotID & "' and b.filename = to_char(a.id) and a.mpn_desc = '" & strCusPN & "' and instr(b.substrateid, '+') = 0 and not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"
    Else
        rs.Source = "select distinct '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, b.productid as productid,a.reticle_level_71 from customeroitbl_test a,mappingdatatest b where b.lotid = '" & strLotID & "' and b.filename = to_char(a.id) and a.mpn_desc = '" & strCusPN & "' and not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"

    End If

Else
    If Frm_ProductionPlan.cbWOType = "重工工单" Or Frm_ProductionPlan.cbWOType = "FO_CSP工单" Then
        rs.Source = "select distinct  '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, replace(b.productid,'_','') as productid, a.reticle_level_71  from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and b.filename = to_char(a.id) and a.mpn_desc = '" & strCusPN & "' and b.substrateid like '%+%' and  not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"
    ElseIf Frm_ProductionPlan.cbWOType = "硅基工单" Then
        rs.Source = "select distinct '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, replace(b.productid,'_','') as productid, a.reticle_level_71  from customeroitbl_test a,mappingdatatest b where b.lotid = '" & strLotID & "' and b.filename = to_char(a.id) and a.mpn_desc = '" & strCusPN & "' and instr(b.substrateid, '+') = 0 and not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"
    Else
        rs.Source = "select distinct '" & strWOID & "' as woid, b.lotid,b.substrateid,b.passbincount+b.failbincount, b.passbincount, replace(b.productid,'_','') as productid, a.reticle_level_71  from customeroitbl_test a,mappingdatatest b where b.lotid = '" & strLotID & "' and b.filename = to_char(a.id) and a.mpn_desc = '" & strCusPN & "' and not exists(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by woid, b.substrateid"

    End If

End If

rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount > 0 Then

    For i = 1 To rs.RecordCount
        strMarkingCode = Trim("" & rs!PRODUCTID)
        strLotID = Trim$("" & rs!LOTID)
        strLotWaferID = Trim$("" & rs!SUBSTRATEID)
        strWaferID = Replace(Replace$(strLotWaferID, strLotID, ""), "+", "")
        '打标码更新
        Call updateMarkingCode2(strCustCode, strCustPN, strHTPN, strLotID, strWaferID, strLotWaferID, strMarkingCode)
        '增加最终获取函数
        'Call updateMarkingCodeFinal(strHTPN, strLotWaferID)
        
        '增加检查
        If Check_MarkingcodeByHT(strHTPN, strLotWaferID) = False Then
            Exit Sub
        End If
        
        'EU010特殊检查
        If CheckOrderData() = False Then
            Exit Sub

        End If

        rs.MoveNext
    Next

End If

rs.Close
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount > 0 Then
    rs.MoveFirst

    For i = 1 To rs.RecordCount
        strMarkingCode = Trim("" & rs!PRODUCTID)
        strLotID = Trim$("" & rs!LOTID)
        strLotWaferID = Trim$("" & rs!SUBSTRATEID)
        strWaferID = Replace(Replace$(strLotWaferID, strLotID, ""), "+", "")
        '打标码校验
        If Frm_ProductionPlan.cbWOType <> "Dummy工单" And Frm_ProductionPlan.cbWOType <> "FO_CSP工单" And Frm_ProductionPlan.cbWOType <> "硅基工单" And Frm_ProductionPlan.cbWOType <> "玻璃工单" Then
            strCheckMarkingCodeRes = Get_OracleStr("select CHECK_MARKINGCODE('" & strMarkingCode & "','" & txtCusCode.text & "','" & txtCusPN.text & "','" & txtHTPN.text & "','" & strLotID & "','" & strWaferID & "','" & strLotWaferID & "') from dual ")
            If strCheckMarkingCodeRes <> "0" Then
                MsgBox strCheckMarkingCodeRes, vbCritical, "提示"
                Exit Sub

            End If

            If Trim(txtCusCode.text) = "SX" Or Trim(txtCusCode.text) = "HJ" Or Trim(txtCusCode.text) = "TJ003" Or Trim(txtCusCode.text) = "JS140" Or Trim(txtCusCode.text) = "BJ153" Then
                Set listRS2 = GetMARK2(Trim$("" & rs("substrateid")), TxtPN.text)
                If listRS2.RecordCount > 0 Then
                    MsgBox "打标码长度异常"
                    Exit Sub

                End If

                If Get_OracleCnt("select * from tbltsvnpiproduct where customerptno1 = '" & txtCusPN & "' and qtechptno2 = '" & TxtPN & "'  and  marking_code > 0 ") > 0 Then
                    Set listRS3 = GetMARK3(Trim$("" & rs("productid")))
                    If listRS3.RecordCount > 0 Then
                        MsgBox "存在重复打标码"
                        Exit Sub

                    End If

                End If

            End If

        End If

        ' 显示ListViewData
        Set ListItem = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)

        For j = 1 To ListView1.ColumnHeaders.count - 1
            If IsNull(rs(j - 1)) Then
                strt = " "
            Else
                strt = Trim(rs(j - 1))

            End If

            ListItem.ListSubItems.Add , , strt
        Next j

        rs.MoveNext
    Next
Else
    MsgBox "没有查询到明细", vbInformation, "提示"

End If

If Frm_ProductionPlan.cbWOType = "FO_CSP工单" Then
    Dim RsNew As New ADODB.Recordset

    Set RsNew.ActiveConnection = OraConnect
    RsNew.Source = "select distinct '' as 选择,'" & strWOID & "' as 工单号,b.lotid as LotID,b.substrateid as WaferID, c.customerdieqty as TotalDie数量,replace(b.productid, '_', '') as 打标码 " & "from customeroitbl_test a, mappingdatatest b, tbltsvnpiproduct c where b.lotid = '" & strLotID & "' and b.filename = to_char(a.id) and c.customerptno1 = '" & strCusPN & "' " & "and c.qtechptno2 = '" & TxtPN.text & "' and c.customershortname = '" & txtCusCode & "' and a.mpn_desc = '" & strCusPN & "' and b.substrateid like '%+%' and not exists " & "(select 1 from ib_waferlist c where c.waferid = b.substrateid) order by 工单号, b.substrateid"
    RsNew.Open , , adOpenStatic, adLockReadOnly, adCmdText

    With Fps(0)
        RsNew.MoveFirst

        For i = 1 To RsNew.RecordCount
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1
            .text = "1"
            .Col = 2
            .text = Trim$("" & RsNew(1))
            .Col = 3
            .text = Trim$("" & RsNew(2))
            .Col = 4
            .text = Trim$("" & RsNew(3))
            .Col = 5
            .text = Trim$("" & RsNew(4))
            .Col = 6
            .text = Trim$("" & RsNew(5))
            RsNew.MoveNext
        Next

    End With

End If

Screen.MousePointer = 0
rs.Close
Set rs = Nothing
Exit Sub
there:
Screen.MousePointer = 0
MsgBox "查询失败！" + Chr(13) + "原因:" + Err.DESCRIPTION, vbInformation, Me.Caption

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       updateMarkingCode2
' Description:       更新打标码
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/11/19-11:16:50
'
' Parameters :       strCustCode (String)
'                    strCustPN (String)
'                    strHTPN (String)
'                    strLotID (String)
'                    strWaferID (String)
'                    strLotWaferID (String)
'                    strOldMarkingCode (String)
'--------------------------------------------------------------------------------
Private Sub updateMarkingCode2(strCustCode As String, _
                               strCustPN As String, _
                               strHTPN As String, _
                               strLotID As String, _
                               strWaferID As String, _
                               strLotWaferID As String, _
                               strOldMarkingCode As String)
Dim strNewMarkingCode As String
Dim strSql            As String

If Frm_ProductionPlan.cbWOType = "Dummy工单" Then
    Exit Sub
End If

Select Case strHTPN

    Case "X76006B"
        If CLng(strWaferID) < 13 Or CLng(strWaferID) > 19 Then
            MsgBox "waferID不可小于13或大于19,请联系IT", vbCritical, "错误"
            Exit Sub

        End If

        If CLng(strWaferID) >= 13 And CLng(strWaferID) <= 15 Then
            strNewMarkingCode = "DC-1" & "\\" & Mid$("KABCDE", Year(Now) - 2018, 1) & Mid("0123456789ACDEFHKLNPRSTUXYZ", (DatePart("WW", Now) + 1) \ 2, 1) & "P2"
        ElseIf CLng(strWaferID) >= 16 And CLng(strWaferID) <= 17 Then
            strNewMarkingCode = "DC-2" & "\\" & Mid$("KABCDE", Year(Now) - 2018, 1) & Mid("0123456789ACDEFHKLNPRSTUXYZ", (DatePart("WW", Now) + 1) \ 2, 1) & "P2"
        Else
            strNewMarkingCode = "DC-3" & "\\" & Mid$("KABCDE", Year(Now) - 2018, 1) & Mid("0123456789ACDEFHKLNPRSTUXYZ", (DatePart("WW", Now) + 1) \ 2, 1) & "P2"

        End If

    Case "X76008B"
        strNewMarkingCode = "6D" & Mid(strLotID, 5, 2) & "\\" & Mid$("KABCDE", Year(Now) - 2018, 1) & Mid("0123456789ACDEFHKLNPRSTUXYZ", (DatePart("WW", Now) + 1) \ 2, 1) & "P2"

    Case "X76007B"
        strNewMarkingCode = "VJ" & Mid(strLotID, 9, 2) & "\\" & Mid$("KABCDE", Year(Now) - 2018, 1) & Mid("0123456789ACDEFHKLNPRSTUXYZ", (DatePart("WW", Now) + 1) \ 2, 1) & "P2"

    Case "X76010B"
        strNewMarkingCode = "6F" & Mid(strLotID, 5, 2) & "\\" & Mid$("KABCDE", Year(Now) - 2018, 1) & Mid("0123456789ACDEFHKLNPRSTUXYZ", (DatePart("WW", Now) + 1) \ 2, 1) & "P2"

    Case "YHW74007B"
        strNewMarkingCode = "REDPINE" & "\\" & "SIGNALS" & "\\" & "RS8112-W0" & "\\" & strLotID & strWaferID & "\\" & Right$(Year(Now), 2) & DatePart("WW", Now) & " " & "1.0"

    Case "X33006B"
        strNewMarkingCode = "CVSMicro" & "\\" & "CV8035D" & "\\" & Right$(Year(Now), 2) & DatePart("WW", Now)

        '王雪 2020-01-10
    Case "XAH017A5B"
        If Len(DatePart("WW", Now)) = 1 Then
            strNewMarkingCode = "2105" & "\\" & Right$(Year(Now), 2) & "0" & DatePart("WW", Now) & "\\" & Mid("ABCDEFGHIJKLMNOPQRSTUVWXY", Right(strWaferID, 2), 1) & Mid(strLotID, InStr(1, strLotID, ".") - 3, 3) & "\\" & "WYY"
        ElseIf Len(DatePart("WW", Now)) = 2 Then
            strNewMarkingCode = "2105" & "\\" & Right$(Year(Now), 2) & DatePart("WW", Now) & "\\" & Mid("ABCDEFGHIJKLMNOPQRSTUVWXY", Right(strWaferID, 2), 1) & Mid(strLotID, InStr(1, strLotID, ".") - 3, 3) & "\\" & "WYY"

        End If

    Case "XSH48B07B", "XSH48A07B", "XSH48009B", "XSH48A09B"
        strNewMarkingCode = "6" & Right$(Year(Now), 1) & Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", DatePart("WW", Now), 1)

    Case "XSH48A02B", "XSH48A04B", "XSH48008B"
        strNewMarkingCode = "7" & Right$(Year(Now), 1) & Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", DatePart("WW", Now), 1)

    Case "XAH017B2B"
        If Len(DatePart("WW", Now)) = 1 Then
            strNewMarkingCode = "1651" & "\\" & Right$(Year(Now), 2) & "0" & DatePart("WW", Now) & "\\" & Mid("ABCDEFGHIJKLMNOPQRSTUVWXY", Right(strWaferID, 2), 1) & Mid(strLotID, InStr(1, strLotID, ".") - 3, 3) & "\\" & "WFF"
        ElseIf Len(DatePart("WW", Now)) = 2 Then
            strNewMarkingCode = "1651" & "\\" & Right$(Year(Now), 2) & DatePart("WW", Now) & "\\" & Mid("ABCDEFGHIJKLMNOPQRSTUVWXY", Right(strWaferID, 2), 1) & Mid(strLotID, InStr(1, strLotID, ".") - 3, 3) & "\\" & "WFF"

        End If

    Case "XSH48002B", "XSH48B02B" '王雪 2020-01-19
        strNewMarkingCode = "7" & Right$(Year(Now), 1) & Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmenopqrstuvwxyz", DatePart("WW", Now), 1)
    
    Case "XAH017C2B", "XAH01702B", "XAH017D2B"
        strNewMarkingCode = "1619" & "\\" & Right$(Year(Now), 2) & Right("00" & DatePart("WW", Now), 2) & "\\" & Mid("ABCDEFGHIJKLMNOPQRSTUVWXY", strWaferID, 1) & Mid(strLotID, InStr(strLotID, ".") - 3, 3) & "\\" & "WFF"
    
    Case "YJS16001C"
        strNewMarkingCode = "SRXC311E" & "\\" & "StorMicro" & "\\" & Replace(strLotWaferID, "+", "")
    
    Case Else
        Exit Sub

End Select

AddSql ("update mappingdatatest set productid = '" & strNewMarkingCode & "' where substrateid = '" & strLotWaferID & "' ")
AddSql2 ("update ERPBASE..tblmappingData set productid = '" & strNewMarkingCode & "' where substrateid = '" & strLotWaferID & "' ")

End Sub

Private Sub updateMarkingCodeFinal(strHTPN As String, strLotWaferID As String)

Dim strSql As String
Dim strNewMarkingCode As String

strSql = "SELECT DEFINED_FLAG FROM TBL_MARKINGCODE_REP WHERE HT_PN = '" & strHTPN & "'"
If Get_OracleStr(strSql) = "N" Then   ' 说明需要更新工单DateCode
    strSql = "select Get_MarkingCode_YF('" & strLotWaferID & "','', '', '', '', '', '', '', '', '', '', '', '', '', '','','','','') from dual"
    strNewMarkingCode = Get_OracleStr(strSql)
    
    AddSql ("update mappingdatatest set productid = '" & strNewMarkingCode & "' where substrateid = '" & strLotWaferID & "' ")
    AddSql2 ("update ERPBASE..tblmappingData set productid = '" & strNewMarkingCode & "' where substrateid = '" & strLotWaferID & "' ")
End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckOrderData
' Description:       检查工单数据
' Created by :       Project Administrator
' Machine    :       0-354AD8C194ED4
' Date-Time  :       2019-11-25-13:45:31
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function CheckOrderData() As Boolean
Dim strSql    As String
Dim strCustPN As String

CheckOrderData = False
If txtCusCode.text = "EU010" Then
    strCustPN = Replace(txtCusPN.text, "@", "")
    strSql = "select * from EU010_REFERENCE where cust_device = '" & strCustPN & "'"
    If Get_OracleCnt(strSql) = 0 Then
        MsgBox "EU010机种特殊信息未维护,请联系NPI维护好再开工单", vbInformation, "警告"
        Exit Function

    End If

End If

CheckOrderData = True

End Function

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
Frm_ProductionPlan.tb1.Buttons("SEARCH").Enabled = True
Frm_ProductionPlan.tb1.Buttons("PREVIEW").Enabled = False
Dim strWOID    As String
Dim bRtn       As Boolean
Dim i          As Integer
Dim iPart      As Integer
Dim rs1        As New ADODB.Recordset
Dim Rs2        As New ADODB.Recordset
Dim thisDies   As String
Dim strLotID   As String
Dim strTaxType As String
Dim strlot     As String
Dim j          As Integer

j = 1
i = 0
iPart = 0
bRtn = False
ProgressBar1.Value = 1
strWOID = ""
strLotID = ""
'37 ('HCLAMP1831ZBTFT','RCLAMP5021ZATFT','UCLAMP5031Z2CTKT','RCLAMP1031ZCTFT')机种对应
If InStr(txtCusPN.text, "HCLAMP1831ZBTFT") Or InStr(txtCusPN.text, "RCLAMP5021ZATFT") Or InStr(txtCusPN.text, "UCLAMP5031Z2CTKT") Or InStr(txtCusPN.text, "RCLAMP1031ZCTFT") Then

    For i = 1 To ListView1.ListItems.count
        If ListView1.ListItems(i).Checked And ListView1.ListItems(i).SubItems(2) <> strLotID Then
            strLotID = ListView1.ListItems(i).SubItems(2)
            Call SaveWOLot_37(strLotID, ListView1.ListItems(i).SubItems(1), Trim(txtCusPN.text))

        End If

    Next

End If

'转换
If Frm_ProductionPlan.cbWOType = "FO_CSP工单" Then

    With Fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .text = 1 Then
                ListView1.ListItems(i).Checked = True
                .Col = 5
                thisDies = Trim$(.text)
                ListView1.ListItems(i).SubItems(4) = thisDies
                ListView1.ListItems(i).SubItems(5) = thisDies
                .Col = 4
                AddSql ("update mappingdatatest set passbincount ='" & thisDies & "' where  substrateid = '" & Trim(.text) & "'  ")
                AddSql2 ("update [ERPBASE].[dbo].[tblmappingData] set passbincount ='" & thisDies & "' where  substrateid = '" & Trim(.text) & "'  ")
            Else
                ListView1.ListItems(i).Checked = False

            End If

        Next

    End With

End If

For i = 1 To ListView1.ListItems.count
    If ListView1.ListItems(i).Checked Then
        bRtn = True

    End If

Next i

If bRtn = False Then
    MsgBox "请勾选要开立的WAFER ID", vbCritical, "警告"
    Toolbar1.Buttons("SAVE").Enabled = True
    ProgressBar1.Value = 0
    Toolbar1.Buttons("SAVE").Enabled = True
    lPart = 0
    Exit Sub

End If

ListView1.Visible = True
Fps(0).Visible = False
' PO价格卡控
If Frm_ProductionPlan.cbWOType <> "Dummy工单" And Frm_ProductionPlan.cbWOType <> "FO_CSP工单" And Frm_ProductionPlan.cbWOType <> "硅基工单" And Frm_ProductionPlan.cbWOType <> "玻璃工单" Then
    If Trim(txtCusCode.text) <> "AA" And Trim(txtCusCode.text) <> "76" Then

        For i = 1 To ListView1.ListItems.count
            If ListView1.ListItems(i).Checked And (InStr(ListView1.ListItems(i).SubItems(3), "+") = 0) Then
                If chkPOPrice(ListView1.ListItems(i).SubItems(3)) = False Then
                    Toolbar1.Buttons("SAVE").Enabled = True
                    lPart = 0
                    Exit Sub

                End If

            End If

        Next i

    End If

End If

' 37PO, JOB确认
If Frm_ProductionPlan.cbWOType <> "Dummy工单" And Frm_ProductionPlan.cbWOType <> "FO_CSP工单" And Frm_ProductionPlan.cbWOType <> "硅基工单" And Frm_ProductionPlan.cbWOType <> "玻璃工单" Then
    If Trim(txtCusCode.text) = "37" Then
        If Frm_ProductionPlan.cbWOType <> "重工工单" Then

            For N = 1 To ListView1.ListItems.count
                If ListView1.ListItems(N).Checked Then
                    If Check37JOB(ListView1.ListItems(N).SubItems(3)) = False Then
                        MsgBox " 此JOB已开立过工单，请确认同JOB其他WAFER的工单信息", vbCritical, "警告"
                        Toolbar1.Buttons("SAVE").Enabled = True
                        lPart = 0
                        Exit Sub

                    End If

                End If

            Next N

        End If

        For i = 1 To ListView1.ListItems.count
            If ListView1.ListItems(i).Checked Then
                If Check37PO(ListView1.ListItems(i).SubItems(3)) = False Then
                    MsgBox "37客户, 市场部未维护订单PO信息或JOBID, 禁止开立工单, 请联系市场部确认", vbCritical, "警告"
                    Toolbar1.Buttons("SAVE").Enabled = True
                    lPart = 0
                    Exit Sub

                End If

            End If

        Next i

    End If

End If

' 保税非保税卡控
For i = 1 To ListView1.ListItems.count
    If ListView1.ListItems(i).Checked Then
        strTaxType = Get_OracleStr("select distinct SUBSTRATETYPE from mappingdatatest where substrateid = '" & ListView1.ListItems(i).SubItems(3) & "'")

        Select Case strTaxType

            Case "A"    ' 保税
                If Left$(ListView1.ListItems(i).SubItems(1), 1) <> "A" Then
                    MsgBox "WaferID: " & ListView1.ListItems(i).SubItems(3) & " 市场部维护的是保税" & vbCrLf & "但本次工单号前缀不是A，工单号不正确", vbExclamation, "警告"
                    Exit Sub

                End If

            Case "B"    ' 非保税
                If Left$(ListView1.ListItems(i).SubItems(1), 1) <> "B" Then
                    MsgBox "WaferID: " & ListView1.ListItems(i).SubItems(3) & " 市场部维护的是非保税" & vbCrLf & "但本次工单号前缀不是B，工单号不正确", vbExclamation, "警告"
                    Exit Sub

                End If

            Case Else

        End Select

    End If

Next i

'37DateCode确认
'If Frm_ProductionPlan.cbCusCode.Text = "37" And Frm_ProductionPlan.cbWOType <> "Dummy工单" Then
'    Set rs1.ActiveConnection = OraConnect
'    rs1.Source = " select * from tbltsvnpiproduct a where a.customershortname in ( '37')  and   instr(a.struckstr1,'TR') >0  and a.customerptno1 = '" & Trim$(Frm_ProductionPlan.txtCusPN.Text) & "'"
'    rs1.Open , , adOpenStatic, adLockReadOnly, adCmdText
'    If rs1.RecordCount > 0 Then
'
'        For i = 1 To ListView1.ListItems.Count
'            If ListView1.ListItems(i).Checked And (InStr(ListView1.ListItems(i).SubItems(3), "+") > 0) Then
'                If Check37DATECODE(Replace(ListView1.ListItems(i).SubItems(3), "+", "")) = False Then
'                    MsgBox "37客户, 市场部未维护DATECODE信息, 禁止开立工单, 请联系市场部确认", vbCritical, "警告"
'                    Toolbar1.Buttons("SAVE").Enabled = True
'                    lPart = 0
'                    Exit Sub
'
'                End If
'
'            End If
'
'        Next i
'
'    End If
'
'End If
'Mapping确认
If txtMapping.text = "Y" And (Frm_ProductionPlan.cbWOType <> "重工工单" And Frm_ProductionPlan.cbWOType <> "玻璃工单" And Frm_ProductionPlan.cbWOType <> "Dummy工单" And Frm_ProductionPlan.cbWOType <> "FO_CSP工单" And Frm_ProductionPlan.cbWOType <> "硅基工单") Then
    Dim checkLot As String

    checkLot = ""

    For i = 1 To ListView1.ListItems.count
        If ListView1.ListItems(i).Checked Then
            If ListView1.ListItems(i).SubItems(2) <> checkLot Then
                If InStr(ListView1.ListItems(i).SubItems(3), "+") = 0 Then
                    If CheckMapping(ListView1.ListItems(i).SubItems(2)) = True Then
                        MsgBox "市场部未维护订单MAPPING信息, 禁止开立工单, 请联系市场部确认", vbCritical, "警告"
                        Toolbar1.Buttons("SAVE").Enabled = True
                        lPart = 0
                        Exit Sub

                    End If

                    checkLot = ListView1.ListItems(i).SubItems(2)

                End If

            End If

        End If

    Next i

End If

'GrossDies确认
If (Frm_ProductionPlan.cbWOType <> "重工工单" And Frm_ProductionPlan.cbWOType <> "FO_CSP工单" And Frm_ProductionPlan.cbWOType <> "玻璃工单" And Frm_ProductionPlan.cbCusCode.text <> "37" And Frm_ProductionPlan.cbWOType <> "Dummy工单" And Frm_ProductionPlan.cbWOType <> "硅基工单") Then

    For i = 1 To ListView1.ListItems.count
        If ListView1.ListItems(i).Checked Then
            If Frm_ProductionPlan.cbCusCode.text = "KR001" And InStr(ListView1.ListItems(i).SubItems(3), "+") > 0 Then
                Exit For

            End If

            If ListView1.ListItems(i).SubItems(4) <> txtCusDieQty.text Then
                MsgBox "GROSSDIE数量错误", vbCritical, "警告"
                Toolbar1.Buttons("SAVE").Enabled = True
                lPart = 0
                Exit Sub

            End If

        End If

    Next i

    '校验WAFER库存
    strlot = ""

End If

For i = 1 To ListView1.ListItems.count
    If ListView1.ListItems(i).Checked Then
        If ListView1.ListItems(i).SubItems(1) <> strWOID Then
            strWOID = ListView1.ListItems(i).SubItems(1)
            iPart = iPart + 1

        End If

    End If

Next i

strWOID = ""
' SG005/US026 OVT_JOB检查
If Frm_ProductionPlan.cbCusCode.text = "SG005" Or Frm_ProductionPlan.cbCusCode.text = "US026" Then
    Dim strPWOVT_JOB As String

    strPWOVT_JOB = ""

    For i = 1 To ListView1.ListItems.count
        If ListView1.ListItems(i).Checked Then
            If ListView1.ListItems(i).SubItems(1) <> strWOID Then
                strWOID = ListView1.ListItems(i).SubItems(1)
                strPWOVT_JOB = ListView1.ListItems(i).SubItems(7)
            Else
                If ListView1.ListItems(i).SubItems(7) <> strPWOVT_JOB Then
                    MsgBox "工单:" & ListView1.ListItems(i).SubItems(1) & "包含多种OVT_JOB, 不允许开立在一张工单, 请确认", vbInformation, "警告"
                    Toolbar1.Buttons("SAVE").Enabled = True
                    lPart = 0
                    Exit Sub

                End If

            End If

        End If

    Next i

End If

lPart = 20 * (1 / iPart)
strWOID = ""

For i = 1 To ListView1.ListItems.count
    If ListView1.ListItems(i).Checked Then
        If ListView1.ListItems(i).SubItems(1) <> strWOID Then
            strWOID = ListView1.ListItems(i).SubItems(1)
            Call SaveWOID(strWOID)
            Call addLogTxt(strWOID & ":new", txtLog.text)

        End If

    End If

Next i

ProgressBar1.Value = 100
SaveAsExcel

End Sub
    
Private Sub SaveWOID(strWOID As String)
AddLog strWOID & ": 工单正在开立...."
If ProgressBar1.Value < 100 Then
    ProgressBar1.Value = ProgressBar1.Value + lPart

End If

If SaveOldData(strWOID) = False Then
    Exit Sub

End If

ProgressBar1.Value = ProgressBar1.Value + lPart
If SaveNewData(strWOID) = False Then
    Exit Sub

End If

ProgressBar1.Value = ProgressBar1.Value + lPart
Sleep (500)
ProgressBar1.Value = ProgressBar1.Value + lPart
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
Dim strora As String
Dim strSql As String
Dim lRtn   As Long
Dim i      As Integer

On Error GoTo DealError

Cnn.BeginTrans
INIadoCon.BeginTrans
' 1. 表shop_order_detail
strora = "insert into shop_order_detail(SHOP_ORDER,CUST_LOT_ID,WAFER_ID,GROSS_DIE_QTY,GOOD_DIE_QTY,MARK_CODE)" & "select a.ordername as SHOP_ORDER,a.waferlot as CUST_LOT_ID,a.waferid as WAFER_ID,a.dieqty      as GROSS_DIE_QTY,a.fgdieqty    as GOOD_DIE_QTY, " & "decode(b.customer,'AA',d.markingcodefirst || a.markingcode,a.markingcode) as MARK_CODE from ib_waferlist a inner join ib_workorder b on b.ordername = a.ordername " & "left join  CUSTOMERMPNAttributes d on d.part = b.mpn where  b.ordername  in ('" & strWOID & "')"
If AddSql(strora) = 0 Then
    AddLog strWOID & ": shop_order_detail插入失败"
    GoTo DealError

End If

' 2. 表SHOP_ORDER_PROPERTY
strora = "select SHOP_ORDER_PROPERTY_PKG.SHOP_ORDER_PROPERTY('" & strWOID & "')  from dual"
If AddSql(strora) = 0 Then
    AddLog strWOID & ": SHOP_ORDER_PROPERTY插入失败"
    GoTo DealError

End If

' 3. 表shop_order
strora = "insert into shop_order(SHOP_ORDER,PRD_ID,PRD_VER,ERP_ROUTING,ORDER_QTY,CUST_LOT_QTY,PLAN_STAR_DATE,PLAN_END_DATE,MANF_DEPT,MANF_DEPT_DESC,LOT_TYPE,PRIORITY,PKG,CUST_ID,ERP_CREATE_DATE,CREATOR,flag,ht_device,RELEASE_TYPE) " & _
   " select trim(a.ordername) as SHOP_ORDER, trim(b.product) as PRD_ID, 'A' as PRD_VER, '' as ERP_ROUTING,trim(COUNT(distinct A.WAFERID)) as ORDER_QTY, trim(COUNT(distinct A.WAFERLOT)) as CUST_LOT_QTY, trim(B.PLANSTARTDATE) AS PLAN_STAR_DATE, " & _
   " trim(B.PLANENDDATE) AS PLAN_END_DATE,  trim(B.PARA8) AS MANF_DEPT, trim(g.manf_dept_desc) AS MANF_DEPT_DESC,trim(e.lot_type) as LOT_TYPE,trim(decode(e.pri, 'Hot Lot', 1, 'Super Hot Lot', 1, 4)) as PRIORITY,trim(f.pkg_type) as PKG, " & _
   " trim(shop_order_property_pkg.SHOP_ORDER_CUSTOMER(b.customer)) AS CUST_ID, trim(b.erpcreatedate) as ERP_CREATE_DATE,'" & gUserName & "' as CREATOR, '8' as flag, trim(f.qtechptno) as ht_device, '1' as RELEASE_TYPE " & _
   " from ib_waferlist a, ib_workorder b, MAPPINGDATATEST  C, PJ_WO_PRI e,tbltsvnpiproduct f, MES_DEPT g " & _
   " where b.ordername = a.ordername and f.customerptno1 = b.mpn AND B.CUSTOMER = F.CUSTOMERSHORTNAME and b.ordername = '" & strWOID & "'  AND C.SUBSTRATEID = a.waferid  and e.wo = b.ordername and f.qtechptno2 = b.product and  rownum = 1  and g.manf_dept = substr(b.para8, 1, instr(b.para8, '_') - 1) " & _
   " group by a.ordername, b.product, B.PLANSTARTDATE, B.PLANENDDATE, B.PARA8,e.lot_type, e.pri,f.pkg_type,shop_order_property_pkg.SHOP_ORDER_CUSTOMER(b.customer), b.erpcreatedate, e.creat_by, f.qtechptno, g.manf_dept_desc"
If AddSql(strora) = 0 Then
    AddLog strWOID & ": shop_order插入失败"
    GoTo DealError

End If

' 4. 表ERPBASE..TblERPFLToME
strSql = "insert into ERPBASE..TblERPFLToME (STOCK_TYPE,STOCK_ID,PRD_ID,PRD_VER,QTY,PRD_DATE,EFF_DATE,SHOP_ORDER,SupSN,Flag,Memo,CreateDate,FStauts,HeaderID) " & "select 'W',upper(b.ORDERNAME) + upper(c.WAFERLOT), b.PRODUCT,'A',COUNT(*),GETDATE() - 1,GETDATE() + 300,b.ORDERNAME,c.WAFERLOT,0,'',GETDATE(),'','' from erpdata .. tblTSVworkorder b,erpdata .. tblTSVwaferlist c where c.ORDERNAME = b.ORDERNAME and b.ORDERNAME in ( '" & strWOID & "' ) group by b.PRODUCT, b.ORDERNAME, c.WAFERLOT"
If AddSql2(strSql) = 0 Then
    AddLog strWOID & ": ERPBASE..TblERPFLToME插入失败"
    GoTo DealError

End If

Cnn.CommitTrans
INIadoCon.CommitTrans
strSql = "insert into erpdata..shop_order select * from (select * from  OPENQUERY(ORACLEDB, 'SELECT * from shop_order' )) AA where  AA.shop_order = '" & strWOID & "'  "
AddSql2 (strSql)
SaveNewData = True
Exit Function
DealError:
Cnn.RollbackTrans
INIadoCon.RollbackTrans
MsgBox "新工单接口保存失败！" + Chr(13) + "原因:" + Err.DESCRIPTION, vbInformation, Me.Caption
AddLog strWOID & ": 新工单接口保存失败插入失败"

End Function

Private Sub SaveWOPri(ByVal strWOID As String)
Dim strora     As String, strora2 As String
Dim strPri     As String
Dim strPri2    As String
Dim strLotType As String
Dim struser    As String

strPri = Trim(txt37Pri.text)
strLotType = Trim$(txtLotType.text)
strPri2 = Trim$(Frm_ProductionPlan.cb37Pri2.text)
struser = gUserName
strora = "insert into PJ_WO_PRI(wo,pri,great_date,lot_type,Creat_by,return) values('" & strWOID & "','" & strPri & "',to_char(sysdate,'YYYY-MM-DD'),'" & strLotType & "', '" & struser & "','" & strPri2 & "')"
strora2 = "insert into erpbase..PJ_WO_PRI(wo,pri,great_date,lot_type,Creat_by,""return"") values('" & strWOID & "','" & strPri & "',CONVERT(varchar(30),getdate(),23 ),'" & strLotType & "', '" & struser & "','" & strPri2 & "')"
If AddSql(strora) = 0 Or AddSql2(strora2) = 0 Then
    AddLog strWOID & ": Pri插入失败"

End If

End Sub

Private Sub SaveWOHead(strWOID As String, lDieQty As Long)
Dim stHeader As BillHeader
Dim typeId   As Integer

Select Case txtWOType.text

    Case "一般工单"
        typeId = 1

    Case "再加工工单"
        typeId = 5

    Case "委外工单"
        typeId = 7

    Case "重工委外工单"
        typeId = 8

    Case "拆件式工单"
        typeId = 11

    Case "预测工单"
        typeId = 13

    Case "试产工单"
        typeId = 15

    Case Else
        typeId = 0

End Select

stHeader.id = GetSeqID()
stHeader.QTY = lDieQty
stHeader.ORDERNAME = strWOID
stHeader.ORDERTYPE = typeId
stHeader.EVENTTYPE = "CREATED"
stHeader.ERPUSER = txtNPIOwner.text
stHeader.product = Trim(TxtPN.text)
stHeader.RequestDate = Now
stHeader.PLANSTARTDATE = Trim(txtOpenDate.text)
stHeader.PLANENDDATE = Trim(txtCloseDate.text)
stHeader.CUSTOMER = Trim(txtCusCode.text)
stHeader.SALESORDER = Trim(txtPO.text)
stHeader.MODIFYFLAG = 0
stHeader.CustomerERPN = Trim(txtCusPN.text)
If stHeader.CUSTOMER = "GC" And InStr(stHeader.CUSTOMER, "-") > 0 Then
    stHeader.CustomerERPN = Mid(stHeader.CustomerERPN, 1, InStr(stHeader.CustomerERPN, "-") - 1)

End If

stHeader.FABFACILITY = Trim(txtFABDevice.text)
stHeader.IMAGERREV = Trim(txtICR.text)
stHeader.DESIGNID = Trim(txtDesignedID.text)
stHeader.MLEVEL235 = Trim(txt235.text)
stHeader.MLEVEL260 = Trim(txt260.text)
stHeader.NGFLAG = Val(txtNGFlag.text)
stHeader.PARA1 = Trim(txtMarkingCode.text)
stHeader.PARA2 = Trim(txtPercent.text)
stHeader.PARA3 = Trim(txtCtyFab.text)
stHeader.PARA4 = Trim(txtType.text)
stHeader.PARA5 = Trim(txtPOItem.text)
stHeader.PARA6 = Trim(txtShipSite.text)
stHeader.PARA8 = Trim(txtWODept.text)
stHeader.PARA10 = Trim(txtDateCode.text)
stHeader.PROTECTIVE_FILM_APLD = Trim(txtPFA.text)
stHeader.Lot_Stauts = Trim(txtLotStatus.text)
stHeader.MPN = Trim(txtMPN.text)
Call AddWOHead(stHeader)

End Sub

Private Sub AddWOHead(stHeader As BillHeader)
Dim strora       As String
Dim strSql       As String
Dim strOrderDept As String

strOrderDept = Right(stHeader.PARA8, Len(stHeader.PARA8) - InStr(stHeader.PARA8, "_"))
If Len(strOrderDept) < 3 Then
    strOrderDept = "ERROR"

End If

strora = "insert into ib_workorder (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
   " CUstOMER ,SALESORDER,CUstOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
   "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PARA10,PROTECTIVE_FILM_APLD ,LOT_stATUS ,MPN)" & _
   " Values (" & stHeader.id & ",'" & stHeader.ORDERNAME & "','" & stHeader.ORDERTYPE & "' ,'CREATED','" & stHeader.ERPUSER & "','" & stHeader.product & "'," & stHeader.QTY & ",sysdate,to_date('" & stHeader.PLANSTARTDATE & "','yyyy-mm-dd'),to_date('" & stHeader.PLANENDDATE & "','yyyy-mm-dd')," & _
   " '" & stHeader.CUSTOMER & "','" & stHeader.SALESORDER & "','" & stHeader.CustomerERPN & "','" & stHeader.FABFACILITY & "','" & stHeader.IMAGERREV & "','" & stHeader.DESIGNID & "','" & stHeader.MLEVEL235 & "','" & stHeader.MLEVEL260 & "','" & stHeader.NGFLAG & "','" & stHeader.PARA1 & "'," & _
   "  '" & stHeader.PARA2 & "','" & stHeader.PARA3 & "','" & stHeader.PARA4 & "','" & stHeader.PARA5 & "','" & stHeader.PARA6 & "','" & stHeader.RequestDate & "','" & stHeader.PARA8 & "','" & stHeader.PARA10 & "','" & stHeader.PROTECTIVE_FILM_APLD & "','" & stHeader.Lot_Stauts & "'," & _
   " '" & stHeader.MPN & "')"
If AddSql(strora) = 0 Then
    AddLog strWOID & ": ib_workorder插入失败"

End If

strora = "insert into ib_wohistory (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
   " CUstOMER ,SALESORDER,CUstOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
   "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PARA10,PROTECTIVE_FILM_APLD ,LOT_stATUS ,MPN)" & _
   " Values (" & stHeader.id & ",'" & stHeader.ORDERNAME & "','" & stHeader.ORDERTYPE & "' ,'CREATED','" & stHeader.ERPUSER & "','" & stHeader.product & "'," & stHeader.QTY & ",sysdate,to_date('" & stHeader.PLANSTARTDATE & "','yyyy-mm-dd'),to_date('" & stHeader.PLANENDDATE & "','yyyy-mm-dd')," & _
   " '" & stHeader.CUSTOMER & "','" & stHeader.SALESORDER & "','" & stHeader.CustomerERPN & "','" & stHeader.FABFACILITY & "','" & stHeader.IMAGERREV & "','" & stHeader.DESIGNID & "','" & stHeader.MLEVEL235 & "','" & stHeader.MLEVEL260 & "','" & stHeader.NGFLAG & "','" & stHeader.PARA1 & "'," & _
   "  '" & stHeader.PARA2 & "','" & stHeader.PARA3 & "','" & stHeader.PARA4 & "','" & stHeader.PARA5 & "','" & stHeader.PARA6 & "','" & stHeader.RequestDate & "','" & stHeader.PARA8 & "','" & stHeader.PARA10 & "','" & stHeader.PROTECTIVE_FILM_APLD & "','" & stHeader.Lot_Stauts & "'," & _
   " '" & stHeader.MPN & "')"
If AddSql(strora) = 0 Then
    AddLog strWOID & ": ib_wohistory插入失败"

End If

strSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
   " CUstOMER ,SALESORDER,CUstOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
   "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PARA10,PROTECTIVE_FILM_APLD ,LOT_stATUS ,MPN)" & _
   " Values (" & stHeader.id & ",'" & stHeader.ORDERNAME & "','" & stHeader.ORDERTYPE & "' ,'CREATED','" & stHeader.ERPUSER & "','" & stHeader.product & "'," & stHeader.QTY & ",GetDate(),convert(datetime,'" & stHeader.PLANSTARTDATE & "'),convert(datetime,'" & stHeader.PLANENDDATE & "')," & _
   " '" & stHeader.CUSTOMER & "','" & stHeader.SALESORDER & "','" & stHeader.CustomerERPN & "','" & stHeader.FABFACILITY & "','" & stHeader.IMAGERREV & "','" & stHeader.DESIGNID & "','" & stHeader.MLEVEL235 & "','" & stHeader.MLEVEL260 & "','" & stHeader.NGFLAG & "','" & stHeader.PARA1 & "'," & _
   "  '" & stHeader.PARA2 & "','" & stHeader.PARA3 & "','" & stHeader.PARA4 & "','" & stHeader.PARA5 & "','" & stHeader.PARA6 & "','" & stHeader.RequestDate & "','" & strOrderDept & "', '" & stHeader.PARA10 & "','" & stHeader.PROTECTIVE_FILM_APLD & "','" & stHeader.Lot_Stauts & "'," & _
   " '" & stHeader.MPN & "')"
If AddSql2(strSql) = 0 Then
    AddLog strWOID & ": tblTSVworkorder插入失败"

End If

End Sub

Private Sub AddWODetail(stDetail As BillDetail)
Dim strora As String
Dim strSql As String

strora = "insert into ib_waferlist(ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & stDetail.ORDERNAME & "'," & " '" & stDetail.waferid & "'," & stDetail.DIEQTY & "," & stDetail.FGDIEQTY & ",'" & stDetail.WAFERLOT & "',100,'" & stDetail.MARKINGCODE & "')"
strSql = "insert into [erpdata].[dbo].[tblTSVwaferlist](ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & stDetail.ORDERNAME & "'," & " '" & stDetail.waferid & "'," & stDetail.DIEQTY & "," & stDetail.FGDIEQTY & ",'" & stDetail.WAFERLOT & "',100,'" & stDetail.MARKINGCODE & "')"
If AddSql(strora) = 0 Then
    AddLog strWOID & ": ib_waferlist插入失败"

End If

If AddSql2(strSql) = 0 Then
    AddLog strWOID & ": tblTSVwaferlist插入失败"

End If

Call addLogTxt(stDetail.ORDERNAME, " 插入表:ib_waferlist: " & stDetail.waferid)

End Sub

Private Sub SaveWOLot(strLotID As String)
Dim strCusCode As String
Dim strpo      As String

strCusCode = Replace(Trim(txtCusCode.text), "(ICI)", "")
strpo = Trim(txtPO.text)
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
Dim strSql   As String
Dim lID      As Long
Dim strDC    As String
Dim strJob   As String
Dim strSql2  As String
Dim strJob2  As String
Dim strJob3  As String

For i = 1 To ListView1.ListItems.count
    If ListView1.ListItems(i).Checked And ListView1.ListItems(i).SubItems(1) = strWOID Then
        stDetail.ORDERNAME = ListView1.ListItems(i).SubItems(1)
        stDetail.WAFERLOT = ListView1.ListItems(i).SubItems(2)
        stDetail.waferid = ListView1.ListItems(i).SubItems(3)
        stDetail.DIEQTY = ListView1.ListItems(i).SubItems(4)
        stDetail.FGDIEQTY = ListView1.ListItems(i).SubItems(5)
        stDetail.MARKINGCODE = ListView1.ListItems(i).SubItems(6)
        '37RETURNTOLINE
        If Frm_ProductionPlan.cbCusCode.text = "37" And Frm_ProductionPlan.cbWOType.text = "重工工单" And Frm_ProductionPlan.cb37Pri2.text = "Y" And InStr(stDetail.waferid, "+") > 0 And Frm_ProductionPlan.ComRe.text = "Y" Then
            strJob3 = Get_OracleStr(" select test_mtrl_desc from customeroitbl_test where to_char(id) in (select filename from mappingdatatest where substrateid = '" & stDetail.waferid & "' )")
            If Right$(strJob3, 1) <> "R" Then
                strSql = "update customeroitbl_test set test_mtrl_desc = test_mtrl_desc || 'R' where id in (select filename from mappingdatatest where substrateid = '" & stDetail.waferid & "' )"
                AddSql (strSql)
                strSql = "update [ERPBASE].[dbo].[tblCustomerOI] set test_mtrl_desc = test_mtrl_desc + 'R' where id in (select filename from  [ERPBASE].[dbo].[tblmappingData] where substrateid = '" & stDetail.waferid & "' )"
                AddSql2 (strSql)

            End If

        End If

        '保存37DC
        If Frm_ProductionPlan.cbCusCode.text = "37" Then
            strJob = Get_OracleStr("select t4.source_mtrl_sloc from mappingdatatest t3 inner join customeroitbl_test t4 on t4.SOURCE_BATCH_ID = t3.LOTID  and to_char(t4.ID) = t3.FILENAME where t3.substrateid = '" & stDetail.waferid & "'")
            strJob2 = Get_OracleStr("select t4.test_mtrl_desc from mappingdatatest t3 inner join customeroitbl_test t4 on t4.SOURCE_BATCH_ID = t3.LOTID  and to_char(t4.ID) = t3.FILENAME where t3.substrateid = '" & stDetail.waferid & "'")

            Do While Right$(strJob2, 1) = "R"
                strJob2 = Left(strJob2, Len(strJob2) - 1)
            Loop
            If strJob2 <> strJob Then
                strDC = Get_OracleStr("select DATE_CODE_CONVERT.DC_CONVERT(to_char(sysdate, 'YYYY-MM-DD'),1) from dual ")
                strSql2 = "select * from TBL37TESTDC where JOBID = '" & strJob2 & "' "
                If Get_OracleCnt(strSql2) = 0 Then
                    AddSql ("insert into TBL37TESTDC(JOBID,DC) values('" & strJob2 & "','" & strDC & "')")
                    AddSql2 ("insert into erptemp..TBL37TESTDC(JOBID,DC) values('" & strJob2 & "','" & strDC & "')")
                    AddSql ("insert into mes_reference values ('37JOB_DC','" & strJob2 & "','NULL','NULL','37JOB_DC','" & strDC & "',0,'" & gUserName & "',SYSDATE) ")

                End If

            End If

        End If

        Call AddWODetail(stDetail)
        If ListView1.ListItems(i).SubItems(2) <> strLotID Then
            strLotID = ListView1.ListItems(i).SubItems(2)
            Call SaveWOLot(strLotID)

        End If

        lWaferQty = lWaferQty + 1
        lDieQty = lDieQty + ListView1.ListItems(i).SubItems(4)

    End If

Next i

End Sub

Private Sub SaveWOBom(strWOID As String, lWaferQty As Long, lDieQty As Long)
Dim strSql As String
Dim strPN  As String

strPN = Trim$(TxtPN.text)
If Frm_ProductionPlan.cbWOType = "重工工单" Or Frm_ProductionPlan.cbWOType = "Dummy工单" Or Frm_ProductionPlan.cbWOType = "玻璃工单" Then
    strSql = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记)" & "select b.ORDERNAME ,a.物料编号,'1','主选材料',SUM(convert(int, c.DIEQTY)),'1' " & "from erpdata..tblSmainM2 a ,erpdata..tblTSVworkorder b, erpdata .. tblTSVwaferlist c " & "where b.ORDERNAME = '" & strWOID & "'  and  b.PRODUCT = a.料号 and c.ORDERNAME = b.ORDERNAME " & "group by b.ORDERNAME ,a.物料编号"
Else

    strSql = "INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) " & _
       " SELECT  DISTINCT xx.* FROM ( SELECT distinct  '" & strWOID & "' AS 工单 ,X.物料编号,'1' AS 序组 ,'主选材料' AS 材料," & _
       " CAST( (CAST(X.用量 AS DECIMAL(18,8)) * '" & lWaferQty & "' ) AS  DECIMAL(18,3)) AS 用量 ,1 AS 产线标记 " & _
       " FROM ( SELECT b.材料规范编号, b.物料编号,sum(b.每只用量) as 用量 " & _
       " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
       " WHERE a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & strPN & "'  AND SUBSTRING(b.料号,1,2)  NOT  IN ('18','19')  " & _
       " GROUP by b.材料规范编号, b.物料编号 )  X  " & " UNION SELECT distinct  '" & strWOID & "',y.物料编号,'1','主选材料',  y.用量  ,1 FROM ( SELECT b.材料规范编号, b.物料编号,SUM(CONVERT(INT,d.DIEQTY)) as 用量 " & _
       " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b ,erpdata .. tblTSVworkorder c, erpdata .. tblTSVwaferlist d  " & _
       " WHERE a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & strPN & "'  " & " AND SUBSTRING(b.料号,1,2)   IN ('18','19')  AND c.PRODUCT = a.物料编号 " & _
       " AND c.ORDERNAME = '" & strWOID & "' AND d.ORDERNAME = c.ORDERNAME " & " GROUP by b.材料规范编号, b.物料编号 ) y   ) xx  ,AIS20141114094336.dbo.t_ICItem yy, " & _
       " erptemp..BOM_DEVICE zz   WHERE yy.FNumber = xx.物料编号  AND( zz.product = yy.F_101 OR SUBSTRING(yy.F_101,1,2)   IN ('60','66','18','19','21') ) "

End If

If AddSql2(strSql) = 0 Then
    AddLog strWOID & ": Bom表插入失败"

End If

End Sub

Private Sub SaveAsExcel()
If Me.ListView1.ListItems.count = 0 Then Exit Sub

On Error Resume Next

Dim xlApp As Object, Wb As Object, i&, j&, T&, h&, ar()

T = ListView1.ListItems.count
h = ListView1.ColumnHeaders.count
If T = 0 Then Exit Sub
DoEvents
Set xlApp = CreateObject("Excel.Application")
If xlApp Is Nothing Then
    MsgBox "pls install Microsoft Excel."
    Exit Sub
Else
    ReDim ar(1 To T + 1, 1 To (h + 2))

    For i = 1 To h
        ar(1, i) = ListView1.ColumnHeaders(i)
    Next
    ar(1, h + 1) = "厂内机种"
    ar(1, h + 2) = "成品料号"

    For i = 2 To T + 1
        If ListView1.ListItems(i - 1).Checked And Get_OracleCnt("select * from shop_order where shop_order = '" & ListView1.ListItems(i - 1).SubItems(1) & "'") Then
            ar(i, 1) = ListView1.ListItems(i - 1)

            For j = 1 To h
                ar(i, j + 1) = ListView1.ListItems(i - 1).SubItems(j)
            Next
            ar(i, h + 1) = txtHTPN.text
            ar(i, h + 2) = TxtPN.text

        End If

    Next
    Set Wb = xlApp.Workbooks.Add
    Wb.ActiveSheet.Range("A:AA").NumberFormatLocal = "@"
    Wb.ActiveSheet.Range("a1").Resize(UBound(ar), h + 2) = ar
    Wb.ActiveSheet.Cells.Columns.AutoFit
    xlApp.Visible = True
    Set Wb = Nothing
    Set xlApp = Nothing

End If

End Sub

Private Sub AddLog(strLog As String)
txtLog = txtLog & time() & "->" & strLog & vbCrLf

End Sub

Private Function Check_MarkingcodeByHT(strHTPN As String, _
                                       strLotWaferID As String) As Boolean
Dim strSql         As String
Dim strKeyWord     As String
Dim i              As Integer
Dim keyChar1       As String
Dim keyChar2       As String
Dim strMarkingCode As String
Dim currentLen     As Integer
Dim realLen        As Integer
Check_MarkingcodeByHT = False

If Frm_ProductionPlan.cbWOType <> "Dummy工单" And Frm_ProductionPlan.cbWOType <> "FO_CSP工单" And Frm_ProductionPlan.cbWOType <> "硅基工单" And Frm_ProductionPlan.cbWOType <> "玻璃工单" And Frm_ProductionPlan.cbWOType <> "重工工单" Then
    strSql = "SELECT * FROM erpbase..tbltoinrec_wafer where 晶圆ID = '" & strLotWaferID & "' "

    If Get_SqlStr(strSql) = "" Then
        MsgBox strLotWaferID & "晶圆未入库，无法开立工单", vbCritical, "警告"
        Exit Function

    End If

    strMarkingCode = Get_OracleStr("select productid from mappingdatatest where substrateid = '" & strLotWaferID & "'")
    strKeyWord = Get_OracleStr("SELECT REMARK FROM TBL_MARKINGCODE_REP  WHERE HT_PN = '" & strHTPN & "'  and APPLY_FLAG = 'Y' ")

    If strKeyWord <> "" Then
        currentLen = Len(strMarkingCode)
        realLen = Len(strKeyWord)

        If currentLen <> realLen Then
            MsgBox "打标码长度错误,规定长度:" & realLen & vbCrLf & "当前长度:" & currentLen, vbCritical, "警告"
            Exit Function

        End If

        For i = 1 To Len(strKeyWord)
            keyChar1 = Mid$(strMarkingCode, i, 1)
            keyChar2 = Mid$(strKeyWord, i, 1)

            If keyChar2 <> "*" Then
                If keyChar1 <> keyChar2 Then
                    MsgBox strHTPN & "规定的第" & i & "位是字符:" & keyChar2 & vbCrLf & "当前Wafer:" & strLotWaferID & "打标码的第" & i & "位是字符:" & keyChar1 & vbCrLf & "打标码不符合规范", vbCritical, "警告"
                    Exit Function

                End If

            End If

        Next

    End If

End If

Check_MarkingcodeByHT = True

End Function
