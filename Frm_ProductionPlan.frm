VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_ProductionPlan 
   Caption         =   "工单开立"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8400
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
   MinButton       =   0   'False
   ScaleHeight     =   10263.16
   ScaleMode       =   0  'User
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ProductionPlan.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ProductionPlan.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ProductionPlan.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ProductionPlan.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ProductionPlan.frx":2368
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ProductionPlan.frx":2C42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   1535
      ButtonWidth     =   2408
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "       查找        "
            Key             =   "SEARCH"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "预览"
            Key             =   "PREVIEW"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刷新"
            Key             =   "INIT"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "EXIT"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CommandButton cmd_shop_order_del 
         BackColor       =   &H000080FF&
         Caption         =   "PMC工单删除"
         Height          =   360
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "LOT明细"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7575
      Left            =   4080
      TabIndex        =   3
      Top             =   960
      Width           =   4455
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "检索LOTID"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtSel 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "全选/反选"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   870
         ItemData        =   "Frm_ProductionPlan.frx":351C
         Left            =   120
         List            =   "Frm_ProductionPlan.frx":3523
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   6600
         Visible         =   0   'False
         Width           =   4335
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5055
         Left            =   360
         TabIndex        =   37
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   8916
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NO."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "LOT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "FAB-DEVICE"
            Object.Width           =   4410
         EndProperty
      End
   End
   Begin VB.Frame Fra 
      Caption         =   "工单选项"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   3975
      Begin VB.ComboBox ComRe 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Frm_ProductionPlan.frx":352C
         Left            =   2520
         List            =   "Frm_ProductionPlan.frx":3536
         TabIndex        =   39
         Top             =   5040
         Width           =   735
      End
      Begin VB.TextBox txtfab 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1080
         TabIndex        =   35
         Top             =   1450
         Width           =   2775
      End
      Begin VB.TextBox txtEP 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2880
         TabIndex        =   34
         Top             =   3203
         Width           =   975
      End
      Begin VB.TextBox txtNPIOwner 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   31
         Top             =   3945
         Width           =   1215
      End
      Begin VB.ComboBox cb37Pri2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Frm_ProductionPlan.frx":3540
         Left            =   2520
         List            =   "Frm_ProductionPlan.frx":354A
         TabIndex        =   30
         Top             =   4725
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "批量"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   3218
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.ComboBox cbWO 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1080
         TabIndex        =   26
         Top             =   3180
         Width           =   1095
      End
      Begin VB.TextBox cbHTPN 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2250
         Width           =   2775
      End
      Begin VB.ComboBox cbCusCode 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1080
         TabIndex        =   22
         Top             =   765
         Width           =   2775
      End
      Begin VB.TextBox txtCusPN 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1080
         TabIndex        =   20
         Top             =   1100
         Width           =   2775
      End
      Begin VB.ComboBox cbPN 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1860
         Width           =   2775
      End
      Begin VB.ComboBox cb37Pri 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1080
         TabIndex        =   18
         Top             =   4725
         Width           =   1455
      End
      Begin VB.ComboBox cbLotType 
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "Frm_ProductionPlan.frx":3554
         Left            =   1080
         List            =   "Frm_ProductionPlan.frx":3564
         TabIndex        =   17
         Top             =   5640
         Width           =   1455
      End
      Begin VB.TextBox txtWODept 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3525
         Width           =   2775
      End
      Begin VB.ComboBox cbWOType 
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "Frm_ProductionPlan.frx":3595
         Left            =   1080
         List            =   "Frm_ProductionPlan.frx":3597
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   420
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   21
         Top             =   6225
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
         Format          =   102432769
         CurrentDate     =   43271
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   23
         Top             =   6645
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
         CalendarForeColor=   16744576
         CalendarTitleBackColor=   16744703
         CalendarTitleForeColor=   8438015
         CalendarTrailingForeColor=   16777215
         Format          =   102432769
         CurrentDate     =   43271
      End
      Begin VB.Label lblre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "重测标记"
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
         TabIndex        =   40
         Top             =   5160
         Width           =   840
      End
      Begin VB.Label lblFAB_DEVICE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FAB机种"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   36
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblNPIOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NPI 负责"
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
         TabIndex        =   33
         Top             =   3960
         Width           =   795
      End
      Begin VB.Label lblNPIName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "负责人"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2520
         TabIndex        =   32
         Top             =   3990
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码     "
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
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   825
         Width           =   885
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单类型"
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
         Index           =   11
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单部门"
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
         Index           =   9
         Left            =   240
         TabIndex        =   12
         Top             =   3600
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单前缀"
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
         Index           =   8
         Left            =   240
         TabIndex        =   11
         Top             =   3240
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOT_TYPE"
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
         Index           =   7
         Left            =   240
         TabIndex        =   10
         Top             =   5640
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "37_PRI"
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
         Index           =   6
         Left            =   240
         TabIndex        =   9
         Top             =   4785
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "厂内机种     "
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
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   885
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "完工日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   6720
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开工日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   6315
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "成品料号     "
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
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   1920
         Width           =   885
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户机种     "
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
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   1185
         Width           =   885
      End
   End
End
Attribute VB_Name = "Frm_ProductionPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SetParent _
                Lib "user32.dll" (ByVal hWndChild As Long, _
                                  ByVal hWndNewParent As Long) As Long

Private Sub ClearData()
cbPn.Clear
cbHTPN.text = ""
cbWO.text = ""
txtWODept.text = ""
txtfab.text = ""
List1.Clear
ListView1.ListItems.Clear

End Sub

Private Sub cbHTPN_Change()
' 带出NPI负责人
Dim strNPIOwner As String
Dim strCustCode As String
Dim strCustPN   As String
Dim strHTPN     As String
Dim strSql      As String

strCustCode = Trim(cbCusCode.text)
strCustPN = Trim$(txtCusPN.text)
strHTPN = Trim$(cbHTPN.text)
strSql = "select residual from tbltsvnpiproduct where customerptno1 = '" & strCustPN & "' and qtechptno = '" & strHTPN & "' and customershortname = '" & strCustCode & "' and residual is not null"
strNPIOwner = Get_OracleStr(strSql)
txtNPIOwner.text = strNPIOwner

End Sub

Private Sub cbPn_Click()
' 带出厂内机种
Dim rs       As New ADODB.Recordset
Dim strCusPN As String
Dim strPN    As String
Dim strSql   As String

cbHTPN.text = ""
strCusPN = Trim(txtCusPN.text)
strPN = Trim$(cbPn.text)
Set rs.ActiveConnection = OraConnect
rs.Source = "select distinct qtechptno from tbltsvnpiproduct where customerptno1 = '" & strCusPN & "' and qtechptno2 = '" & strPN & "' "
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount > 0 Then
    If rs.RecordCount > 1 Then
        MsgBox "料号带出了多个厂内机种, 请NPI确认", vbCritical, "警告"
        Exit Sub

    End If

    rs.MoveFirst

    For i = 1 To rs.RecordCount
        cbHTPN.text = Trim(rs("qtechptno"))
        rs.MoveNext
    Next i

Else
    MsgBox "该机种:" & strCusPN & "查询不到厂内机种, 请NPI确认", vbCritical, "警告"
    Exit Sub

End If

rs.Close
Set rs = Nothing
' 37判断厂内机种和料号关系
If Trim(cbCusCode.text) = "37" And cbHTPN.text = "X37B" Then
    If Left(Right(strPN, 2), 1) <> "B" Then
        MsgBox "NPI维护错误, X37B对应料号倒数第二位必须是B, 请NPI确认", vbCritical, "错误"
        cbHTPN.text = ""
        Exit Sub

    End If

End If

' 带出工单部门
Dim sProductDept As String
Dim sProductCode As String

txtWODept.text = ""
sProductDept = GetWoDept(cbPn.text)
sProductCode = GetGWoDeptID(sProductDept)
If sProductDept <> "" And sProductCode <> "" Then
    txtWODept.text = sProductDept & "_" & sProductCode

End If

' 带出NPI负责人
Dim strNPIOwner As String
Dim strCustCode As String
Dim strCustPN   As String
Dim strHTPN     As String

strCustCode = Trim(cbCusCode.text)
strCustPN = Trim$(txtCusPN.text)
strHTPN = Trim$(cbHTPN.text)
strSql = "select residual from tbltsvnpiproduct where customerptno1 = '" & strCustPN & "' and qtechptno = '" & strHTPN & "' and customershortname = '" & strCustCode & "' and residual is not null"
strNPIOwner = Get_OracleStr(strSql)
If strNPIOwner = "" Then
    MsgBox "NPI未维护该料号:" & strHTPN & "的负责人,请联系NPI维护好", vbInformation, "提示"

End If

'带出工程属性
Dim strPE As String

strPE = Get_OracleStr("select p_e from tbltsvnpiproduct where qtechptno2 = '" & strPN & "'")
txtEP.text = IIf(strPE = "", "NPI未维护", strPE)

End Sub

Private Sub cbWO_Change()

Select Case Mid$(Trim(cbWO.text), 2, 1)

    Case "P", "T"
        cbLotType.ListIndex = 0

    Case "S"
        If Left(UCase(Trim(cbWO.text)), 3) = "BSM" Then
            cbLotType.ListIndex = 1
        Else
            cbLotType.ListIndex = 2

        End If

End Select

End Sub

Private Sub cbWO_Click()

Select Case Mid$(Trim(cbWO.text), 2, 1)

    Case "P", "T"
        cbLotType.text = "量产(M)"

    Case "S"
        If Left(UCase(Trim(cbWO.text)), 3) = "BSM" Then
            cbLotType.text = "工程DC(E)"
        Else
            cbLotType.text = "样品(Q)"

        End If

End Select

End Sub

Private Sub cbWOType_Click()

Select Case cbWOType.text

    Case "重工工单", "Dummy工单", "玻璃工单", "硅基工单"
        Unload Frm_ReWO
        Frm_ReWO.Show 1

End Select

End Sub

Private Sub Check1_Click()
Dim i    As Integer
Dim rs   As New ADODB.Recordset
Dim rs1  As New ADODB.Recordset
Dim cust As String

'If Check1.Value = 1 Then
'
'    With List1
'
'        For i = 0 To .ListCount - 1
'            .Selected(i) = True
'        Next
'
'    End With
'
'ElseIf Check1.Value = 0 Then
'
'    With List1
'
'        For i = 0 To .ListCount - 1
'            .Selected(i) = False
'        Next
'
'    End With
'
'End If
cust = "SELECT * FROM erptemp..CONFIG a WHERE a.CUSTOMER = '" & UCase(Trim(cbCusCode.text)) & "'  AND a.REMARK1 = 'Y'"
If rs1.State = adStateOpen Then rs1.Close
rs1.Open cust, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If Not rs1.EOF Then
    If Check1.Value = 1 Then

        With ListView1

            For i = 1 To .ListItems.count
                .ListItems(i).Checked = True
            Next

        End With

    ElseIf Check1.Value = 0 Then

        With ListView1

            For i = 1 To .ListItems.count
                .ListItems(i).Checked = False
            Next

        End With

    End If

    For i = 1 To ListView1.ListItems.count
        If ListView1.ListItems(i).Checked Then
            If UCase(Trim(txtfab.text)) = "" Then
                txtfab.text = ListView1.ListItems(i).SubItems(2)
            Else
                If txtfab.text <> ListView1.ListItems(i).SubItems(2) Then
                    MsgBox "请输入工单前三位!" & ListView1.ListItems(i).SubItems(2) & "的FAB_DEVICE和其他LOT不一致，请确认！"
                    tb1.Buttons("PREVIEW").Enabled = False
                    Exit Sub

                End If

            End If

        End If

    Next

End If

Set rs.ActiveConnection = OraConnect
If UCase(Trim(cbPn.text)) = "" Then
    If UCase(Trim(txtfab.text)) <> "" Then
        If cbWOType.text = "Dummy工单" Then
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & Trim(cbCusCode.text) & "' and customerptno1 = '" & Trim(txtCusPN.text) & "'  and  customerptno2 = '" & UCase(Trim(txtfab.text)) & "' and substr(qtechptno2, -3, 1) <> 'W' "
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & Trim(cbCusCode.text) & "'  and  customerptno1 =  '" & Trim(txtCusPN.text) & "' and  customerptno2 = '" & UCase(Trim(txtfab.text)) & "' and substr(qtechptno2, -3, 1) = 'W' "
        Else
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname ='" & Trim(cbCusCode.text) & "'  and customerptno1 = '" & Trim(txtCusPN.text) & "' and  customerptno2 = '" & UCase(Trim(txtfab.text)) & "' and substr(qtechptno2, -3, 1) <> 'W' "

        End If

    Else
        If cbWOType.text = "Dummy工单" Then
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname ='" & Trim(cbCusCode.text) & "' and customerptno1 =  '" & Trim(txtCusPN.text) & "' and substr(qtechptno2, -3, 1) <> 'W' "
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & Trim(cbCusCode.text) & "'  and  customerptno1 =  '" & Trim(txtCusPN.text) & "'and substr(qtechptno2, -3, 1) = 'W' "
        Else
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & Trim(cbCusCode.text) & "' and customerptno1 =  '" & Trim(txtCusPN.text) & "' and substr(qtechptno2, -3, 1) <> 'W' "

        End If

    End If

    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
    cbPn.Clear
    If rs.RecordCount > 0 Then
        If rs.RecordCount > 1 Then
            MsgBox "请注意,该客户机种包含多个成品料号, 请确认是否有误", vbInformation, "提示"

        End If

        rs.MoveFirst

        For i = 1 To rs.RecordCount
            cbPn.AddItem Trim(rs("qtechptno2"))
           ' cbPN.text = Trim(rs("qtechptno2"))
            If rs.RecordCount = 1 Then
                cbPn.text = Trim(rs("qtechptno2"))
            End If
            rs.MoveNext
        Next i


    Else
        MsgBox "该客户代码:" & UCase(Trim(cbCusCode.text)) & "机种:" & UCase(Trim(txtCusPN.text)) & ": NPI未维护对应关系, 查询不到料号", vbCritical, "警告"
        Exit Sub

    End If

    rs.Close
    Set rs = Nothing

End If

tb1.Buttons("SEARCH").Enabled = False
tb1.Buttons("PREVIEW").Enabled = True
If Check1.Value = 1 Then

    With ListView1

        For i = 1 To .ListItems.count
            .ListItems(i).Checked = True
        Next

    End With

ElseIf Check1.Value = 0 Then

    With ListView1

        For i = 1 To .ListItems.count
            .ListItems(i).Checked = False
        Next

    End With

End If

End Sub

Private Sub cmd_shop_order_del_Click()
frmPMC_delshop_order.Show

'Frm_ProductionPlan.Hide
End Sub

Private Sub Command1_Click()
Dim strKey As String
Dim rs     As New ADODB.Recordset
Dim rs1    As New ADODB.Recordset
Dim cust   As String

strKey = Trim$(txtSel)
If strKey = "" Then
    MsgBox "请输入LOT ID", vbInformation, "提示:"
    Exit Sub

End If

'With List1
'
'    For i = 0 To .ListCount - 1
'        If strKey = .List(i) Then
'            .Selected(i) = True
'
'        End If
'
'    Next
'
'End With
cust = "SELECT * FROM erptemp..CONFIG a WHERE a.CUSTOMER = '" & UCase(Trim(cbCusCode.text)) & "'  AND a.REMARK1 = 'Y'"
If rs1.State = adStateOpen Then rs1.Close
rs1.Open cust, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If Not rs1.EOF Then

    With ListView1

        For i = 1 To .ListItems.count
            If strKey = Trim$(.ListItems(i).SubItems(1)) Then
                .ListItems(i).Checked = True

            End If

        Next

    End With

    For i = 1 To ListView1.ListItems.count
        If ListView1.ListItems(i).Checked Then
            If UCase(Trim(txtfab.text)) = "" Then
                txtfab.text = ListView1.ListItems(i).SubItems(2)
            Else
                If txtfab.text <> ListView1.ListItems(i).SubItems(2) Then
                    MsgBox "请输入工单前三位!" & ListView1.ListItems(i).SubItems(2) & "的FAB_DEVICE和其他LOT不一致，请确认！"
                    tb1.Buttons("PREVIEW").Enabled = False
                    Exit Sub

                End If

            End If

        End If

    Next

End If

Set rs.ActiveConnection = OraConnect
If UCase(Trim(cbPn.text)) = "" Then
    If UCase(Trim(txtfab.text)) <> "" Then
        If cbWOType.text = "Dummy工单" Then
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & UCase(Trim(cbCusCode.text)) & "' and customerptno1 = '" & UCase(Trim(txtCusPN.text)) & "'  and  customerptno2 = '" & UCase(Trim(txtfab.text)) & "' and substr(qtechptno2, -3, 1) <> 'W' "
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & UCase(Trim(cbCusCode.text)) & "'  and  customerptno1 =  '" & UCase(Trim(txtCusPN.text)) & "' and  customerptno2 = '" & UCase(Trim(txtfab.text)) & "' and substr(qtechptno2, -3, 1) = 'W' "
        Else
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname ='" & UCase(Trim(cbCusCode.text)) & "'  and customerptno1 = '" & UCase(Trim(txtCusPN.text)) & "' and  customerptno2 = '" & UCase(Trim(txtfab.text)) & "' and substr(qtechptno2, -3, 1) <> 'W' "

        End If

    Else
        If cbWOType.text = "Dummy工单" Then
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname ='" & UCase(Trim(cbCusCode.text)) & "' and customerptno1 =  '" & UCase(Trim(txtCusPN.text)) & "' and substr(qtechptno2, -3, 1) <> 'W' "
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & UCase(Trim(cbCusCode.text)) & "'  and  customerptno1 =  '" & UCase(Trim(txtCusPN.text)) & "'and substr(qtechptno2, -3, 1) = 'W' "
        Else
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & UCase(Trim(cbCusCode.text)) & "' and customerptno1 =  '" & UCase(Trim(txtCusPN.text)) & "' and substr(qtechptno2, -3, 1) <> 'W' "

        End If

    End If

    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
    cbPn.Clear
    If rs.RecordCount > 0 Then
        If rs.RecordCount > 1 Then
            MsgBox "请注意,该客户机种包含多个成品料号, 请确认是否有误", vbInformation, "提示"

        End If

        rs.MoveFirst

        For i = 1 To rs.RecordCount
            cbPn.AddItem Trim(rs("qtechptno2"))
           ' cbPN.text = Trim(rs("qtechptno2"))
             If rs.RecordCount = 1 Then
                cbPn.text = Trim(rs("qtechptno2"))
            End If
            rs.MoveNext
        Next i

    Else
        MsgBox "该客户代码:" & UCase(Trim(cbCusCode.text)) & "机种:" & UCase(Trim(txtCusPN.text)) & ": NPI未维护对应关系, 查询不到料号", vbCritical, "警告"
        Exit Sub

    End If

    rs.Close
    Set rs = Nothing

End If

tb1.Buttons("SEARCH").Enabled = False
tb1.Buttons("PREVIEW").Enabled = True

With ListView1

    For i = 1 To .ListItems.count
        If strKey = Trim$(.ListItems(i).SubItems(1)) Then
            .ListItems(i).Checked = True

        End If

    Next

End With

End Sub

Private Sub Form_Load()
' 初始化
Init
If gUserName = "07885" Or gUserName = "17607" Then
    cmd_shop_order_del.Visible = True
Else
    cmd_shop_order_del.Visible = False

End If

End Sub

Private Sub Init()
' 时间
DTPicker1(1).Value = Format(Year(Now()) & "-" & Month(Now()) & "-" & "28", "yyyy-MM-dd")
DTPicker1(0).Value = Format(Now(), "yyyy-MM-dd")
' 客户代码
InitCustomerCode
' 37Pri
Init37Pri
' 批量工单前缀
InitLotWO
' 工单类型
InitWOType

End Sub

Private Sub InitCustomerCode()
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = SqlConnect
rs.Source = "SELECT 客户代码 as PID,客户代码 as productname FROM erpdata.dbo.tblXCustomer " & " union  select 'JX117' as PID,'JX117' as productname " & " union  select 'AA(ON)' as PID,'AA(ON)' as productname " & " union  select '37(ICI)' as PID,'37(ICI)' as productname " & " union  select 'AB18(2)' as PID,'AB18(2)' as productname " & " union  select 'BUMPINGDM' as PID,'BUMPINGDM' as productname " & " union select 'YZ22(2)' as PID,'YZ22(2)' as productname order by 客户代码 "
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
cbCusCode.Clear
If rs.RecordCount > 0 Then
    rs.MoveFirst

    For i = 1 To rs.RecordCount
        cbCusCode.AddItem Trim(rs("productname"))
        rs.MoveNext
    Next i

End If

rs.Close
Set rs = Nothing

End Sub

Private Sub Init37Pri()
cb37Pri.AddItem ("Hot Lot")
cb37Pri.AddItem ("Normal Lot")
cb37Pri.AddItem ("Super Hot Lot")
cb37Pri.text = "Normal Lot"
cb37Pri2.ListIndex = 0

End Sub

Private Sub InitLotWO()
Dim strSql As String
Dim rs     As New ADODB.Recordset

strSql = "select distinct a.NAME from TBLWORKORDERNAME a where a.flag = '1' order by a.name "
If rs.State = adStateOpen Then rs.Close
rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
cbWO.Clear
If Not rs.EOF Then

    Do While Not rs.EOF
        cbWO.AddItem UCase(Trim$("" & rs!name))
        rs.MoveNext
    Loop

End If

End Sub

Private Sub InitWOType()
cbWOType.AddItem ("一般工单")
cbWOType.AddItem ("小批量试产工单")
cbWOType.AddItem ("Dummy工单")
cbWOType.AddItem ("玻璃工单")
cbWOType.AddItem ("FO_CSP工单")
cbWOType.AddItem ("硅基工单")
cbWOType.AddItem ("样品工单")
cbWOType.AddItem ("重工工单")
cbWOType.ListIndex = 0

End Sub

Public Sub ForSearch()
Dim strCusCode As String
Dim strCusPN   As String
Dim strdevice  As String
Dim rs8        As New ADODB.Recordset
Dim strac70 As String
Dim rsac70 As New ADODB.Recordset



If strCusCode = "AC70" Then

strac70 = " select * from EU010_reference C  where c.cust_device = '" & strCusPN & "'   "


If rsac70.State = adStateOpen Then rsac70.Close
rsac70.Open strac70, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If rsac70.RecordCount < 1 Then
    MsgBox "该客户代码:" & strCusCode & "机种:" & strCusPN & ":客户PKG信息未维护 ", vbCritical, "警告"
   Exit Sub

End If
rsac70.Close

End If




ListView1.ListItems.Clear
If cbWOType.text = "" Then
    MsgBox "请选择工单类型", vbCritical, "警告"
    Exit Sub

End If

If cbWOType.text = "玻璃工单" Then
    If InStr(txtCusPN.text, "-CV") = 0 Then
        txtCusPN.text = txtCusPN.text & "-CV"

    End If

End If

If cbWOType.text = "硅基工单" Then
    If InStr(txtCusPN.text, "-FO") = 0 Then
        MsgBox "硅基工单的客户机种后缀必须为'-FO'", vbCritical, "警告"
        Exit Sub

    End If

End If

If cbWOType.text = "FO_CSP工单" Then
    If InStr(txtCusPN.text, "-FO") > 0 Then
        MsgBox "FO_CSP工单的客户机种后缀不可以包含'-FO'", vbCritical, "警告"
        Exit Sub

    End If

End If

strCusCode = cbCusCode.text
strCusPN = Trim$(txtCusPN.text)
If strCusCode = "" Then
    MsgBox "客户代码不可为空", vbCritical, "警告"
    Exit Sub

End If

If strCusPN = "" Then
    MsgBox "客户机种不可为空", vbCritical, "警告"
    Exit Sub

End If

strdevice = "select * from tbltsvnpiproduct a ,ib_wohistory b where a.customerptno1 = '" & strCusPN & "' and a.customershortname = '" & strCusCode & "' and b.product = a.qtechptno2 and TO_CHAR(B.ERPCREATEDATE,'YYYY-MM-DD') > to_char( sysdate -180,'YYYY-MM-DD')  "
If rs8.State = adStateOpen Then rs8.Close
rs8.Open strdevice, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If rs8.RecordCount < 1 Then
    MsgBox "该客户代码:" & strCusCode & "机种:" & strCusPN & ":半年内没开过工单 ", vbCritical, "警告"
    MsgBox "该客户代码:" & strCusCode & "机种:" & strCusPN & ":半年内没开过工单 ", vbCritical, "警告"

End If

rs8.Close
Call SearchByCPN(strCusCode, strCusPN)

End Sub

Private Sub SearchByCPN(strCusCode As String, strCusPN As String)
Dim rs  As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset

' 特殊校验
If (cbCusCode.text = "37" Or cbCusCode.text = "EU010" Or cbCusCode.text = "HK075") And cbWOType.text <> "重工工单" Then
    Set rs1.ActiveConnection = OraConnect
    rs1.Source = " select * from tbltsvnpiproduct a where a.customershortname in ( '37','EU010','HK075')  and   instr(a.struckstr1,'ASSY') >0  and a.customerptno1 = '" & Trim$(txtCusPN.text) & "'"
    rs1.Open , , adOpenStatic, adLockReadOnly, adCmdText
    If rs1.RecordCount > 0 Then
        rs1.Close
        Set rs1 = Nothing
        Set Rs2.ActiveConnection = OraConnect
        Rs2.Source = "select * from code37 d where d.device = '" & Trim$(txtCusPN.text) & "' "
        Rs2.Open , , adOpenStatic, adLockReadOnly, adCmdText
        If Rs2.RecordCount < 1 Then
            MsgBox "该客户代码:" & strCusCode & "机种:" & strCusPN & ": 没有维护阴极线", vbCritical, "警告"
            Exit Sub

        End If

    End If

End If

' 带出料号
'Set rs.ActiveConnection = OraConnect
'
'If UCase(Trim(txtfab.Text)) <> "" Then
'
'If cbWOType.Text = "Dummy工单" Then
'    rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & strCusCode & "' and customerptno1 = '" & strCusPN & "'  and  customerptno2 = '" & UCase(Trim(txtfab.Text)) & "' and substr(qtechptno2, -3, 1) <> 'W' "
'    rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & strCusCode & "' and  customerptno1 = '" & strCusPN & "' and  customerptno2 = '" & UCase(Trim(txtfab.Text)) & "' and substr(qtechptno2, -3, 1) = 'W' "
'Else
'    rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & strCusCode & "' and customerptno1 = '" & strCusPN & "' and  customerptno2 = '" & UCase(Trim(txtfab.Text)) & "' and substr(qtechptno2, -3, 1) <> 'W' "
'
'End If
'
'
'Else
'
'
'If cbWOType.Text = "Dummy工单" Then
'    rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & strCusCode & "' and customerptno1 = '" & strCusPN & "' and substr(qtechptno2, -3, 1) <> 'W' "
'    rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & strCusCode & "' and  customerptno1 = '" & strCusPN & "' and substr(qtechptno2, -3, 1) = 'W' "
'Else
'    rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & strCusCode & "' and customerptno1 = '" & strCusPN & "' and substr(qtechptno2, -3, 1) <> 'W' "
'
'End If
'
'End If
'
'rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
'cbPN.Clear
'If rs.RecordCount > 0 Then
'    If rs.RecordCount > 1 Then
'        MsgBox "请注意,该客户机种包含多个成品料号, 请确认是否有误", vbInformation, "提示"
'
'    End If
'
'    rs.MoveFirst
'
'    For i = 1 To rs.RecordCount
'        cbPN.AddItem Trim(rs("qtechptno2"))
'        cbPN.Text = Trim(rs("qtechptno2"))
'        rs.MoveNext
'    Next i
'
'Else
'    MsgBox "该客户代码:" & strCusCode & "机种:" & strCusPN & ": NPI未维护对应关系, 查询不到料号", vbCritical, "警告"
'    Exit Sub
'
'End If
'
'rs.Close
'Set rs = Nothing
' 查询此机种,带出LotID
If strCusCode = "AA" And cbWOType.text <> "Dummy工单" And cbWOType.text <> "玻璃工单" Then
    Call GetAALotID(rs, strCusCode, strCusPN)
    If rs.RecordCount = 0 Then
        rs.Close
        Call GetLotID(rs, strCusCode, strCusPN)

    End If

Else
    Call GetLotID(rs, strCusCode, strCusPN)

End If

List1.Clear
If rs.RecordCount > 0 Then
    rs.MoveFirst

    For i = 1 To rs.RecordCount
        List1.AddItem Trim(rs("source_batch_id"))
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
    Next i

Else
    MsgBox "该机种:" & strCusPN & "查询不到订单信息, 请确认;硅基,玻璃,dummy工单请手动维护数据", vbCritical, "警告"
    Exit Sub

End If

rs.Close
Set rs = Nothing

'tb1.Buttons("SEARCH").Enabled = False
'tb1.Buttons("PREVIEW").Enabled = True
End Sub

Private Sub GetLotID(ByRef rs As ADODB.Recordset, _
                     strCusCode As String, _
                     strCusPN As String)
Set rs.ActiveConnection = OraConnect
If UCase(Trim(txtfab.text)) <> "" Then
    If cbWOType.text = "重工工单" Or cbWOType.text = "委外工单" Then
        rs.Source = "select distinct a.source_batch_id ,A.fab_conv_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "'  AND A.fab_conv_id  = '" & UCase(Trim(txtfab.text)) & "' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0   and instr(b.substrateid, '+') > 0 and  not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"
    ElseIf cbWOType.text = "硅基工单" Then
        rs.Source = "select distinct a.source_batch_id ,A.fab_conv_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "'  AND  A.fab_conv_id  = '" & UCase(Trim(txtfab.text)) & "' and a.flag = 'T'  and instr(b.substrateid,'+') = 0 and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and a.source_batch_id like 'SI%' and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"
    ElseIf cbWOType.text = "玻璃工单" Then
        rs.Source = "select distinct a.source_batch_id ,A.fab_conv_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "'  AND  A.fab_conv_id  = '" & UCase(Trim(txtfab.text)) & "' and a.flag = 'T' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and a.source_batch_id like 'G%' and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"
    ElseIf cbWOType.text = "Dummy工单" Then
        rs.Source = "select distinct a.source_batch_id ,A.fab_conv_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "'  AND  A.fab_conv_id  = '" & UCase(Trim(txtfab.text)) & "' and a.flag = 'T' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and (a.source_batch_id like 'D%' or a.source_batch_id like 'SI%') and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"
    ElseIf cbWOType.text = "FO_CSP工单" Then
        rs.Source = "select distinct a.source_batch_id ,A.fab_conv_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "'  AND  A.fab_conv_id  = '" & UCase(Trim(txtfab.text)) & "' and a.flag = 'T' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and a.source_batch_id like 'SI%' and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"
    Else
        rs.Source = "select distinct a.source_batch_id ,A.fab_conv_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "'  AND  A.fab_conv_id  = '" & UCase(Trim(txtfab.text)) & "' and a.flag = 'Y' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"

    End If

Else
    If cbWOType.text = "重工工单" Or cbWOType.text = "委外工单" Then
        rs.Source = "select distinct a.source_batch_id ,A.fab_conv_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0   and instr(b.substrateid, '+') > 0 and  not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"
    ElseIf cbWOType.text = "硅基工单" Then
        rs.Source = "select distinct a.source_batch_id,A.fab_conv_id  from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "' and a.flag = 'T'  and instr(b.substrateid,'+') = 0 and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and a.source_batch_id like 'SI%' and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"
    ElseIf cbWOType.text = "玻璃工单" Then
        rs.Source = "select distinct a.source_batch_id ,A.fab_conv_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "' and a.flag = 'T' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and a.source_batch_id like 'G%' and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"
    ElseIf cbWOType.text = "Dummy工单" Then
        rs.Source = "select distinct a.source_batch_id ,A.fab_conv_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "' and a.flag = 'T' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and (a.source_batch_id like 'D%' or a.source_batch_id like 'SI%') and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"
    ElseIf cbWOType.text = "FO_CSP工单" Then
        rs.Source = "select distinct a.source_batch_id ,A.fab_conv_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "' and a.flag = 'T' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and a.source_batch_id like 'SI%' and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"
    Else
        rs.Source = "select distinct a.source_batch_id ,A.fab_conv_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "' and a.flag = 'Y' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"

    End If

End If

rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

End Sub

Private Sub GetAALotID(ByRef rs As ADODB.Recordset, _
                       strCusCode As String, _
                       strCusPN As String)
Dim customerPTTemp As String
Dim opnTemp        As String

opnTemp = strCusPN
customerPTTemp = GetONOPN_WSG(opnTemp)
Set rs.ActiveConnection = OraConnect
rs.Source = " select distinct b.batchid as source_batch_id , ''  as  fab_conv_id  from  ( select * from (select * from CUSTOMERFORECASTTBL order by ID desc) where   out_part_id = '" & customerPTTemp & "'  and rownum = 1 ) a ,CustomerBCtbl b " & "  where a.out_part_id='" & customerPTTemp & "' and a.comments='" & opnTemp & "' and a.flag='Y' and a.start_part_id=b.mtrlnum and b.batchid not in (select lotid from  On_WO_HisTory where flag='Y')  "
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

End Sub

Private Sub ForInit()
tb1.Buttons("SEARCH").Enabled = True
tb1.Buttons("PREVIEW").Enabled = False
ClearData

End Sub

Private Sub ForExit()
Unload Me
Unload Frm_ProductionPlanDetail

End Sub

Private Sub ForPreview()
Screen.MousePointer = 11
Unload Frm_ProductionPlanDetail
If CheckPowerInfo = True Then
    If List1.SelCount > 0 Or ListView1.selectedItem > 0 Then
        ' SetParent Frm_ProductionPlanDetail.hWnd, Me.hWnd
        Frm_ProductionPlanDetail.Show 1
    Else
        MsgBox "请选择LOT", vbCritical, "警告"
        Screen.MousePointer = 0
        Exit Sub

    End If

End If

Screen.MousePointer = 0

End Sub

Private Sub List1_Click()
Dim lot As String

End Sub

Private Sub ListView1_Click()
Dim rs   As New ADODB.Recordset
Dim rs1  As New ADODB.Recordset
Dim cust As String

cust = "SELECT * FROM erptemp..CONFIG a WHERE a.CUSTOMER = '" & UCase(Trim(cbCusCode.text)) & "'  AND a.REMARK1 = 'Y'"
If rs1.State = adStateOpen Then rs1.Close
rs1.Open cust, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If Not rs1.EOF Then

    For i = 1 To ListView1.ListItems.count
        If ListView1.ListItems(i).Checked Then
            If UCase(Trim(txtfab.text)) = "" Then
                txtfab.text = ListView1.ListItems(i).SubItems(2)
            Else
                If txtfab.text <> ListView1.ListItems(i).SubItems(2) Then
                    MsgBox "请输入工单前三位!" & ListView1.ListItems(i).SubItems(2) & "的FAB_DEVICE和其他LOT不一致，请确认！"
                    tb1.Buttons("PREVIEW").Enabled = False
                    Exit Sub

                End If

            End If

        End If

    Next

End If

Set rs.ActiveConnection = OraConnect
If UCase(Trim(cbPn.text)) = "" Then
    If UCase(Trim(txtfab.text)) <> "" Then
        If cbWOType.text = "Dummy工单" Then
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & Trim(cbCusCode.text) & "' and customerptno1 = '" & Trim(txtCusPN.text) & "'  and  customerptno2 = '" & UCase(Trim(txtfab.text)) & "' and substr(qtechptno2, -3, 1) <> 'W' "
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & Trim(cbCusCode.text) & "'  and  customerptno1 =  '" & Trim(txtCusPN.text) & "' and  customerptno2 = '" & UCase(Trim(txtfab.text)) & "' and substr(qtechptno2, -3, 1) = 'W' "
        Else
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname ='" & Trim(cbCusCode.text) & "'  and customerptno1 = '" & Trim(txtCusPN.text) & "' and  customerptno2 = '" & UCase(Trim(txtfab.text)) & "' and substr(qtechptno2, -3, 1) <> 'W' "

        End If

    Else
        If cbWOType.text = "Dummy工单" Then
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname ='" & Trim(cbCusCode.text) & "' and customerptno1 =  '" & Trim(txtCusPN.text) & "' and substr(qtechptno2, -3, 1) <> 'W' "
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & Trim(cbCusCode.text) & "'  and  customerptno1 =  '" & Trim(txtCusPN.text) & "'and substr(qtechptno2, -3, 1) = 'W' "
        Else
            rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & Trim(cbCusCode.text) & "' and customerptno1 =  '" & Trim(txtCusPN.text) & "' and substr(qtechptno2, -3, 1) <> 'W' "

        End If

    End If

    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
    cbPn.Clear
    If rs.RecordCount > 0 Then
        If rs.RecordCount > 1 Then
            MsgBox "请注意,该客户机种包含多个成品料号, 请确认是否有误", vbInformation, "提示"

        End If

        rs.MoveFirst

        For i = 1 To rs.RecordCount
            cbPn.AddItem Trim(rs("qtechptno2"))
           ' cbPN.text = Trim(rs("qtechptno2"))
            If rs.RecordCount = 1 Then
                cbPn.text = Trim(rs("qtechptno2"))
            End If
            rs.MoveNext
        Next i

    Else
        MsgBox "该客户代码:" & UCase(Trim(cbCusCode.text)) & "机种:" & UCase(Trim(txtCusPN.text)) & ": NPI未维护对应关系, 查询不到料号", vbCritical, "警告"
        Exit Sub

    End If

    rs.Close
    Set rs = Nothing

End If

tb1.Buttons("SEARCH").Enabled = False
tb1.Buttons("PREVIEW").Enabled = True

End Sub

Private Sub tb1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "SEARCH"
        ForSearch

    Case "INIT"
        ForInit

    Case "EXIT"
        ForExit

    Case "PREVIEW"
        ForPreview

End Select

End Sub

Public Function GetWOID() As String
Dim FirstChar    As String
Dim SeqChar      As String
Dim typenameTemp As String
Dim yMonthTemp   As String
Dim seqTemp      As Integer
Dim headChar     As String
Dim mdChar       As String
Dim id           As Long
Dim strWOID      As String

FirstChar = UCase(Trim(cbWO.text))
If Len(FirstChar) <> 3 Then
    MsgBox "请输入工单前三位!"
    cbWO.text = ""
    Exit Function

End If

headChar = FirstChar
SeqChar = GetWoIDTemp(FirstChar)
mdChar = Right(Year(DateTime.DATE), 2) & Right("0" & Month(DateTime.DATE), 2)
FirstChar = FirstChar & "-" & mdChar
SeqChar = Right("000" & CStr(CInt(SeqChar)), 4)
id = CLng(SeqChar)
strWOID = FirstChar & SeqChar
cmdStr = "insert into TSV_WO_SEQ_TAB(wotype,ymonth,sequenceid,flag) values ( '" & headChar & "','" & mdChar & "'," & id & ", 'Y' ) "
AddSql (cmdStr)
GetWOID = strWOID

End Function

Private Function CheckPowerInfo() As Boolean
Dim strdevice As String



CheckPowerInfo = False
If txtNPIOwner.text = "" Then
    MsgBox "工单对应NPI负责人工号不可以为空", vbCritical, "警告"
    Exit Function
Else
    If IsNumeric(Trim(txtNPIOwner.text)) = False Then
        MsgBox "请填写工单负责人工号", vbCritical, "警告"
        Exit Function

    End If

End If

If cbWO.text = "" Then
    MsgBox "工单前缀不可以为空", vbCritical, "警告"
    Exit Function
Else
    If Len(Trim$(cbWO.text)) <> 3 Then
        MsgBox "工单前缀必须是3位", vbCritical, "警告"
        Exit Function

    End If

End If

If txtWODept.text = "" Or txtWODept.text = "_" Then
    MsgBox "工单部门不可以为空", vbCritical, "警告"
    Exit Function

End If

If cbCusCode.text = "" Then
    MsgBox "客户代码不可以为空", vbCritical, "警告"
    Exit Function

End If

If txtCusPN.text = "" Then
    MsgBox "客户机种不可以为空", vbCritical, "警告"
    Exit Function

End If

If cbPn.text = "" Then
    MsgBox "成品料号不可以为空", vbCritical, "警告"
    Exit Function
Else
    If cbWOType.text <> "Dummy工单" And cbWOType.text <> "玻璃工单" And cbWOType.text <> "FO_CSP工单" And cbWOType.text <> "硅基工单" And cbWOType.text <> "重工工单" Then
        If CheckPN(Trim$(cbPn.text), Trim(txtWODept.text)) = False Then
            Exit Function

        End If

    End If

End If

If cbHTPN.text = "" Then
    MsgBox "厂内机种不可以为空", vbCritical, "警告"
    Exit Function

End If

If cb37Pri.text = "" Then
    MsgBox "37PRI不可以为空", vbCritical, "警告"
    Exit Function

End If

If cbLotType.text = "" Then
    MsgBox "LOT_TYPE不可以为空", vbCritical, "警告"
    Exit Function

End If

If cbWOType.text = "" Then
    MsgBox "工单类型不可以为空", vbCritical, "警告"
    Exit Function

End If

If DTPicker1(0).Value > DTPicker1(1).Value Then
    MsgBox "开工日期必须先于完工日期", vbCritical, "警告"
    Exit Function
ElseIf DTPicker1(0).Value = DTPicker1(1).Value Then
    MsgBox "开工日期不可以等于完工日期", vbCritical, "警告"
    Exit Function

End If

If cbWOType.text = "玻璃工单" Then
    If CheckBLWO(Trim(cbCusCode.text), Trim(txtCusPN.text), Trim(cbHTPN.text)) = False Then
        MsgBox "玻璃工单没有维护特定的信息(清洗步骤, CV高度, 清洗程序, 玻璃规格), 请联系NPI维护对应机种的信息", vbCritical, "提示"
        Exit Function

    End If

End If

If cbLotType.ListIndex = 1 Or cbLotType.ListIndex = 2 Then
    If txtNPIOwner.text = "" Then
        MsgBox "工程DC(E)工单或NPI样品(Q)工单必须有对应的NPI机种负责人" & vbCrLf & "请联系NPI维护对照表的机种负责人栏位,否则无法开立工单", vbInformation, "提示"
        Exit Function

    End If

End If







CheckPowerInfo = True

End Function

Private Function CheckBLWO(strCusCode, strCusPN, strHTPN) As Boolean
Dim strSql As String

CheckBLWO = False
strSql = "select * from tbltsvnpiproduct where customershortname = '" & strCusCode & "' and customerptno1 = '" & strCusPN & "' and qtechptno = '" & strHTPN & "' and  customerptno3 is not null and customerptno4 is not null and customerptno5 is not null and customerptno6 is not null"
If Get_OracleCnt(strSql) = 0 Then
    Exit Function

End If

CheckBLWO = True

End Function

Private Function CheckPN(strPN As String, strdept As String) As Boolean
CheckPN = False
Dim bomRS2 As New ADODB.Recordset

Set bomRS2 = GetProductBom(strPN)
If bomRS2.RecordCount <= 0 Then
    MsgBox "新系统中这料号的Bom不存在！请联系相关的人，先维护Bom ！"
    Exit Function

End If

'
'    Set bomRS2 = GetProductJDObject(strPN)
'
'    If bomRS2.RecordCount <= 0 Then
'        MsgBox "此料号在金碟系统中无成本对象，请找相关人员确认 ！"
'
'        Exit Function
'
'    End If
'
'    If InStr(UCase(strDept), "BUMPING") = 0 And InStr(UCase$(strDept), "SSP") = 0 And InStr(UCase(strDept), "WLP") = 0 Then
'        Set bomRS2 = GetProduct_Check(strPN)
'
'        If bomRS2.RecordCount <= 0 Then
'            MsgBox "料号不存在！请联系相关的人，先维护料号 ！"
'
'            Exit Function
'
'        End If
'
'    End If
Set bomRS2 = GetProductBomERpSign(strPN)
If bomRS2.RecordCount <= 0 Then
    MsgBox "新系统中这料号的Bom没有被审核通过，请联系工程部！"
    Exit Function

End If

CheckPN = True

End Function

Private Sub txtNPIOwner_Change()
Dim strSql As String

strSql = "select EmpName from XTW..employee where empno = '" & Trim$(txtNPIOwner.text) & "'"
lblNPIName.Caption = Get_SqlStr2(strSql)

End Sub
