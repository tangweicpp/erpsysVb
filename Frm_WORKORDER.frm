VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_WORKORDER 
   Caption         =   "生产工单维护2.0"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   450
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10545
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "工单明细"
      ForeColor       =   &H00800000&
      Height          =   11415
      Left            =   4080
      TabIndex        =   4
      Top             =   840
      Width           =   13935
      Begin FPSpreadADO.fpSpread fpSDetail 
         Height          =   10575
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   13575
         _Version        =   524288
         _ExtentX        =   23945
         _ExtentY        =   18653
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxRows         =   0
         SpreadDesigner  =   "Frm_WORKORDER.frx":0000
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "工单选项"
      ForeColor       =   &H00800000&
      Height          =   11415
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   3975
      Begin VB.TextBox txtWOCreater 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "检索"
         Height          =   285
         Left            =   3000
         TabIndex        =   33
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox txtLotID 
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Top             =   5400
         Width           =   1695
      End
      Begin VB.CheckBox chkLotSelect 
         Caption         =   "全选/反选"
         Height          =   255
         Left            =   1200
         TabIndex        =   31
         Top             =   5160
         Width           =   1335
      End
      Begin VB.ListBox lstLotID 
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5190
         Left            =   1200
         Style           =   1  'Checkbox
         TabIndex        =   29
         Top             =   5760
         Width           =   2655
      End
      Begin VB.TextBox txtNPIOwner 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         TabIndex        =   27
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ComboBox cb37Pri 
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
         Index           =   1
         ItemData        =   "Frm_WORKORDER.frx":0410
         Left            =   2880
         List            =   "Frm_WORKORDER.frx":041A
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3795
         Width           =   975
      End
      Begin VB.ComboBox cb37Pri 
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
         Index           =   0
         ItemData        =   "Frm_WORKORDER.frx":0424
         Left            =   1200
         List            =   "Frm_WORKORDER.frx":0431
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3795
         Width           =   1695
      End
      Begin VB.TextBox txtWODept 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   3105
         Width           =   2655
      End
      Begin VB.CheckBox chkLots 
         Caption         =   "批量工单"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   3480
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.ComboBox cbWOName 
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
         Left            =   1200
         TabIndex        =   15
         Top             =   3435
         Width           =   1215
      End
      Begin VB.ComboBox cbProduct 
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
         Left            =   1200
         TabIndex        =   13
         Top             =   1995
         Width           =   2655
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
         Left            =   1200
         TabIndex        =   11
         Top             =   1635
         Width           =   2655
      End
      Begin VB.ComboBox cbWOType 
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
         Index           =   1
         ItemData        =   "Frm_WORKORDER.frx":0459
         Left            =   1200
         List            =   "Frm_WORKORDER.frx":0469
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   2655
      End
      Begin VB.ComboBox cbWOType 
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
         Index           =   0
         ItemData        =   "Frm_WORKORDER.frx":049A
         Left            =   1200
         List            =   "Frm_WORKORDER.frx":04B0
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2655
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
         Left            =   1200
         TabIndex        =   6
         Top             =   1275
         Width           =   2655
      End
      Begin VB.ComboBox cbCustCode 
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
         Left            =   1200
         TabIndex        =   3
         Top             =   915
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dTBegin 
         Height          =   375
         Left            =   1200
         TabIndex        =   24
         Top             =   4260
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   247791617
         CurrentDate     =   43271
      End
      Begin MSComCtl2.DTPicker dTEnd 
         Height          =   375
         Left            =   1200
         TabIndex        =   25
         Top             =   4680
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   247791617
         CurrentDate     =   43271
      End
      Begin VB.Label lblCreater 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开立人员"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2640
         TabIndex        =   38
         Top             =   2760
         Width           =   780
      End
      Begin VB.Label lblWOCreater 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单开立"
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
         TabIndex        =   36
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label lblWOType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单用途"
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
         Index           =   1
         Left            =   240
         TabIndex        =   34
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblLotID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户批号"
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
         TabIndex        =   30
         Top             =   5400
         Width           =   900
      End
      Begin VB.Label lblNPIName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "负责人"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2640
         TabIndex        =   28
         Top             =   2445
         Width           =   540
      End
      Begin VB.Label lblNPIOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NPI负责人"
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
         Top             =   2400
         Width           =   960
      End
      Begin VB.Label lblWOEndDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预计完工"
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
         TabIndex        =   23
         Top             =   4800
         Width           =   900
      End
      Begin VB.Label lblWOBeginDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预计开工"
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
         TabIndex        =   22
         Top             =   4320
         Width           =   900
      End
      Begin VB.Label lbl37PRI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "37_PRI"
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
         TabIndex        =   19
         Top             =   3840
         Width           =   600
      End
      Begin VB.Label lblWODept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单部门"
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
         TabIndex        =   17
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label lblWOName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单前缀"
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
         TabIndex        =   14
         Top             =   3480
         Width           =   900
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
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label lblHTPN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "厂内机种"
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
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label lblWOType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单类型"
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
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   285
         Width           =   900
      End
      Begin VB.Label lblCustPN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户机种"
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
         TabIndex        =   5
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label lblCustCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码"
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
         TabIndex        =   2
         Top             =   960
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   1535
      ButtonWidth     =   1455
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   " 打印 "
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "输出"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "  新  增 "
            Key             =   "INSERT"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查询订单"
            Key             =   "READ"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "生成工单"
            Key             =   "CREATE"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改工单"
            Key             =   "UPDATE"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "删除工单"
            Key             =   "DELETE"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "导出工单"
            Key             =   "EXPORT"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "  生 成"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "A004"
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "待审核"
            Key             =   "WAIT_PASS"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "  审  核"
            Key             =   "PASS"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "反审核"
            Key             =   "CANCEL_PASS"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  退  出"
            Key             =   "EXIT"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8400
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
               Picture         =   "Frm_WORKORDER.frx":04F3
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WORKORDER.frx":262D
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WORKORDER.frx":54B7
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WORKORDER.frx":7C69
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WORKORDER.frx":9DA3
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WORKORDER.frx":C555
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WORKORDER.frx":ED07
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WORKORDER.frx":11D89
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WORKORDER.frx":1453B
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WORKORDER.frx":14855
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WORKORDER.frx":1552F
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WORKORDER.frx":185B1
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WORKORDER.frx":1AD63
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Frm_WORKORDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : Frm_WORKORDER
'    Project    : 正式工程1
'
'    Description: PMC工单开立及维护
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Private Enum E_WO_DETAIL

    E_CHOOSE = 1
    E_LOTID
    E_WAFERNO
    E_WAFERID
    E_GROSSDIES
    E_GOODDIES
    E_NGDIES
    E_MARKINGCODE
    E_END

End Enum

Private Type T_WO_HEADER

    SEQ_IBWO As String
    ORDERNAME As String
    ORDERTYPE As String
    DESCRIPTION As String
    EVENTTYPE As String
    ERPUSER As String
    product As String
    PRODUCTREVISION As String
    QTY As Long
    PRODUCTBOM As String
    ERPCREATEDATE As String
    PLANSTARTDATE As String
    PLANENDDATE As String
    CUSTOMER As String
    SALESORDER As String
    PRODUCTFAMILY As String
    MODIFYFLAG As String
    CUSTOMERPN As String
    FABFACILITY As String
    IMAGERREV As String
    DESIGNID As String
    MLEVEL235 As String
    MLEVEL260 As String
    NGFLAG As String
    PARA1 As String
    PARA2 As String
    PARA3 As String
    PARA4 As String
    PARA5 As String
    PARA6 As String
    PARA7 As String
    PARA8 As String
    PARA9 As String
    PARA10 As String
    PROTECTIVE_FILM_APLD As String
    LOT_STATUS As String
    MPN As String

End Type

Private Type T_WO_DETAIL

    ORDERNAME As String
    waferid As String
    COMPLETEFLAG As String
    DIEQTY As Long
    FGDIEQTY As Long
    WAFERLOT As String
    WAFERSEQUENCE As String
    MARKINGCODE As String

End Type

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       Form_Load
' Description:       窗体加载
' Created by :       汤威
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-9:39:22
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Load()
InitCtrls
InitData

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       initCtrls
' Description:       初始化控件状态
' Created by :       汤威
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-9:40:14
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitCtrls()
InitFps
InitDT
InitCB_WOType
InitCB_WOName
initCB_CustCode
InitCB_37PRI
InitWOCreater

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       InitFps
' Description:       初始化FPS
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/11-16:40:25
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitFps()

'工单明细
With fpSDetail
    .ReDraw = False
    .MaxCols = E_WO_DETAIL.E_END - 1
    .MaxRows = 0
    .FontBold = False
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Row = 0
    .TypeVAlign = TypeVAlignCenter
    .TypeHAlign = TypeHAlignLeft
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .TypeVAlign = TypeVAlignCenter
    '.TypeHAlign = TypeVAlignCenter
    .TypeHAlign = TypeHAlignLeft
    .SelForeColor = &HFF8080
    .Col = E_WO.E_CHOOSE
    .CellType = CellTypeCheckBox
    .Lock = False
    .SetText 0, 0, "序号"
    .ColWidth(0) = 4
    .SetText E_WO_DETAIL.E_CHOOSE, 0, "√"
    .SetText E_WO_DETAIL.E_LOTID, 0, "客批号"
    .SetText E_WO_DETAIL.E_WAFERNO, 0, "晶圆序号"
    .SetText E_WO_DETAIL.E_WAFERID, 0, "晶圆ID"
    .SetText E_WO_DETAIL.E_GROSSDIES, 0, "总DIES"
    .SetText E_WO_DETAIL.E_GOODDIES, 0, "良品DIES"
    .SetText E_WO_DETAIL.E_NGDIES, 0, "不良品DIES"
    .SetText E_WO_DETAIL.E_MARKINGCODE, 0, "打标码"
    .ColWidth(E_WO_DETAIL.E_CHOOSE) = 4
    .ColWidth(E_WO_DETAIL.E_LOTID) = 12
    .ColWidth(E_WO_DETAIL.E_WAFERNO) = 10
    .ColWidth(E_WO_DETAIL.E_WAFERID) = 16
    .ColWidth(E_WO_DETAIL.E_GROSSDIES) = 8
    .ColWidth(E_WO_DETAIL.E_GOODDIES) = 8
    .ColWidth(E_WO_DETAIL.E_NGDIES) = 8
    .ColWidth(E_WO_DETAIL.E_MARKINGCODE) = 20
    .ReDraw = True

End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       initDT
' Description:       初始化日期
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-11:37:22
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitDT()
dTBegin.Value = Format(Now() + 1, "yyyy-MM-dd")
dTEnd.Value = Format(Now() + 15, "yyyy-MM-dd")

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       initCB_WOType
' Description:       初始化工单类型列表
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-10:12:06
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitCB_WOType()
cbWOType(0).ListIndex = 0
cbWOType(1).ListIndex = 0

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       initCB_CustCode
' Description:       初始化客户代码列表
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-9:44:13
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub initCB_CustCode()
Dim rs     As New ADODB.Recordset
Dim strSql As String

strSql = "select distinct 客户代码 from erpdata..tblxcustomer where 客户代码 is not null"
Set rs = Get_SqlserveRs(strSql)
cbCustCode.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbCustCode.AddItem Trim("" & rs!客户代码)
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       initCB_WOName
' Description:       初始化工单前缀列表
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-11:05:01
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitCB_WOName()
Dim rs     As New ADODB.Recordset
Dim strSql As String

strSql = "select distinct substr(trim(ordername),1,3) as prefix from ib_wohistory where ordername is not null order by prefix"
Set rs = Get_OracleRs(strSql)
cbWOName.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbWOName.AddItem Trim("" & rs!prefix)
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       initCB_37PRI
' Description:       初始化37PRI列表
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-11:28:41
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitCB_37PRI()
cb37Pri(0).ListIndex = 1
cb37Pri(1).ListIndex = 1

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       InitWOCreater
' Description:       初始化工单开立人员
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/7-17:09:08
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitWOCreater()
Dim strSql As String

txtWOCreater.text = gUserName
strSql = "select EmpName from XTW..employee where empno = '" & Trim$(txtWOCreater.text) & "'"
lblCreater.Caption = Get_SqlStr2(strSql)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       initData
' Description:       初始化数据
' Created by :       汤威
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-9:40:23
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitData()

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cbCustCode_LostFocus
' Description:       客户代码转大写
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-10:27:48
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cbCustCode_LostFocus()
cbCustCode.text = UCase(cbCustCode.text)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cbCustCode_Change
' Description:       客户代码改变带出客户机种/厂内机种列表,清空lstLotID
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-10:22:50
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cbCustCode_Change()
Dim rs          As New ADODB.Recordset
Dim strSql      As String
Dim strCustCode As String

lstLotID.Clear
strCustCode = UCase(Trim$(cbCustCode.text))
strSql = "select distinct customerptno1  from tbltsvnpiproduct where customershortname = '" & strCustCode & "' and customerptno1 is not null"
Set rs = Get_OracleRs(strSql)
cbCustPN.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbCustPN.AddItem Trim("" & rs!CustomerPTNo1)
        rs.MoveNext
    Loop

End If

strSql = "select distinct qtechptno  from tbltsvnpiproduct where customershortname = '" & strCustCode & "' and qtechptno is not null"
Set rs = Get_OracleRs(strSql)
cbHTPN.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbHTPN.AddItem Trim("" & rs!qtechPTNo)
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cbCustCode_DropDown
' Description:       客户代码点击带出客户机种/厂内机种列表,清空lstLotID
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-10:23:42
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cbCustCode_Click()
Dim rs          As New ADODB.Recordset
Dim strSql      As String
Dim strCustCode As String

lstLotID.Clear
strCustCode = UCase(Trim$(cbCustCode.text))
strSql = "select distinct customerptno1  from tbltsvnpiproduct where customershortname = '" & strCustCode & "' and customerptno1 is not null"
Set rs = Get_OracleRs(strSql)
cbCustPN.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbCustPN.AddItem Trim("" & rs!CustomerPTNo1)
        rs.MoveNext
    Loop

End If

strSql = "select distinct qtechptno  from tbltsvnpiproduct where customershortname = '" & strCustCode & "' and qtechptno is not null"
Set rs = Get_OracleRs(strSql)
cbHTPN.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbHTPN.AddItem Trim("" & rs!qtechPTNo)
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cbHTPN_Change
' Description:       厂内机种变更带出唯一的客户机种/成品料号
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-10:45:50
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cbHTPN_Change()
Dim strSql  As String
Dim strHTPN As String
Dim rs      As New ADODB.Recordset

strHTPN = UCase(Trim$(cbHTPN.text))
strSql = "select distinct customerptno1  from tbltsvnpiproduct where qtechptno = '" & strHTPN & "' and customerptno1 is not null"
cbCustPN.text = Get_OracleStr(strSql)
strSql = "select distinct qtechptno2  from tbltsvnpiproduct where qtechptno = '" & strHTPN & "' and qtechptno2 is not null"
Set rs = Get_OracleRs(strSql)
cbProduct.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbProduct.AddItem Trim("" & rs!QtechPTNo2)
        cbProduct.text = Trim("" & rs!QtechPTNo2)
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cbHTPN_DropDown
' Description:       厂内机种变更带出唯一的客户机种/成品料号
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-10:52:05
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cbHTPN_Click()
Dim strSql  As String
Dim strHTPN As String
Dim rs      As New ADODB.Recordset

strHTPN = UCase(Trim$(cbHTPN.text))
strSql = "select distinct customerptno1  from tbltsvnpiproduct where qtechptno = '" & strHTPN & "' and customerptno1 is not null"
cbCustPN.text = Get_OracleStr(strSql)
strSql = "select distinct qtechptno2  from tbltsvnpiproduct where qtechptno = '" & strHTPN & "' and qtechptno2 is not null"
Set rs = Get_OracleRs(strSql)
cbProduct.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbProduct.AddItem Trim("" & rs!QtechPTNo2)
        cbProduct.text = Trim("" & rs!QtechPTNo2)
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cbCustPN_Change
' Description:       厂内机种变更带出客户机种变更,清空lstLotID
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-26-10:23:16
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cbCustPN_Change()
lstLotID.Clear

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cbCustPN_Click
' Description:       客户机种变更,清空lstLotID
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-26-10:24:50
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cbCustPN_Click()
lstLotID.Clear

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cbProductNO_Change
' Description:       成品料号变更带出工单部门/NPI负责人
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-11:17:14
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cbProduct_Change()
Dim strSql       As String
Dim strProductNO As String
Dim strPart1     As String
Dim strPart2     As String

' 工单部门
strProductNO = Trim(cbProduct.text)
strSql = "select Get_Product_Dept('" & strProductNO & "') qtechpt from dual"
strPart1 = Get_OracleStr(strSql)
strSql = "select FNumber from AIS20141114094336.dbo.t_Department where FName='" & strPart1 & "' "
strPart2 = Get_SqlStr(strSql)
txtWODept.text = strPart1 & strPart2
' NPI负责人
strSql = "select residual from tbltsvnpiproduct where qtechptno2 = '" & strProductNO & "'"
txtNPIOwner.text = Get_OracleStr(strSql)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cbProductNO_Click
' Description:       成品料号变更带出工单部门/NPI负责人
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-11:21:25
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cbProduct_Click()
Dim strSql       As String
Dim strProductNO As String
Dim strPart1     As String
Dim strPart2     As String

' 工单部门
strProductNO = Trim(cbProduct.text)
strSql = "select Get_Product_Dept('" & strProductNO & "') qtechpt from dual"
strPart1 = Get_OracleStr(strSql)
strSql = "select FNumber from AIS20141114094336.dbo.t_Department where FName='" & strPart1 & "' "
strPart2 = Get_SqlStr(strSql)
txtWODept.text = strPart1 & strPart2
' NPI负责人
strSql = "select residual from tbltsvnpiproduct where qtechptno2 = '" & strProductNO & "'"
txtNPIOwner.text = Get_OracleStr(strSql)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cbWOType_Change
' Description:       NPI负责人-工程批(E)+客户实验(Q)
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-26-9:58:29
'
' Parameters :       Index (Integer)
'--------------------------------------------------------------------------------
Private Sub cbWOType_Change(Index As Integer)

Select Case cbWOType(1).ListIndex

    Case 1, 2
        lblNPIOwner.Visible = True
        txtNPIOwner.Visible = True
        lblNPIName.Visible = True

    Case Else
        lblNPIOwner.Visible = False
        txtNPIOwner.Visible = False
        lblNPIName.Visible = False

End Select

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cbWOType_Click
' Description:       PI负责人-工程批(E)+客户实验(Q)
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-26-10:00:15
'
' Parameters :       Index (Integer)
'--------------------------------------------------------------------------------
Private Sub cbWOType_Click(Index As Integer)

Select Case cbWOType(1).ListIndex

    Case 1, 2
        lblNPIOwner.Visible = True
        txtNPIOwner.Visible = True
        lblNPIName.Visible = True

    Case Else
        lblNPIOwner.Visible = False
        txtNPIOwner.Visible = False
        lblNPIName.Visible = False

End Select

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       txtNPIOwner_Change
' Description:       负责人工号带出姓名
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-11:45:13
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub txtNPIOwner_Change()
Dim strSql As String

strSql = "select EmpName from XTW..employee where empno = '" & Trim$(txtNPIOwner.text) & "'"
lblNPIName.Caption = Get_SqlStr2(strSql)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cbWOName_Change
' Description:       工单号带出工单用途
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-26-10:08:05
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cbWOName_Change()

Select Case Mid$(Trim(cbWOName.text), 2, 1)

    Case "P", "T"
        cbWOType(1).ListIndex = 0

    Case "S"
        If Left(UCase(Trim(cbWOName.text)), 3) = "BSM" Then
            cbWOType(1).ListIndex = 1
        Else
            cbWOType(1).ListIndex = 2

        End If

End Select

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       chkLotSelect_Click
' Description:       LOTID全选/反选
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-11:54:28
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub chkLotSelect_Click()
Dim i As Integer

If chkLotSelect.Value = 1 Then

    With lstLotID

        For i = 0 To .ListCount - 1
            .Selected(i) = True
        Next

    End With

ElseIf chkLotSelect.Value = 0 Then

    With lstLotID

        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next

    End With

End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cmdQuery_Click
' Description:       检索LOTID
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-12:00:47
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdquery_Click()
Dim strKey As String
Dim i      As Integer
Dim bRet   As Boolean

bRet = False
strKey = Trim$(txtLotID.text)
If strKey = "" Then
    MsgBox "请输入LOT ID", vbInformation, "提示:"
    Exit Sub

End If

With lstLotID

    For i = 0 To .ListCount - 1
        If strKey = .List(i) Then
            .Selected(i) = True
            bRet = True

        End If

    Next

End With

If bRet = False Then
    MsgBox "查询不到该LOTID", vbInformation, "提示"

End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "READ"
        Call ReadSalesOrder

    Case "CREATE"
        Call CreateWorkOrder

    Case "UPDATE"
        Call UpdateWorkOrder

    Case "DELETE"
        Call DeleteWorkOrder

    Case "EXPORT"
        Call ExportWOData

    Case "EXIT"
        Unload Me

End Select

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ReadData
' Description:       读取订单
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-25-9:52:37
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub ReadSalesOrder()
Dim strCustCode As String
Dim strCustPN   As String

strCustCode = Trim$(cbCustCode.text)
strCustPN = Trim$(cbCustPN.text)
If strCustCode = "" Then
    MsgBox "请输入客户代码", vbInformation, "提示"
    Exit Sub

End If

If strCustPN = "" Then
    MsgBox "请输入客户机种", vbInformation, "提示"
    Exit Sub

End If

Call ShowLotList(strCustCode, strCustPN)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ShowLotList
' Description:       查询LOTID
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-25-10:06:22
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub ShowLotList(strCustCode As String, strCustPN As String)
Dim strSql             As String
Dim rs                 As New ADODB.Recordset
Dim strSqlPart_Flag    As String
Dim strSqlPart_WaferID As String
Dim strSqlPart_LotID   As String
Dim strSqlPart_OrderBy As String

fpSDetail.MaxRows = 0
'Read
strSql = "select distinct aa.lotid from mappingdatatest aa inner join customeroitbl_test bb on to_char(bb.id) = aa.filename " & "and aa.lotid = bb.source_batch_id and aa.customershortname = bb.customershortname " & "where bb.customershortname = '" & strCustCode & "' and bb.mpn_desc = '" & strCustPN & "' " & "and not exists (select 1 from ib_waferlist cc where cc.waferid = aa.substrateid)"
strSqlPart_OrderBy = " order by aa.lotid"

Select Case cbWOType(0).text

    Case "普通工单"
        strSqlPart_Flag = " and aa.flag = 'Y'"

    Case "重工工单"
        strSqlPart_WaferID = " and instr(aa.substrateid,'+') > 0"

    Case "Dummy工单"
        strSqlPart_Flag = " and aa.flag = 'T'"
        strSqlPart_LotID = " and (aa.lotid like 'D%' or aa.lotid like 'SI%')"

    Case "玻璃工单"
        strSqlPart_Flag = " and aa.flag = 'T'"
        strSqlPart_LotID = " and aa.lotid like 'G%' "

    Case "硅基工单"
        strSqlPart_Flag = " and aa.flag = 'T'"
        strSqlPart_LotID = " and aa.lotid like 'SI%' "
        strSqlPart_WaferID = " and instr(aa.substrateid,'+') = 0"

    Case "FO_CSP工单"
        strSqlPart_Flag = " and aa.flag = 'T'"
        strSqlPart_LotID = " and aa.lotid like 'SI%' "
        strSqlPart_WaferID = " and instr(aa.substrateid,'+') > 0"

End Select

strSql = strSql & strSqlPart_Flag & strSqlPart_WaferID & strSqlPart_LotID & strSqlPart_OrderBy
Set rs = Get_OracleRs(strSql)
'Show
lstLotID.Clear
If Not rs.EOF Then

    Do While Not rs.EOF
        lstLotID.AddItem Trim("" & rs!LOTID)
        rs.MoveNext
    Loop
Else

    Select Case cbWOType(0).text

        Case "普通工单"
            MsgBox "查询不到该客户机种的上传WO,或者上传的WO已经开了工单" & vbCrLf & "请查询LotID是否存在,以及开立工单的情况", vbInformation, "提示"

        Case "重工工单"
            MsgBox "查询不到该客户机种的重工WO" & vbCrLf & "请手动维护", vbInformation, "提示"

        Case "Dummy工单"
            MsgBox "查询不到该客户机种的DummyWO" & vbCrLf & "请手动维护", vbInformation, "提示"

        Case "玻璃工单"
            MsgBox "查询不到该客户机种的玻璃WO" & vbCrLf & "请手动维护", vbInformation, "提示"

        Case "硅基工单"
            MsgBox "查询不到该客户机种的硅基WO" & vbCrLf & "请手动维护", vbInformation, "提示"

        Case "FO_CSP工单"
            MsgBox "查询不到该客户机种的FO_CSP WO" & vbCrLf & "请手动维护", vbInformation, "提示"

    End Select

End If

rs.Close
Set rs = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       lstLotID_Click
' Description:       根据LOTID展开Wafer列表
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/11-16:25:33
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub lstLotID_Click()
Dim i        As Integer
Dim strLotID As String

With lstLotID

    For i = 0 To .ListCount - 1
        strLotID = Trim$("" & .List(i))
        If .Selected(i) = True Then
            Call ShowWaferList(strLotID, 1)
        Else
            Call ShowWaferList(strLotID, 2)

        End If

    Next

End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ShowWaferList
' Description:       显示工单明细
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/11-16:49:23
'
' Parameters :       strLotID (String)
'--------------------------------------------------------------------------------
Private Sub ShowWaferList(strLotID As String, intBJ As Integer)
Dim strSql   As String
Dim i        As Long
Dim strWOID  As String
Dim rsDetail As New ADODB.Recordset

If intBJ = 1 Then

    With fpSDetail

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_WO_DETAIL.E_LOTID
            If strLotID = Trim$("" & .text) Then
                Exit Sub

            End If

        Next
        '查询资料
        strSql = "select a.lotid,a.wafer_id,a.substrateid,(a.passbincount+a.failbincount) GROSSDIES,a.passbincount GOODDIES,a.failbincount NGDIES ,a.productid from mappingdatatest a where a.lotid = '" & strLotID & "' and not exists(select 1 from ib_waferlist b where b.waferid = a.substrateid) order by a.substrateid"
        Set rsDetail = Get_OracleRs(strSql)
        If Not rsDetail.EOF Then

            For i = 1 To rsDetail.RecordCount
                .MaxRows = .MaxRows + 1
                .SetText E_WO_DETAIL.E_CHOOSE, .MaxRows, 1
                .SetText E_WO_DETAIL.E_LOTID, .MaxRows, Trim$("" & rsDetail!LOTID)
                .SetText E_WO_DETAIL.E_WAFERNO, .MaxRows, Trim$("" & rsDetail!wafer_id)
                .SetText E_WO_DETAIL.E_WAFERID, .MaxRows, Trim$("" & rsDetail!SUBSTRATEID)
                .SetText E_WO_DETAIL.E_GROSSDIES, .MaxRows, Trim$("" & rsDetail!GROSSDIES)
                .SetText E_WO_DETAIL.E_GOODDIES, .MaxRows, Trim$("" & rsDetail!GOODDIES)
                .SetText E_WO_DETAIL.E_NGDIES, .MaxRows, Trim$("" & rsDetail!NGDIES)
                .SetText E_WO_DETAIL.E_MARKINGCODE, .MaxRows, Trim$("" & rsDetail!PRODUCTID)
                rsDetail.MoveNext
            Next

        End If

        rsDetail.Close
        Set rsDetail = Nothing

    End With

End If

If intBJ = 2 Then

    With fpSDetail
        Set .DataSource = Nothing

        For i = .MaxRows To 1 Step -1
            .Row = i
            .Col = E_WO_DETAIL.E_LOTID
            If strLotID = Trim$("" & .text) Then
                .DeleteRows i, 1
                .MaxRows = .MaxRows - 1

            End If

        Next

    End With

End If

'刷新数量
End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CreateWorkOrder
' Description:       创建工单
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-25-9:52:53
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub CreateWorkOrder()
If Not CheckHandler Then Exit Sub
Call SaveHandler
End Sub

Private Function CheckHandler() As Boolean
CheckHandler = False
If Not CheckByWO Then Exit Function     '工单层级数据检查
If Not CheckByLot Then Exit Function    'Lot层级数据检查
If Not CheckByWafer Then Exit Function  'Wafer层级数据检查
CheckHandler = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckByWO
' Description:       检查工单层级数据
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/7-17:35:23
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function CheckByWO() As Boolean
Dim strWOType   As String
Dim strCustCode As String
Dim strCustPN   As String
Dim strHTPN     As String
Dim strproduct  As String
Dim strSql      As String

CheckByWO = False
If cbWOType(0).text = "" Then
    MsgBox "请输入工单类型", vbCritical, "提示"
    Exit Function

End If

If cbWOType(1).text = "" Then
    MsgBox "请输入工单用途", vbCritical, "提示"
    Exit Function

End If

If cbCustCode.text = "" Then
    MsgBox "请输入客户代码", vbCritical, "提示"
    Exit Function

End If

If cbCustPN.text = "" Then
    MsgBox "请输入客户机种", vbCritical, "提示"
    Exit Function

End If

If cbHTPN.text = "" Then
    MsgBox "请输入厂内机种", vbCritical, "提示"
    Exit Function

End If

If cbProduct.text = "" Then
    MsgBox "请输入成品料号", vbCritical, "提示"
    Exit Function

End If

If txtWODept.text = "" Then
    MsgBox "请输入工单部门", vbCritical, "提示"
    Exit Function

End If

If cbWOName.text = "" Then
    MsgBox "请输入工单前缀", vbCritical, "提示"
    Exit Function

End If

If cb37Pri(0).text = "" Then
    MsgBox "请输入PRI", vbCritical, "提示"
    Exit Function

End If

strWOType = cbWOType(0).text
strCustCode = Trim$(cbCustCode.text)
strCustPN = Trim$(cbCustPN.text)
strHTPN = Trim$(cbHTPN.text)
strproduct = Trim$(cbProduct.text)
'NPI机种对照表卡控
strSql = "select * from tbltsvnpiproduct where customershortname = '" & strCustCode & "' and customerptno1 = '" & strCustPN & "' and qtechptno = '" & strHTPN & "' and qtechptno2 = '" & strproduct & "'"
If Get_OracleCnt(strSql) = 0 Then
    MsgBox "NPI未维护相关对照记录,请联系NPI维护", vbCritical, "警告"
    Exit Function

End If

'工单日期卡控
If Not (dTEnd.Value > dTBegin.Value) Then
    MsgBox "完工日期必须大于开工日期", vbCritical, "警告"
    Exit Function

End If

'成品料号卡控
If strWOType = "普通工单" Then
    strSql = "SELECT b.料号 FROM [erpdata].[dbo].[TSVtblSetMRule] a inner join [erpdata].[dbo].[TSVtblMRuleData] b on a.材料规范编号 = b.材料规范编号 where a.物料编号='" & strproduct & "'  and a.审核日期 is not null "
    If Get_SqlserverCnt(strSql) = 0 Then
        MsgBox "系统中该料号的BOM不存在或未审核,请联系相关的人,先维护并审核BOM", vbCritical, "警告"
        Exit Function

    End If

End If

'玻璃工单卡控
If strWOType = "玻璃工单" Then
    strSql = "select * from tbltsvnpiproduct where customershortname = '" & strCustCode & "' and customerptno1 = '" & strCustPN & "' and qtechptno = '" & strHTPN & "' and  customerptno3 is not null and customerptno4 is not null and customerptno5 is not null and customerptno6 is not null"
    If Get_OracleCnt(strSql) = 0 Then
        MsgBox "玻璃工单没有维护特定的信息(清洗步骤,CV高度,清洗程序,玻璃规格)" & vbCrLf & "请联系NPI维护对应机种的信息", vbCritical, "警告"
        Exit Function

    End If

End If

CheckByWO = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckByLot
' Description:       检查Lot层级数据
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/9-9:41:53
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function CheckByLot() As Boolean
CheckByLot = False
CheckByLot = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckByWafer
' Description:       检查Wafer层级数据
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/7-17:35:33
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function CheckByWafer() As Boolean
Dim strSql        As String
Dim strWOType     As String
Dim strCustCode   As String
Dim strCustPN     As String
Dim strHTPN       As String
Dim strproduct    As String
Dim lNpiGrossDies As Long
Dim i             As Integer
Dim bChoose       As Boolean

CheckByWafer = False
bChoose = False
strWOType = cbWOType(0).text
strCustCode = Trim(cbCustCode.text)
strCustPN = Trim$(cbCustPN.text)
strHTPN = Trim$(cbHTPN.text)
strproduct = Trim$(cbProduct.text)
strSql = "select customerdieqty from tbltsvnpiproduct where  customershortname = '" & strCustCode & "' and customerptno1 = '" & strCustPN & "' and qtechptno = '" & strHTPN & "' and qtechptno2 = '" & strproduct & "' and customerdieqty is not null "
lNpiGrossDies = Get_OracleNo(strSql)
If lNpiGrossDies = 0 Then
    MsgBox "NPI对照表未维护正确的GROSSDIES,请联系NPI重新维护", vbCritical, "警告"
    Exit Function

End If

With fpSDetail

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_WO_DETAIL.E_CHOOSE
        If .Value = 1 Then
            bChoose = True
            'GrossDies卡控
            If strWOType = "普通工单" Then
                .Col = E_WO_DETAIL.E_GROSSDIES
                If CLng(.text) <> lNpiGrossDies Then
                    MsgBox "NPI维护的GROSSDIES为: " & lNpiGrossDies & vbCrLf & "WO维护的GROSSDIES为: " & .text & vbCrLf & "二者不一致,请联系双方确认", vbCritical, "警告"
                    Exit Function

                End If

            End If

        End If

    Next i

End With

If Not bChoose Then
    MsgBox "请选择需要开立工单的WaferID", vbCritical, "警告"
    Exit Function

End If

CheckByWafer = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       SaveHandler
' Description:       保存数据
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/9-13:06:33
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub SaveHandler()
Dim tWOData As T_WO_HEADER
Dim tWaferData As T_WO_DETAIL

If chkLots.Value = 1 Then
    
Else
    
End If


Call GetWOData(tWOData)
Call SaveWOData(tWOData)

Call GetLotData
Call SaveLotData

Call GetWaferData(tWaferData)
Call SaveWaferData(tWaferData)

End Sub


'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetWOData
' Description:       获取工单层级数据
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/12-8:46:39
'
' Parameters :       tWOData (T_WO_HEADER)
'--------------------------------------------------------------------------------
Private Sub GetWOData(ByRef tWOData As T_WO_HEADER)



End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       SaveWOData
' Description:       保存工单层级数据
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/12-8:47:04
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub SaveWOData(ByRef tWOData As T_WO_HEADER)


End Sub

Private Sub GetLotData()

End Sub

Private Sub SaveLotData()


End Sub

Private Sub GetWaferData(tWaferData As T_WO_DETAIL)



End Sub

Private Sub SaveWaferData(tWaferData As T_WO_DETAIL)


End Sub
'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       UpdateData
' Description:       修改工单
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-25-9:53:06
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub UpdateWorkOrder()

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       DeleteData
' Description:       删除工单
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-25-9:53:12
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub DeleteWorkOrder()

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ExportData
' Description:       导出工单
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-25-9:53:17
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub ExportWOData()

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetNewWOID
' Description:       获取新工单号
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/11-17:36:10
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function GetNewWOID() As String
Dim strPre3     As String
Dim strseq      As String
Dim seqTemp     As Integer
Dim strHeadChar As String
Dim strDateChar As String
Dim lSeq        As Long
Dim strNewWOID  As String
Dim strSql      As String

strPre3 = UCase(Trim(cbWOName.text))
strHeadChar = strPre3
strseq = GetWoIDTemp(strPre3)
strDateChar = Right(Year(DateTime.DATE), 2) & Right("0" & Month(DateTime.DATE), 2)
strPre3 = strPre3 & "-" & strDateChar
strseq = Right("000" & CStr(CInt(strseq)), 4)
lSeq = CLng(strseq)
strNewWOID = strPre3 & strseq
strSql = "insert into TSV_WO_SEQ_TAB(wotype,ymonth,sequenceID,flag,WOID) values ( '" & strHeadChar & "','" & strDateChar & "'," & lSeq & ", 'Y','" & strNewWOID & "' ) "
AddSql (strSql)
GetNewWOID = strNewWOID

End Function
