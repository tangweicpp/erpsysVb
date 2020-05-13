VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPOPriceSys_NEW 
   Caption         =   "市场部订单价格维护"
   ClientHeight    =   11400
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14610
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
   ScaleHeight     =   11400
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   20000
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   17805
      _ExtentX        =   31406
      _ExtentY        =   35269
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "维护界面"
      TabPicture(0)   =   "FrmPOPriceSysNew.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "待审核界面"
      TabPicture(1)   =   "FrmPOPriceSysNew.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chk"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      Begin VB.CheckBox chk 
         Caption         =   "全选/反选"
         Height          =   255
         Left            =   -74760
         TabIndex        =   41
         Top             =   480
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "订单价格维护待审核明细"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   7575
         Left            =   -74880
         TabIndex        =   39
         Top             =   840
         Width           =   17295
         Begin FPSpreadADO.fpSpread fps 
            Height          =   7095
            Index           =   1
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   17055
            _Version        =   524288
            _ExtentX        =   30083
            _ExtentY        =   12515
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
            SpreadDesigner  =   "FrmPOPriceSysNew.frx":0038
            TextTip         =   2
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "订单价格维护明细"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   7575
         Left            =   240
         TabIndex        =   37
         Top             =   3840
         Width           =   17295
         Begin FPSpreadADO.fpSpread fps 
            Height          =   7095
            Index           =   0
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   17055
            _Version        =   524288
            _ExtentX        =   30083
            _ExtentY        =   12515
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
            SpreadDesigner  =   "FrmPOPriceSysNew.frx":04A8
            TextTip         =   2
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "订单价格设定参数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3375
         Left            =   4920
         TabIndex        =   16
         Top             =   360
         Width           =   4335
         Begin VB.TextBox txtDiePrice 
            BackColor       =   &H00FFC0FF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3120
            TabIndex        =   25
            Top             =   2850
            Width           =   975
         End
         Begin VB.TextBox txtRate 
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   3120
            TabIndex        =   24
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtPcePrice 
            BackColor       =   &H00FFC0FF&
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
            Left            =   3120
            TabIndex        =   23
            Top             =   2400
            Width           =   975
         End
         Begin VB.ComboBox cmbUnit 
            BackColor       =   &H00FFC0FF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "FrmPOPriceSysNew.frx":0918
            Left            =   1560
            List            =   "FrmPOPriceSysNew.frx":0922
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox txtDies 
            BackColor       =   &H00FFC0FF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1560
            TabIndex        =   21
            Top             =   2850
            Width           =   1095
         End
         Begin VB.TextBox txtPces 
            BackColor       =   &H00FFC0FF&
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
            Left            =   1560
            TabIndex        =   20
            Top             =   2400
            Width           =   1095
         End
         Begin VB.TextBox txtFileName 
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   960
            TabIndex        =   19
            Top             =   240
            Width           =   3015
         End
         Begin VB.TextBox txtBJDH 
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   960
            TabIndex        =   18
            Top             =   540
            Width           =   3015
         End
         Begin VB.ComboBox cbPOType 
            BackColor       =   &H00FFC0FF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "FrmPOPriceSysNew.frx":0934
            Left            =   960
            List            =   "FrmPOPriceSysNew.frx":0944
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1200
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DT3 
            Height          =   375
            Left            =   960
            TabIndex        =   26
            Top             =   825
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12582912
            CalendarForeColor=   49152
            CalendarTitleBackColor=   255
            CalendarTitleForeColor=   65535
            CalendarTrailingForeColor=   16711935
            Format          =   161415169
            CurrentDate     =   43308
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单价"
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
            Index           =   15
            Left            =   2640
            TabIndex        =   36
            Top             =   2895
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单价"
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
            Index           =   14
            Left            =   2640
            TabIndex        =   35
            Top             =   2445
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单价单位/币别"
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
            Index           =   10
            Left            =   120
            TabIndex        =   34
            Top             =   1965
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "数量(&Dies)"
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
            Index           =   9
            Left            =   240
            TabIndex        =   33
            Top             =   2895
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "数量(&Wafers)"
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
            Index           =   7
            Left            =   240
            TabIndex        =   32
            Top             =   2445
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "汇率"
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
            Left            =   2640
            TabIndex        =   31
            Top             =   1965
            Width           =   420
         End
         Begin VB.Shape Shape1 
            Height          =   495
            Index           =   1
            Left            =   120
            Top             =   2760
            Width           =   4095
         End
         Begin VB.Shape Shape1 
            Height          =   495
            Index           =   0
            Left            =   120
            Top             =   2280
            Width           =   4095
         End
         Begin VB.Shape Shape1 
            DrawMode        =   1  'Blackness
            Height          =   495
            Index           =   2
            Left            =   120
            Top             =   1800
            Width           =   4095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "订单日期"
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
            Index           =   11
            Left            =   120
            TabIndex        =   30
            Top             =   900
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "文件名"
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
            Index           =   13
            Left            =   210
            TabIndex        =   29
            Top             =   285
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "报价单号"
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
            Index           =   12
            Left            =   165
            TabIndex        =   28
            Top             =   585
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "订单类型"
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
            Index           =   5
            Left            =   120
            TabIndex        =   27
            Top             =   1230
            Width           =   840
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "订单价格维护选项"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   4575
         Begin VB.ComboBox cbCustCode 
            BackColor       =   &H00FFC0FF&
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
            Height          =   300
            Left            =   1080
            TabIndex        =   7
            Top             =   510
            Width           =   1575
         End
         Begin VB.TextBox txtCustFullName 
            BackColor       =   &H00FFC0FF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   935
            Width           =   3015
         End
         Begin VB.TextBox txtCustPO 
            BackColor       =   &H00FFC0FF&
            Height          =   300
            Left            =   1080
            TabIndex        =   5
            Top             =   1345
            Width           =   3015
         End
         Begin VB.TextBox txtCustShortName 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   518
            Width           =   1215
         End
         Begin VB.ComboBox txtCustPN 
            BackColor       =   &H00FFC0FF&
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
            Height          =   300
            Left            =   1080
            TabIndex        =   3
            Top             =   1800
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtEnd 
            Height          =   300
            Left            =   2880
            TabIndex        =   8
            Top             =   2880
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
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
            CalendarBackColor=   12582912
            CalendarForeColor=   49152
            CalendarTitleBackColor=   255
            CalendarTitleForeColor=   65535
            CalendarTrailingForeColor=   16711935
            Format          =   161415169
            CurrentDate     =   43308
         End
         Begin MSComCtl2.DTPicker dtStart 
            Height          =   300
            Left            =   1080
            TabIndex        =   9
            Top             =   2880
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
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
            Format          =   161415169
            CurrentDate     =   43271
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "查询日期"
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
            Index           =   4
            Left            =   240
            TabIndex        =   15
            Top             =   2925
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "至"
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
            Left            =   2640
            TabIndex        =   14
            Top             =   2925
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "客户PO号"
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
            Index           =   3
            Left            =   240
            TabIndex        =   13
            Top             =   1410
            Width           =   840
         End
         Begin VB.Label Label1 
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
            Index           =   2
            Left            =   240
            TabIndex        =   12
            Top             =   1830
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "客户全称"
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
            Left            =   240
            TabIndex        =   11
            Top             =   972
            Width           =   840
         End
         Begin VB.Label Label1 
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
            Index           =   0
            Left            =   240
            TabIndex        =   10
            Top             =   555
            Width           =   840
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14610
      _ExtentX        =   25770
      _ExtentY        =   1402
      ButtonWidth     =   1138
      ButtonHeight    =   1349
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查询 "
            Key             =   "QUERY"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "ADD"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "UPDATE"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "删除"
            Key             =   "DELETE"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "导出"
            Key             =   "EXPORT"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "待审核"
            Key             =   "WAIT_PASS"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "审核"
            Key             =   "PASS"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "反审核"
            Key             =   "CANCEL_PASS"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "EXIT"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   10320
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPOPriceSysNew.frx":0970
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPOPriceSysNew.frx":2AAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPOPriceSysNew.frx":36FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPOPriceSysNew.frx":6586
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPOPriceSysNew.frx":8D38
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPOPriceSysNew.frx":AE72
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPOPriceSysNew.frx":D624
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPOPriceSysNew.frx":FDD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPOPriceSysNew.frx":12E58
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPOPriceSysNew.frx":1560A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPOPriceSysNew.frx":15924
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPOPriceSysNew.frx":165FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPOPriceSysNew.frx":19680
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPOPriceSysNew.frx":1BE32
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmPOPriceSys_NEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Enum E_POPRI

    E_CHECK = 1
    E_CUSTCODE
    E_CUSTNAME
    E_POID
    E_PODATE
    E_PODATE_CREATE
    E_POTYPE
    E_CUSTPN
    E_Rate
    E_PECS
    E_PEC_PRI
    E_DIES
    E_DIE_PRI
    E_UNIT
    E_POFILENAME
    E_CUSTSHORTNAME
    E_BJDH
    E_END

End Enum



Private Sub cbPOType_Click()
Dim strner As String
Dim strnerdevice As String
Dim nerqty  As String
Dim strdevice As String
Dim rs          As New ADODB.Recordset


If cbPOType.text = "NRE订单" Then
    
   If Trim(cbCustCode.text) = "" Or Trim(txtCustPN.text) = "" Or Trim$(txtCustPO.text) = "" Then
    MsgBox " 请输入客户代码和PO及客户机种! "
    Exit Sub
    
   End If
   
   strdevice = Get_SqlStr("  SELECT COUNT(*) FROM erptemp..HT_PRICE_CONTROL A WHERE A.CUST_DEVICE = '" & Trim(txtCustPN.text) & "' and a.cust_id = '" & Trim(cbCustCode.text) & "' and flag = 0 ")
   
   If Val(strdevice) > 0 Then
    
   nerqty = Get_SqlStr(" SELECT a.product FROM erptemp..HT_PRICE_CONTROL A  WHERE A.CUST_DEVICE = '" & Trim(txtCustPN.text) & "' and a.cust_id = '" & Trim(cbCustCode.text) & "'and flag = 0  ")
    
   strner = Get_OracleStr(" select nvl(sum(a.peaceqty),0) from TSV_MD_POPrice a where a.po_type = 'NRE订单' and a.customershortname = '" & Trim(cbCustCode.text) & "' and a.pt = '" & Trim(txtCustPN.text) & "' ")
   
   If Val(nerqty) <= Val(strner) Then
   
      MsgBox " 机种 " & Trim(txtCustPN.text) & " NER订单数量超过定义, 不允许维护!"
      Exit Sub
      
   End If
   
   txtPces.text = nerqty
   txtDIES.text = nerqty
   txtPcePrice.text = 0
   txtDiePrice.text = 0

End If
End If


End Sub

Private Sub chk_Click()
Dim i As Integer

If chk.Value = 1 Then

    For i = 1 To fps(1).MaxRows

        With fps(1)
            .Row = i
            .Col = 1
            .text = 1

        End With

    Next i

ElseIf chk.Value = 0 Then

    For i = 1 To fps(1).MaxRows

        With fps(1)
            .Row = i
            .Col = 1
            .text = 0

        End With

    Next i

End If

End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab

    Case 0
        Toolbar1.Buttons("QUERY").Enabled = True
        Toolbar1.Buttons("ADD").Enabled = True
        Toolbar1.Buttons("DELETE").Enabled = True
        Toolbar1.Buttons("UPDATE").Enabled = True
        Toolbar1.Buttons("EXPORT").Enabled = True
        Toolbar1.Buttons("WAIT_PASS").Enabled = False
        Toolbar1.Buttons("PASS").Enabled = False
    Case 1
        Toolbar1.Buttons("QUERY").Enabled = False
        Toolbar1.Buttons("ADD").Enabled = False
        Toolbar1.Buttons("DELETE").Enabled = False
        Toolbar1.Buttons("UPDATE").Enabled = False
        Toolbar1.Buttons("EXPORT").Enabled = False
        Toolbar1.Buttons("WAIT_PASS").Enabled = True
        Toolbar1.Buttons("PASS").Enabled = True
End Select
End Sub

Private Sub txtCustPN_Change()
Dim strSql      As String
Dim rs          As New ADODB.Recordset
Dim strCustPN   As String
Dim strPOID     As String
Dim strCustCode As String

strPOID = Trim$(txtCustPO.text)
strCustCode = Trim$(cbCustCode.text)
strCustPN = Trim$(txtCustPN.text)

txtPces.text = ""
txtDIES.text = ""

If strPOID = "" Or strCustCode = "" Or strCustPN = "" Then Exit Sub

strSql = "select Sum(bb.PASSBINCOUNT + bb.FailBinCount) qty, " & "      Count(bb.SUBSTRATEID) qty2 " & " from customeroitbl_test aa, mappingdatatest bb " & " Where bb.filename = to_char(aa.id) " & "  And bb.LOTID = aa.source_batch_id " & "  and bb.flag = 'Y' " & "  and aa.flag = 'Y' " & "  and aa.po_num = '" & strPOID & "' " & "  and aa.mpn_desc = '" & strCustPN & "' " & "  and aa.customershortname = '" & strCustCode & "' " & " GROUP BY AA.MPN_DESC "
Set rs = Get_OracleRs(strSql)
If rs.EOF Then Exit Sub

txtPces.text = Trim$("" & rs("qty2").Value)
txtDIES.text = Trim$("" & rs("qty").Value)
End Sub

Private Sub txtCustPN_Click()
Dim strSql      As String
Dim rs          As New ADODB.Recordset
Dim strCustPN   As String
Dim strPOID     As String
Dim strCustCode As String

strPOID = Trim$(txtCustPO.text)
strCustCode = Trim$(cbCustCode.text)
strCustPN = Trim$(txtCustPN.text)

txtPces.text = ""
txtDIES.text = ""

If strPOID = "" Or strCustCode = "" Or strCustPN = "" Then Exit Sub

strSql = "select to_char(aa.qtech_created_date, 'YYYY-MM-DD') qtech_created_date, " & "      Sum(bb.PASSBINCOUNT + bb.FailBinCount) qty, " & "      Count(bb.SUBSTRATEID) qty2 " & " from customeroitbl_test aa, mappingdatatest bb " & " Where bb.filename = to_char(aa.id) " & "  And bb.LOTID = aa.source_batch_id " & "  and bb.flag = 'Y' " & "  and aa.flag = 'Y' " & "  and aa.po_num = '" & strPOID & "' " & "  and aa.mpn_desc = '" & strCustPN & "' " & "  and aa.customershortname = '" & strCustCode & "' " & " GROUP BY to_char(aa.qtech_created_date, 'YYYY-MM-DD'), AA.MPN_DESC "
Set rs = Get_OracleRs(strSql)
If rs.EOF Then Exit Sub

txtPces.text = Trim$("" & rs("qty2").Value)
txtDIES.text = Trim$("" & rs("qty").Value)
End Sub

Private Sub txtCustPN_DropDown()
Dim strSql      As String
Dim rs          As New ADODB.Recordset
Dim strCustPN   As String
Dim strPOID     As String
Dim strCustCode As String

strPOID = Trim$(txtCustPO.text)
strCustCode = Trim$(cbCustCode.text)
'strCustPN = Trim$(txtCustPN.Text)

strSql = "select distinct mpn_desc from customeroitbl_test where po_num = '" & strPOID & "'"
Set rs = Get_OracleRs(strSql)

If Not rs.EOF Then
    txtCustPN.Clear
    Do While Not rs.EOF
        txtCustPN.AddItem ("" & rs(0))
        rs.MoveNext
    Loop
    
End If

'strSql = "select to_char(aa.qtech_created_date, 'YYYY-MM-DD') qtech_created_date, " & "      Sum(bb.PASSBINCOUNT + bb.FailBinCount) qty, " & "      Count(bb.SUBSTRATEID) qty2 " & " from customeroitbl_test aa, mappingdatatest bb " & " Where bb.filename = to_char(aa.id) " & "  And bb.LOTID = aa.source_batch_id " & "  and bb.flag = 'Y' " & "  and aa.flag = 'Y' " & "  and aa.po_num = '" & strPOID & "' " & "  and aa.mpn_desc = '" & strCustPN & "' " & "  and aa.customershortname = '" & strCustCode & "' " & " GROUP BY to_char(aa.qtech_created_date, 'YYYY-MM-DD'), AA.MPN_DESC "
'Set rs = Get_OracleRs(strSql)
'txtPieces.Text = Trim$("" & rs("qty2").Value)
'txtDies.Text = Trim$("" & rs("qty").Value)

End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
InitCuscode
InitDate
InitFps
txtCustPN.text = ""
End Sub

Private Sub InitCuscode()
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = SqlConnect
rs.Source = "select distinct 客户代码 from tblxcustomer"
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
cbCustCode.Clear
If rs.RecordCount > 0 Then
    rs.MoveFirst

    For i = 1 To rs.RecordCount
        cbCustCode.AddItem Trim(rs("客户代码"))
        rs.MoveNext
    Next i

End If

End Sub

Private Sub InitDate()
dtStart.Value = Now - 1
dTEnd.Value = Now
DT3.Value = Now

End Sub

Private Sub InitFps()

With fps(0)
    .ReDraw = False
    .MaxCols = E_POPRI.E_END - 1
    .MaxRows = 0
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .TypeVAlign = TypeVAlignCenter
    .SelForeColor = &HFF8080
    .SetText E_POPRI.E_CHECK, 0, "√"
    .SetText E_POPRI.E_CUSTCODE, 0, "客户代码"
    .SetText E_POPRI.E_CUSTNAME, 0, "客户全称"
    .SetText E_POPRI.E_POID, 0, "订单号"
    .SetText E_POPRI.E_PODATE, 0, "提供订单日期"
    .SetText E_POPRI.E_PODATE_CREATE, 0, "录入日期"
    .SetText E_POPRI.E_POTYPE, 0, "订单类型"
    .SetText E_POPRI.E_CUSTPN, 0, "机种"
    .SetText E_POPRI.E_Rate, 0, "汇率"
    .SetText E_POPRI.E_PECS, 0, "订单片数"
    .SetText E_POPRI.E_PEC_PRI, 0, "单片价格"
    .SetText E_POPRI.E_DIES, 0, "订单DIE数"
    .SetText E_POPRI.E_DIE_PRI, 0, "单DIE价格"
    .SetText E_POPRI.E_UNIT, 0, "单价单位"
    .SetText E_POPRI.E_POFILENAME, 0, "返点文件名"
    .SetText E_POPRI.E_CUSTSHORTNAME, 0, "客户简写"
    .SetText E_POPRI.E_BJDH, 0, "报价单号"
    
    .Col = E_POPRI.E_CHECK
    .Lock = False
    .CellType = CellTypeCheckBox
    
    .Col = E_POPRI.E_PECS
    .Lock = False
    .BackColor = vbGreen
    
    .Col = E_POPRI.E_PEC_PRI
    .Lock = False
    .BackColor = vbGreen
    
    .Col = E_POPRI.E_DIES
    .Lock = False
    .BackColor = vbGreen
    
    .Col = E_POPRI.E_DIE_PRI
    .Lock = False
    .BackColor = vbGreen
    
    .ColWidth(E_POPRI.E_CHECK) = 4
    .RowHeight(0) = 20
    .RowHeight(-1) = 15
    .ReDraw = True

End With

With fps(1)
    .ReDraw = False
    .MaxCols = E_POPRI.E_END - 1
    .MaxRows = 0
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .TypeVAlign = TypeVAlignCenter
    .SelForeColor = &HFF8080
    .SetText E_POPRI.E_CHECK, 0, "√"
    .SetText E_POPRI.E_CUSTCODE, 0, "客户代码"
    .SetText E_POPRI.E_CUSTNAME, 0, "客户全称"
    .SetText E_POPRI.E_POID, 0, "订单号"
    .SetText E_POPRI.E_PODATE, 0, "提供订单日期"
    .SetText E_POPRI.E_PODATE_CREATE, 0, "录入日期"
    .SetText E_POPRI.E_POTYPE, 0, "订单类型"
    .SetText E_POPRI.E_CUSTPN, 0, "机种"
    .SetText E_POPRI.E_Rate, 0, "汇率"
    .SetText E_POPRI.E_PECS, 0, "订单片数"
    .SetText E_POPRI.E_PEC_PRI, 0, "单片价格"
    .SetText E_POPRI.E_DIES, 0, "订单DIE数"
    .SetText E_POPRI.E_DIE_PRI, 0, "单DIE价格"
    .SetText E_POPRI.E_UNIT, 0, "单价单位"
    .SetText E_POPRI.E_POFILENAME, 0, "返点文件名"
    .SetText E_POPRI.E_CUSTSHORTNAME, 0, "客户简写"
    .SetText E_POPRI.E_BJDH, 0, "报价单号"
    
    .Col = E_POPRI.E_CHECK
    .Lock = False
    .CellType = CellTypeCheckBox
    
    .ColWidth(E_POPRI.E_CHECK) = 4
    .RowHeight(0) = 20
    .RowHeight(-1) = 15
    .ReDraw = True

End With

End Sub

Private Sub cbCustCode_Change()
ListCustInfo

End Sub

Private Sub cbCustCode_Click()
ListCustInfo

End Sub

Private Sub ListCustInfo()
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = SqlConnect
rs.Source = "select distinct 客户名称 from tblxcustomer where 客户代码='" & Trim(cbCustCode) & "'"
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount > 0 Then
    rs.MoveFirst
    txtCustFullName.text = Trim(rs!客户名称)

End If

txtCustShortName.text = GetCustomerNameSqlServer1(Trim(cbCustCode.text))
If GetCustomerNameSqlServer2(Trim(cbCustCode.text)) = "01" Then
    cmbUnit.ListIndex = 0
Else
    cmbUnit.ListIndex = 1

End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "QUERY"
        ForQuery
        Toolbar1.Buttons("PASS").Enabled = False
    Case "ADD"
        ForAdd

    Case "UPDATE"
        ForUpdate

    Case "DELETE"
        ForDelete

    Case "EXPORT"
        ForExport

    Case "WAIT_PASS"
        ForWaitPass
    
    Case "PASS"
        ForPass
        
    Case "EXIT"
        Unload Me

End Select

End Sub

Private Sub ForWaitPass()
Dim strSql       As String
Dim rs           As New ADODB.Recordset

SSTab1.Tab = 1
strSql = "select '' ,CUSTOMERSHORTNAME,CUSTOMERNAME ,PO_NUM ,PO_DATE ,qtech_created_date, PO_TYPE ,PT , rate, PeaceQty,PRICE,QTY,to_char(DIE_PRICE,'fm999999990.999999999'),UNIT, FILENAME,CUSTAA,BJ from TSV_MD_POPrice_TMP where flag='Y' order by qtech_created_date desc"

Set rs = Get_OracleRs(strSql)

With fps(1)
    .MaxRows = 0
    If rs.RecordCount > 0 Then
        Toolbar1.Buttons("PASS").Enabled = True
        Set .DataSource = rs
        
    Else
        MsgBox "查询不到需要审核的记录", vbInformation, "提示"

    End If

End With

rs.Close

End Sub

Private Sub ForPass()
Dim i          As Integer
Dim bChecked   As Boolean
Dim strid As String
Dim strPECS    As String
Dim strPEC_PRI As String
Dim strDies    As String
Dim strDIE_PRI As String
Dim strPONUM As String
Dim strCustPN As String
Dim strSql As String

If gUserName <> "16642" And gUserName <> "07885" Then
    MsgBox "你没有权限审核", vbInformation, "提示"
    Exit Sub
End If

bChecked = False

With fps(1)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .Value = 1 Then
            bChecked = True
            
            .Col = E_POPRI.E_POID
            strPONUM = Trim$(.text)
            
            .Col = E_POPRI.E_CUSTPN
            strCustPN = Trim$(.text)
            
            .Col = E_POPRI.E_CHECK
            strid = Trim$(.text)
            .Col = E_POPRI.E_PECS
            strPECS = Trim$(.text)
            .Col = E_POPRI.E_PEC_PRI
            strPEC_PRI = Trim$(.text)
            .Col = E_POPRI.E_DIES
            strDies = Trim$(.text)
            .Col = E_POPRI.E_DIE_PRI
            strDIE_PRI = Trim$(.text)
            
            strSql = "insert into TSV_MD_POPrice select * from TSV_MD_POPrice_TMP where PO_NUM = '" & strPONUM & "' and PT =  '" & strCustPN & "' "
            AddSql (strSql)
            
            strSql = "insert into erptemp..tblBB_CSRPO select * from erptemp..tblBB_CSRPO_TMP where PO_NUM = '" & strPONUM & "' and FAB_DEVICE = '" & strCustPN & "' "
            AddSql2 (strSql)
            
            strSql = "delete from TSV_MD_POPrice_TMP where PO_NUM = '" & strPONUM & "' and PT =  '" & strCustPN & "'"
            AddSql (strSql)
            
            strSql = "delete from  erptemp..tblBB_CSRPO_TMP where PO_NUM = '" & strPONUM & "' and FAB_DEVICE = '" & strCustPN & "' "
            AddSql2 (strSql)
        End If

    Next

End With

If bChecked = False Then
    MsgBox "请勾选需要审核批准的项目", vbInformation, "提示"
    Exit Sub

Else
    ShowPOPrice
    MsgBox "PO价格审核成功", vbInformation, "提示"
    Toolbar1.Buttons("PASS").Enabled = False
    
End If

End Sub

Private Sub ForQuery()
ShowPOPrice

End Sub

Private Sub ShowPOPrice()
Dim strStartDate As String
Dim strEndDate   As String
Dim strSql       As String
Dim rs           As New ADODB.Recordset


SSTab1.Tab = 0

If txtCustPO.text <> "" Then
    strSql = "select '' ,CUSTOMERSHORTNAME,CUSTOMERNAME ,PO_NUM ,PO_DATE ,qtech_created_date, PO_TYPE ,PT , rate, PeaceQty,PRICE,QTY,to_char(DIE_PRICE,'fm999999990.999999999'),UNIT, FILENAME,CUSTAA,BJ from TSV_MD_POPrice where flag='Y' and PO_NUM ='" & UCase$(Trim(txtCustPO.text)) & "' "
    If txtCustPN.text <> "" Then
        strSql = strSql & " and PT = '" & UCase(Trim(txtCustPN.text)) & "'"

    End If

    strSql = strSql & " order by qtech_created_date desc"
Else
    strStartDate = Format(dtStart.Value, "YYYY-MM-DD")
    strEndDate = Format(dTEnd.Value + 1, "YYYY-MM-DD")
    strSql = " select '' ,CUSTOMERSHORTNAME,CUSTOMERNAME ,PO_NUM ,PO_DATE ,qtech_created_date, PO_TYPE ,PT , rate,PeaceQty,PRICE,QTY,to_char(DIE_PRICE,'fm999999990.999999999'),UNIT,FILENAME,CUSTAA,BJ from TSV_MD_POPrice where flag='Y' and PO_DATE >=to_date('" & strStartDate & "','YYYY-MM-DD')  and  PO_DATE <to_date('" & strEndDate & "','YYYY-MM-DD')  "
    If cbCustCode.text <> "" Then
        strSql = strSql & " and  CUSTOMERSHORTNAME = '" & UCase(Trim(cbCustCode.text)) & "' "

    End If

    If cbPOType.text <> "" Then
        strSql = strSql & " and  PO_TYPE = '" & UCase(Trim(cbPOType.text)) & "' "

    End If

    If txtCustPN.text <> "" Then
        strSql = strSql & " and  PT = '" & UCase(Trim(txtCustPN.text)) & "' "

    End If

    strSql = strSql & " order by qtech_created_date desc"

End If

Set rs = Get_OracleRs(strSql)

With fps(0)
    .MaxRows = 0
    If rs.RecordCount > 0 Then
        Set .DataSource = rs
    Else
        MsgBox "查询不到订单价格维护记录,请确认是否输入错误", vbInformation, "提示"

    End If

End With

rs.Close

End Sub

Private Sub ForAdd()
Dim nPOTemp As POPrice
Dim rs      As ADODB.Recordset
Dim strSql  As String
Dim strrebate As String
Dim rsrebate  As New ADODB.Recordset

Dim rsdevice  As New ADODB.Recordset
Dim strprdevice As String
Dim rsprice  As New ADODB.Recordset
Dim strprice As String



If Trim(cbCustCode.text) = "" Then
    MsgBox "客户代码不可以为空！", vbExclamation, "警告"
    Exit Sub

End If

If Trim(txtCustFullName.text) = "" Then
    MsgBox "客户名称不可以为空！", vbExclamation, "警告"
    Exit Sub

End If

If Trim(txtCustPO.text) = "" Then
    MsgBox "客户PO不可以为空！", vbExclamation, "警告"
    Exit Sub

End If

If txtCustPN.text = "" Then
    MsgBox "客户机种不可以为空！", vbExclamation, "警告"
    Exit Sub

End If

If Trim(txtPcePrice.text) = "" And Trim(txtDiePrice.text) = "" Then
    MsgBox "单价不可以为空!", vbExclamation, "警告"
    Exit Sub

End If

If Trim(txtDIES.text) = "" Then
    MsgBox "DIE数量 不可以为空！"
    Exit Sub

End If

If Trim(cmbUnit.text) = "" Then
    MsgBox "币别单位 不可以为空！"
    Exit Sub

End If

If cbPOType.text = "" Then
    MsgBox "订单类型 不可以为空！"
    Exit Sub

End If

If UCase(Trim(cbCustCode.text)) = "KR001" Or UCase(Trim(cbCustCode.text)) = "81" Then
    If UCase(Trim(txtBJDH.text)) = "" Then
        MsgBox "报价单号不能为空！"
        Exit Sub

    End If

End If

nPOTemp.CreateBy = UCase(gUserName)
nPOTemp.id = GetPOPriceID()
nPOTemp.customerName = UCase(Trim(txtCustFullName.text))
nPOTemp.CUSTOMERSHORTNAME = UCase(Trim(cbCustCode.text))
nPOTemp.PODATE = Format(DT3.Value, "YYYY-MM-DD")
nPOTemp.PONo = UCase(Trim(txtCustPO.text))
nPOTemp.POType = UCase(Trim(cbPOType.text))
nPOTemp.pt = UCase(Trim(txtCustPN.text))
nPOTemp.QTY = UCase(Trim(txtDIES.text))
nPOTemp.SingDie = Trim$(txtDiePrice.text)
nPOTemp.peaseQty = UCase(Trim(txtPces.text))
nPOTemp.Price = UCase(Trim(txtPcePrice.text))
nPOTemp.unit = UCase(Trim(cmbUnit.text))
nPOTemp.File = UCase(Trim(txtFileName.text))
nPOTemp.bj = UCase(Trim(txtBJDH.text))
nPOTemp.custAA = UCase(Trim(txtCustShortName))
nPOTemp.SingWafer = Trim$(txtRate.text)


If cbPOType.text <> "NRE订单" Then

strprdevice = "  SELECT convert(varchar(100),a.wafer_price),convert(varchar(100),a.die_price),a.currency FROM erptemp..ht_price_control a  WHERE a.cust_id = '" & UCase(Trim(cbCustCode.text)) & "' AND a.cust_device = '" & UCase(Trim(txtCustPN.text)) & "' AND a.flag = 0 "

 If rsdevice.State = adStateOpen Then rsrebate.Close
    rsdevice.Open strprdevice, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If Not rsdevice.EOF Then
      
      If Trim(txtDiePrice.text) <> rsdevice.Fields(1).Value Or Trim(txtPcePrice.text) <> rsdevice.Fields(0).Value Or Trim(cmbUnit.text) <> rsdevice.Fields(2).Value Then
      
         MsgBox "PO价格与产品定义价格不一致 , 请确认产品价格！"
         Exit Sub
      
        
      End If


        
    End If

End If



strrebate = " SELECT a.cust ,a.rebate_waf,a.rebate_die FROM erptemp..ht_cust_rebate a WHERE a.cust = '" & nPOTemp.CUSTOMERSHORTNAME & "'  AND flag = 0"

 If rsrebate.State = adStateOpen Then rsrebate.Close
    rsrebate.Open strrebate, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If Not rsrebate.EOF Then
      
      nPOTemp.Price = nPOTemp.Price * Val(rsrebate.Fields(1).Value) / 100
      nPOTemp.SingDie = nPOTemp.SingDie * Val(rsrebate.Fields(2).Value) / 100
        
    End If



strSql = "select * from TSV_MD_POPrice where customershortname = '" & nPOTemp.CUSTOMERSHORTNAME & "' and PO_NUM= '" & nPOTemp.PONo & "' and PT = '" & nPOTemp.pt & "'"
Set rs = Get_OracleRs(strSql)
If rs.RecordCount > 0 Then
    MsgBox "PO: " & Trim(txtCustPO.text) & ",客户机种: " & Trim$(txtCustPN.text) & "已经存在维护记录,不可新增" & vbCrLf & "请点击修改或退出", vbInformation, "提示"

    Exit Sub
End If

strSql = "select * from TSV_MD_POPrice_tmp where customershortname = '" & nPOTemp.CUSTOMERSHORTNAME & "' and PO_NUM= '" & nPOTemp.PONo & "' and PT = '" & nPOTemp.pt & "'"
Set rs = Get_OracleRs(strSql)
If rs.RecordCount > 0 Then
    MsgBox "PO: " & Trim(txtCustPO.text) & ",客户机种: " & Trim$(txtCustPN.text) & "已经存在相同的待审核的维护记录，无法重复维护", vbInformation, "提示"

    Exit Sub
End If

Call AddPOPrice(nPOTemp)
MsgBox "新增PO价格维护成功", vbInformation, "提示"

'ForWaitPass
If cbPOType.text = "NRE订单" Then

 cbPOType.text = ""
 
End If

End Sub

Private Sub ForUpdate()
Dim i          As Integer
Dim bChecked   As Boolean
Dim strid As String
Dim strPECS    As String
Dim strPEC_PRI As String
Dim strDies    As String
Dim strDIE_PRI As String
Dim strPONUM As String
Dim strCustPN As String
Dim strSql As String


Dim strCust As String
Dim strunit As String
Dim rsdevice  As New ADODB.Recordset
Dim strprdevice As String
Dim rsprice  As New ADODB.Recordset
Dim strprice As String


bChecked = False



With fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .Value = 1 Then
            bChecked = True
            
            .Col = E_CUSTCODE
             strCust = Trim(.text)
             .Col = E_UNIT
             strunit = Trim(.text)
             
            .Col = E_POPRI.E_POID
            strPONUM = Trim$(.text)
            
            .Col = E_POPRI.E_CUSTPN
            strCustPN = Trim$(.text)
            
            .Col = E_POPRI.E_CHECK
            strid = Trim$(.text)
            .Col = E_POPRI.E_PECS
            strPECS = Trim$(.text)
            .Col = E_POPRI.E_PEC_PRI
            strPEC_PRI = Trim$(.text)
            .Col = E_POPRI.E_DIES
            strDies = Trim$(.text)
            .Col = E_POPRI.E_DIE_PRI
            strDIE_PRI = Trim$(.text)
            
            
            
    strprdevice = "  SELECT convert(varchar(100),a.wafer_price),convert(varchar(100),a.die_price),a.currency FROM erptemp..ht_price_control a  WHERE a.cust_id = '" & strCust & "' AND a.cust_device = '" & strCustPN & "' AND a.flag = 0 "

    If rsdevice.State = adStateOpen Then rsrebate.Close
    rsdevice.Open strprdevice, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If Not rsdevice.EOF Then
      
      If strDIE_PRI <> rsdevice.Fields(1).Value Or strPEC_PRI <> rsdevice.Fields(0).Value Or strunit <> rsdevice.Fields(2).Value Then
      
         MsgBox "PO价格与产品定义价格不一致 , 请确认产品价格！"
         Exit Sub
      
        
      End If


        
    End If
            
            
            
            
            strSql = "update TSV_MD_POPrice set PeaceQty = '" & strPECS & "',PRICE = '" & strPEC_PRI & "',QTY='" & strDies & "',DIE_PRICE = '" & strDIE_PRI & "' where PO_NUM = '" & strPONUM & "'  and PT =  '" & strCustPN & "' "
            AddSql (strSql)
            
            strSql = "update erptemp..tblBB_CSRPO set QTY = '" & strPECS & "',WAFER_PRICE = '" & strPEC_PRI & "',QTY_DIE = '" & strDies & "',DIE_PRICE = '" & strDIE_PRI & "' where PO_NUM = '" & strPONUM & "' and FAB_DEVICE = '" & strCustPN & "' "
            AddSql2 (strSql)
        End If

    Next

End With

If bChecked = False Then
    MsgBox "请勾选需要修改的项目", vbInformation, "提示"
    Exit Sub

Else
    ShowPOPrice
    MsgBox "PO价格更新成功", vbInformation, "提示"
    
End If

End Sub

Private Sub ForDelete()
Dim i          As Integer
Dim bChecked   As Boolean
Dim strid      As String
Dim strPECS    As String
Dim strPEC_PRI As String
Dim strDies    As String
Dim strDIE_PRI As String
Dim strPONUM   As String
Dim strCustPN  As String
Dim strSql     As String

bChecked = False

With fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .Value = 1 Then
            bChecked = True
            .Col = E_POPRI.E_POID
            strPONUM = Trim$(.text)
            .Col = E_POPRI.E_CUSTPN
            strCustPN = Trim$(.text)
            
            strSql = "insert into TSV_MD_POPrice_bak select * from TSV_MD_POPrice where PO_NUM = '" & strPONUM & "'  and PT =  '" & strCustPN & "' "
            AddSql (strSql)
            
            strSql = "delete from TSV_MD_POPrice where PO_NUM = '" & strPONUM & "'  and PT =  '" & strCustPN & "'"
            AddSql (strSql)
            
            strSql = "delete from erptemp..tblBB_CSRPO where PO_NUM = '" & strPONUM & "' and FAB_DEVICE = '" & strCustPN & "' "
            AddSql2 (strSql)

        End If

    Next

End With

If bChecked = False Then
    MsgBox "请勾选需要删除的项目", vbInformation, "提示"
    Exit Sub
Else
    ShowPOPrice
    MsgBox "PO价格删除成功", vbInformation, "提示"

End If

End Sub

Private Sub ForExport()
Dim xlsApp      As Excel.Application
Dim xlsBook     As Excel.Workbook
Dim xlsSheet    As Excel.Worksheet
Dim i           As Long
Dim j           As Long
Dim strFileName As String
Dim strPartName As String

On Error GoTo Ert

If fps(0).MaxRows = 0 Then
    MsgBox "没有查询数据,无法导出", vbInformation, "提示"
    Exit Sub

End If

Set xlsApp = CreateObject("Excel.Application")
Set xlsBook = xlsApp.Workbooks.Add
Set xlsSheet = xlsBook.Worksheets(1)

With xlsApp
    .Rows(1).Font.Bold = True

End With

With fps(0)

    For i = 0 To .MaxRows
        For j = 1 To .MaxCols
            .Col = j
            .Row = i
            xlsSheet.Cells(i + 1, j) = Trim$(("" & .text))
        Next j
    Next i

End With

xlsApp.Visible = True
strFileName = "C:\test\" & "PO_PRICE" & Format(Now, "YYYYMMDD") & GetGC_FileNoNew(UCase(Trim(cbCustCode.text))) & ".xlsx"
xlsBook.SaveAs strFileName
Set xlsApp = Nothing
Set xlsSheet = Nothing
Set xlsBook = Nothing
AddSql2 ("insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('','" & Format(Now, "YYYY-MM-DD") & "','Y','Auto',getdate(),'" & UCase(Trim(cbCustCode.text)) & "') ")
Exit Sub
Ert:
If Not (xlsApp Is Nothing) Then
    Set xlsApp = Nothing
    Set xlsSheet = Nothing
    Set xlsBook = Nothing

End If

End Sub



