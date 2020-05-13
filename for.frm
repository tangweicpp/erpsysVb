VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_GWZLWH 
   Caption         =   "关务维护"
   ClientHeight    =   10935
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   17940
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
   ScaleWidth      =   17940
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   12015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17775
      _ExtentX        =   31353
      _ExtentY        =   21193
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "关务维护"
      TabPicture(0)   =   "for.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lb1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lb2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lb3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lb4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lb5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lb8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lb9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lb6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lb7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fpS(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Toolbar1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ImageList1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Combo1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "fpss(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Combo2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Combo3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "DTPicker1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "DTPicker2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text5"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Command4"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Command5"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "手册号维护"
      TabPicture(1)   =   "for.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text3"
      Tab(1).Control(1)=   "Command3"
      Tab(1).Control(2)=   "fpsss(0)"
      Tab(1).Control(3)=   "CommonDialog1"
      Tab(1).Control(4)=   "Command2"
      Tab(1).Control(5)=   "Command1"
      Tab(1).Control(6)=   "Text2"
      Tab(1).Control(7)=   "Label2"
      Tab(1).Control(8)=   "Label1"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "晶圆OPEN PO查询"
      TabPicture(2)   =   "for.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fpS_Clear"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Optpatial"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Optall"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Cmd_Query"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "CmdExport"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "TxtCust"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "TxtPn"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Optpatial2"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      Begin VB.OptionButton Optpatial2 
         Caption         =   "关务已维护"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -66600
         TabIndex        =   40
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox TxtPn 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -73320
         TabIndex        =   39
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox TxtCust 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -73320
         TabIndex        =   37
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "导出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -72480
         TabIndex        =   35
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Cmd_Query 
         Caption         =   "Open PO查询"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   33
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Optall 
         Caption         =   "显示所有"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69840
         TabIndex        =   32
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Optpatial 
         Caption         =   "关务未维护"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -68400
         TabIndex        =   31
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "删 除"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9720
         TabIndex        =   30
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "添加一行"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9720
         TabIndex        =   29
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         TabIndex        =   26
         Top             =   2880
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   8880
         TabIndex        =   25
         Top             =   2400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   8040
         TabIndex        =   22
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   106758145
         CurrentDate     =   43531
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5400
         TabIndex        =   21
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   106758145
         CurrentDate     =   43531
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   4680
         TabIndex        =   20
         Top             =   2400
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "for.frx":0054
         Left            =   1320
         List            =   "for.frx":0056
         TabIndex        =   18
         Top             =   2400
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   -73800
         TabIndex        =   16
         Top             =   2520
         Width           =   3615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "查询"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69960
         TabIndex        =   14
         Top             =   2520
         Width           =   990
      End
      Begin FPSpreadADO.fpSpread fpsss 
         Height          =   7815
         Index           =   0
         Left            =   -74400
         TabIndex        =   13
         Top             =   3120
         Width           =   17535
         _Version        =   524288
         _ExtentX        =   30930
         _ExtentY        =   13785
         _StockProps     =   64
         DAutoCellTypes  =   0   'False
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
         SpreadDesigner  =   "for.frx":0058
         AppearanceStyle =   0
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -72000
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command2 
         Caption         =   "上传"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74400
         TabIndex        =   12
         Top             =   1920
         Width           =   990
      End
      Begin FPSpreadADO.fpSpread fpss 
         Height          =   2775
         Index           =   0
         Left            =   11040
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   6495
         _Version        =   524288
         _ExtentX        =   11456
         _ExtentY        =   4895
         _StockProps     =   64
         DAutoCellTypes  =   0   'False
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
         SpreadDesigner  =   "for.frx":0486
         AppearanceStyle =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   ".."
         Height          =   600
         Left            =   -64560
         TabIndex        =   10
         Top             =   1200
         Width           =   435
      End
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   -74400
         TabIndex        =   9
         Top             =   1200
         Width           =   9735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "for.frx":08B4
         Left            =   1320
         List            =   "for.frx":08C4
         TabIndex        =   7
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   1920
         Width           =   8295
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7080
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "for.frx":0904
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "for.frx":1556
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "for.frx":21A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "for.frx":2DFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "for.frx":3A4C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "for.frx":469E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "for.frx":52F0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   870
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   1535
         ButtonWidth     =   1032
         ButtonHeight    =   1482
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查询"
               Key             =   "QUE"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "新增"
               Key             =   "ADD"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "MOD"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "DEL"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "EXIT"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "返回"
               Key             =   "RET"
               ImageIndex      =   2
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin FPSpreadADO.fpSpread fpS 
         Height          =   7455
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   3360
         Width           =   17295
         _Version        =   524288
         _ExtentX        =   30506
         _ExtentY        =   13150
         _StockProps     =   64
         DAutoCellTypes  =   0   'False
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
         SpreadDesigner  =   "for.frx":5F42
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread fpS_Clear 
         Height          =   7725
         Left            =   -74880
         TabIndex        =   34
         Top             =   2760
         Width           =   18015
         _Version        =   524288
         _ExtentX        =   31776
         _ExtentY        =   13626
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
         SpreadDesigner  =   "for.frx":6370
      End
      Begin VB.Label Label6 
         Caption         =   "料       号"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   38
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "客户代码"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   36
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lb7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总数量"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   7920
         TabIndex        =   28
         Top             =   2880
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lb6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总金额"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   7920
         TabIndex        =   27
         Top             =   2400
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lb9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         TabIndex        =   24
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label lb8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   23
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label lb5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手册编号"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   19
         Top             =   2400
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lb4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "贸易方式"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手册号码"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74880
         TabIndex        =   15
         Top             =   2520
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择上传的资料:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74400
         TabIndex        =   8
         Top             =   600
         Width           =   2010
      End
      Begin VB.Label lb3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ps:输入多个采购单号时以 - 为间隔符 示例:C120905027-C120905028"
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
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Visible         =   0   'False
         Width           =   7395
      End
      Begin VB.Label lb2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询条件"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   960
      End
      Begin VB.Label lb1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "维护类型"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   960
      End
   End
End
Attribute VB_Name = "Frm_GWZLWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Global variable
Public strval  As String
Public strtet  As String
Public strstate As Boolean
Public stridid   As String
Public strstate1 As Boolean
Public stridid1   As String
'import

Private Enum F_fp
        
    F_gx = 1
    F_no                    '批次
    F_purchaseno            '采购单号 null
    F_partno                '料号
    F_modelno               '型号
    F_modetrade             '类别
    F_orderqty              '订单数量 null
    F_die                   '标准die
    F_totaldie              '总die数
    F_manualno              '手册编号
    F_itemno                '项号
    F_name                  '品名
    F_baoguanqty            '报关量
    F_unit                  '计量单位
    F_indate                '入场日期
    F_invoice               '发票号
    F_caseqty               '件数
    F_currency              '币别
    F_unitprice             '采购单价
    F_baoguanvalue          '报关金额
    F_rate                  '汇率
    F_tariffrate            '关税率
    F_tariff                '关税
    F_addtaxrate            '增值税率
    F_addtax                '增值税
    F_declarationno         '报关单号
    F_awb                   'AWB#
    F_freight               '货代
    F_chargebackdate        '退单日期
    F_mark                  '备注
    F_id                    'id
     
End Enum


'export
Private Enum E_FPS
        
    E_gx = 1
    e_NO                    '批次
    E_exportno              '出货单据
    E_partno                '料号
    E_modetrade             '类别
    e_Invoice               '发票号
    E_exportdate            '出货日期
    E_exportquantity        '出货数量
    E_declarationno         '报关单号
    E_manualno              '手册编号
    E_itemno                '手册项号
    E_name                  '品名
    E_UNIT                  '计量单位
    E_currency              '币别
    E_totalprice            '总价
    E_unitprice             '单价
    E_AWB                   'AWB#
    E_destination           '目的地
    E_freight               '货代
    E_chargebackdate        '退单日期
    E_mark                  '备注
    e_ID                    'id
    E_flienum               '文件编号
End Enum

Private Enum E_FPS0          '
   ' E_CHOOSE = 1
    e_ID = 1
    E_CGDBH = 2 '采购单编号
    E_CGDITEM '采购单序号
    E_PODATE 'PO生效日期
    E_CUSTOMSCLEARDATE 'PO清关日期
  'E_CreateDate '维护日期
  ' E_Createby '维护人员
    E_cust '客户代码
    E_PN '料号
    E_SUPPLIERNAME '供应商名称
    E_SUPPLIERCODE '供应商编号
    e_device 'Device
    E_POqty '数量
    E_Entryqty '数量
    E_Lastqty '数量
    E_END
    
End Enum



Private Sub cmdExport_Click()
Call ExportExcel(fpS_Clear)
End Sub

Private Sub Combo1_Click()

    Select Case Combo1.text
            
        Case "出口明细表"
        
            Command4.Visible = False
            
            Command5.Visible = False
            
            lb2.Visible = True
            
            Text1.Visible = True
        
            lb2 = "发票单号"
            
            lb3.Visible = True
            
            lb3 = "ps:输入多个发票号编号时以/为间隔符 示例:S1902200012/S1902200013"
            
            lb4.Visible = True
            
            comBo2.Visible = True
            
            comBo2.Clear
            
            comBo2.AddItem ("进料对口")
            comBo2.AddItem ("一般贸易")
            comBo2.AddItem ("其他进出口免费")
            comBo2.AddItem ("进料料件复出")
            comBo2.AddItem ("进料成品退换")
            comBo2.AddItem ("修理物品")
            comBo2.AddItem ("设备退运")
            comBo2.AddItem ("其他")
              
            lb5.Visible = False
            
            Combo3.Visible = False
            
            fpss(0).Visible = False
            
            fpS(0).MaxRows = 0
            fpS(0).MaxCols = 0
            fpss(0).MaxRows = 0
            fpss(0).MaxCols = 0
            
        Case "出口明细表(特殊)"
        
            Command4.Visible = False
            
            Command5.Visible = False
            
            lb2.Visible = True
            
            Text1.Visible = True
            
            lb2 = "发票单号"
            
            lb3.Visible = False
            
            lb4.Visible = True
            
            comBo2.Visible = True
            
            comBo2.Clear
            
            comBo2.AddItem ("进料对口")
            comBo2.AddItem ("一般贸易")
            comBo2.AddItem ("其他进出口免费")
            comBo2.AddItem ("进料料件复出")
            comBo2.AddItem ("进料成品退换")
            comBo2.AddItem ("修理物品")
            comBo2.AddItem ("设备退运")
            comBo2.AddItem ("其他")
              
            lb5.Visible = False
            
            Combo3.Visible = False
            
            fpss(0).Visible = False
            
            fpS(0).MaxRows = 0
            fpS(0).MaxCols = 0
            fpss(0).MaxRows = 0
            fpss(0).MaxCols = 0
            
        Case "进口明细表"
            
            Command4.Visible = False
            
            Command5.Visible = False
            
            lb2.Visible = True
            
            Text1.Visible = True
        
            lb2 = "采购单号"
            
            lb3.Visible = True
            
            lb3 = "ps:输入多个采购单号时以/为间隔符 示例:C120905027/C120905028"
            
            lb4.Visible = False
            
            comBo2.Visible = False
            
            lb5.Visible = False
            
            Combo3.Visible = False
            
            fpS(0).MaxRows = 0
            fpS(0).MaxCols = 0
            fpss(0).MaxRows = 0
            fpss(0).MaxCols = 0
            
        Case "进口明细表(特殊)"
            
            Command4.Visible = False
            
            Command5.Visible = False
        
            lb2.Visible = False
            
            Text1.Visible = False
            Text1.text = ""
            
            lb3.Visible = False
            
            lb4.Visible = False
            
            comBo2.Visible = False
            
            lb5.Visible = False
            
            Combo3.Visible = False
            
            fpS(0).MaxRows = 0
            fpS(0).MaxCols = 0
            fpss(0).MaxRows = 0
            fpss(0).MaxCols = 0
        

    End Select

End Sub

Private Sub Combo2_Click()

    Dim rs     As New ADODB.Recordset

    Dim strsql As String
    
    Dim j      As Integer
        
    If comBo2.text = "" Then
    
        MsgBox "请选择贸易方式", vbInformation, "提示"
        Exit Sub
    
    End If
    
    If comBo2.text = "进料对口" Or comBo2.text = "进料成品退换" Or comBo2.text = "进料料件复出" Then
    
        lb5.Visible = True
            
        Combo3.Visible = True
            
        Combo3.Clear
            
        strsql = "select distinct 手册编号 from erptemp.dbo.ksmanual where 1 = 1 and flag = '2'"

        If rs.State = 1 Then rs.Close
        rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

        rs.MoveFirst

        For j = 1 To rs.RecordCount

            Combo3.AddItem (rs("手册编号"))
            rs.MoveNext
                
        Next

        rs.Clone

        Set rs = Nothing

    Else
            
        lb5.Visible = False
            
        Combo3.Visible = False
            
        Combo3.text = ""

    End If

End Sub

Private Sub Command1_Click()

    Dim FName As String

    CommonDialog1.Filter = "EXCEL文件(*.xlsx)|*.xlsx|EXCEL文件(*.xls)|*.xls"
    CommonDialog1.ShowOpen
    '得到文件名
    FName = CommonDialog1.filename

    If FName <> "" Then
    
        Text2.text = FName

    End If

End Sub

Private Sub Command2_Click()

    Dim i           As Integer

    Dim j           As Integer

    Dim tempVal     As String

    Dim temp1       As String

    Dim temp2       As String

    Dim temp3       As String

    Dim temp4       As String

    Dim temp5       As String

    Dim temp6       As String

    Dim strChar     As String
    
    Dim SumCount    As Integer
    
    Dim SumDelCount As Integer
    
    Dim VBExcel     As Excel.Application

    Dim xlBook      As Excel.Workbook

    Dim xlSheet     As Excel.Worksheet
    
    If Text2.text = "" Then
        MsgBox "先选择待上传的文件"
        Exit Sub

    End If
    
    SumCount = 0
    SumDelCount = 0

    'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text2.text)    '打开文件

    Set xlSheet = xlBook.Worksheets("sheet1")        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 6 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    For i = 1 To xlSheet.Range("A1").CurrentRegion.Rows.count
   
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
        
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
        
            If j = 1 Then
            
                temp1 = tempVal
            
            End If
        
            If j = 2 Then
                temp2 = tempVal
        
            End If

            If j = 3 Then

                temp3 = tempVal

            End If

            If j = 4 Then
                temp4 = tempVal

            End If

            If j = 5 Then

                temp5 = tempVal

            End If

            If j = 6 Then
                temp6 = tempVal

            End If
        
        Next j
        
        If Get_SqlserverCnt("select * from erptemp.dbo.ksmanual where 手册编号 = '" & temp1 & "' and flag = '" & temp2 & "' and 序号 = '" & temp3 & "' ") <> 0 Then

            AddSql2 ("DELETE FROM erptemp.dbo.ksmanual where 手册编号 = '" & temp1 & "' and flag = '" & temp2 & "' and 序号 = '" & temp3 & "'")

            SumDelCount = SumDelCount + 1

        End If

        AddSql2 ("insert into erptemp.dbo.ksmanual values('" & temp1 & "','" & temp2 & "','" & temp3 & "','" & temp4 & "','" & temp5 & "','" & temp6 & "')")
    
        
        SumCount = SumCount + 1
        
    Next i

    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing
    
    If SumCount > 0 Then
    
        MsgBox "已成功上传" & SumCount & "笔！", , "友情提醒"
    Else

        MsgBox "无资料上传成功", vbInformation, "提示"
    
    End If
    
    If SumDelCount > 0 Then

        MsgBox "覆盖数据库资料" & SumDelCount & "笔！", , "友情提醒"

    End If

End Sub

Private Sub Command3_Click()

    Dim aflag     As String
    
    Dim strmanual As String
    
    Dim strsql    As String
    
    Dim rs        As New ADODB.Recordset
    
    strmanual = Trim$(Text3.text)
    
    If Text3.text <> "" Then
    
        strsql = "select 手册编号,case flag when '1' then '料件表' when '2' then '成品表' end as 类型,序号,商品名称,计量单位,申报数量 from erptemp.dbo.ksmanual where 手册编号 = '" & strmanual & "' order by 手册编号,flag,序号"
    
    Else
    
        strsql = "select 手册编号,case flag when '1' then '料件表' when '2' then '成品表' end as 类型,序号,商品名称,计量单位,申报数量 from erptemp.dbo.ksmanual where 1 = 1 order by 手册编号,flag,序号 "

    End If

    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType2(rs)
    Else
        
        MsgBox "查询不到该手册信息", vbInformation, "提示"
        Exit Sub

    End If
    
End Sub

Private Sub Command4_Click()

    Dim stridd As String
    
    Dim i      As Integer
    
    Dim j      As Integer
        
    Dim rs     As New ADODB.Recordset
       
    Dim strsql As String
    
    stridd = Createid
    
    Select Case Combo1.text
    
        Case "出口明细表(特殊)"
    
            With fpS(0)

                .MaxRows = .MaxRows + 1
                i = .MaxRows
        
                .Row = i
                .Col = E_FPS.e_NO
                .text = stridd
        
                .Col = E_FPS.e_Invoice
                .text = Trim$(Text1.text)
        
                .Col = E_FPS.E_modetrade
                .text = Trim$(comBo2.text)
                .CellType = CellTypeComboBox
                       
                .TypeComboBoxList = .TypeComboBoxList & "进料对口"
            
                .TypeComboBoxList = .TypeComboBoxList & "一般贸易"
            
                .TypeComboBoxList = .TypeComboBoxList & "其他进出口免费"
            
                .TypeComboBoxList = .TypeComboBoxList & "进料料件复出"
            
                .TypeComboBoxList = .TypeComboBoxList & "进料成品退换"
            
                .TypeComboBoxList = .TypeComboBoxList & "修理物品"
            
                .TypeComboBoxList = .TypeComboBoxList & "设备退运"
            
                .TypeComboBoxList = .TypeComboBoxList & "其他"
                
                .Col = E_FPS.E_currency
                .SetText E_FPS.E_currency, i, "USD"
                .CellType = CellTypeComboBox
                
                .TypeComboBoxList = "USD"
            
                .TypeComboBoxList = .TypeComboBoxList & "JPY"

                .TypeComboBoxList = .TypeComboBoxList & "EUR"

                .TypeComboBoxList = .TypeComboBoxList & "RMB"
                
                .Col = E_FPS.E_manualno
                .text = Trim$(Combo3.text)
                
                .Col = E_FPS.E_exportno
                .Lock = True
                
                .Col = E_FPS.E_chargebackdate
                .Lock = True
                
                .Col = E_FPS.E_mark
                .Lock = True
                
                .Col = E_FPS.e_ID
                .Lock = True
                
                .Col = E_FPS.E_unitprice
                .Lock = True
                
                .LockBackColor = vbYellow
                
                .Col = E_FPS.E_gx

                If .text = 1 Then
            
                    .Col = E_FPS.E_exportquantity
                
                    strtet = Val(.text) + Val(strtet)
                
                    .Col = E_FPS.E_totalprice
                
                    strval = Val(.text) + Val(strval)
            
                End If
                
                strtet = Format(Trim$(strtet), "0.000")
            
                strval = Format(Trim$(strval), "0.0000")
                
                Text4.text = strval
    
                Text5.text = strtet
        
            End With
            
        Case "进口明细表(特殊)"

            With fpS(0)

                .MaxRows = .MaxRows + 1
                i = .MaxRows
        
                .Row = i
                .Col = F_fp.F_no
                .text = stridd
                
                .Row = i
                .Col = F_fp.F_modetrade
                .CellType = CellTypeComboBox
                       
                .TypeComboBoxList = .TypeComboBoxList & "进料对口"
            
                .TypeComboBoxList = .TypeComboBoxList & "一般贸易"
            
                .TypeComboBoxList = .TypeComboBoxList & "其他进出口免费"
            
                .TypeComboBoxList = .TypeComboBoxList & "成品复进"
            
                .TypeComboBoxList = .TypeComboBoxList & "维修物品"
            
                .TypeComboBoxList = .TypeComboBoxList & "料件复进"
            
                .TypeComboBoxList = .TypeComboBoxList & "进料成品退换"
            
                .TypeComboBoxList = .TypeComboBoxList & "其他"
                
                .Row = i

                For j = F_fp.F_rate To F_fp.F_addtax
            
                    .Col = j
                    .Lock = True
        
                Next
                
                .Col = F_fp.F_currency
                .SetText F_fp.F_currency, i, "USD"
                .CellType = CellTypeComboBox
                
                .TypeComboBoxList = .TypeComboBoxList & "USD"
            
                .TypeComboBoxList = .TypeComboBoxList & "JPY"

                .TypeComboBoxList = .TypeComboBoxList & "EUR"

                .TypeComboBoxList = .TypeComboBoxList & "RMB"
                
                .LockBackColor = vbYellow
                
                .Col = F_fp.F_gx

                If .text = 1 Then
            
                    .Col = F_fp.F_baoguanqty
                
                    strtet = Val(.text) + Val(strtet)
                
                    .Col = F_fp.F_baoguanvalue
                
                    strval = Val(.text) + Val(strval)
            
                End If
                  
                strtet = Format(Trim$(strtet), "0.000")
            
                strval = Format(Trim$(strval), "0.0000")
        
                Text4.text = strval
    
                Text5.text = strtet
                
                strsql = "select distinct 手册编号 from erptemp.dbo.ksmanual where 1 = 1 and flag = '1'"

                If rs.State = 1 Then rs.Close
                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                .Col = F_fp.F_manualno
                .ColWidth(F_fp.F_manualno) = 12
                .CellType = CellTypeComboBox

                rs.MoveFirst

                For i = 1 To rs.RecordCount

                    .TypeComboBoxList = .TypeComboBoxList & rs("手册编号")
                    rs.MoveNext
                Next
        
                rs.Clone
        
                Set rs = Nothing

            End With

    End Select

End Sub

Private Sub Command5_Click()

    Dim j      As Integer
    
    Dim strsum As Integer
    
    Dim strsum1 As Integer
    
    strsum = 0
    strtet = 0
    
    strval = 0
    
    With fpS(0)
        
        If .MaxRows > 0 Then
        
            '            .DeleteRows .ActiveRow, 1
             strsum1 = .MaxRows
            
            For j = 1 To strsum1
            
                .Row = j
            
                .Col = 1
                If .text <> 1 Then
            
                    .DeleteRows j, 1
            
                    strsum = strsum + 1
                    
                    j = j - 1

                End If
            
            Next

            If strsum = 0 Then
            
                MsgBox "没有需要删除的行", vbInformation, "提示"
            
                Exit Sub
            
            End If
            
            .MaxRows = strsum1 - strsum
            '            .MaxRows = .MaxRows - 1

            For j = 1 To .MaxRows
            
                Select Case Combo1.text
            
                    Case "出口明细表(特殊)"
                    
                        .Row = j
                        .Col = E_FPS.E_gx

                        If .text = 1 Then
            
                            .Col = E_FPS.E_exportquantity
                
                            strtet = Val(.text) + Val(strtet)
                
                            .Col = E_FPS.E_totalprice
                
                            strval = Val(.text) + Val(strval)
            
                        End If
                  
                        strtet = Format(Trim$(strtet), "0.000")
            
                        strval = Format(Trim$(strval), "0.0000")
        
                        Text4.text = strval
    
                        Text5.text = strtet
                
                    Case "进口明细表(特殊)"
                        .Row = j
                        .Col = F_fp.F_gx

                        If .text = 1 Then
            
                            .Col = F_fp.F_baoguanqty
                
                            strtet = Val(.text) + Val(strtet)
                
                            .Col = F_fp.F_baoguanvalue
                
                            strval = Val(.text) + Val(strval)
            
                        End If
                  
                        strtet = Format(Trim$(strtet), "0.000")
            
                        strval = Format(Trim$(strval), "0.0000")
        
                        Text4.text = strval
    
                        Text5.text = strtet
            
                End Select

            Next
            
        Else
            
            MsgBox "已经无资料可删除，请确认", vbInformation, "提示"
            
            strtet = 0
            
            strval = 0
            
            strtet = Format(Trim$(strtet), "0.000")
            
            strval = Format(Trim$(strval), "0.0000")
        
            Text4.text = strval
    
            Text5.text = strtet
            
            Exit Sub
        
        End If
         
    End With
    
End Sub

Private Sub Form_Load()

    With fpS(0)
    
        .Col = -1
        .Row = -1
        .Lock = True

    End With
    
    With fpss(0)
    
        .Col = -1
        .Row = -1
        .Lock = True

    End With
    
    With fpsss(0)
    
        .Col = -1
        .Row = -1
        .Lock = True

    End With
    
    DTPicker1.Value = DateTime.DATE
    DTPicker2.Value = DateTime.DATE

End Sub

Private Sub fps_Change(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    
    Dim rs       As New ADODB.Recordset
    
    Dim strbaog  As String
    
    Dim strsql   As String

    Dim i        As Integer
    
    Dim stritem  As String
    
    Dim strflag  As Integer
    
    Dim strsty   As String
    
    Dim strno    As String
    
    Dim strInv1  As String
    
    Dim strInv2  As String
    
    Dim strInv5  As String
    
    Dim strNo1   As String

    Dim strNo2   As String

    Dim strNo3   As String
    
    Dim strcool2 As String
    
    Dim strcool1 As String
    
    Dim strdie   As String
    
    Dim strtdie  As String
    
    Dim strunitq As String
    
    Dim strunitv As String

    'Me.MousePointer = 11
    
    If Combo1.text = "进口明细表" Or Combo1.text = "进口明细表(特殊)" Then
        If Text1.text <> "" Then
            fpscopy '自动带出下几列函数 ZYF 20200331
        End If
    End If
    
    Select Case Combo1.text
   
        Case "进口明细表"
            
            strflag = 1

            If Col <> 11 And Col <> 6 And Col <> 1 And Col <> 20 And Col <> 7 And Col <> 8 And Col <> 13 Then Exit Sub
            
            If Index = 0 Then
            
                If Col = 1 Then
                 
                    With fpS(0)

                        .Row = Row

                        .Col = 1

                        If .text <> 1 Then

                            .Col = 13

                            strtet = Val(strtet) - Val(.text)

                            strtet = Format(Trim$(strtet), "0.000")

                            .Col = 20

                            strval = Val(strval) - Val(.text)
                            strval = Format(Trim$(strval), "0.0000")

                        End If

                    End With
                    
                    With fpS(0)

                        .Row = Row

                        .Col = 1

                        If .text = 1 Then

                            .Col = 13

                            strtet = Val(strtet) + Val(.text)

                            strtet = Format(Trim$(strtet), "0.000")

                            .Col = 20

                            strval = Val(strval) + Val(.text)
                            strval = Format(Trim$(strval), "0.0000")

                        End If

                    End With

                    Text4.text = strval

                    Text5.text = strtet
                    
                End If
    
                If Col = 6 Then
            
                    With fpS(0)
                        .Row = Row
                        '类别里只有进料对口&成品复进两种情况的才会输入手册号，其他类别均无需选择手册号与项号
                        .Col = 3
                        
                        strcool1 = Trim$(.text)
                        
                        .Col = 6

                        If .text = "" Then
                
                            MsgBox "请输入类别", vbInformation, "提示"
                            Exit Sub

                        End If

                        strsty = Trim$(.text)
                        
                        For i = 1 To .MaxRows
                            
                            .Row = i
                            
                            .Col = 3
                            
                            strcool2 = Trim$(.text)
                            
                            If strcool2 = strcool1 Then
                                
                                .Row = i
                                
                                .Col = 6
    
                                .text = strsty
        
                                If strsty = "进料对口" Or strsty = "成品复进" Then
                
                                    .Col = 10
                                    .Lock = False
                    
                                    .Col = 11
                                    .Lock = False
                        
                                    .Col = 12
                                    .Lock = True
                
                                    .Col = 14
                                    .Lock = True
                        
                                    .Col = 21
                                    .Lock = True
                                    .SetText 21, Row, Trim$("")
                    
                                    .Col = 22
                                    .Lock = True
                                    .SetText 22, Row, Trim$("")
                    
                                    .Col = 23
                                    .Lock = True
                                    .SetText 23, Row, Trim$("")
                    
                                    .Col = 24
                                    .Lock = True
                                    .SetText 24, Row, Trim$("")
                    
                                    .Col = 25
                                    .Lock = True
                                    .SetText 25, Row, Trim$("")
                                Else
                    
                                    .Col = 10
                                    .Lock = True
                                    .SetText 10, Row, Trim$("")
                    
                                    .Col = 11
                                    .Lock = True
                                    .SetText 11, Row, Trim$("")
                    
                                    .Col = 12
                                    .text = ""
                                    .Lock = False
                
                                    .Col = 14
                                    .text = ""
                                    .Lock = False
                    
                                    .Col = 21
                                    .Lock = False
                    
                                    .Col = 22
                                    .Lock = False
                    
                                End If

                            End If

                        Next

                    End With
        
                End If
                
                If Col = 7 Then
                    
                    With fpS(0)
                        
                        .Row = Row
                        
                        .Col = 3
                        
                        strInv1 = Trim(.text)
                        
                        .Col = 4
                        
                        strInv2 = Trim(.text)
                    
                        .Col = 7
                        
                        .text = Format(Trim$(.text), "0.00")
                        
                        strInv5 = Trim(.text)
                
                        strNo1 = Get_SqlStr("SELECT isnull(SUM(a.批准采购数量),0) FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b WHERE a.采购单编号 = '" & strInv1 & "' and a.物料编号 = b.物料编号 and b.料号 = '" & strInv2 & "' ")
                
                        strNo2 = Get_SqlStr("SELECT isnull(SUM(订单数量),0) FROM erptemp.dbo.ksimport where 采购单号 = '" & strInv1 & "' and 料号 = '" & strInv2 & "' and flag = '0'")
                
                        strNo3 = Val(strNo1) - Val(strNo2)
                
                        If Val(strInv5) > Val(strNo3) Then
                        
                            MsgBox "该笔料号" & strInv2 & "批准采购数量: " & strNo1 & ",已经维护订单数量：" & strNo2 & ",最大数量只能维护：" & strNo3 & "", vbInformation, "提示"
                            Exit Sub

                        End If
                
                        If Val(strInv5) <= 0 Then
                
                            MsgBox "订单数量不可小于等于0", vbInformation, "提示"
                            Exit Sub

                        End If
                       
                        .Col = 8
                        
                        strdie = Format(Trim$(.text), "0.00")
                        
                        .Col = 9
                        
                        strtdie = Val(strdie) * Val(strInv5)
                        
                        .text = Format(Trim$(strtdie), "0.00")
                    
                    End With
                    
                End If
                
                If Col = 13 Then
                      
                    With fpS(0)
                    
                        strtet = 0

                        For i = 1 To .MaxRows
                            
                            .Row = i
                            
                            .Col = 1
                            
                            If .text = 1 Then
                                
                                .Col = 13
                            
                                strtet = Val(strtet) + Val(.text)
                                
                            End If
                            
                        Next
                    
                    End With
                            
                    strtet = Format(Trim$(strtet), "0.000")
                    
                    Text5.text = strtet
                        
                End If
                
                If Col = 8 Then
                    
                    With fpS(0)
                        
                        .Row = Row
                        
                        .Col = 7
                        
                        strInv5 = Format(Trim$(.text), "0.00")
                        
                        .Col = 8
                        
                        strdie = Format(Trim$(.text), "0.00")
                        
                        .Col = 9
                        
                        strtdie = Val(strdie) * Val(strInv5)
                        
                        .text = Format(Trim$(strtdie), "0.00")
                        
                    End With
                        
                End If
        
                If Col = 11 Then
    
                    With fpS(0)
                        .Row = Row
                        .Col = 6
                        strsty = Trim(.text)
                
                        '只有选择了手册号才可以出现选择项号的功能，才能带出品名，否则就需要手工输入品名与单位
                
                        If strsty = "进料对口" Or strsty = "成品复进" Then
                
                            .Col = 10
                            stritem = Trim(.text)

                            If stritem = "" Then
                
                                MsgBox "请输入手册编号", vbInformation, "提示"
                                Exit Sub

                            End If

                            .Col = 11

                            If Trim$(.text) <> "" Then
                            
                                If Get_SqlserverCnt("SELECT 商品名称 as 品名,计量单位 " & " FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & stritem & "' and flag = '" & strflag & "' and 序号= '" & Trim$(.text) & "'") = 0 Then
                                    
                                    MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"
                                    
                                    .SetText 12, Row, Trim$("")
                                    .SetText 14, Row, Trim$("")
                                    
                                    Exit Sub
                                
                                End If

                                strsql = "SELECT 商品名称 as 品名,计量单位 " & " FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & stritem & "' and flag = '" & strflag & "' and 序号= '" & Trim$(.text) & "'"
                                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                                If Not rs.EOF Then

                                    With fpS(0)
                                        .SetText 12, Row, Trim$("" & rs!品名)
                                        .SetText 14, Row, Trim$("" & rs!计量单位)

                                    End With

                                End If

                                rs.Close

                            End If

                        End If

                    End With
           
                End If
                
                If Col = 20 Then
                    
                    With fpS(0)
                        
                        strval = 0
                    
                        For i = 1 To .MaxRows
                            
                            .Row = i
                            
                            .Col = 1
                            
                            If .text = 1 Then
                                
                                .Col = 20
                            
                                strval = Val(strval) + Val(.text)
                                
                            End If
                            
                        Next
                    
                    End With
                            
                    strval = Format(Trim$(strval), "0.0000")
                    
                    Text4.text = strval
                    
                End If

            End If
    
        Case "出口明细表"
        
            strflag = 2
                
            '进料对口/进料成品退换/进料料件复出
            
            If Col <> 11 And Col <> 5 And Col <> 15 And Col <> 9 And Col <> 17 And Col <> 18 And Col <> 19 And Col <> 20 And Col <> 1 Then Exit Sub

            If Index = 0 Then
                 
                If Col = 1 Then
                 
                    With fpS(0)

                        .Row = Row

                        .Col = 1

                        If .text <> 1 Then

                            .Col = 8

                            strtet = Val(strtet) - Val(.text)

                            strtet = Format(Trim$(strtet), "0.000")

                            .Col = 15

                            strval = Val(strval) - Val(.text)
                            strval = Format(Trim$(strval), "0.0000")

                        End If

                    End With
                    
                    With fpS(0)

                        .Row = Row

                        .Col = 1

                        If .text = 1 Then

                            .Col = 8

                            strtet = Val(strtet) + Val(.text)

                            strtet = Format(Trim$(strtet), "0.000")

                            .Col = 15

                            strval = Val(strval) + Val(.text)
                            strval = Format(Trim$(strval), "0.0000")

                        End If

                    End With

                    Text4.text = strval

                    Text5.text = strtet
                    
                End If
                 
                If Col = 9 Or Col = 17 Or Col = 18 Or Col = 19 Or Col = 20 Then
                 
                    With fpS(0)
                    
                        .Row = Row
                        .Col = Col
                        strbaog = Trim$(.text)
                        
                        For i = 1 To .MaxRows
                        
                            .Row = i
                            .Col = Col
                            .text = strbaog
                            
                        Next
                    
                    End With
                    
                End If
 
                If Col = 5 Then

                    With fpS(0)
                        .Row = Row
                        .Col = 5

                        If .text = "" Then

                            MsgBox "请输入类别", vbInformation, "提示"
                            Exit Sub

                        End If

                        strsty = Trim(.text)

                        If strsty = "进料对口" Or strsty = "进料成品退换" Or strsty = "进料料件复出" Then

                            .Col = 10
                            .Lock = False

                            .Col = 11
                            .Lock = False

                            .Col = 12
                            .Lock = True
                            .SetText 12, Row, Trim$("")

                            .Col = 13
                            .Lock = True
                            .SetText 13, Row, Trim$("")

                        Else

                            .Col = 10
                            .Lock = True
                            .SetText 10, Row, Trim$("")

                            .Col = 11
                            .Lock = True
                            .SetText 11, Row, Trim$("")

                            .Col = 12
                            .text = ""
                            .Lock = False

                            .Col = 13
                            .text = ""
                            .Lock = False

                        End If

                    End With

                End If

                If Col = 11 Then
    
                    With fpS(0)
                        .Row = Row
                        .Col = 5
                        strsty = Trim(.text)
                
                        '只有选择了手册号才可以出现选择项号的功能，才能带出品名，否则就需要手工输入品名与单位
                
                        If strsty = "进料对口" Or strsty = "进料成品退换" Or strsty = "进料料件复出" Then
                
                            .Col = 10
                            stritem = Trim(.text)

                            .Col = 11

                            If Trim$(.text) <> "" Then
                                    
                                If strsty = "进料对口" Or strsty = "进料成品退换" Then
                                
                                    If Get_SqlserverCnt("SELECT 商品名称 as 品名,计量单位 " & " FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & stritem & "' and flag = '" & strflag & "' and 序号= '" & Trim$(.text) & "'") = 0 Then
                                    
                                        MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"
                                    
                                        .SetText 12, Row, Trim$("")
                                        .SetText 13, Row, Trim$("")
                                    
                                        Exit Sub
                                
                                    End If

                                    strsql = "SELECT 商品名称 as 品名,计量单位 " & " FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & stritem & "' and flag = '" & strflag & "' and 序号= '" & Trim$(.text) & "'"
                                Else

                                    If Get_SqlserverCnt("SELECT 商品名称 as 品名,计量单位 " & " FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & stritem & "' and flag = '1' and 序号= '" & Trim$(.text) & "'") = 0 Then
                                    
                                        MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"
                                    
                                        .SetText 12, Row, Trim$("")
                                        .SetText 13, Row, Trim$("")
                                    
                                        Exit Sub
                                
                                    End If

                                    strsql = "SELECT 商品名称 as 品名,计量单位 " & " FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & stritem & "' and flag = '1' and 序号= '" & Trim$(.text) & "'"

                                End If

                                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                                If Not rs.EOF Then

                                    With fpS(0)
                                    
                                        .SetText 12, Row, Trim$("" & rs!品名)
                                        .SetText 13, Row, Trim$("" & rs!计量单位)

                                    End With

                                End If

                                rs.Close

                            End If

                        End If

                    End With
           
                End If

                If Col = 15 Then
                    
                    With fpS(0)
                    
                        .Row = Row
                        
                        .Col = 8
                        strno = Trim$(.text)
                    
                        .Col = 15

                        If Trim$(.text) = "" Then
                            MsgBox "请输入总价", vbInformation, "提示"
                            Exit Sub
                                
                        Else
                            strNo1 = Trim$(.text)
                            
                            strNo2 = Val(strNo1) / Val(strno)
                            
                            strNo3 = Format(Trim$(strNo2), "0.000000")
                    
                            .SetText 16, Row, Trim$("" & strNo3)

                        End If

                    End With
                    
                    With fpS(0)
                        
                        strval = 0
                    
                        For i = 1 To .MaxRows
                            
                            .Row = i
                            
                            .Col = 1
                            
                            If .text = 1 Then
                                
                                .Col = 15
                            
                                strval = Val(strval) + Val(.text)
                                
                            End If
                            
                        Next
                    
                    End With
                    
                    strval = Format(Trim$(strval), "0.0000")
                            
                    Text4.text = strval
    
                End If
                
            End If
            
        Case "进口明细表(特殊)"
            
            strflag = 1

            If Col <> F_fp.F_baoguanqty And Col <> F_fp.F_modetrade And Col <> F_fp.F_gx And Col <> F_fp.F_itemno And Col <> F_fp.F_baoguanvalue Then Exit Sub
            
            If Index = 0 Then
            
                If Col = F_fp.F_gx Then
                 
                    With fpS(0)

                        .Row = Row

                        .Col = F_fp.F_gx

                        If .text <> 1 Then

                            .Col = F_fp.F_baoguanqty

                            strtet = Val(strtet) - Val(.text)

                            strtet = Format(Trim$(strtet), "0.000")

                            .Col = F_fp.F_baoguanvalue

                            strval = Val(strval) - Val(.text)
                            strval = Format(Trim$(strval), "0.0000")

                        End If

                    End With
                    
                    With fpS(0)

                        .Row = Row

                        .Col = 1

                        If .text = 1 Then

                            .Col = F_fp.F_baoguanqty

                            strtet = Val(strtet) + Val(.text)

                            strtet = Format(Trim$(strtet), "0.000")

                            .Col = F_fp.F_baoguanvalue

                            strval = Val(strval) + Val(.text)
                            strval = Format(Trim$(strval), "0.0000")

                        End If

                    End With

                    Text4.text = strval

                    Text5.text = strtet
                    
                End If
    
                If Col = F_fp.F_modetrade Then
            
                    With fpS(0)
                        .Row = Row
                        
                        .Col = F_fp.F_modetrade
                        
                        If .text = "" Then
                
                            MsgBox "请输入类别", vbInformation, "提示"
                            Exit Sub

                        End If

                        strsty = Trim$(.text)
        
                        If strsty = "进料对口" Or strsty = "成品复进" Then
                
                            .Col = F_fp.F_manualno
                            .Lock = False
                    
                            .Col = F_fp.F_itemno
                            .Lock = False
                        
                            .Col = F_fp.F_name
                            .Lock = True
                
                            .Col = F_fp.F_unit
                            .Lock = True
                        
                            .Col = F_fp.F_rate
                            .Lock = True
                            .SetText F_fp.F_rate, Row, Trim$("")
                    
                            .Col = F_fp.F_tariffrate
                            .Lock = True
                            .SetText F_fp.F_tariffrate, Row, Trim$("")
                    
                            .Col = F_fp.F_tariff
                            .Lock = True
                            .SetText F_fp.F_tariff, Row, Trim$("")
                    
                            .Col = F_fp.F_addtaxrate
                            .Lock = True
                            .SetText F_fp.F_addtaxrate, Row, Trim$("")
                    
                            .Col = F_fp.F_addtax
                            .Lock = True
                            .SetText F_fp.F_addtax, Row, Trim$("")
                        Else
                    
                            .Col = F_fp.F_manualno
                            .Lock = True
                            .SetText F_fp.F_manualno, Row, Trim$("")
                    
                            .Col = F_fp.F_itemno
                            .Lock = True
                            .SetText F_fp.F_itemno, Row, Trim$("")
                    
                            .Col = F_fp.F_name
                            .text = ""
                            .Lock = False
                
                            .Col = F_fp.F_unit
                            .text = ""
                            .Lock = False
                    
                            .Col = F_fp.F_rate
                            .Lock = False
                    
                            .Col = F_fp.F_tariffrate
                            .Lock = False
                    
                        End If

                    End With
        
                End If
                
                If Col = F_fp.F_baoguanqty Then
                      
                    With fpS(0)
                    
                        strtet = 0

                        For i = 1 To .MaxRows
                            
                            .Row = i
                            
                            .Col = 1
                            
                            If .text = 1 Then
                                
                                .Col = F_fp.F_baoguanqty
                            
                                strtet = Val(strtet) + Val(.text)
                                
                            End If
                            
                        Next
                    
                    End With
                            
                    strtet = Format(Trim$(strtet), "0.000")
                    
                    Text5.text = strtet
                        
                End If
                        
                If Col = F_fp.F_itemno Then
    
                    With fpS(0)
                        .Row = Row
                        .Col = F_fp.F_modetrade
                        strsty = Trim(.text)
                
                        If strsty = "进料对口" Or strsty = "成品复进" Then
                
                            .Col = F_fp.F_manualno
                            stritem = Trim(.text)

                            If stritem = "" Then
                
                                MsgBox "请输入手册编号", vbInformation, "提示"
                                Exit Sub

                            End If

                            .Col = F_fp.F_itemno

                            If Trim$(.text) <> "" Then
                            
                                If Get_SqlserverCnt("SELECT 商品名称 as 品名,计量单位 " & " FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & stritem & "' and flag = '" & strflag & "' and 序号= '" & Trim$(.text) & "'") = 0 Then
                                    
                                    MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"
                                    
                                    .SetText F_fp.F_name, Row, Trim$("")
                                    .SetText F_fp.F_unit, Row, Trim$("")
                                    
                                    Exit Sub
                                
                                End If

                                strsql = "SELECT 商品名称 as 品名,计量单位 " & " FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & stritem & "' and flag = '" & strflag & "' and 序号= '" & Trim$(.text) & "'"
                                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                                If Not rs.EOF Then

                                    With fpS(0)
                                        .SetText F_fp.F_name, Row, Trim$("" & rs!品名)
                                        .SetText F_fp.F_unit, Row, Trim$("" & rs!计量单位)

                                    End With

                                End If

                                rs.Close

                            End If

                        End If

                    End With
           
                End If
                
                If Col = F_fp.F_baoguanvalue Then

                    With fpS(0)
                    
                        .Row = Row
                        .Col = F_fp.F_baoguanqty

                        If .text = "" Then
                        
                            MsgBox "请输入报关数量", vbInformation, "提示"
                            Exit Sub
                        
                        End If

                        strunitq = Trim$(.text)
                        
                        .Col = F_fp.F_baoguanvalue
                        strunitv = Trim$(.text)
                        
                        .Col = F_fp.F_unitprice
                        .text = Format(Trim$(Val(strunitv) / Val(strunitq)), "0.000")
                    
                    End With

                    With fpS(0)
                        
                        strval = 0
                    
                        For i = 1 To .MaxRows
                            
                            .Row = i
                            
                            .Col = 1
                            
                            If .text = 1 Then
                                
                                .Col = F_fp.F_baoguanvalue
                            
                                strval = Val(strval) + Val(.text)
                                
                            End If
                            
                        Next
                    
                    End With
                            
                    strval = Format(Trim$(strval), "0.0000")
                    
                    Text4.text = strval
                    
                End If

            End If
            
        Case "出口明细表(特殊)"
        
            strflag = 2
                
            '进料对口/进料成品退换/进料料件复出
            
            If Col <> E_FPS.E_exportquantity And Col <> E_FPS.E_itemno And Col <> E_FPS.E_modetrade And Col <> E_FPS.E_totalprice And Col <> E_FPS.E_declarationno And Col <> E_FPS.E_AWB And Col <> E_FPS.E_destination And Col <> E_FPS.E_freight And Col <> E_FPS.E_gx Then Exit Sub

            If Index = 0 Then
                 
                If Col = E_FPS.E_gx Then
                 
                    With fpS(0)

                        .Row = Row

                        .Col = E_FPS.E_gx

                        If .text <> 1 Then

                            .Col = E_FPS.E_exportquantity

                            strtet = Val(strtet) - Val(.text)

                            strtet = Format(Trim$(strtet), "0.000")

                            .Col = E_FPS.E_totalprice

                            strval = Val(strval) - Val(.text)
                            strval = Format(Trim$(strval), "0.0000")

                        End If

                    End With
                    
                    With fpS(0)

                        .Row = Row

                        .Col = E_FPS.E_gx

                        If .text = 1 Then

                            .Col = E_FPS.E_exportquantity

                            strtet = Val(strtet) + Val(.text)

                            strtet = Format(Trim$(strtet), "0.000")

                            .Col = E_FPS.E_totalprice

                            strval = Val(strval) + Val(.text)
                            strval = Format(Trim$(strval), "0.0000")

                        End If

                    End With

                    Text4.text = strval

                    Text5.text = strtet
                    
                End If
                 
                If Col = E_FPS.E_declarationno Or Col = E_FPS.E_AWB Or Col = E_FPS.E_destination Or Col = E_FPS.E_freight Then
                 
                    With fpS(0)
                    
                        .Row = Row
                        .Col = Col
                        strbaog = Trim$(.text)
                        
                        For i = 1 To .MaxRows
                        
                            .Row = i
                            .Col = Col
                            .text = strbaog
                            
                        Next
                    
                    End With
                    
                End If
 
                If Col = E_FPS.E_modetrade Then

                    With fpS(0)
                        .Row = Row
                        .Col = E_FPS.E_modetrade

                        If .text = "" Then

                            MsgBox "请输入类别", vbInformation, "提示"
                            Exit Sub

                        End If

                        strsty = Trim(.text)

                        If strsty = "进料对口" Or strsty = "进料成品退换" Or strsty = "进料料件复出" Then

                            .Col = E_FPS.E_manualno
                            .Lock = False

                            .Col = E_FPS.E_itemno
                            .Lock = False

                            .Col = E_FPS.E_name
                            .Lock = True
                            .SetText E_FPS.E_name, Row, Trim$("")

                            .Col = E_FPS.E_UNIT
                            .Lock = True
                            .SetText E_FPS.E_UNIT, Row, Trim$("")

                        Else

                            .Col = E_FPS.E_manualno
                            .Lock = True
                            .SetText E_FPS.E_manualno, Row, Trim$("")

                            .Col = E_FPS.E_itemno
                            .Lock = True
                            .SetText E_FPS.E_itemno, Row, Trim$("")

                            .Col = E_FPS.E_name
                            .text = ""
                            .Lock = False

                            .Col = E_FPS.E_UNIT
                            .text = ""
                            .Lock = False

                        End If

                    End With

                End If

                If Col = E_FPS.E_itemno Then
    
                    With fpS(0)
                        .Row = Row
                        .Col = E_FPS.E_modetrade
                        strsty = Trim(.text)
                
                        '只有选择了手册号才可以出现选择项号的功能，才能带出品名，否则就需要手工输入品名与单位
                
                        If strsty = "进料对口" Or strsty = "进料成品退换" Or strsty = "进料料件复出" Then
                
                            .Col = E_FPS.E_manualno
                            stritem = Trim(.text)

                            .Col = E_FPS.E_itemno
                            
                            If Trim$(.text) <> "" Then
                                
                                If strsty = "进料对口" Or strsty = "进料成品退换" Then
                                
                                    If Get_SqlserverCnt("SELECT 商品名称 as 品名,计量单位 " & " FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & stritem & "' and flag = '" & strflag & "' and 序号= '" & Trim$(.text) & "'") = 0 Then
                                    
                                        MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"
                                    
                                        .SetText E_FPS.E_name, Row, Trim$("")
                                        .SetText E_FPS.E_UNIT, Row, Trim$("")
                                    
                                        Exit Sub
                                
                                    End If

                                    strsql = "SELECT 商品名称 as 品名,计量单位 " & " FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & stritem & "' and flag = '" & strflag & "' and 序号= '" & Trim$(.text) & "'"
                                
                                Else
                                
                                    If Get_SqlserverCnt("SELECT 商品名称 as 品名,计量单位 " & " FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & stritem & "' and flag = '1' and 序号= '" & Trim$(.text) & "'") = 0 Then
                                    
                                        MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"
                                    
                                        .SetText E_FPS.E_name, Row, Trim$("")
                                        .SetText E_FPS.E_UNIT, Row, Trim$("")
                                    
                                        Exit Sub
                                
                                    End If

                                    strsql = "SELECT 商品名称 as 品名,计量单位 " & " FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & stritem & "' and flag = '1' and 序号= '" & Trim$(.text) & "'"
                                
                                End If

                                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                                If Not rs.EOF Then

                                    With fpS(0)
                                    
                                        .SetText E_FPS.E_name, Row, Trim$("" & rs!品名)
                                        .SetText E_FPS.E_UNIT, Row, Trim$("" & rs!计量单位)

                                    End With

                                End If

                                rs.Close

                            End If

                        End If

                    End With
           
                End If

                If Col = E_FPS.E_totalprice Then
                    
                    With fpS(0)
                    
                        .Row = Row
                        
                        .Col = E_FPS.E_exportquantity
                        strno = Trim$(.text)
                        
                        If Trim$(.text) = "" Then
                        
                            MsgBox "请输入出货数量", vbInformation, "提示"
                            Exit Sub
                            
                        End If
                    
                        .Col = E_FPS.E_totalprice

                        If Trim$(.text) = "" Then
                        
                            MsgBox "请输入总价", vbInformation, "提示"
                            Exit Sub
                                
                        Else
                            strNo1 = Trim$(.text)
                            
                            strNo2 = Val(strNo1) / Val(strno)
                            
                            strNo3 = Format(Trim$(strNo2), "0.000000")
                    
                            .SetText E_FPS.E_unitprice, Row, Trim$("" & strNo3)

                        End If

                    End With
                    
                    With fpS(0)
                        
                        strval = 0
                    
                        For i = 1 To .MaxRows
                            
                            .Row = i
                            
                            .Col = E_FPS.E_gx
                            
                            If .text = 1 Then
                                
                                .Col = E_FPS.E_totalprice
                            
                                strval = Val(strval) + Val(.text)
                                
                            End If
                            
                        Next
                    
                    End With
                    
                    strval = Format(Trim$(strval), "0.0000")
                            
                    Text4.text = strval
    
                End If
                
                If Col = E_FPS.E_exportquantity Then
                    
                    With fpS(0)
                        
                        strtet = 0
                    
                        For i = 1 To .MaxRows
                            
                            .Row = i
                            
                            .Col = E_FPS.E_gx
                            
                            If .text = 1 Then
                                
                                .Col = E_FPS.E_exportquantity
                            
                                strtet = Val(strtet) + Val(.text)
                                
                            End If
                            
                        Next
                    
                    End With
                    
                    strtet = Format(Trim$(strtet), "0.000")
                            
                    Text5.text = strtet
    
                End If
                
            End If

    End Select

    '    Me.MousePointer = 0

End Sub


Private Sub fpS_Clear_Click(ByVal Col As Long, ByVal Row As Long)


If Col <> 1 Then Exit Sub
With fpS_Clear
    .Col = 1
    .Row = Row
    .Value = Abs(Val(.Value) - 1)

    If Val(.Value) = 1 Then
    
        .Row = Row
        .Col = -1
        .BackColor = &HC0C0FF

          
    Else

        .Row = Row
        .Col = -1
        .BackColor = &H8000000F
        
        
    End If
End With
End Sub

Private Sub Fps_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
  
    Dim i      As Long
    
    Dim j      As Long
    
    Dim strRow As Long

    With fpS(0)
        
        .Row = .ActiveRow
       
        strRow = .Row
    
        If .MaxRows > 1 Then
       
            For i = 1 To .MaxRows
        
                If i <> strRow Then
            
                    .Row = i
                
                    For j = 2 To .MaxCols
                    
                        .Col = j
                        .BackColor = vbWhite
                    
                    Next

                End If
        
            Next
        
            .Row = strRow
       
            For i = 2 To .MaxCols
       
                .Col = i
            
                '.ForeColor = &HFF8080

                .BackColor = vbGreen
        
            Next

        End If

    End With

End Sub

Private Sub fps_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If Index = 0 Then
    
  ' enter 键
'        If KeyCode = 13 Then
                   
        If KeyCode = vbKeyBack Then
        
            
            With fpS(0)
        
                .Row = .ActiveRow
                .Col = .ActiveCol
                
                If .Lock = False Then
                
                    .text = ""
                    
                End If
    
            End With
        
        End If

    End If

End Sub

Private Sub Optall_Click()
QueryData
End Sub

Private Sub Optpatial_Click()
QueryData
End Sub



Private Sub Optpatial2_Click()
QueryData
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'    Dim a As Integer
'    Dim b As Integer
'    a = 100
'    b = -300
'    MsgBox a + b
    Select Case Button.Key
        
        Case "QUE"
        
            strstate = False
            
            strstate1 = False
            
            ForQuery

        Case "ADD"
            ForAdd
        
        Case "MOD"

            Select Case Combo1.text
                    
                Case "出口明细表"
                    ForMod5
                
                Case "进口明细表"
                    ForMod6
                    
                Case "进口明细表(特殊)"
                    ForMod2
                
                Case "出口明细表(特殊)"
                    ForMod1

            End Select
        
        Case "DEL"
            
            ForDel

        Case "EXIT"
            Unload Me
            
        Case "RET"
            Toolbar1.Buttons(1).Enabled = True
            Toolbar1.Buttons(1).Caption = "查询"
            Toolbar1.Buttons(1).Image = 1
            Toolbar1.Buttons(3).Enabled = True
            Toolbar1.Buttons(3).Caption = "新增"
            Toolbar1.Buttons(3).Image = 3
            Toolbar1.Buttons(5).Enabled = True
            Toolbar1.Buttons(5).Caption = "修改"
            Toolbar1.Buttons(5).Image = 4
            Toolbar1.Buttons(7).Enabled = True
            Toolbar1.Buttons(7).Caption = "删除"
            Toolbar1.Buttons(7).Image = 5
            
            lb4.Visible = False
            
            Command4.Visible = False
            
            Command5.Visible = False
        
            comBo2.Visible = False
            
            lb5.Visible = False
            
            Combo3.Visible = False
            
            comBo2.text = ""
            
            Combo3.text = ""
            
            '            Combo1.Text = ""
            
            lb6.Visible = False
            
            lb7.Visible = False
            
            Text4.Visible = False
            
            Text5.Visible = False
        
            Text4.text = ""
            
            Text5.text = ""
            
            strval = 0
            
            strtet = 0

            Select Case Combo1.text
                    
                Case "出口明细表(特殊)"
                
                    lb4.Visible = True
                    comBo2.Visible = True

            End Select
                
            fpS(0).MaxRows = 0
            fpS(0).MaxCols = 0
            fpss(0).MaxRows = 0
            fpss(0).MaxCols = 0
            fpss(0).Visible = False
            
    End Select

End Sub

Private Sub ForQuery()

    If Combo1.text = "" Then
        MsgBox "请选择维护类型", vbInformation, "提示"
        Exit Sub

    End If
          
    lb6.Visible = True
            
    lb7.Visible = True
            
    Text4.Visible = True
            
    Text5.Visible = True

    Select Case Combo1.text
        
        Case "出口明细表"
        
            QueType5
                
        Case "进口明细表"
        
            QueType6
            
        Case "出口明细表(特殊)"
        
            QueType5
                
        Case "进口明细表(特殊)"
        
            QueType6

    End Select

End Sub

Private Sub QueType5()
    
    Dim rs       As New ADODB.Recordset

    Dim strInv   As String
    
    Dim strInv1  As String
    
    Dim strflag1 As Integer
    
    Dim strsssql As String
    
    Dim strflag2 As Integer
    
    Dim i        As Integer

    Dim strsql   As String
    
    Dim a()      As String
    
    Dim leni     As Integer
    
    Dim strstart As String
    
    Dim strend   As String
    
    '    If Text1.Text = "" Then
    '        MsgBox "请输入发票号", vbInformation, "提示"
    '        Exit Sub
    '
    '    End If
    strInv = Trim$(Text1.text)

    a = Split(strInv, "/")
    
    leni = UBound(a) - LBound(a) + 1
    
    AddSql2 ("delete from erptemp.dbo.ksexport_temp where 1 = 1")

    For i = 0 To leni - 1
        
        If Get_SqlserverCnt("SELECT * FROM erpdata..tblsale A WHERE A.销售单编号 = '" & a(i) & "'") = 0 Then
            
            strflag1 = 1
            
        End If
        
        strsssql = "select delivery from erpbase.dbo.tblCustomerShippingUp where delivery = '" & a(i) & "'"
        
        If Get_SqlserverCnt(strsssql) = 0 Then
            
            strflag2 = 1
            
        End If
        
        Select Case Combo1.text
                
            Case "出口明细表"

                If strflag1 = 1 And strflag2 = 1 Then
        
                    MsgBox "没有此发票号" & a(i) & ",请重新输入", vbInformation, "提示"
                    Exit Sub
        
                End If

        End Select
        
        strflag1 = 0
    
        strflag2 = 0
        
        AddSql2 (" insert into erptemp.dbo.ksexport_temp values('" & a(i) & "') ")
        
    Next
    
    strstart = Format(DTPicker1.Value, "yyyy-MM-dd")
    
    strend = Format(DTPicker2.Value, "yyyy-MM-dd")
    
    If strstart > strend Then
    
        MsgBox "开始日期不可选择大于结束日期", vbInformation, "提示"
            
        Exit Sub
    
    End If

    If Text1.text = "" Then
    
        If strstate = True Then
        
            strsql = "select '' as '√',批次,出货单据,料号,类别,发票号,出货日期,数量,报关单号,手册编号,手册项号,品名,计量单位,币别,总价,单价,AWB#,目的地,货代,退单日期,备注,id from erptemp.dbo.ksexport where flag = '0' and CONVERT(varchar(100),键入时间, 23) >= '" & strstart & "' and CONVERT(varchar(100),键入时间, 23) <= '" & strend & "' and 批次 = '" & stridid & "' order by 批次,id"
        
        Else
            Select Case Combo1.text
                
                Case "出口明细表"
                
                    strsql = "select '' as '√',批次,出货单据,料号,类别,发票号,出货日期,数量,报关单号,手册编号,手册项号,品名,计量单位,币别,总价,单价,AWB#,目的地,货代,退单日期,备注,id from erptemp.dbo.ksexport where flag = '0' and CONVERT(varchar(100),键入时间, 23) >= '" & strstart & "' and CONVERT(varchar(100),键入时间, 23) <= '" & strend & "' and 出货单据 <> '' order by 批次,id"
            
                Case "出口明细表(特殊)"
                
                    strsql = "select '' as '√',批次,出货单据,料号,类别,发票号,出货日期,数量,报关单号,手册编号,手册项号,品名,计量单位,币别,总价,单价,AWB#,目的地,货代,退单日期,备注,id from erptemp.dbo.ksexport where flag = '0' and CONVERT(varchar(100),键入时间, 23) >= '" & strstart & "' and CONVERT(varchar(100),键入时间, 23) <= '" & strend & "' and 出货单据 = ''  order by 批次,id"
                
            End Select
            
        
        End If
    
    Else
        
        strsql = "select '' as '√',批次,出货单据,料号,类别,发票号,出货日期,数量,报关单号,手册编号,手册项号,品名,计量单位,币别,总价,单价,AWB#,目的地,货代,退单日期,备注,id from erptemp.dbo.ksexport where 发票号  in (select distinct purchase from erptemp.dbo.ksexport_temp where 1 = 1)  and flag = '0' and CONVERT(varchar(100),键入时间, 23) >= '" & strstart & "' and CONVERT(varchar(100),键入时间, 23) <= '" & strend & "' order by 批次,id "

    End If
    
    fpS(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType5(rs)
    Else
            
        strtet = 0
        
        strval = 0
        
        lb7 = "出货总量"
        
        strtet = Format(Trim$(strtet), "0.000")
        
        lb6 = "出货总额"
        
        strval = Format(Trim$(strval), "0.0000")
        
        Text4.text = strval
    
        Text5.text = strtet
        
        MsgBox "查询不到该出口单据信息", vbInformation, "提示"
        Exit Sub

    End If
    
    strstate = False
    
    AddSql2 ("delete from erptemp.dbo.ksexport_temp where 1 = 1")
    
End Sub

Private Sub QueType6()

    Dim rs       As New ADODB.Recordset

    Dim strInv   As String

    Dim strsql   As String
    
    Dim a()      As String
    
    Dim strstart As String
    
    Dim strend   As String
    
    Dim i        As Integer
    
    Dim leni     As Integer
    
    strInv = Trim$(Text1.text)
    
    a = Split(strInv, "/")
    
    leni = UBound(a) - LBound(a) + 1
    
    AddSql2 ("delete from erptemp.dbo.ksimport_temp where 1 = 1")

    For i = 0 To leni - 1
        
        Select Case Combo1.text
                
            Case "进口明细表"
        
                If Get_SqlserverCnt("SELECT * FROM erpbase..tblCPurDataSub WHERE 采购单编号 = '" & a(i) & "'") = 0 Then
                    MsgBox "没有此采购单号" & a(i) & ",请重新输入", vbInformation, "提示"
                    Exit Sub
    
                End If

        End Select
        
        AddSql2 (" insert into erptemp.dbo.ksimport_temp values('" & a(i) & "') ")

    Next
    
    strstart = Format(DTPicker1.Value, "yyyy-MM-dd")
    
    strend = Format(DTPicker2.Value, "yyyy-MM-dd")
    
    If strstart > strend Then
    
        MsgBox "开始日期不可选择大于结束日期", vbInformation, "提示"
            
        Exit Sub
    
    End If
    
    If Text1.text = "" Then
    
        
        If strstate1 = True Then
    
            strsql = "select '' as '√',批次,采购单号,料号,型号,类别,订单数量,标准die,总die数,手册编号,项号,品名,报关量,计量单位,入场日期,发票号,件数,币别,采购单价,报关金额,汇率,关税率,关税,增值税率,增值税,报关单号,AWB#,货代,退单日期,备注,id from erptemp.dbo.ksimport where  flag = '0' and CONVERT(varchar(100),键入时间, 23) >= '" & strstart & "' and CONVERT(varchar(100),键入时间, 23) <= '" & strend & "' and 批次 = '" & stridid1 & "' order by 批次,id"
        
        Else
            
            Select Case Combo1.text
                
                Case "进口明细表"
            
                    strsql = "select '' as '√',批次,采购单号,料号,型号,类别,订单数量,标准die,总die数,手册编号,项号,品名,报关量,计量单位,入场日期,发票号,件数,币别,采购单价,报关金额,汇率,关税率,关税,增值税率,增值税,报关单号,AWB#,货代,退单日期,备注,id from erptemp.dbo.ksimport where  flag = '0' and CONVERT(varchar(100),键入时间, 23) >= '" & strstart & "' and CONVERT(varchar(100),键入时间, 23) <= '" & strend & "' and 采购单号 <> '' order by 批次,id"
                
                Case "进口明细表(特殊)"
                
                    strsql = "select '' as '√',批次,采购单号,料号,型号,类别,订单数量,标准die,总die数,手册编号,项号,品名,报关量,计量单位,入场日期,发票号,件数,币别,采购单价,报关金额,汇率,关税率,关税,增值税率,增值税,报关单号,AWB#,货代,退单日期,备注,id from erptemp.dbo.ksimport where  flag = '0' and CONVERT(varchar(100),键入时间, 23) >= '" & strstart & "' and CONVERT(varchar(100),键入时间, 23) <= '" & strend & "' and 采购单号 = ''  order by 批次,id"
            
            End Select
        End If
    
    Else
        
        strsql = "select '' as '√',批次,采购单号,料号,型号,类别,订单数量,标准die,总die数,手册编号,项号,品名,报关量,计量单位,入场日期,发票号,件数,币别,采购单价,报关金额,汇率,关税率,关税,增值税率,增值税,报关单号,AWB#,货代,退单日期,备注,id from erptemp.dbo.ksimport where 采购单号 in (select distinct purchase from erptemp.dbo.ksimport_temp where 1 = 1) and flag = '0' and CONVERT(varchar(100),键入时间, 23) >= '" & strstart & "' and CONVERT(varchar(100),键入时间, 23) <= '" & strend & "' order by 批次,id "
        
    End If
    
    fpS(0).MaxRows = 0
    fpS(0).MaxCols = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
       
        Call ListDataType6(rs)
       
    Else
            
        strtet = 0
        
        strval = 0
        
        lb7 = "报关总量"
        
        strtet = Format(Trim$(strtet), "0.000")
        
        lb6 = "报关总额"
        
        strval = Format(Trim$(strval), "0.0000")
        
        Text4.text = strval
    
        Text5.text = strtet

        MsgBox "查询不到该采购单据信息", vbInformation, "提示"
        Exit Sub

    End If
  
    Select Case Combo1.text
                
        Case "进口明细表"
        
            fpss(0).Visible = True
                
        Case "进口明细表(特殊)"
        
            fpss(0).Visible = False

    End Select
    
    If Text1.text = "" Then
    
        strsql = "SELECT  采购单号,料号,isnull(sum(订单数量),0) as 已收数量 from erptemp.dbo.ksimport where 1 = 1 and flag = '0' and CONVERT(varchar(100),键入时间, 23) >= '" & strstart & "' and CONVERT(varchar(100),键入时间, 23) <= '" & strend & "' AND 采购单号 <> '' group by 采购单号,料号 order by 采购单号,料号"
    
    Else
        
        strsql = "SELECT  采购单号,料号,isnull(sum(订单数量),0) as 已收数量 from erptemp.dbo.ksimport where 1 = 1 and flag = '0' and 采购单号 in (select distinct purchase from erptemp.dbo.ksimport_temp where 1 = 1) group by 采购单号,料号 order by 采购单号,料号 "
    
    End If
    
    fpss(0).MaxRows = 0
    fpss(0).MaxCols = 0
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
    Call ListDataType1(rs)
    
    strstate1 = False
    
    AddSql2 ("delete from erptemp.dbo.ksimport_temp where 1 = 1")

End Sub

Private Sub ListDataType1(rs As ADODB.Recordset)

    Dim i As Long
   
    With fpss(0)
        
        .MaxRows = 0
 
        Set .DataSource = rs

        For i = 1 To .MaxRows
        
            .Row = i
            .Col = 15
            
            .text = Format(Trim$(.text), "0.00")
            
        Next

    End With

End Sub

Private Sub ListDataType2(rs As ADODB.Recordset)

 Dim i As Long
  
   
    With fpsss(0)
        
        .MaxRows = 0
        
        Set .DataSource = rs

    End With

End Sub

Private Sub ListDataType5(rs As ADODB.Recordset)

    Dim i As Long
   
    Dim j As Long
    
    With fpS(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
    With fpS(0)
        
        strtet = 0
        
        strval = 0
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = E_FPS.E_gx
            .ColWidth(E_FPS.E_gx) = 2
            .CellType = CellTypeCheckBox
            .text = 1
            
            .Col = E_FPS.E_gx
            .Lock = False
            
             .Col = E_FPS.e_NO
            .Lock = False
            
            For j = E_FPS.E_exportno To E_FPS.e_ID
                        
                .Col = j
                .Lock = True
                        
            Next
            .LockBackColor = vbYellow
            
            .Col = 1

            If .text = 1 Then
            
                .Col = 8
                    strtet = Val(.text) + Val(strtet)

                
                
                .Col = 15
                
                strval = Val(.text) + Val(strval)
            
            End If
            
            .Col = 2
            .ColWidth(2) = 8
            
            .Col = 14
            .CellType = CellTypeComboBox
            
            .TypeComboBoxList = .TypeComboBoxList & "USD"
            
            .TypeComboBoxList = .TypeComboBoxList & "JPY"

            .TypeComboBoxList = .TypeComboBoxList & "EUR"

            .TypeComboBoxList = .TypeComboBoxList & "RMB"
            
            .Col = 8
            
            .text = Format(Trim$(.text), "0.000")
            
            .Col = 15
            
            .text = Format(Trim$(.text), "0.00")
            
            .Col = 16
            
            .text = Format(Trim$(.text), "0.000000")
            
            .Col = 22
            .ColWidth(22) = 4
            
        Next
        
        lb7 = "出货总量"
        strtet = Format(Trim$(strtet), "0.000")
        
        lb6 = "出货总额"
        
        strval = Format(Trim$(strval), "0.0000")
        
        Text4.text = strval
    
        Text5.text = strtet

    End With

End Sub

Private Sub ListDataType6(rs As ADODB.Recordset)

    Dim i      As Long
    
    Dim j      As Long
    
    Dim strsql As String

    With fpS(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
    With fpS(0)
    
        strtet = 0
        
        strval = 0

        For i = 1 To .MaxRows
            .Row = i
             
            .Col = F_fp.F_gx
            .ColWidth(1) = 2
            .CellType = CellTypeCheckBox
            .text = 1
            
            .Col = F_fp.F_gx
            .Lock = False

            '            Select Case Combo1.Text
            '
            '                Case "进口明细表(特殊)"
                    
            '          .DAutoSizeCols = DAutoSizeColsMax
            .Col = F_fp.F_no
            .Lock = False
            For j = F_fp.F_purchaseno To F_fp.F_id
                        
                .Col = j
                .Lock = True
                        
            Next
            .LockBackColor = vbYellow
                    
            .Col = F_fp.F_partno
            .ColWidth(F_fp.F_partno) = 16
        
            .Col = F_fp.F_modelno
            .ColWidth(F_fp.F_modelno) = 10
        
            .Col = F_fp.F_freight
            .ColWidth(F_fp.F_freight) = 10
        
            .Col = F_fp.F_modetrade
            .ColWidth(F_fp.F_modetrade) = 14
                    
            .Col = F_fp.F_manualno
            .ColWidth(F_fp.F_manualno) = 12
                    
            .Col = F_fp.F_currency
            .ColWidth(F_fp.F_currency) = 6

            .Col = F_fp.F_declarationno
            .ColWidth(F_fp.F_declarationno) = 18
        
            .Col = F_fp.F_invoice
            .ColWidth(F_fp.F_invoice) = 14
        
            .Col = F_fp.F_awb
            .ColWidth(F_fp.F_awb) = 14
            '
            '            End Select
        
            .Col = F_fp.F_gx

            If .text = 1 Then
            
                .Col = F_fp.F_baoguanqty
                
                strtet = Val(.text) + Val(strtet)
                
                .Col = F_fp.F_baoguanvalue
                
                strval = Val(.text) + Val(strval)
            
            End If
            
            .Col = F_fp.F_no
            .ColWidth(2) = 8
        
            .Col = F_fp.F_modetrade
            .CellType = CellTypeComboBox
                       
            .TypeComboBoxList = .TypeComboBoxList & "进料对口"
            
            .TypeComboBoxList = .TypeComboBoxList & "一般贸易"
            
            .TypeComboBoxList = .TypeComboBoxList & "其他进出口免费"
            
            .TypeComboBoxList = .TypeComboBoxList & "成品复进"
            
            .TypeComboBoxList = .TypeComboBoxList & "维修物品"
            
            .TypeComboBoxList = .TypeComboBoxList & "料件复进"
            
            .TypeComboBoxList = .TypeComboBoxList & "进料成品退换"
            
            .TypeComboBoxList = .TypeComboBoxList & "其他"

            .Col = F_fp.F_orderqty
            .text = Format(Trim$(.text), "0.000")
            
            .Col = F_fp.F_die
            .text = Format(Trim$(.text), "0.000")
            
            .Col = F_fp.F_totaldie
            .text = Format(Trim$(.text), "0.000")
            
            .Col = F_fp.F_baoguanqty
            .text = Format(Trim$(.text), "0.000")
            
            .Col = F_fp.F_currency
            .CellType = CellTypeComboBox
            
            .TypeComboBoxList = .TypeComboBoxList & "USD"
            
            .TypeComboBoxList = .TypeComboBoxList & "JPY"

            .TypeComboBoxList = .TypeComboBoxList & "EUR"

            .TypeComboBoxList = .TypeComboBoxList & "RMB"
            
            .Col = F_fp.F_unitprice
            .text = Format(Trim$(.text), "0.0000")
            
            .Col = F_fp.F_baoguanvalue
            .text = Format(Trim$(.text), "0.0000")
            
            .Col = F_fp.F_rate
            .text = Format(Trim$(.text), "0.0000")
            
            .Col = F_fp.F_tariffrate
            .text = Format(Trim$(.text), "0.0000")
            
            .Col = F_fp.F_tariff
            .text = Format(Trim$(.text), "0.000")
            
            .Col = F_fp.F_addtaxrate
            .text = Format(Trim$(.text), "0.0000")
            
            .Col = F_fp.F_addtax
            .text = Format(Trim$(.text), "0.00")
            
            .Col = F_fp.F_id
            .ColWidth(F_fp.F_id) = 3
        
        Next
        
        lb7 = "报关总量"
        
        strtet = Format(Trim$(strtet), "0.000")
        
        lb6 = "报关总额"
        
        strval = Format(Trim$(strval), "0.0000")
        
        Text4.text = strval
    
        Text5.text = strtet

    End With
     
End Sub

Private Sub ForAdd()

    If Toolbar1.Buttons(3).Caption = "提交" Then
        
        Select Case Combo1.text
             
            Case "出口明细表"
                ForCommit1
                
            Case "进口明细表"
                ForCommit2

            Case "出口明细表(特殊)"
                ForCommit1

            Case "进口明细表(特殊)"
                ForCommit2

        End Select
        
        Exit Sub

    End If

    If Combo1.text = "" Then
        MsgBox "请选择维护类型", vbInformation, "提示"
        Exit Sub

    End If

    Select Case Combo1.text

        Case "出口明细表"
            AddType5

        Case "进口明细表"
            AddType6
            
        Case "出口明细表(特殊)"
            AddType1
        
        Case "进口明细表(特殊)"
            AddType2

    End Select

End Sub

Private Function Createid() As String

    Dim stridd   As String
    
    Dim strtime  As String
    
    Dim stridd1  As String
    
    Dim stridd2  As String
    
    Dim stritime As String

    strtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    strtime = Left$(strtime, 4)
    
    Select Case Combo1.text

        Case "出口明细表(特殊)"
        
            stridd = Get_SqlStr("select isnull(max(批次),0) from erptemp.dbo.ksexport where flag = '0'")

        Case "进口明细表(特殊)"
        
            stridd = Get_SqlStr("select isnull(max(批次),0) from erptemp.dbo.ksimport where flag = '0'")
            
        Case "出口明细表"
        
            stridd = Get_SqlStr("select isnull(max(批次),0) from erptemp.dbo.ksexport where flag = '0'")

        Case "进口明细表"
        
            stridd = Get_SqlStr("select isnull(max(批次),0) from erptemp.dbo.ksimport where flag = '0'")

    End Select
    
    stridd = Format(Trim$(stridd), "00000000")
    
    stridd1 = Left$(stridd, 4)
    
    stridd2 = Right$(stridd, 4)
    
    If stridd1 <> strtime Then
    
        stridd1 = strtime
        
        stridd2 = "0001"
        
        stridd = stridd1 & stridd2
        
    Else
        
        stridd2 = Format(Val(stridd2) + 1, "0000")
        
        stridd = stridd1 & stridd2
        
    End If

    Createid = stridd

End Function

Private Sub AddType1()
    
    Dim stridd As String
    
    Dim rs     As New ADODB.Recordset
    
    Dim i      As Integer
    
    Dim strsql As String
    
    strtet = 0
        
    strval = 0
    
    stridd = Createid
    
    Command4.Visible = True
    
    Command5.Visible = True
    
    With fpS(0)
        '
        '        .ReDraw = False
        .MaxCols = E_FPS.e_ID
        .MaxRows = 0
        
        '        .DAutoHeadings = False
        '        .DAutoCellTypes = True
        '        .DAutoSizeCols = DAutoSizeColsBest
        '        .DAutoSizeCols = DAutoSizeColsMax
        .Col = -1
        .Row = -1
        .Lock = False

        '        .OperationMode = OperationModeNormal
        '        .TypeVAlign = TypeVAlignCenter
        '        .SelForeColor = &HFF8080
        
        .SetText E_FPS.E_gx, 0, "√"
        .SetText E_FPS.e_NO, 0, "批次"
        .SetText E_FPS.E_exportno, 0, "出货单据"
        .SetText E_FPS.E_partno, 0, "料号"
        .SetText E_FPS.E_modetrade, 0, "类别"
        .SetText E_FPS.e_Invoice, 0, "发票号"
        .SetText E_FPS.E_exportdate, 0, "出货日期"
        .SetText E_FPS.E_exportquantity, 0, "出货数量"
        .SetText E_FPS.E_declarationno, 0, "报关单号"
        .SetText E_FPS.E_manualno, 0, "手册编号"
        .SetText E_FPS.E_itemno, 0, "手册项号"
        .SetText E_FPS.E_name, 0, "品名"
        .SetText E_FPS.E_UNIT, 0, "计量单位"
        .SetText E_FPS.E_currency, 0, "币别"
        .SetText E_FPS.E_totalprice, 0, "总价"
        .SetText E_FPS.E_unitprice, 0, "单价"
        .SetText E_FPS.E_AWB, 0, "AWB#"
        .SetText E_FPS.E_destination, 0, "目的地"
        .SetText E_FPS.E_freight, 0, "货代"
        .SetText E_FPS.E_chargebackdate, 0, "退单日期"
        .SetText E_FPS.E_mark, 0, "备注"
        .SetText E_FPS.e_ID, 0, "id"
        
        '        .RowHeight(0) = 22
        '        .RowHeight(-1) = 22

        .Col = E_FPS.E_gx    '选择
        .CellType = CellTypeCheckBox
        .ColWidth(1) = 2
        .Lock = False
        .text = 1
        
        .Col = E_FPS.e_ID
        
        .ColWidth(E_FPS.e_ID) = 3
     
        .Col = E_FPS.E_declarationno
        .ColWidth(E_FPS.E_declarationno) = 18
        
        .Col = E_FPS.E_destination
        .ColWidth(E_FPS.E_destination) = 8
        
        .Col = E_FPS.E_manualno
        .ColWidth(E_FPS.E_manualno) = 12
        
        .Col = E_FPS.e_Invoice
        .ColWidth(E_FPS.e_Invoice) = 14
        
        .Col = E_FPS.E_AWB
        .ColWidth(E_FPS.E_AWB) = 14

        .Col = E_FPS.E_partno
        .ColWidth(E_FPS.E_partno) = 16

        .Col = E_FPS.E_freight
        .ColWidth(E_FPS.E_freight) = 10
        
        .Col = E_FPS.E_modetrade
        .ColWidth(E_FPS.E_modetrade) = 14
        .CellType = CellTypeComboBox
                       
        .TypeComboBoxList = .TypeComboBoxList & "进料对口"
            
        .TypeComboBoxList = .TypeComboBoxList & "一般贸易"
            
        .TypeComboBoxList = .TypeComboBoxList & "其他进出口免费"
            
        .TypeComboBoxList = .TypeComboBoxList & "进料料件复出"
            
        .TypeComboBoxList = .TypeComboBoxList & "进料成品退换"
            
        .TypeComboBoxList = .TypeComboBoxList & "修理物品"
            
        .TypeComboBoxList = .TypeComboBoxList & "设备退运"
            
        .TypeComboBoxList = .TypeComboBoxList & "其他"
                
        .Col = E_FPS.E_currency
        
        .CellType = CellTypeComboBox
            
        .TypeComboBoxList = "USD"
            
        .TypeComboBoxList = .TypeComboBoxList & "JPY"

        .TypeComboBoxList = .TypeComboBoxList & "EUR"

        .TypeComboBoxList = .TypeComboBoxList & "RMB"
        
        strsql = "select distinct 手册编号 from erptemp.dbo.ksmanual where 1 = 1 and flag = '2'"

        If rs.State = 1 Then rs.Close
        rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

        .Col = E_FPS.E_manualno
        .ColWidth(E_FPS.E_manualno) = 12
        .CellType = CellTypeComboBox

        rs.MoveFirst

        For i = 1 To rs.RecordCount

            .TypeComboBoxList = .TypeComboBoxList & rs("手册编号")
            rs.MoveNext
        Next
        
        rs.Clone
        
        Set rs = Nothing
        
        '        .ReDraw = True
        
    End With
    
    Toolbar1.Buttons(3).Caption = "提交"
    Toolbar1.Buttons(3).Image = 6
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    
    With fpS(0)

        .MaxRows = .MaxRows + 1
        '        .DAutoSizeCols = DAutoSizeColsMax
        
        .SetText E_FPS.E_gx, 1, "1"
        .SetText E_FPS.e_NO, 1, Trim$(stridd)
        .SetText E_FPS.E_exportno, 1, ""
        .SetText E_FPS.E_partno, 1, ""
        .SetText E_FPS.E_modetrade, 1, Trim$(comBo2.text)
        .SetText E_FPS.e_Invoice, 1, Trim$(Text1.text)
        .SetText E_FPS.E_exportdate, 1, ""
        .SetText E_FPS.E_exportquantity, 1, ""
        .SetText E_FPS.E_declarationno, 1, ""
        .SetText E_FPS.E_manualno, 1, Trim$(Combo3.text)
        .SetText E_FPS.E_itemno, 1, ""
        .SetText E_FPS.E_name, 1, ""
        .SetText E_FPS.E_UNIT, 1, ""
        .SetText E_FPS.E_currency, 1, "USD"
        .SetText E_FPS.E_totalprice, 1, ""
        .SetText E_FPS.E_unitprice, 1, ""
        .SetText E_FPS.E_AWB, 1, ""
        .SetText E_FPS.E_destination, 1, ""
        .SetText E_FPS.E_freight, 1, ""
        .SetText E_FPS.E_chargebackdate, 1, ""
        .SetText E_FPS.E_mark, 1, ""
        .SetText E_FPS.e_ID, 1, ""
        
        .Col = E_FPS.E_exportno
        .Lock = True
                
        .Col = E_FPS.E_chargebackdate
        .Lock = True
                
        .Col = E_FPS.E_mark
        .Lock = True
        
        .Col = E_FPS.e_ID
        .Lock = True
        
        .Col = E_FPS.E_unitprice
        .Lock = True
        
        If Trim$(comBo2.text) = "进料对口" Or Trim$(comBo2.text) = "进料成品退换" Or Trim$(comBo2.text) = "进料料件复出" Then
                
            .Col = E_FPS.E_itemno
            .Lock = False
            
            .Col = E_FPS.E_name
            .Lock = True
                
            .Col = E_FPS.E_UNIT
            .Lock = True
        
        Else
            
            .Col = E_FPS.E_manualno
            .Lock = True
            
            .Col = E_FPS.E_itemno
            .Lock = True
            
            .Col = E_FPS.E_name
            .Lock = False
                
            .Col = E_FPS.E_UNIT
            .Lock = False

        End If
        
        .Row = 1
        .LockBackColor = vbYellow
        
        .Row = 1
        .Col = E_FPS.E_gx

        If .text = 1 Then
            
            .Col = E_FPS.E_exportquantity
                
            strtet = Val(.text) + Val(strtet)
                
            .Col = E_FPS.E_totalprice
                
            strval = Val(.text) + Val(strval)
            
        End If
        
    End With
    
    lb7 = "出货总量"
    
    lb6.Visible = True
    
    lb6 = "出货总额"
    
    lb7.Visible = True
    
    Text4.Visible = True
    
    Text5.Visible = True
  
    strtet = Format(Trim$(strtet), "0.000")
            
    strval = Format(Trim$(strval), "0.0000")
        
    Text4.text = strval
    
    Text5.text = strtet

End Sub

Private Sub AddType2()
    
    Dim stridd As String
    
    Dim rs     As New ADODB.Recordset
    
    Dim i      As Integer
    
    Dim strsql As String
    
    Command4.Visible = True
    
    Command5.Visible = True
    
    fpss(0).Visible = False
    
    strtet = 0
        
    strval = 0
    
    stridd = Createid
    
    With fpS(0)
    
        .MaxCols = F_fp.F_id
        .MaxRows = 0
        
        .Col = -1
        .Row = -1
        .Lock = False
        
        .SetText F_fp.F_gx, 0, "√"
        .SetText F_fp.F_no, 0, "批次"
        .SetText F_fp.F_purchaseno, 0, "采购单号"
        .SetText F_fp.F_partno, 0, "料号"
        .SetText F_fp.F_modelno, 0, "型号"
        .SetText F_fp.F_modetrade, 0, "类别"
        .SetText F_fp.F_orderqty, 0, "订单数量"
        .SetText F_fp.F_die, 0, "标准die"
        .SetText F_fp.F_totaldie, 0, "总die数"
        .SetText F_fp.F_manualno, 0, "手册编号"
        .SetText F_fp.F_itemno, 0, "项号"
        .SetText F_fp.F_name, 0, "品名"
        .SetText F_fp.F_baoguanqty, 0, "报关量"
        .SetText F_fp.F_unit, 0, "计量单位"
        .SetText F_fp.F_indate, 0, "入场日期"
        .SetText F_fp.F_invoice, 0, "发票号"
        .SetText F_fp.F_caseqty, 0, "件数"
        .SetText F_fp.F_currency, 0, "币别"
        .SetText F_fp.F_unitprice, 0, "单价"
        .SetText F_fp.F_baoguanvalue, 0, "报关金额"
        .SetText F_fp.F_rate, 0, "汇率"
        .SetText F_fp.F_tariffrate, 0, "关税率"
        .SetText F_fp.F_tariff, 0, "关税"
        .SetText F_fp.F_addtaxrate, 0, "增值税率"
        .SetText F_fp.F_addtax, 0, "增值税"
        .SetText F_fp.F_declarationno, 0, "报关单号"
        .SetText F_fp.F_awb, 0, "AWB#"
        .SetText F_fp.F_freight, 0, "货代"
        .SetText F_fp.F_chargebackdate, 0, "退单日期"
        .SetText F_fp.F_mark, 0, "备注"
        .SetText F_fp.F_id, 0, "id"

        .Col = F_fp.F_gx    '选择
        .CellType = CellTypeCheckBox
        .ColWidth(F_fp.F_gx) = 2
        .Lock = False
        .text = 1
        
        .Col = F_fp.F_no
        .Lock = False
        
        .Col = F_fp.F_purchaseno
        .Lock = True
        
        .Col = F_fp.F_partno
        .ColWidth(F_fp.F_partno) = 16
        
        .Col = F_fp.F_modelno
        .ColWidth(F_fp.F_modelno) = 10
        
        .Col = F_fp.F_freight
        .ColWidth(F_fp.F_freight) = 10
        
        .Col = F_fp.F_modetrade
        .ColWidth(F_fp.F_modetrade) = 14
        .CellType = CellTypeComboBox
                       
        .TypeComboBoxList = .TypeComboBoxList & "进料对口"
            
        .TypeComboBoxList = .TypeComboBoxList & "一般贸易"
            
        .TypeComboBoxList = .TypeComboBoxList & "其他进出口免费"
            
        .TypeComboBoxList = .TypeComboBoxList & "成品复进"
            
        .TypeComboBoxList = .TypeComboBoxList & "维修物品"
            
        .TypeComboBoxList = .TypeComboBoxList & "料件复进"
            
        .TypeComboBoxList = .TypeComboBoxList & "进料成品退换"
            
        .TypeComboBoxList = .TypeComboBoxList & "其他"

        .Col = F_fp.F_orderqty
        .Lock = True
        
        .Col = F_fp.F_declarationno
        .ColWidth(F_fp.F_declarationno) = 18
        
        .Col = F_fp.F_invoice
        .ColWidth(F_fp.F_invoice) = 14
        
        .Col = F_fp.F_awb
        .ColWidth(F_fp.F_awb) = 14
        
        For i = F_fp.F_rate To F_fp.F_addtax
            
            .Col = i
            .Lock = True
        
        Next
        
        For i = F_fp.F_manualno To F_fp.F_name
        
            .Col = i
            .Lock = True
        
        Next
        
        .Col = F_fp.F_unit
        .Lock = True
        
        .Col = F_fp.F_id
        
        .ColWidth(F_fp.F_id) = 3
        .Lock = True
        
        .Col = F_fp.F_unitprice
        .Lock = True

        .Col = F_fp.F_currency
        
        .CellType = CellTypeComboBox
        
        .TypeComboBoxList = "USD"
        
        .TypeComboBoxList = .TypeComboBoxList & "JPY"
        
        .TypeComboBoxList = .TypeComboBoxList & "EUR"
        
        .TypeComboBoxList = .TypeComboBoxList & "RMB"
        
        strsql = "select distinct 手册编号 from erptemp.dbo.ksmanual where 1 = 1 and flag = '1'"

        If rs.State = 1 Then rs.Close
        rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

        .Col = F_fp.F_manualno
        .ColWidth(F_fp.F_manualno) = 12
        .CellType = CellTypeComboBox

        rs.MoveFirst

        For i = 1 To rs.RecordCount

            .TypeComboBoxList = .TypeComboBoxList & rs("手册编号")
            rs.MoveNext
        Next
        
        rs.Clone
        
        Set rs = Nothing

    End With
    
    Toolbar1.Buttons(3).Caption = "提交"
    Toolbar1.Buttons(3).Image = 6
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    
    With fpS(0)

        .MaxRows = .MaxRows + 1
'        .DAutoSizeCols = DAutoSizeColsMax
        
        .SetText F_fp.F_gx, 1, "1"
        .SetText F_fp.F_no, 1, Trim$(stridd)
        .SetText F_fp.F_purchaseno, 1, ""
        .SetText F_fp.F_partno, 1, ""
        .SetText F_fp.F_modelno, 1, ""
        .SetText F_fp.F_modetrade, 1, ""
        .SetText F_fp.F_orderqty, 1, ""
        .SetText F_fp.F_die, 1, ""
        .SetText F_fp.F_totaldie, 1, ""
        .SetText F_fp.F_manualno, 1, ""
        .SetText F_fp.F_itemno, 1, ""
        .SetText F_fp.F_name, 1, ""
        .SetText F_fp.F_baoguanqty, 1, ""
        .SetText F_fp.F_unit, 1, ""
        .SetText F_fp.F_indate, 1, ""
        .SetText F_fp.F_invoice, 1, ""
        .SetText F_fp.F_caseqty, 1, ""
        .SetText F_fp.F_currency, 1, "USD"
        .SetText F_fp.F_unitprice, 1, ""
        .SetText F_fp.F_baoguanvalue, 1, ""
        .SetText F_fp.F_rate, 1, ""
        .SetText F_fp.F_tariffrate, 1, ""
        .SetText F_fp.F_tariff, 1, ""
        .SetText F_fp.F_addtaxrate, 1, ""
        .SetText F_fp.F_addtax, 1, ""
        .SetText F_fp.F_declarationno, 1, ""
        .SetText F_fp.F_awb, 1, ""
        .SetText F_fp.F_freight, 1, ""
        .SetText F_fp.F_chargebackdate, 1, ""
        .SetText F_fp.F_mark, 1, ""
        .SetText F_fp.F_id, 1, ""
        
        .Row = 1
        .LockBackColor = vbYellow
        
        .Row = 1
        .Col = F_fp.F_gx
        If .text = 1 Then
            
            .Col = F_fp.F_baoguanqty
                
            strtet = Val(.text) + Val(strtet)
                
            .Col = F_fp.F_baoguanvalue
                
            strval = Val(.text) + Val(strval)
            
        End If
        
    End With

    lb7 = "报关总量"
    
    lb6.Visible = True
    
    lb6 = "报关总额"
    
    lb7.Visible = True
    
    Text4.Visible = True
    
    Text5.Visible = True
  
    strtet = Format(Trim$(strtet), "0.000")
            
    strval = Format(Trim$(strval), "0.0000")
        
    Text4.text = strval
    
    Text5.text = strtet

End Sub

Private Sub AddType5()

    Dim rs     As New ADODB.Recordset

    Dim i      As Integer
    
    Dim j      As Integer

    Dim m      As Integer
    
    Dim strInv As String
    
    Dim strcom1 As String
    
    Dim strcom2 As String

    Dim strsql As String

    Dim a()    As String
    
    Dim strsssql As String
    
    Dim strflag1 As Integer
    
    Dim strflag2 As Integer
    
    Dim leni   As Integer
    
    Dim stridd As String

    If Text1.text = "" Then
        MsgBox "请填写要维护的发票号", vbInformation, "提示"
        Exit Sub

    End If
    
    If comBo2.text = "" Then
          
        MsgBox "请选择贸易方式", vbInformation, "提示"
        Exit Sub
    
    End If
    
    
    If comBo2.text = "进料对口" Or comBo2.text = "进料成品退换" Or comBo2.text = "进料料件复出" Then
        
        If Combo3.text = "" Then
    
            MsgBox "请选择手册号码", vbInformation, "提示"
            Exit Sub
    
        End If
        
    End If
    
    fpS(0).MaxRows = 0

    strInv = Trim$(Text1.text)
    
    strcom1 = Trim$(comBo2.text)
    
    strcom2 = Trim$(Combo3.text)
    
    a = Split(strInv, "/")
    
    leni = UBound(a) - LBound(a) + 1
    
    AddSql2 ("delete from erptemp.dbo.ksexport_temp where 1 = 1")
    
    stridd = Createid
    
    strflag1 = 0
    
    strflag2 = 0

    For i = 0 To leni - 1

        If Get_SqlserverCnt("SELECT * FROM erpdata..tblsale A WHERE A.销售单编号 = '" & a(i) & "'") = 0 Then
            
            strflag1 = 1
            
        End If
        
        strsssql = "select DN from erpdata..tblStockNumTree where DN = '" & a(i) & "'"
        
        If Get_SqlserverCnt(strsssql) = 0 Then
            
            strflag2 = 1
            
        End If
        
        If strflag1 = 1 And strflag2 = 1 Then
        
            MsgBox "没有此发票号" & a(i) & ",请重新输入", vbInformation, "提示"
            Exit Sub
        
        End If
        
        strflag1 = 0
    
        strflag2 = 0
        
        AddSql2 (" insert into erptemp.dbo.ksexport_temp values('" & a(i) & "') ")
        
    Next

    strsql = " select '' as '√','" & stridd & "' as 批次 ,b.单据编号 as 出货单据,c.料号,'" & strcom1 & "' as 类别,e.delivery as 发票号,  " & _
    " CONVERT(varchar(100), b.操作日期, 23) as 出货日期,CONVERT(decimal(19,3),(SUM(b.实发良品数+b.实发不良数+b.实发制程不良数)/1000.00)) as 数量,'' as 报关单号, " & _
    " '" & strcom2 & "' as 手册编号,'' as 手册项号,'' as 品名,'' as 计量单位,'USD' as 币别, " & _
    " CONVERT(decimal(19,2),e.总价) as 总价, " & _
    " CONVERT(decimal(19,6),e.总价/CONVERT(decimal(19,3),(SUM(b.实发良品数+b.实发不良数+b.实发制程不良数)/1000.00))) as 单价,'' as AWB#,'' as 目的地,'' as 货代,'' as 退单日期,'' as 备注,'' as id " & _
    " from  erpdata..tblSmainM2 c,erpdata..tblStockMove b " & _
    " LEFT JOIN (SELECT distinct p1.单据编号,d.DN as delivery,sum((ISNULL(p1.单价, 0) + ISNULL(p1.客供材料单价, 0)) * p1.数量) AS 总价,p1.料号 FROM erpdata..tblSaleRec p1  " & _
    " LEFT JOIN (select distinct p3.DN,p1.单据编号 from erpdata..tblSaleRec p1  inner join erpdata..tblStocksqfhsub p2 on p1.单据编号 = p2.单据编号 and p1.单据项次 = p2.单据项次 " & _
    " inner join erpdata..tblStockNumTree p3 on p3.箱号 = p2.箱号  where p3.DN in (select distinct purchase from erptemp.dbo.ksexport_temp where 1 = 1)) d on d.单据编号 = p1.单据编号 " & _
    " where p1.单据编号 in( select distinct p1.单据编号 from erpdata..tblSaleRec p1 inner join erpdata..tblStocksqfhsub p2 on p1.单据编号 = p2.单据编号 and p1.单据项次 = p2.单据项次 " & _
    " inner join erpdata..tblStockNumTree p3 on p3.箱号 = p2.箱号  where p3.DN in (select distinct purchase from erptemp.dbo.ksexport_temp where 1 = 1)) " & _
    " group by p1.单据编号,p1.料号,d.DN UNION ALL " & _
    " SELECT distinct RTRIM(b.单据编号) as 单据编号,a.销售单编号 as delivery,sum(b.数量 * (b.客供材料单价 + b.单价)) AS 总价,b.料号 FROM erpdata..tblsale a INNER JOIN erpdata..tblSaleRec b ON a.销售单编号 = b.销售单编号 " & _
    " where a.销售单编号 in (select distinct purchase from erptemp.dbo.ksexport_temp where 1 = 1) group by b.单据编号,a.销售单编号,b.料号) e " & _
    " ON  e.单据编号 = b.单据编号 AND  e.delivery in (select distinct purchase from erptemp.dbo.ksexport_temp where 1 = 1) " & _
    " where e.单据编号 = b.单据编号 AND e.delivery in (select distinct purchase from erptemp.dbo.ksexport_temp where 1 = 1) AND c.物料编号 = b.物料编号 and e.料号 = c.料号 and c.料号 not in (select distinct 料号 from erptemp.dbo.ksexport where 出货单据 = e.单据编号 and flag = '0' )" & _
    " group by b.单据编号,c.料号,CONVERT(varchar(100), b.操作日期, 23),e.delivery,e.总价 "
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType5(rs)
    Else
            
        If Get_SqlserverCnt("select 出货单据  from erptemp.dbo.ksexport where 发票号 in (select distinct purchase from erptemp.dbo.ksexport_temp where 1 = 1)") <> 0 Then
            
                    
             MsgBox "此笔已经新增过,请核实！", vbInformation, "提示"
             Exit Sub
                    
         Else
         
         
             MsgBox "查询不到该出口单据信息", vbInformation, "提示"
             Exit Sub
        
                    
        End If
        
       
    End If
    
    Toolbar1.Buttons(3).Caption = "提交"
    Toolbar1.Buttons(3).Image = 6
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
   
    With fpS(0)
            
        strtet = 0
        
        strval = 0
        
        For i = 1 To .MaxRows
            .Row = i
            .LockBackColor = vbYellow
            
            .Col = E_FPS.E_gx
            
            If .text = 1 Then

                .Col = E_FPS.E_totalprice
                        Dim je As Long
                        je = Val(.text)
                .Col = E_FPS.E_exportquantity
                        If je < 0 Then
                            strtet = Val(strtet) - Val(.text)
                        Else
                            strtet = Val(strtet) + Val(.text)
                        End If
                .Col = E_FPS.E_totalprice

                    strval = Val(.text) + Val(strval)
                
                
            
            End If
            
            .Col = E_FPS.E_gx
            .Lock = False

            .Col = E_FPS.E_declarationno
            .Lock = False
            
            .Col = E_FPS.E_modetrade
            If .text = "进料对口" Or .text = "进料成品退换" Or .text = "进料料件复出" Then
                
                .Col = E_FPS.E_itemno
                .Lock = False
            
            Else
            
                .Col = E_FPS.E_name
                .Lock = False
                
                .Col = E_FPS.E_UNIT
                .Lock = False

            End If
            
            
            .Col = E_FPS.E_currency
            .Lock = False
            
            .Col = E_FPS.E_totalprice
            .Lock = False
            
            For m = E_FPS.E_AWB To E_FPS.E_mark
            
                .Col = m
                .Lock = False
      
            Next

'
'            strSql = "select distinct 手册编号 from erptemp.dbo.ksmanual where 1 = 1 and flag = '2'"
'
'            If Rs.State = 1 Then Rs.Close
'            Rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
'
'            .Col = 10
'
'            .CellType = CellTypeComboBox
'
'           ' .TypeComboBoxList = ""
'
'            Rs.MoveFirst
'
'            For j = 1 To Rs.RecordCount
'
'                .TypeComboBoxList = .TypeComboBoxList & Rs("手册编号")
'                Rs.MoveNext
'            Next
'
'            Rs.Clone
'
'            Set Rs = Nothing
'
        Next

    End With
    
    lb7 = "出货总量"
    
    lb6.Visible = True
    
    lb6 = "出货总额"
    
    lb7.Visible = True
    
    Text4.Visible = True
    
    Text5.Visible = True
  
    strtet = Format(Trim$(strtet), "0.000")
            
    strval = Format(Trim$(strval), "0.0000")
        
    Text4.text = strval
    
    Text5.text = strtet
    
    AddSql2 ("delete from erptemp.dbo.ksexport_temp where 1 = 1")
    
End Sub

Private Sub AddType6()

    Dim rs     As New ADODB.Recordset

    Dim i      As Integer
    
    Dim j      As Integer

    Dim m      As Integer
    
    Dim id     As Integer
    
    Dim strInv As String

    Dim strsql As String
    
    Dim a()    As String
    
    Dim leni   As Integer
    
    Dim stridd As String

    If Text1.text = "" Then
    
        MsgBox "请填写要维护的采购单号", vbInformation, "提示"
        Exit Sub

    End If
    
    fpS(0).MaxRows = 0
    
    stridd = Createid

    strInv = Trim$(Text1.text)
    
    a = Split(strInv, "/")
    
    leni = UBound(a) - LBound(a) + 1
    
    AddSql2 ("delete from erptemp.dbo.ksimport_temp where 1 = 1")

    For i = 0 To leni - 1
    
        If Get_SqlserverCnt("SELECT * FROM erpbase..tblCPurDataSub WHERE 采购单编号 = '" & a(i) & "'") = 0 Then
            MsgBox "没有此采购单号" & a(i) & ",请重新输入", vbInformation, "提示"
            Exit Sub
    
        End If
        
        AddSql2 (" insert into erptemp.dbo.ksimport_temp values('" & a(i) & "') ")

    Next
    
    strsql = "SELECT '' as '√','" & stridd & "' as 批次,a.采购单编号,b.料号,b.规格型号 as 型号,'' AS 类别,ceiling(sum(a.批准采购数量) - isnull(c.订单数量,0)) as 订单数量,t7.qty1 as 标准die," & _
    "ceiling(sum(a.批准采购数量) - isnull(c.订单数量,0))* t7.qty1 as 总die数,'' as 手册编号,'' as 项号,'' as 品名,'0' as 报关量,'' as 计量单位,'' as 入场日期,'' as 发票号,'' as 件数,'USD' as 币别,a.单价 as 采购单价,((sum(a.批准采购数量) - isnull(c.订单数量,0))) * a.单价 as 报关金额," & _
    " '' as 汇率,'' as 关税率,'' as 关税,'' as 增值税率,'' as 增值税,'' as 报关单号,'' as AWB#,'' as 货代,'' as 退单日期,'' as 备注," & _
    " '' as id FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b  left join (select 采购单号,料号,isnull(sum(订单数量),0) as 订单数量, " & _
    " flag from erptemp.dbo.ksimport where flag = '0' group by 采购单号,料号,flag) c on c.料号 = b.料号 and flag = '0' and c.采购单号 in (select distinct purchase from erptemp.dbo.ksimport_temp " & _
    " where 1 = 1) left join (select t1.采购单编号,t6.料号,isnull(t8.qty,0) as qty1 ,t1.请购单编号,t1.请购单项次 ,t1.采购单项次  from  erpbase..tblCPurDataSub t1  inner join  erpdata..tblSmainM2 t6  on t1.物料编号 = t6.物料编号  " & _
    " left join  (select m2.料号,max(m1.QTECHDIEQTY) as qty  from erptemp..TBLTSVNPIPRODUCT m1,erpdata..TSVtblMRuleData m2 where 1=1  " & _
    " and m1.QTECHPTNO2 = m2.工序号 group by m2.料号) t8 on t8.料号 = t6.料号 where 1=1 ) t7  on t7.料号 = b.料号  and  " & _
    " t7.采购单编号 in (select distinct purchase from erptemp.dbo.ksimport_temp where 1 = 1) " & _
    " WHERE a.是否禁用 = '0' and a.采购单编号 in (select distinct purchase from erptemp.dbo.ksimport_temp where 1 = 1) and t7.采购单编号 = a.采购单编号  AND t7.请购单编号 = a.请购单编号  AND t7.请购单项次 = a.请购单项次  AND t7.采购单项次 = a.采购单项次 " & _
    " and a.物料编号 = b.物料编号 GROUP by a.单价,b.规格型号,a.采购单编号,b.料号,c.订单数量,t7.qty1,a.采购单项次 order by a.采购单项次 "
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
    Call ListDataType6(rs)
    
    fpss(0).Visible = True
    
    strsql = "SELECT  采购单号,料号,isnull(sum(订单数量),0) as 已收数量 from erptemp.dbo.ksimport where 1 = 1 and flag = '0' and 采购单号 in (select distinct purchase from erptemp.dbo.ksimport_temp where 1 = 1) group by 采购单号,料号"
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
    Call ListDataType1(rs)
    
    Toolbar1.Buttons(3).Caption = "提交"
    Toolbar1.Buttons(3).Image = 6
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(7).Enabled = False

    With fpS(0)
        
        strtet = 0
        
        strval = 0
        
        For i = 1 To .MaxRows
            .Row = i
            .LockBackColor = vbYellow
            
            .Col = F_fp.F_gx
            .Lock = False
            
            .Col = F_fp.F_gx
            If .text = 1 Then
            
                .Col = F_fp.F_baoguanqty
                
                strtet = Val(.text) + Val(strtet)
                
                .Col = F_fp.F_baoguanvalue
                
                strval = Val(.text) + Val(strval)
            
            End If
        
            
            For m = 5 To 8
            
                .Col = m
                .Lock = False
                
            Next
            
            .Col = 10
            .Lock = False
            
             .Col = 11
            .Lock = False
            
             .Col = 13
            .Lock = False

            For m = 15 To 18
            
                .Col = m
                .Lock = False
      
            Next
                
            For m = 20 To 22
            
                .Col = m
                .Lock = False
      
            Next
                           
            For m = 26 To 30
            
                .Col = m
                .Lock = False
      
            Next
            
            strsql = "select distinct 手册编号 from erptemp.dbo.ksmanual where 1 = 1 and flag = '1'"

            If rs.State = 1 Then rs.Close
            rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

            .Col = F_fp.F_manualno

            .CellType = CellTypeComboBox

            rs.MoveFirst

            For j = 1 To rs.RecordCount

                .TypeComboBoxList = .TypeComboBoxList & rs("手册编号")
                rs.MoveNext
            Next
        
            rs.Clone
        
            Set rs = Nothing
            
        Next

    End With
    
    lb7 = "报关总量"
    
    lb6.Visible = True
    
    lb6 = "报关总额"
    
    lb7.Visible = True
    
    Text4.Visible = True
    
    Text5.Visible = True
  
    strtet = Format(Trim$(strtet), "0.0000")
            
    strval = Format(Trim$(strval), "0.0000")
        
    Text4.text = strval
    
    Text5.text = strtet
    
    AddSql2 ("delete from erptemp.dbo.ksimport_temp where 1 = 1")
    

End Sub


Private Sub ForCommit1()

    Dim strInv1     As String

    Dim strInv2     As String

    Dim strInv3     As String
    
    Dim strInv4     As String

    Dim strInv5     As String

    Dim strInv6     As String

    Dim strInv7     As String

    Dim strInv8     As String

    Dim strInv9     As String

    Dim strInv10    As String

    Dim strInv11    As String

    Dim strInv12    As String

    Dim strInv13    As String

    Dim strInv14    As String

    Dim strInv15    As String

    Dim strInv16    As String
    
    Dim strInv17    As String
    
    Dim strInv18    As String
    
    Dim strInv19    As String
    
    Dim strInv20    As String
    
    Dim strInv21    As String

    Dim strsql      As String

    Dim stritemname As String

    Dim strunit     As String

    Dim i           As Integer

    Dim j           As Integer

    Dim bFlag       As Boolean
    
    Dim stridd      As String

    bFlag = False
            
    stridd = Createid

    With fpS(0)
    
        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
        
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_FPS.E_gx
    
            j = 0
                
            If .text = "1" Then
                
                j = j + 1
                bFlag = True
                
                .Col = E_FPS.e_NO
                '批次
                strInv21 = Trim$(.text)
                
                If Trim$(stridd) <> Trim$(strInv21) Then
                
                    MsgBox "批次有变动,更改为 " & strInv21 & "", vbInformation, "提示"
                    
'                    strInv21 = stridd
                
                End If
                
                .Col = E_FPS.E_exportno
                
                Select Case Combo1.text
                
                
                    Case "出口明细表"
                        
                        If .text = "" Then
                            
                            MsgBox "请输入出货单据", vbInformation, "提示"
                            Exit Sub

                        End If
                    
                    Case "出口明细表(特殊)"
                            
                        .text = ""
                    
                End Select
    
                strInv1 = Trim$(.text)
    
                .Col = E_FPS.E_partno

                If .text = "" Then
                    MsgBox "请输入料号", vbInformation, "提示"
                    Exit Sub

                End If
    
                strInv2 = Trim$(.text)
                
                Select Case Combo1.text
                
                
                    Case "出口明细表"
                        
                       If Get_SqlserverCnt("select * from erptemp.dbo.ksexport where 出货单据 = '" & strInv1 & "' and 料号 = '" & strInv2 & "' and flag = '0'") > 0 Then
                        
                            MsgBox "该笔资料已经新增过", vbInformation, "提示"
                            Exit Sub

                       End If
                    
                End Select
                
                .Col = E_FPS.E_modetrade
                
                If .text = "" Then
                
                    MsgBox "请选择类别", vbInformation, "提示"
                    Exit Sub

                End If
                
                strInv3 = Trim$(.text)
                
                .Col = E_FPS.e_Invoice
                
                strInv4 = Trim$(.text)
        
                .Col = E_FPS.E_exportdate
                
                If .text = "" Then
                
                    MsgBox "请输入出货日期", vbInformation, "提示"
                    Exit Sub

                End If
                
                strInv5 = Trim$(.text)
        
                .Col = E_FPS.E_exportquantity
                
                If .text = "" Then
                    
                    MsgBox "请输入出货数量", vbInformation, "提示"
                    Exit Sub
                    
                End If
                
                strInv6 = Format(Trim$(.text), "0.0000")
        
                .Col = E_FPS.E_declarationno
                strInv7 = Trim$(.text)
        
                .Col = E_FPS.E_manualno
                
                strInv8 = Trim$(.text)
        
                .Col = E_FPS.E_itemno
                
                strInv9 = Trim$(.text)
        
                .Col = E_FPS.E_name
                
                If strInv3 = "进料对口" Or strInv3 = "进料成品退换" Or strInv3 = "进料料件复出" Then
                '品名
                    If strInv3 = "进料对口" Or strInv3 = "进料成品退换" Then
                    
                        If Get_SqlserverCnt("SELECT 商品名称 FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & strInv8 & "' and flag = '2' and 序号= '" & strInv9 & "'") = 0 Then
                                    
                            MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"
                                    
                            Exit Sub
                                
                        End If
                    
                        stritemname = Get_SqlStr("select distinct 商品名称 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '2' and  序号 = '" & strInv9 & "'")

                    Else
                    
                        If Get_SqlserverCnt("SELECT 商品名称 FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & strInv8 & "' and flag = '1' and 序号= '" & strInv9 & "'") = 0 Then
                                    
                            MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"
                                    
                            Exit Sub
                                
                        End If
                    
                        stritemname = Get_SqlStr("select distinct 商品名称 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '1' and  序号 = '" & strInv9 & "'")
                    
                    End If
                    
                    .text = stritemname


                End If
                
                strInv10 = Trim$(.text)
        
                .Col = E_FPS.E_UNIT
                
                '计量单位
                If strInv3 = "进料对口" Or strInv3 = "进料成品退换" Or strInv3 = "进料料件复出" Then
                    
                    If strInv3 = "进料对口" Or strInv3 = "进料成品退换" Then
                
                        strunit = Get_SqlStr("select distinct 计量单位 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '2' and  序号 = '" & strInv9 & "'")

                    Else
                    
                        strunit = Get_SqlStr("select distinct 计量单位 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '1' and  序号 = '" & strInv9 & "'")
                        
                    End If
                    
                    .text = strunit
                    
                End If

                strInv11 = Trim$(.text)
        
                .Col = E_FPS.E_currency
                strInv12 = Trim$(.text)
        
                .Col = E_FPS.E_totalprice
                '总价
                If .text = "" Then
                    MsgBox "请输入总价", vbInformation, "提示"
                    Exit Sub

                End If
    
                strInv13 = Trim$(.text)
        
                .Col = E_FPS.E_unitprice
                
                If strInv13 <> "" And .text = "" Then
                    
                    .text = Format(Val(strInv13) / Val(strInv6), "0.000000")
                
                End If
                '单价
                
                strInv14 = Trim$(.text)
        
                .Col = E_FPS.E_AWB
                strInv15 = Trim$(.text)
        
                .Col = E_FPS.E_destination
                strInv16 = Trim$(.text)
                
                .Col = E_FPS.E_freight
                strInv17 = Trim$(.text)
                
                .Col = E_FPS.E_chargebackdate
                
                Select Case Combo1.text
                
                
                    Case "出口明细表(特殊)"
                        
                        .text = ""
                        
                End Select
                
                strInv18 = Trim$(.text)
                
                .Col = E_FPS.E_mark
                
                Select Case Combo1.text
                
                
                    Case "出口明细表(特殊)"
                    
                        .text = ""
                    
                End Select
                
                strInv19 = Trim$(.text)
                
                .Col = E_FPS.e_ID
                strInv20 = Trim$(i)
                
                
                AddSql2 ("insert into erptemp.dbo.ksexport( 批次,出货单据,料号,类别,发票号,出货日期,数量,报关单号,手册编号,手册项号,品名,计量单位,币别,总价,单价,AWB#,目的地,货代,退单日期,备注,键入时间,修改状态,修改时间,删除时间,flag,id) values('" & strInv21 & "','" & strInv1 & "','" & strInv2 & "','" & strInv3 & "','" & strInv4 & "','" & strInv5 & "','" & strInv6 & "','" & strInv7 & "','" & strInv8 & "','" & strInv9 & "','" & strInv10 & "','" & strInv11 & "','" & strInv12 & "','" & strInv13 & "','" & strInv14 & "','" & strInv15 & "','" & strInv16 & "','" & strInv17 & "','" & strInv18 & "','" & strInv19 & "',GetDate(),NULL,NULL,NULL,'0','" & strInv20 & "')")

            End If

        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "请选择要新增的行", vbInformation, "提示"
            Exit Sub
            
        End If

    End With
    
    MsgBox "新增成功", vbInformation, "提示"
    Toolbar1.Buttons(3).Caption = "新增"
    Toolbar1.Buttons(3).Image = 3
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    
    
    stridid = stridd
    
    strstate = True
    
    ForQuery

End Sub

Private Sub ForCommit2()

    Dim strInv1     As String

    Dim strInv2     As String

    Dim strInv3     As String
    
    Dim strInv4     As String

    Dim strInv5     As String

    Dim strInv6     As String

    Dim strInv7     As String

    Dim strInv8     As String

    Dim strInv9     As String

    Dim strInv10    As String

    Dim strInv11    As String
    
    Dim strInv12    As String

    Dim strInv13    As String

    Dim strInv14    As String

    Dim strInv15    As String

    Dim strInv16    As String
    
    Dim strInv17    As String
    
    Dim strInv18    As String
    
    Dim strInv19    As String
    
    Dim strInv20    As String
    
    Dim strInv21    As String
    
    Dim strInv22    As String
    
    Dim strInv23    As String

    Dim strInv24    As String
    
    Dim strInv25    As String

    Dim strInv26    As String
    
    Dim strInv27    As String

    Dim strInv28    As Integer
    
    Dim strInv29    As String
    
    Dim strInv30    As String
    
    Dim strunit     As String

    Dim strNo1      As String

    Dim strNo2      As String

    Dim strNo3      As String
    
    Dim stritemname As String
    
    Dim strbaono1   As Double
    
    Dim strbaono2   As Double
    
    Dim strbaono3   As Double

    Dim strsql      As String
    
    Dim stridd      As String

    Dim i           As Integer

    Dim j           As Integer

    Dim bFlag       As Boolean

    bFlag = False
    
    stridd = Createid
    
    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
    
            .Row = i
            .Col = 1
    
            j = 0

            If .text = "1" Then
            
                j = j + 1
                bFlag = True
                
                .Col = F_fp.F_no

                strInv29 = Trim$(.text)
                
                If Trim$(stridd) <> Trim$(strInv29) Then
                
                    MsgBox "批次有变动,更改为 " & strInv29 & "", vbInformation, "提示"
                    
'                    strInv29 = stridd

                End If
                
                .Col = F_fp.F_purchaseno
                
                Select Case Combo1.text
                
                    Case "进口明细表(特殊)"
                    
                        .text = ""
                        
                    Case "进口明细表"
                    
                        If .text = "" Then
                            
                            MsgBox "请输入采购单号", vbInformation, "提示"
                            Exit Sub

                        End If
                    
                End Select
    
                strInv1 = Trim$(.text)
    
                .Col = F_fp.F_partno

                strInv2 = Trim$(.text)
                
                .Col = F_fp.F_modelno
                
                strInv3 = Trim$(.text)
    
                .Col = F_fp.F_modetrade
                
                If .text = "" Then
                    MsgBox "请输入类别", vbInformation, "提示"
                    Exit Sub

                End If
                
                strInv4 = Trim$(.text)
                              
                .Col = F_fp.F_orderqty
                
                Select Case Combo1.text
                
                    Case "进口明细表(特殊)"
                    
                        .text = 0
                        
                    Case "进口明细表"
                        
                        strInv5 = Trim$(.text)
                
                        strNo1 = Get_SqlStr("SELECT isnull(SUM(a.批准采购数量),0) FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b WHERE a.采购单编号 = '" & strInv1 & "' and a.物料编号 = b.物料编号 and b.料号 = '" & strInv2 & "' ")
                    
                        strNo2 = Get_SqlStr("SELECT isnull(SUM(订单数量),0) FROM erptemp.dbo.ksimport where 采购单号 = '" & strInv1 & "' and 料号 = '" & strInv2 & "' and flag = '0'")
                    
                        strNo3 = Val(strNo1) - Val(strNo2)
                    
                        If Val(strInv5) > Val(strNo3) Then
                            
                            MsgBox "该笔料号" & strInv2 & "批准采购数量: " & strNo1 & ",已经维护订单数量：" & strNo2 & ",最大数量只能维护：" & strNo3 & "", vbInformation, "提示"
                            Exit Sub

                        End If
                    
                        If Val(strInv5) <= 0 Then
                    
                            MsgBox "订单数量不可小于等于0", vbInformation, "提示"
                            Exit Sub
                        
                        End If
            
                End Select
                
                strInv5 = Format(Trim$(.text), "0.00")
                
                .Col = F_fp.F_die
                '标准die
                
                Select Case Combo1.text
                
                    Case "进口明细表(特殊)"
                    
                        If .text = "" Then
                    
                            .text = 0
                    
                        End If

                End Select
                
                strInv6 = Format(Trim$(.text), "0.00")
        
                .Col = F_fp.F_totaldie
                '总die数量
                
                Select Case Combo1.text
                
                    Case "进口明细表(特殊)"
                    
                        If .text = "" Then
                    
                            .text = 0
                    
                        End If
                         
                        strInv7 = Format(Trim$(.text), "0.00")
                    
                    Case "进口明细表"
                        
                        strInv7 = Val(strInv5) * Val(strInv6)

                End Select
        
                .Col = F_fp.F_manualno
                '手册号
                strInv8 = Trim$(.text)
                
                .Col = F_fp.F_itemno
                '项号
                
                strInv9 = Trim$(.text)
        
                .Col = F_fp.F_name

                '品名
                If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                
                    If Get_SqlserverCnt("SELECT 商品名称 FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & strInv8 & "' and flag = '1' and 序号= '" & strInv9 & "'") = 0 Then
                                    
                        MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"

                        Exit Sub
                    
                    End If
                                
                    stritemname = Get_SqlStr("select distinct 商品名称 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '1' and  序号 = '" & strInv9 & "'")

                    .text = stritemname

                End If
                
                strInv10 = Trim$(.text)
                
                .Col = F_fp.F_baoguanqty
                
                Select Case Combo1.text
                
                    Case "进口明细表(特殊)"
                    
                        If .text = "" Then
                    
                            MsgBox "请输入报关数量", vbInformation, "提示"
                            Exit Sub
                    
                        End If
                    
                        If Val(.text) <= 0 Then
                        
                            MsgBox "报关数量不可小于等于0", vbInformation, "提示"
                            Exit Sub
                    
                        End If
                
                        strInv11 = Format(Trim$(.text), "0.0000")
                
                    Case "进口明细表"

                        '报关数量
                        If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                
                            strbaono1 = Get_SqlStr("select isnull(申报数量,0) from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '1' and  序号 = '" & strInv9 & "'")
                
                            strbaono2 = Get_SqlStr("select isnull(sum(报关量),0) from erptemp.dbo.ksimport where flag = '0' and  采购单号 = ' " & strInv1 & "' and 料号 = '" & strInv2 & "'")
                
                            strbaono3 = strbaono1 - strbaono2
                
                            If .text = "" Then
                
                                MsgBox "请输入报关数量", vbInformation, "提示"
                                Exit Sub
                
                            End If
                
                            If Val(.text) <= 0 Then
                
                                MsgBox "报关数量不可小于等于0", vbInformation, "提示"
                                Exit Sub
                
                            End If
                
                            strInv11 = Format(Trim$(.text), "0.000")
                
                            If Val(strInv11) > Val(strbaono3) Then
                
                                MsgBox "输入的报关量超过可输入的范围,申报数量为" & strbaono1 & ",目前系统已录入数量 " & strbaono2 & "", vbInformation, "提示"
                
                            End If
                
                        Else
                
                            If .text = "" Then
                    
                                MsgBox "请输入报关数量", vbInformation, "提示"
                                Exit Sub
                    
                            End If
                    
                            If Val(.text) <= 0 Then
                        
                                MsgBox "报关数量不可小于等于0", vbInformation, "提示"
                                Exit Sub
                    
                            End If
                
                            strInv11 = Format(Trim$(.text), "0.000")
                    
                        End If

                End Select
                
                .Col = F_fp.F_unit

                '计量单位
                If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                
                    strunit = Get_SqlStr("select distinct 计量单位 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '1' and  序号 = '" & strInv9 & "'")

                    .text = strunit

                End If
                
                strInv12 = Trim$(.text)
            
                .Col = F_fp.F_indate
                If Trim$(.text) <> "" And Len(Trim$(.text)) <> 8 Then
                    MsgBox "进场日期请以YYYYMMDD格式填写,如20200501", vbInformation, "提示"
                    Exit Sub
                End If
                '入场日期
                strInv13 = Trim$(.text)
              
                .Col = F_fp.F_invoice
                '发票号
                strInv14 = Trim$(.text)
                
                .Col = F_fp.F_caseqty
                '件数
                strInv15 = Trim$(.text)
                
                .Col = F_fp.F_currency
                '币别
                strInv16 = Trim$(.text)
                
                .Col = F_fp.F_unitprice

                '采购单价
                If .text = "" Then
                    
                    .text = 0
                    
                End If

                strInv30 = Format(Trim$(.text), "0.000")
                
                .Col = F_fp.F_baoguanvalue

                '报关金额
                If .text = "" Then
                    
                    .text = 0
                    
                End If
                
                strInv17 = Format(Trim$(.text), "0.0000")
                
                Select Case Combo1.text
                
                    Case "进口明细表(特殊)"
    
                            strInv30 = Format(Trim$(Val(strInv17) / Val(strInv11)), "0.000")

                End Select
                
                .Col = F_fp.F_rate

                '汇率
                If .text = "" Then
                    
                    .text = 0
                    
                End If

                strInv18 = Format(Trim$(.text), "0.0000")
    
                .Col = F_fp.F_tariffrate

                '关税率
                If .text = "" Then
                    
                    .text = 0
                    
                End If

                strInv19 = Format(Trim$(.text), "0.0000")
                    
                .Col = F_fp.F_tariff
                '关税
                
                .text = Val(strInv18) * Val(strInv17) * Val(strInv19)
                    
                strInv20 = Format(Trim$(.text), "0.00")
                    
                .Col = F_fp.F_addtaxrate

                '增值税率
                If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                
                    .text = 0
                    strInv21 = Format(Trim$(.text), "0.0000")
                    
                Else
                    .text = 0.13
                    strInv21 = Format(Trim$(.text), "0.0000")
                    
                End If

                .Col = F_fp.F_addtax
                '增值税=（关税+货值*汇率）*0.16
                    
                .text = Val(strInv20) * Val(strInv21) + Val(strInv17) * Val(strInv21) * Val(strInv18)
                    
                strInv22 = Format(Trim$(.text), "0.00")
                            
                .Col = F_fp.F_declarationno
                '报关单号
                
                strInv23 = Trim$(.text)
                
                .Col = F_fp.F_awb
                'AWB#
                
                strInv24 = Trim$(.text)
                
                .Col = F_fp.F_freight
                '货代
        
                strInv25 = Trim$(.text)
                
                .Col = F_fp.F_chargebackdate
                '退单日期
                
                strInv26 = Trim$(.text)
                
                .Col = F_fp.F_mark
                '备注
                strInv27 = Trim$(.text)
                
                .Col = F_fp.F_id
                
                strInv28 = Trim$(i)
                
                AddSql2 ("insert into erptemp.dbo.ksimport( 批次,采购单号,料号,型号,类别,订单数量,标准die,总die数,手册编号,项号,品名,报关量,计量单位,入场日期,发票号,件数,币别,采购单价,报关金额,汇率,关税率,关税,增值税率,增值税,报关单号,AWB#,货代,退单日期,备注,id,键入时间,修改状态,修改时间,删除时间,flag) values('" & strInv29 & "','" & strInv1 & "','" & strInv2 & "','" & strInv3 & "','" & strInv4 & "','" & strInv5 & "','" & strInv6 & "','" & strInv7 & "','" & strInv8 & "','" & strInv9 & "','" & strInv10 & "','" & strInv11 & "','" & strInv12 & "','" & strInv13 & "','" & strInv14 & "','" & strInv15 & "','" & strInv16 & "','" & strInv30 & "','" & strInv17 & "','" & strInv18 & "','" & strInv19 & "','" & strInv20 & "','" & strInv21 & "','" & strInv22 & "','" & strInv23 & "','" & strInv24 & "','" & strInv25 & "','" & strInv26 & "','" & strInv27 & "','" & strInv28 & "',GetDate(),NULL,NULL,NULL,'0')")
            
            End If
            
        Next
        
        'j = 0 获取不到用户需要输入的资料
        If bFlag = False And j = 0 Then
            MsgBox "请选择要新增的行", vbInformation, "提示"
            Exit Sub
            
        End If

    End With
    
    MsgBox "新增成功", vbInformation, "提示"
    Toolbar1.Buttons(3).Caption = "新增"
    Toolbar1.Buttons(3).Image = 3
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    
    stridid1 = stridd
    
    strstate1 = True

    ForQuery

End Sub

Private Sub ForMod1()

    Dim rs          As New ADODB.Recordset

    Dim i           As Integer

    Dim m           As Integer

    Dim j           As Integer

    Dim strInv1     As String

    Dim strInv2     As String

    Dim strInv3     As String

    Dim strInv4     As String

    Dim strInv5     As String

    Dim strInv6     As String

    Dim strInv7     As String

    Dim strInv8     As String

    Dim strInv9     As String

    Dim strInv10    As String

    Dim strInv11    As String

    Dim strInv12    As String

    Dim strInv13    As String

    Dim strInv14    As String

    Dim strInv15    As String

    Dim strInv16    As String
    
    Dim strInv17    As String
    
    Dim strInv18    As String
    
    Dim strInv19    As String
    
    Dim strInv20    As String
    
    Dim strInv21    As String
    
    Dim stritemname As String
    
    Dim strunit     As String

    Dim strtime     As String
    
    Dim strsql      As String
    
    Dim bFlag       As Boolean

    If Toolbar1.Buttons(5).Caption <> "提交" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
                .Col = E_FPS.E_gx
                .Lock = False
    
                For m = E_FPS.E_modetrade To E_FPS.E_freight
            
                    .Col = m
                    .Lock = False
      
                Next
                
                .Col = E_FPS.E_modetrade
                If .text = "进料对口" Or .text = "进料成品退换" Or .text = "进料料件复出" Then
                
                            .Col = E_FPS.E_manualno
                            .Lock = False
                    
                            .Col = E_FPS.E_itemno
                            .Lock = False
                        
                            .Col = E_FPS.E_name
                            .Lock = True
                
                            .Col = E_FPS.E_UNIT
                            .Lock = True
                    
                Else
                    
                            .Col = E_FPS.E_manualno
                            .Lock = True
                    
                            .Col = E_FPS.E_itemno
                            .Lock = True
                    
                            .Col = E_FPS.E_name
                            .Lock = False
                
                            .Col = E_FPS.E_UNIT
                            .Lock = False
        
                End If
                
                            
                .Col = E_FPS.E_modetrade
                .CellType = CellTypeComboBox

                .TypeComboBoxList = .TypeComboBoxList & "进料对口"
    
                .TypeComboBoxList = .TypeComboBoxList & "一般贸易"

                .TypeComboBoxList = .TypeComboBoxList & "其他进出口免费"

                .TypeComboBoxList = .TypeComboBoxList & "进料料件复出"
    
                .TypeComboBoxList = .TypeComboBoxList & "进料成品退换"

                .TypeComboBoxList = .TypeComboBoxList & "修理物品"

                .TypeComboBoxList = .TypeComboBoxList & "设备退运"

                .TypeComboBoxList = .TypeComboBoxList & "其他"
    
                 
                .Col = E_FPS.E_currency
                .CellType = CellTypeComboBox
            
                .TypeComboBoxList = .TypeComboBoxList & "USD"
            
                .TypeComboBoxList = .TypeComboBoxList & "JPY"

                .TypeComboBoxList = .TypeComboBoxList & "EUR"

                .TypeComboBoxList = .TypeComboBoxList & "RMB"
                                
                .Col = E_FPS.E_exportquantity
            
                .text = Format(Trim$(.text), "0.000")
                
                .LockBackColor = vbYellow
                
                strsql = "select distinct 手册编号 from erptemp.dbo.ksmanual where 1 = 1 and flag = '2'"

                If rs.State = 1 Then rs.Close
                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                .Col = E_FPS.E_manualno

                .CellType = CellTypeComboBox

     '           .TypeComboBoxList = ""

                rs.MoveFirst

                For j = 1 To rs.RecordCount

                    .TypeComboBoxList = .TypeComboBoxList & rs("手册编号")
                    rs.MoveNext
                Next
        
                rs.Clone
        
                Set rs = Nothing
    
            Next
        
        End With
    
        Toolbar1.Buttons(5).Caption = "提交"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    bFlag = False
    
    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
           
            .Col = E_FPS.E_gx
    
            j = 0

            If .text = "1" Then
            
                j = j + 1
                bFlag = True
                
                .Col = E_FPS.e_NO
                strInv21 = Trim$(.text)
                
                .Col = E_FPS.E_exportno
                
                strInv1 = Trim$(.text)
    
                .Col = E_FPS.E_partno
                strInv2 = Trim$(.text)
    
                .Col = E_FPS.E_modetrade
                
                If .text = "" Then
                    MsgBox "请输入类别", vbInformation, "提示"
                    Exit Sub

                End If
                
                strInv3 = Trim$(.text)
                
                .Col = E_FPS.e_Invoice
                strInv4 = Trim$(.text)
        
                .Col = E_FPS.E_exportdate
                
                strInv5 = Trim$(.text)
        
                .Col = E_FPS.E_exportquantity
                strInv6 = Trim$(.text)
        
                .Col = E_FPS.E_declarationno
                strInv7 = Trim$(.text)
        
                .Col = E_FPS.E_manualno
                strInv8 = Trim$(.text)
        
                .Col = E_FPS.E_itemno
                strInv9 = Trim$(.text)
        
                .Col = E_FPS.E_name
                
                If strInv3 = "进料对口" Or strInv3 = "进料成品退换" Or strInv3 = "进料料件复出" Then
                '品名
                    If strInv3 = "进料对口" Or strInv3 = "进料成品退换" Then
                        
                        If Get_SqlserverCnt("SELECT 商品名称 FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & strInv8 & "' and flag = '2' and 序号= '" & strInv9 & "'") = 0 Then
                                    
                            MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"
                                    
                            Exit Sub
                                
                        End If
                    
                        stritemname = Get_SqlStr("select distinct 商品名称 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '2' and  序号 = '" & strInv9 & "'")
                    
                    Else
                    
                        If Get_SqlserverCnt("SELECT 商品名称 FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & strInv8 & "' and flag = '1' and 序号= '" & strInv9 & "'") = 0 Then
                                    
                            MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"
                                    
                            Exit Sub
                                
                        End If
                    
                        stritemname = Get_SqlStr("select distinct 商品名称 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '1' and  序号 = '" & strInv9 & "'")
                    
                    End If
                    
                    .text = stritemname

                End If
                
                strInv10 = Trim$(.text)
        
                .Col = E_FPS.E_UNIT
                
                '计量单位
                If strInv3 = "进料对口" Or strInv3 = "进料成品退换" Or strInv3 = "进料料件复出" Then
                
                    If strInv3 = "进料对口" Or strInv3 = "进料成品退换" Then
                    
                        strunit = Get_SqlStr("select distinct 计量单位 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '2' and  序号 = '" & strInv9 & "'")
                    
                    Else
                        
                        strunit = Get_SqlStr("select distinct 计量单位 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '1' and  序号 = '" & strInv9 & "'")
                    
                    End If
                    
                    .text = strunit
                
                End If
                
                strInv11 = Trim$(.text)
        
                .Col = E_FPS.E_currency
                strInv12 = Trim$(.text)
        
                .Col = E_FPS.E_totalprice
                '总价
                
                If .text = "" Then
                
                    MsgBox "请输入总价", vbInformation, "提示"
                    Exit Sub

                End If
                strInv13 = Trim$(.text)
        
                .Col = E_FPS.E_unitprice
                
                If strInv13 <> "" And .text = "" Then
                    
                    .text = Val(strInv13) / Val(strInv6)
                
                End If
                
                strInv14 = Trim$(.text)
        
                .Col = E_FPS.E_AWB
                strInv15 = Trim$(.text)
        
                .Col = E_FPS.E_destination
                strInv16 = Trim$(.text)
                
                .Col = E_FPS.E_freight
                strInv17 = Trim$(.text)
                
                .Col = E_FPS.E_chargebackdate
                strInv18 = Trim$(.text)
                
                .Col = E_FPS.E_mark
                strInv19 = Trim$(.text)
                
                .Col = E_FPS.e_ID
                strInv20 = Trim$(.text)
    
                strtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
                
                AddSql2 ("insert into erptemp.dbo.ksexport (批次,出货单据,料号,类别,发票号,出货日期,数量,报关单号,手册编号,手册项号,品名,计量单位,币别,总价,单价,AWB#,目的地,货代,退单日期,备注,键入时间,修改状态,修改时间,删除时间,flag,id) SELECT 批次,出货单据,料号,类别,发票号,出货日期,数量,报关单号,手册编号,手册项号,品名,计量单位,币别,总价,单价,AWB#,目的地,货代,退单日期,备注,键入时间,'修改前',修改时间,删除时间,'2',id FROM erptemp.dbo.ksexport WHERE 批次 = '" & strInv21 & "' and id = '" & strInv20 & "' AND  flag = '0'")
                
                AddSql2 ("update erptemp.dbo.ksexport set 料号 =  '" & strInv2 & "', 类别 =  '" & strInv3 & "',报关单号 =  '" & strInv7 & "',手册编号 =  '" & strInv8 & "',手册项号 =  '" & strInv9 & "',品名 =  '" & strInv10 & "',计量单位 =  '" & strInv11 & "',币别 =  '" & strInv12 & "',总价 =  '" & strInv13 & "',单价 = '" & strInv14 & "',AWB# =  '" & strInv15 & "',目的地 =  '" & strInv16 & "',货代 =  '" & strInv17 & "',退单日期 =  '" & strInv18 & "',备注 =  '" & strInv19 & "',修改状态 = '修改后',修改时间 = '" & strtime & "' where 批次 = '" & strInv21 & "' and id = '" & strInv20 & "' and flag = '0'")
            
            End If
            
        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "请选择要修改的行", vbInformation, "提示"
            Exit Sub
            
        End If
    
    End With
    
    MsgBox "修改成功", vbInformation, "提示"

    Toolbar1.Buttons(5).Caption = "修改"
    Toolbar1.Buttons(5).Image = 4
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    
    ForQuery
    
End Sub
Private Sub ForMod2()

    Dim rs          As New ADODB.Recordset

    Dim i           As Integer

    Dim m           As Integer

    Dim j           As Integer

    Dim strInv1     As String

    Dim strInv2     As String

    Dim strInv3     As String
    
    Dim strInv4     As String

    Dim strInv5     As String

    Dim strInv6     As String

    Dim strInv7     As String

    Dim strInv8     As String

    Dim strInv9     As String

    Dim strInv10    As String

    Dim strInv11    As String
    
    Dim strInv12    As String

    Dim strInv13    As String

    Dim strInv14    As String

    Dim strInv15    As String

    Dim strInv16    As String
    
    Dim strInv17    As String
    
    Dim strInv18    As String
    
    Dim strInv19    As String
    
    Dim strInv20    As String
    
    Dim strInv21    As String
    
    Dim strInv22    As String
    
    Dim strInv23    As String

    Dim strInv24    As String
    
    Dim strInv25    As String

    Dim strInv26    As String
    
    Dim strInv27    As String

    Dim strInv28    As Integer
    
    Dim strInv29    As String
    
    Dim strInv30    As String
    
    Dim strunit     As String

    Dim strtime     As String
    
    Dim strNo1      As String

    Dim strNo2      As String

    Dim strNo3      As String
    
    Dim strsql      As String

    Dim stritemname As String
    
    Dim strbao1     As String
    
    Dim strbao2     As String
    
    Dim strbao3     As String
    
    Dim strbaono1   As Double
    
    Dim strbaono2   As Double
    
    Dim strbaono3   As Double
    
    Dim bFlag       As Boolean

    If Toolbar1.Buttons(5).Caption <> "提交" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
                
                .Col = F_fp.F_gx
                .Lock = False
                
                For m = F_fp.F_partno To F_fp.F_tariffrate
                    
                    If m = F_fp.F_orderqty Then
                    
                        .Col = m
                        .Lock = True
                        
                    Else
                        
                        .Col = m
                        .Lock = False
                    
                    End If
                    
                Next
            
                For m = F_fp.F_declarationno To F_fp.F_mark
            
                    .Col = m
                    .Lock = False
      
                Next
                
                .Col = F_fp.F_baoguanqty
                .text = Format(Trim$(.text), "0.000")
                
                .Col = F_fp.F_unitprice
                .text = Format(Trim$(.text), "0.000")
                
                .Col = F_fp.F_baoguanvalue
                .text = Format(Trim$(.text), "0.0000")
            
                .Col = F_fp.F_rate
                .text = Format(Trim$(.text), "0.0000")
            
                .Col = F_fp.F_tariffrate
                .text = Format(Trim$(.text), "0.0000")
            
                .Col = F_fp.F_tariff
                .text = Format(Trim$(.text), "0.00")
            
                .Col = F_fp.F_addtaxrate
                .text = Format(Trim$(.text), "0.0000")
            
                .Col = F_fp.F_addtax
                .text = Format(Trim$(.text), "0.00")
                
                .Col = F_fp.F_modetrade
                If .text = "进料对口" Or .text = "成品复进" Then
                
                    .Col = F_fp.F_manualno
                    .Lock = False
                        
                    .Col = F_fp.F_itemno
                    .Lock = False
                            
                    .Col = F_fp.F_name
                    .Lock = True
                    
                    .Col = F_fp.F_unit
                    .Lock = True
                            
                    .Col = F_fp.F_rate
                    .Lock = True
                        
                    .Col = F_fp.F_tariffrate
                    .Lock = True
                        
                    .Col = F_fp.F_tariff
                    .Lock = True
                    
                    .Col = F_fp.F_addtaxrate
                    .Lock = True
                        
                    .Col = F_fp.F_addtax
                    .Lock = True
                Else
                    
                    .Col = F_fp.F_manualno
                    .Lock = True
                        
                    .Col = F_fp.F_itemno
                    .Lock = True
                        
                    .Col = F_fp.F_name
                    .Lock = False
                    
                    .Col = F_fp.F_unit
                    .Lock = False
                        
                    .Col = F_fp.F_rate
                    .Lock = False
                        
                    .Col = F_fp.F_tariffrate
                    .Lock = False
                    
                    '关税
                    .Col = F_fp.F_tariff
                    .Lock = False
                    
                    '增值税
                    .Col = F_fp.F_addtax
                    .Lock = False
                        
                End If

                .LockBackColor = vbYellow
                
                strsql = "select distinct 手册编号 from erptemp.dbo.ksmanual where 1 = 1 and flag = '1'"

                If rs.State = 1 Then rs.Close
                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                .Col = F_fp.F_manualno

                .CellType = CellTypeComboBox

                rs.MoveFirst

                For j = 1 To rs.RecordCount

                    .TypeComboBoxList = .TypeComboBoxList & rs("手册编号")
                    rs.MoveNext
                    
                Next
        
                rs.Clone
        
                Set rs = Nothing
    
            Next
        
        End With
    
        Toolbar1.Buttons(5).Caption = "提交"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    bFlag = False
    
    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            
            .Col = F_fp.F_gx
    
            j = 0

            If .text = "1" Then
            
                j = j + 1
                bFlag = True
                
                .Col = F_fp.F_no
                
                strInv29 = Trim$(.text)
    
                .Col = F_fp.F_purchaseno
                
                strInv1 = Trim$(.text)
    
                .Col = F_fp.F_partno
                strInv2 = Trim$(.text)
    
                .Col = F_fp.F_modelno
                strInv3 = Trim$(.text)
                
                .Col = F_fp.F_modetrade
                
                If .text = "" Then
                    MsgBox "请输入类别", vbInformation, "提示"
                    Exit Sub

                End If
                
                strInv4 = Trim$(.text)
                
                .Col = F_fp.F_orderqty

                strInv5 = Trim$(.text)
    
                .Col = F_fp.F_die

                '标准die
                
                strInv6 = Trim$(.text)
                
                .Col = F_fp.F_totaldie
                '总die数量
                
                strInv7 = Trim$(.text)
        
                .Col = F_fp.F_manualno
                strInv8 = Trim$(.text)
        
                .Col = F_fp.F_itemno
                strInv9 = Trim$(.text)
                
                .Col = F_fp.F_name
                
                If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                
                    If Get_SqlserverCnt("SELECT 商品名称 FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & strInv8 & "' and flag = '1' and 序号= '" & strInv9 & "'") = 0 Then
                                    
                        MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"

                        Exit Sub
                    
                    End If
                                
                    stritemname = Get_SqlStr("select distinct 商品名称 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '1' and  序号 = '" & strInv9 & "'")

                    .text = stritemname

                End If
                
                strInv10 = Trim$(.text)
        
                .Col = F_fp.F_baoguanqty
                '报关数量
             
                If .text = "" Then
                    
                    MsgBox "请输入报关数量", vbInformation, "提示"
                    Exit Sub
                    
                End If
                    
                If Val(.text) <= 0 Then
                        
                    MsgBox "报关数量不可小于等于0", vbInformation, "提示"
                    Exit Sub
                
                End If
                
                strInv11 = Format(Trim$(.text), "0.000")
                    
                .Col = F_fp.F_unit
                
                If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                
                    strunit = Get_SqlStr("select distinct 计量单位 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '1' and  序号 = '" & strInv9 & "'")
                
                    .text = strunit
                                
                End If
                
                strInv12 = Trim$(.text)
        
                .Col = F_fp.F_indate
                If Trim$(.text) <> "" And Len(Trim$(.text)) <> 8 Then
                    MsgBox "进场日期请以YYYYMMDD格式填写,如20200501", vbInformation, "提示"
                    Exit Sub
                End If
                
                '入场日期
                strInv13 = Trim$(.text)

                .Col = F_fp.F_invoice
                strInv14 = Trim$(.text)
        
                .Col = F_fp.F_caseqty
                strInv15 = Trim$(.text)
        
                .Col = F_fp.F_currency
                strInv16 = Trim$(.text)
                
                .Col = F_fp.F_unitprice
                strInv30 = Trim$(.text)
                
                .Col = F_fp.F_baoguanvalue
                strInv17 = Trim$(.text)
                
                strInv30 = Format(Trim$(Val(strInv17) / Val(strInv11)), "0.000")
                
                .Col = F_fp.F_rate
                '汇率
                If .text = "" Then
                    
                    .text = 0
                    
                End If
                strInv18 = Format(Trim$(.text), "0.0000")
                
                .Col = F_fp.F_tariffrate
                '关税率
                If .text = "" Then
                    
                    .text = 0
                    
                End If
                
                strInv19 = Format(Trim$(.text), "0.0000")
                
                .Col = F_fp.F_id
                strInv28 = Trim$(.text)
                
                strbao1 = Get_SqlStr("select distinct 报关金额 from erptemp.dbo.ksimport where 批次 = '" & strInv29 & "' and id = '" & strInv28 & "' and flag = '0' ")
                
                strbao2 = Get_SqlStr("select distinct 汇率 from erptemp.dbo.ksimport where 批次 = '" & strInv29 & "' and id = '" & strInv28 & "' and flag = '0' ")
                
                strbao3 = Get_SqlStr("select distinct 关税率 from erptemp.dbo.ksimport where 批次 = '" & strInv29 & "' and id = '" & strInv28 & "' and flag = '0' ")
                
                .Col = F_fp.F_tariff
                
                '关税
                
                If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                    
                    strInv20 = Format(Trim$(.text), "0.00")
                    
                Else
                    If Val(strbao1) = Val(strInv17) And Val(strbao2) = Val(strInv18) And strbao3 = Val(strInv19) Then
                    
                        strInv20 = Format(Trim$(.text), "0.00")
                        
                    Else
                    
                        .text = Val(strInv18) * Val(strInv17) * Val(strInv19)
                
                        strInv20 = Format(Trim$(.text), "0.00")
                        
                    End If
                    
                End If
                
                .Col = F_fp.F_addtaxrate
                
                '增值税率
                If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                
                    .text = 0
                    strInv21 = Format(Trim$(.text), "0.0000")
                        
                Else
                
                    .text = 0.13
                    strInv21 = Format(Trim$(.text), "0.0000")
                    
                End If
                
                .Col = F_fp.F_addtax
                
                '增值税
                
                If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                
                    strInv22 = Format(Trim$(.text), "0.00")
                
                Else
                
                    If Val(strbao1) = Val(strInv17) And Val(strbao2) = Val(strInv18) Then
                    
                        strInv22 = Format(Trim$(.text), "0.00")
                    
                    Else
                    
                        .text = Val(strInv20) * Val(strInv21) + Val(strInv17) * Val(strInv21) * Val(strInv18)
                    
                        strInv22 = Format(Trim$(.text), "0.00")
                        
                    End If
                End If
                
                .Col = F_fp.F_declarationno
                strInv23 = Trim$(.text)
                
                .Col = F_fp.F_awb
                strInv24 = Trim$(.text)
                
                .Col = F_fp.F_freight
                strInv25 = Trim$(.text)
                
                .Col = F_fp.F_chargebackdate
                strInv26 = Trim$(.text)
                
                .Col = F_fp.F_mark
                strInv27 = Trim$(.text)
        
                
                strtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
                
                AddSql2 ("insert into erptemp.dbo.ksimport(批次,采购单号,料号,型号,类别,订单数量,标准die,总die数,手册编号,项号,品名,报关量,计量单位,入场日期,发票号,件数,币别,采购单价,报关金额,汇率,关税率,关税,增值税率,增值税,报关单号,AWB#,货代,退单日期,备注,id,键入时间,修改状态,修改时间,删除时间,flag) SELECT 批次,采购单号,料号,型号,类别,订单数量,标准die,总die数,手册编号,项号,品名,报关量,计量单位,入场日期,发票号,件数,币别,采购单价,报关金额,汇率,关税率,关税,增值税率,增值税,报关单号,AWB#,货代,退单日期,备注,id,键入时间,'修改前',修改时间,删除时间,'2' FROM erptemp.dbo.ksimport WHERE 批次 =  '" & strInv29 & "' AND id =  '" & strInv28 & "'  AND  flag = '0'")
                
                AddSql2 ("update erptemp.dbo.ksimport set 料号 = '" & strInv2 & "', 型号 = '" & strInv3 & "',类别 = '" & strInv4 & "',订单数量 = '" & strInv5 & "',标准die =  '" & strInv6 & "',总die数 =  '" & strInv7 & "',手册编号 = '" & strInv8 & "',项号 = '" & strInv9 & "',品名 =  '" & strInv10 & "',报关量 =  '" & strInv11 & "', " & " 计量单位  =  '" & strInv12 & "',入场日期 =  '" & strInv13 & "',发票号 =  '" & strInv14 & "',件数 =  '" & strInv15 & "',币别 =  '" & strInv16 & "',报关金额 =  '" & strInv17 & "',汇率 =  '" & strInv18 & "',关税率 =  '" & strInv19 & "',关税 =  '" & strInv20 & "',增值税率 =  '" & strInv21 & "',增值税 =  '" & strInv22 & "',报关单号 =  '" & strInv23 & "',AWB#  =  '" & strInv24 & "',货代 =  '" & strInv25 & "',退单日期 =  '" & strInv26 & "',备注 =  '" & strInv27 & "',修改状态 = '修改后',修改时间 = '" & strtime & "' where 批次 =  '" & strInv29 & "'  and flag = '0'  and id =  '" & strInv28 & "' ")
            
            End If
            
        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "请选择要修改的行", vbInformation, "提示"
            Exit Sub
            
        End If
    
    End With
    
    MsgBox "修改成功", vbInformation, "提示"

    Toolbar1.Buttons(5).Caption = "修改"
    Toolbar1.Buttons(5).Image = 4
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    
    ForQuery
    
End Sub


Private Sub ForMod5()

    Dim rs          As New ADODB.Recordset

    Dim i           As Integer

    Dim m           As Integer

    Dim j           As Integer

    Dim strInv1     As String

    Dim strInv2     As String

    Dim strInv3     As String

    Dim strInv4     As String

    Dim strInv5     As String

    Dim strInv6     As String

    Dim strInv7     As String

    Dim strInv8     As String

    Dim strInv9     As String

    Dim strInv10    As String

    Dim strInv11    As String

    Dim strInv12    As String

    Dim strInv13    As String

    Dim strInv14    As String

    Dim strInv15    As String

    Dim strInv16    As String
    
    Dim strInv17    As String
    
    Dim strInv18    As String
    
    Dim strInv19    As String
    
    Dim strInv20    As String
    
    Dim strInv21    As String
    
    Dim stritemname As String
    
    Dim strunit     As String

    Dim strtime     As String
    
    Dim strsql      As String
    
    Dim bFlag       As Boolean

    If Toolbar1.Buttons(5).Caption <> "提交" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                .Lock = False
    
                .Col = 5
                .Lock = False
    
                For m = 9 To 11
            
                    .Col = m
                    .Lock = False
      
                Next
                
                .Col = 14
                .Lock = False
                
                .Col = 15
                .Lock = False
                
                For m = 17 To 21
            
                    .Col = m
                    .Lock = False
      
                Next
          
                .Col = 5
                If .text = "进料对口" Or .text = "进料成品退换" Or .text = "进料料件复出" Then
                
                            .Col = 10
                            .Lock = False
                    
                            .Col = 11
                            .Lock = False
                        
                            .Col = 12
                            .Lock = True
                
                            .Col = 13
                            .Lock = True
                    
                Else
                    
                            .Col = 10
                            .Lock = True
                    
                            .Col = 11
                            .Lock = True
                    
                            .Col = 12
                            .Lock = False
                
                            .Col = 13
                            .Lock = False
        
                End If
                
                            
                .Col = 5
                .CellType = CellTypeComboBox

                .TypeComboBoxList = .TypeComboBoxList & "进料对口"
    
                .TypeComboBoxList = .TypeComboBoxList & "一般贸易"

                .TypeComboBoxList = .TypeComboBoxList & "其他进出口免费"

                .TypeComboBoxList = .TypeComboBoxList & "进料料件复出"
    
                .TypeComboBoxList = .TypeComboBoxList & "进料成品退换"

                .TypeComboBoxList = .TypeComboBoxList & "修理物品"

                .TypeComboBoxList = .TypeComboBoxList & "设备退运"

                .TypeComboBoxList = .TypeComboBoxList & "其他"
    
                 
                .Col = 14
                .CellType = CellTypeComboBox
            
                .TypeComboBoxList = .TypeComboBoxList & "USD"
            
                .TypeComboBoxList = .TypeComboBoxList & "JPY"

                .TypeComboBoxList = .TypeComboBoxList & "EUR"

                .TypeComboBoxList = .TypeComboBoxList & "RMB"
                                
                .Col = 8
            
                .text = Format(Trim$(.text), "0.000")
                
                .LockBackColor = vbYellow
                
                strsql = "select distinct 手册编号 from erptemp.dbo.ksmanual where 1 = 1 and flag = '2'"

                If rs.State = 1 Then rs.Close
                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                .Col = 10

                .CellType = CellTypeComboBox

     '           .TypeComboBoxList = ""

                rs.MoveFirst

                For j = 1 To rs.RecordCount

                    .TypeComboBoxList = .TypeComboBoxList & rs("手册编号")
                    rs.MoveNext
                Next
        
                rs.Clone
        
                Set rs = Nothing
    
            Next
        
        End With
    
        Toolbar1.Buttons(5).Caption = "提交"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    bFlag = False
    
    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
           
            .Col = 1
    
            j = 0

            If .text = "1" Then
            
                j = j + 1
                bFlag = True
                
                .Col = 2
                strInv21 = Trim$(.text)
                
                .Col = 3
                
'                If .Text = "" Then
'                    MsgBox "请输入出货单据", vbInformation, "提示"
'                    Exit Sub
'
'                End If
                
                strInv1 = Trim$(.text)
    
                .Col = 4
                strInv2 = Trim$(.text)
    
                .Col = 5
                
                If .text = "" Then
                    MsgBox "请输入类别", vbInformation, "提示"
                    Exit Sub

                End If
                
                strInv3 = Trim$(.text)
                
                .Col = 6
                strInv4 = Trim$(.text)
        
                .Col = 7
                
                strInv5 = Trim$(.text)
        
                .Col = 8
                strInv6 = Trim$(.text)
        
                .Col = 9
                strInv7 = Trim$(.text)
        
                .Col = 10
                strInv8 = Trim$(.text)
        
                .Col = 11
                strInv9 = Trim$(.text)
        
                .Col = 12
                
                If strInv3 = "进料对口" Or strInv3 = "进料成品退换" Or strInv3 = "进料料件复出" Then
                '品名
                    If strInv3 = "进料对口" Or strInv3 = "进料成品退换" Then
                    
                        If Get_SqlserverCnt("SELECT 商品名称 FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & strInv8 & "' and flag = '2' and 序号= '" & strInv9 & "'") = 0 Then
                                    
                            MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"
                                    
                            Exit Sub
                                    
                        End If
                    
                        stritemname = Get_SqlStr("select distinct 商品名称 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '2' and  序号 = '" & strInv9 & "'")
                    
                    Else
                    
                        If Get_SqlserverCnt("SELECT 商品名称 FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & strInv8 & "' and flag = '1' and 序号= '" & strInv9 & "'") = 0 Then
                                    
                            MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"
                                    
                            Exit Sub
                                
                        End If
                    
                        stritemname = Get_SqlStr("select distinct 商品名称 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '1' and  序号 = '" & strInv9 & "'")
                    
                    End If
                    
                    .text = stritemname

                End If
                
                strInv10 = Trim$(.text)
        
                .Col = 13
                
                '计量单位
                If strInv3 = "进料对口" Or strInv3 = "进料成品退换" Or strInv3 = "进料料件复出" Then
                
                    If strInv3 = "进料对口" Or strInv3 = "进料成品退换" Then
                    
                        strunit = Get_SqlStr("select distinct 计量单位 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '2' and  序号 = '" & strInv9 & "'")

                    Else
                    
                        strunit = Get_SqlStr("select distinct 计量单位 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '1' and  序号 = '" & strInv9 & "'")
                    
                    End If
                    
                    .text = strunit
                
                End If
                
                strInv11 = Trim$(.text)
        
                .Col = 14
                strInv12 = Trim$(.text)
        
                .Col = 15
                '总价
                
                If .text = "" Then
                
                    MsgBox "请输入总价", vbInformation, "提示"
                    Exit Sub

                End If
                strInv13 = Trim$(.text)
        
                .Col = 16
                
                If strInv13 <> "" And .text = "" Then
                    
                    .text = Val(strInv13) / Val(strInv6)
                
                End If
                
                strInv14 = Trim$(.text)
        
                .Col = 17
                strInv15 = Trim$(.text)
        
                .Col = 18
                strInv16 = Trim$(.text)
                
                .Col = 19
                strInv17 = Trim$(.text)
                
                .Col = 20
                strInv18 = Trim$(.text)
                
                .Col = 21
                strInv19 = Trim$(.text)
                
                .Col = 22
                strInv20 = Trim$(.text)
    
                strtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
                
                AddSql2 ("insert into erptemp.dbo.ksexport (批次,出货单据,料号,类别,发票号,出货日期,数量,报关单号,手册编号,手册项号,品名,计量单位,币别,总价,单价,AWB#,目的地,货代,退单日期,备注,键入时间,修改状态,修改时间,删除时间,flag,id) SELECT 批次,出货单据,料号,类别,发票号,出货日期,数量,报关单号,手册编号,手册项号,品名,计量单位,币别,总价,单价,AWB#,目的地,货代,退单日期,备注,键入时间,'修改前',修改时间,删除时间,'2',id FROM erptemp.dbo.ksexport WHERE 批次 = '" & strInv21 & "' and 出货单据 = '" & strInv1 & "'  AND 料号 =  '" & strInv2 & "' and id = '" & strInv20 & "' AND  flag = '0'")
                
                AddSql2 ("update erptemp.dbo.ksexport set 类别 =  '" & strInv3 & "',报关单号 =  '" & strInv7 & "',手册编号 =  '" & strInv8 & "',手册项号 =  '" & strInv9 & "',品名 =  '" & strInv10 & "',计量单位 =  '" & strInv11 & "',币别 =  '" & strInv12 & "',总价 =  '" & strInv13 & "',单价 = '" & strInv14 & "',AWB# =  '" & strInv15 & "',目的地 =  '" & strInv16 & "',货代 =  '" & strInv17 & "',退单日期 =  '" & strInv18 & "',备注 =  '" & strInv19 & "',修改状态 = '修改后',修改时间 = '" & strtime & "' where 批次 = '" & strInv21 & "' and 出货单据 = '" & strInv1 & "'  and 料号  = '" & strInv2 & "' and id = '" & strInv20 & "' and flag = '0'")
            
            End If
            
        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "请选择要修改的行", vbInformation, "提示"
            Exit Sub
            
        End If
    
    End With
    
    MsgBox "修改成功", vbInformation, "提示"

    Toolbar1.Buttons(5).Caption = "修改"
    Toolbar1.Buttons(5).Image = 4
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    
    ForQuery
    
End Sub

Private Sub ForMod6()

    Dim rs          As New ADODB.Recordset

    Dim i           As Integer

    Dim m           As Integer

    Dim j           As Integer

    Dim strInv1     As String

    Dim strInv2     As String

    Dim strInv3     As String
    
    Dim strInv4     As String

    Dim strInv5     As String

    Dim strInv6     As String

    Dim strInv7     As String

    Dim strInv8     As String

    Dim strInv9     As String

    Dim strInv10    As String

    Dim strInv11    As String
    
    Dim strInv12    As String

    Dim strInv13    As String

    Dim strInv14    As String

    Dim strInv15    As String

    Dim strInv16    As String
    
    Dim strInv17    As String
    
    Dim strInv18    As String
    
    Dim strInv19    As String
    
    Dim strInv20    As String
    
    Dim strInv21    As String
    
    Dim strInv22    As String
    
    Dim strInv23    As String

    Dim strInv24    As String
    
    Dim strInv25    As String

    Dim strInv26    As String
    
    Dim strInv27    As String

    Dim strInv28    As Integer
    
    Dim strInv29    As String
    
    Dim strInv30    As String
    
    Dim strunit     As String

    Dim strtime     As String
    
    Dim strNo1      As String

    Dim strNo2      As String

    Dim strNo3      As String
    
    Dim strssql     As String
    
    Dim strsql      As String

    Dim stritemname As String
    
    Dim strbao1     As String
    
    Dim strbao2     As String
    
    Dim strbao3     As String
    
    Dim strbaono1   As Double
    
    Dim strbaono2   As Double
    
    Dim strbaono3   As Double
    
    Dim bFlag       As Boolean

    If Toolbar1.Buttons(5).Caption <> "提交" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
                
                .Col = 1
                .Lock = False
                
                For m = 5 To 8
            
                    .Col = m
                    .Lock = False
      
                Next
            
                .Col = 10
                .Lock = False
            
                .Col = 11
                .Lock = False
            
                .Col = 13
                .Lock = False
            
                For m = 15 To 18
            
                    .Col = m
                    .Lock = False
      
                Next
                
                For m = 20 To 22
            
                    .Col = m
                    .Lock = False
      
                Next
            
                For m = 26 To 30
            
                    .Col = m
                    .Lock = False
      
                Next
                
                .Col = 13
                .text = Format(Trim$(.text), "0.000")
                
                .Col = 19
                .text = Format(Trim$(.text), "0.000")
                
                .Col = 20
                .text = Format(Trim$(.text), "0.0000")
            
                .Col = 21
                .text = Format(Trim$(.text), "0.0000")
            
                .Col = 22
                .text = Format(Trim$(.text), "0.0000")
            
                .Col = 23
                .text = Format(Trim$(.text), "0.00")
            
                .Col = 24
                .text = Format(Trim$(.text), "0.0000")
            
                .Col = 25
                .text = Format(Trim$(.text), "0.00")
                
                .Col = 6
                If .text = "进料对口" Or .text = "成品复进" Then
                
                    .Col = 10
                    .Lock = False
                        
                    .Col = 11
                    .Lock = False
                            
                    .Col = 12
                    .Lock = True
                    
                    .Col = 14
                    .Lock = True
                            
                    .Col = 21
                    .Lock = True
                        
                    .Col = 22
                    .Lock = True
                        
                    .Col = 23
                    .Lock = True
                    
                    .Col = 24
                    .Lock = True
                        
                    .Col = 25
                    .Lock = True
                Else
                    
                    .Col = 10
                    .Lock = True
                        
                    .Col = 11
                    .Lock = True
                        
                    .Col = 12
                    .Lock = False
                    
                    .Col = 14
                    .Lock = False
                        
                    .Col = 21
                    .Lock = False
                        
                    .Col = 22
                    .Lock = False
                    
                    '关税
                    .Col = 23
                    .Lock = False
                    
                    '增值税
                    .Col = 25
                    .Lock = False
                        
                End If

                
                .LockBackColor = vbYellow
                
                strsql = "select distinct 手册编号 from erptemp.dbo.ksmanual where 1 = 1 and flag = '1'"

                If rs.State = 1 Then rs.Close
                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                .Col = 10

                .CellType = CellTypeComboBox

              '  .TypeComboBoxList = ""

                rs.MoveFirst

                For j = 1 To rs.RecordCount

                    .TypeComboBoxList = .TypeComboBoxList & rs("手册编号")
                    rs.MoveNext
                Next
        
                rs.Clone
        
                Set rs = Nothing
    
            Next
        
        End With
    
        Toolbar1.Buttons(5).Caption = "提交"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    bFlag = False
    
    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            
            .Col = 1
    
            j = 0

            If .text = "1" Then
            
                j = j + 1
                bFlag = True
                
                .Col = 2
                
                strInv29 = Trim$(.text)
    
                .Col = 3
                
                strInv1 = Trim$(.text)
    
                .Col = 4
                strInv2 = Trim$(.text)
    
                .Col = 5
                strInv3 = Trim$(.text)
                
                .Col = 6
                
                If .text = "" Then
                    MsgBox "请输入类别", vbInformation, "提示"
                    Exit Sub

                End If
                
                strInv4 = Trim$(.text)

                '类别是进料对口&成品复进这两种情况，是没有汇率、关税、增值税的，因为是保税的
                
                .Col = 7

                strInv5 = Trim$(.text)
                
                If Val(strInv5) <= 0 Then
                
                     MsgBox "订单数量不可小于等于0", vbInformation, "提示"
                     Exit Sub
                     
                End If
           
                .Col = 8

                '标准die
'                If .Text = "" Then
'
'                    strssql = "select isnull(t8.qty,0) from  erpbase..tblCPurDataSub t1  " & " inner join  erpdata..tblSmainM2 t6 " & " on t1.物料编号 = t6.物料编号  " & " left join  (select m2.料号,max(m1.QTECHDIEQTY) as qty  from erptemp..TBLTSVNPIPRODUCT m1,erpdata..TSVtblMRuleData m2 where 1=1 " & " and m1.QTECHPTNO2 = m2.工序号 group by m2.料号) t8 " & " on t8.料号 = t6.料号 where 1=1 and t1.采购单编号  = '" & strInv1 & "' and t6.料号 = '" & strInv2 & "' "
'
'                    strInv6 = Get_SqlStr(strssql)
'
'                Else
'
'                    strInv6 = Trim$(.Text)
'
'                End If
                
                strInv6 = Trim$(.text)
                
                .Col = 9
                '总die数量
                
                strInv7 = Val(strInv5) * Val(strInv6)
                
                'strInv7 = Trim$(.Text)
        
                .Col = 10
                strInv8 = Trim$(.text)
        
                .Col = 11
                strInv9 = Trim$(.text)
                
                .Col = 12
                
                If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                
                    If Get_SqlserverCnt("SELECT 商品名称 FROM erptemp.dbo.ksmanual WHERE 手册编号 = '" & strInv8 & "' and flag = '1' and 序号= '" & strInv9 & "'") = 0 Then
                                    
                        MsgBox "输入的手册号 + 序号 无对应的品名及计量单位,请确认", vbInformation, "提示"

                        Exit Sub
                    
                    End If
                                
                    stritemname = Get_SqlStr("select distinct 商品名称 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '1' and  序号 = '" & strInv9 & "'")

                    .text = stritemname

                End If
                
                strInv10 = Trim$(.text)
        
                .Col = 13
                '报关数量
                If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                
                    strbaono1 = Get_SqlStr("select distinct isnull(申报数量,0) from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '1' and  序号 = '" & strInv9 & "'")
                
                    strbaono2 = Get_SqlStr("select isnull(sum(报关量),0) from erptemp.dbo.ksimport where flag = '0' and  采购单号 = ' " & strInv1 & "' and 料号 = '" & strInv2 & "'")
                
                    strbaono3 = strbaono1 - strbaono2
                
                    If .text = "" Then
                    
                        MsgBox "请输入报关数量", vbInformation, "提示"
                        Exit Sub
                    
                    End If
                    
                    If Val(.text) <= 0 Then
                        
                        MsgBox "报关数量不可小于等于0", vbInformation, "提示"
                        Exit Sub
                    
                    End If
                
                    strInv11 = Format(Trim$(.text), "0.000")
                
                    If Val(strInv11) > Val(strbaono3) Then
                    
                        MsgBox "输入的报关量超过可输入的范围,申报数量为" & strbaono1 & ",目前系统已录入数量 " & strbaono2 & "", vbInformation, "提示"
                        
                        Exit Sub
                    
                    End If
                    
                Else
                         
                    If .text = "" Then
                    
                        MsgBox "请输入报关数量", vbInformation, "提示"
                        Exit Sub
                    
                    End If
                    
                    If Val(.text) <= 0 Then
                        
                        MsgBox "报关数量不可小于等于0", vbInformation, "提示"
                        Exit Sub
                    
                    End If
                    strInv11 = Format(Trim$(.text), "0.000")
                    
                End If
        
                .Col = 14
                
                If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                
                    strunit = Get_SqlStr("select distinct 计量单位 from erptemp.dbo.ksmanual where 手册编号 = '" & strInv8 & "' and flag = '1' and  序号 = '" & strInv9 & "'")
                
                    .text = strunit
                                
                End If
                
                strInv12 = Trim$(.text)
        
                .Col = 15
                strInv13 = Trim$(.text)
        
                .Col = 16
                strInv14 = Trim$(.text)
        
                .Col = 17
                strInv15 = Trim$(.text)
        
                .Col = 18
                strInv16 = Trim$(.text)
                
                .Col = 19
                strInv30 = Trim$(.text)
                
                .Col = 20
                strInv17 = Trim$(.text)
                
                .Col = 21
                '汇率
                If .text = "" Then
                    
                    .text = 0
                    
                End If
                strInv18 = Format(Trim$(.text), "0.0000")
                
                .Col = 22
                '关税率
                If .text = "" Then
                    
                    .text = 0
                    
                End If
                
                strInv19 = Format(Trim$(.text), "0.0000")
                
                .Col = 31
                strInv28 = Trim$(.text)
                
                strbao1 = Get_SqlStr("select distinct 报关金额 from erptemp.dbo.ksimport where 采购单号 = '" & strInv1 & "' and 料号 = '" & strInv2 & "' and id = '" & strInv28 & "' and flag = '0' ")
                
                strbao2 = Get_SqlStr("select distinct 汇率 from erptemp.dbo.ksimport where 采购单号 = '" & strInv1 & "' and 料号 = '" & strInv2 & "' and id = '" & strInv28 & "' and flag = '0' ")
                
                strbao3 = Get_SqlStr("select distinct 关税率 from erptemp.dbo.ksimport where 采购单号 = '" & strInv1 & "' and 料号 = '" & strInv2 & "' and id = '" & strInv28 & "' and flag = '0' ")
                
                .Col = 23
                
                '关税
                
                If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                    
                    strInv20 = Format(Trim$(.text), "0.00")
                    
                Else
                    If Val(strbao1) = Val(strInv17) And Val(strbao2) = Val(strInv18) And strbao3 = Val(strInv19) Then
                    
                        strInv20 = Format(Trim$(.text), "0.00")
                        
                    Else
                    
                        .text = Val(strInv18) * Val(strInv17) * Val(strInv19)
                
                        strInv20 = Format(Trim$(.text), "0.00")
                        
                    End If
                    
                End If
                
                .Col = 24
                
                '增值税率
                If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                
                    .text = 0
                    strInv21 = Format(Trim$(.text), "0.0000")
                        
                Else
                
                    .text = 0.13
                    strInv21 = Format(Trim$(.text), "0.0000")
                    
                End If
                
                .Col = 25
                
                '增值税
                
                If strInv4 = "进料对口" Or strInv4 = "成品复进" Then
                
                    strInv22 = Format(Trim$(.text), "0.00")
                
                Else
                
                    If Val(strbao1) = Val(strInv17) And Val(strbao2) = Val(strInv18) Then
                    
                        strInv22 = Format(Trim$(.text), "0.00")
                    
                    Else
                    
                        .text = Val(strInv20) * Val(strInv21) + Val(strInv17) * Val(strInv21) * Val(strInv18)
                    
                        strInv22 = Format(Trim$(.text), "0.00")
                        
                    End If
                End If
                
                .Col = 26
                strInv23 = Trim$(.text)
                
                .Col = 27
                strInv24 = Trim$(.text)
                
                .Col = 28
                strInv25 = Trim$(.text)
                
                .Col = 29
                strInv26 = Trim$(.text)
                
                .Col = 30
                strInv27 = Trim$(.text)
                
                     
                strNo1 = Get_SqlStr("SELECT isnull(SUM(a.批准采购数量),0) FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b WHERE a.采购单编号 = '" & strInv1 & "' and a.物料编号 = b.物料编号 and b.料号 = '" & strInv2 & "' ")
                
                strNo2 = Get_SqlStr("SELECT isnull(SUM(订单数量),0) FROM erptemp.dbo.ksimport where 采购单号 = '" & strInv1 & "' and 料号 = '" & strInv2 & "' and id <> '" & strInv28 & "' and flag = '0'")
                
                strNo3 = Val(strNo1) - Val(strNo2)
                
                If Val(strInv5) > Val(strNo3) Then
                
                    MsgBox "该笔料号" & strInv2 & "批准采购数量: " & strNo1 & ",已经维护订单数量：" & strNo2 & ",最大数量只能维护：" & strNo3 & "", vbInformation, "提示"
                    Exit Sub

                End If
                
                strtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
                
                AddSql2 ("insert into erptemp.dbo.ksimport(批次,采购单号,料号,型号,类别,订单数量,标准die,总die数,手册编号,项号,品名,报关量,计量单位,入场日期,发票号,件数,币别,采购单价,报关金额,汇率,关税率,关税,增值税率,增值税,报关单号,AWB#,货代,退单日期,备注,id,键入时间,修改状态,修改时间,删除时间,flag) SELECT 批次,采购单号,料号,型号,类别,订单数量,标准die,总die数,手册编号,项号,品名,报关量,计量单位,入场日期,发票号,件数,币别,采购单价,报关金额,汇率,关税率,关税,增值税率,增值税,报关单号,AWB#,货代,退单日期,备注,id,键入时间,'修改前',修改时间,删除时间,'2' FROM erptemp.dbo.ksimport WHERE 批次 =  '" & strInv29 & "' and 采购单号 = '" & strInv1 & "'  AND 料号 =  '" & strInv2 & "' AND id =  '" & strInv28 & "'  AND  flag = '0'")
                
                AddSql2 ("update erptemp.dbo.ksimport set 型号 = '" & strInv3 & "',类别 = '" & strInv4 & "',订单数量 = '" & strInv5 & "',标准die =  '" & strInv6 & "',总die数 =  '" & strInv7 & "',手册编号 = '" & strInv8 & "',项号 = '" & strInv9 & "',品名 =  '" & strInv10 & "',报关量 =  '" & strInv11 & "', " & " 计量单位  =  '" & strInv12 & "',入场日期 =  '" & strInv13 & "',发票号 =  '" & strInv14 & "',件数 =  '" & strInv15 & "',币别 =  '" & strInv16 & "',报关金额 =  '" & strInv17 & "',汇率 =  '" & strInv18 & "',关税率 =  '" & strInv19 & "',关税 =  '" & strInv20 & "',增值税率 =  '" & strInv21 & "',增值税 =  '" & strInv22 & "',报关单号 =  '" & strInv23 & "',AWB#  =  '" & strInv24 & "',货代 =  '" & strInv25 & "',退单日期 =  '" & strInv26 & "',备注 =  '" & strInv27 & "',修改状态 = '修改后',修改时间 = '" & strtime & "' where 批次 =  '" & strInv29 & "' and 采购单号 = '" & strInv1 & "' and flag = '0' and 料号  = '" & strInv2 & "' and id =  '" & strInv28 & "' ")
            
            End If
            
        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "请选择要修改的行", vbInformation, "提示"
            Exit Sub
            
        End If
    
    End With
    
    MsgBox "修改成功", vbInformation, "提示"

    Toolbar1.Buttons(5).Caption = "修改"
    Toolbar1.Buttons(5).Image = 4
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    
    ForQuery
    
End Sub

Private Sub ForDel()

    Dim i       As Integer

    Dim j       As Integer
    
    Dim bFlag   As Boolean
    
    Dim strInv1 As String
    
    Dim strInv2 As String

    Dim strtime As String

    If Toolbar1.Buttons(7).Caption <> "提交" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
                
                .Col = E_FPS.E_gx
                .Lock = False
              
            Next
        
        End With
    
        Toolbar1.Buttons(7).Caption = "提交"
        Toolbar1.Buttons(7).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        Exit Sub

    End If

    bFlag = False

    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_FPS.E_gx
    
            j = 0

            If .text = "1" Then
            
                j = j + 1
                bFlag = True
                
                .Col = E_FPS.e_NO
                strInv1 = Trim$(.text)
                
                strtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
                
                Select Case Combo1.text
                
                    Case "进口明细表"
                        .Col = F_fp.F_id
                        strInv2 = Trim$(.text)
                        AddSql2 ("update erptemp.dbo.ksimport set flag = '1',删除时间  = '" & strtime & "' where 批次 = '" & strInv1 & "'  and id = '" & strInv2 & "' and flag = '0'")
                 
                    Case "进口明细表(特殊)"
                        .Col = F_fp.F_id
                        strInv2 = Trim$(.text)
                        AddSql2 ("update erptemp.dbo.ksimport set flag = '1',删除时间  = '" & strtime & "' where 批次 = '" & strInv1 & "'  and id = '" & strInv2 & "' and flag = '0'")
                     
                    Case "出口明细表"
                        .Col = E_FPS.e_ID
                        strInv2 = Trim$(.text)
                        AddSql2 ("update erptemp.dbo.ksexport set flag = '1',删除时间  = '" & strtime & "' where 批次 = '" & strInv1 & "' and id = '" & strInv2 & "' and flag = '0'")
                 
                    Case "出口明细表(特殊)"
                
                        .Col = E_FPS.e_ID
                        strInv2 = Trim$(.text)
                        AddSql2 ("update erptemp.dbo.ksexport set flag = '1',删除时间  = '" & strtime & "' where 批次 = '" & strInv1 & "' and id = '" & strInv2 & "' and flag = '0'")
                
                End Select
               
            End If

        Next

    End With

    If bFlag = False And j = 0 Then
        MsgBox "请选择要删除的行", vbInformation, "提示"
        Exit Sub

    End If
    
    MsgBox "删除成功", vbInformation, "提示"

    Toolbar1.Buttons(7).Caption = "删除"
    Toolbar1.Buttons(7).Image = 5
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(5).Enabled = True

    ForQuery
    
End Sub


'--------------------






Private Sub cmd_query_Click()

    QueryData

End Sub









Private Sub ListDataType(rs As ADODB.Recordset, fpS As fpSpread)
    
    With fpS
    .MaxRows = 0
    If rs.RecordCount = 0 Then
         Exit Sub
    End If
    Set .DataSource = rs
    End With
    With fpS
        .ReDraw = False
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
       
        .ColsFrozen = 2
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        '.Row = -1
        '.Col = E_FPS0.E_CHOOSE
        '.Lock = False
        
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        '设定列类型
       ' .Col = E_FPS0.E_CHOOSE   '选择
       ' .CellType = CellTypeCheckBox
       ' .TypeHAlign = TypeVAlignCenter
       ' .TypeVAlign = TypeVAlignCenter
        
        '设定列宽
        .ColWidth(-1) = 10
      '  .ColWidth(E_FPS0.E_CHOOSE) = 4
        .ColWidth(E_FPS0.e_ID) = 4
        .ColWidth(E_FPS0.E_CGDITEM) = 4
         .ColWidth(E_FPS0.E_PN) = 14
        .ColWidth(E_FPS0.E_SUPPLIERNAME) = 20
        .RowHeight(-1) = 10
        '设定是否排序
     '   .UserColAction = UserColActionSort
   '     For i = 1 To .MaxCols
        '    .Col = i
      '      .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
  '      Next
       ' .ZOrder
       ' .ReDraw = True
    End With
    

End Sub







Private Sub QueryData()
Dim strsql As String
Dim rs     As New ADODB.Recordset

On Error GoTo Err_Query
AddSql2 ("delete From erpbase..OPENPO_WAFER")
'同步采购表
strsql = "INSERT INTO erpbase..OPENPO_WAFER(采购单编号,物料编号,PO数量,到货数量,未到货数量,料号) SELECT a.采购单编号, a.物料编号,sum(a.批准采购数量),0,0,c.F_101 FROM erpbase..tblCPurDataSub a, erpbase..tblCPurData b  ,AIS20141114094336..t_ICItem c WHERE a.采购单编号=b.采购单编号 AND a.物料编号=c.FNumber  and  a.采购单编号 like 'c%' and a.物料编号 LIKE '01.01%' and  b.保税标记=1 AND a.是否禁用=0 group by a.采购单编号, a.物料编号 ,c.F_101"
AddSql2 (strsql)


'同步清关日期,清关数量
'strSql = "UPDATE a SET a.清关日期=max(b.入场日期) from erpbase..OPENPO_WAFER a  left JOIN erptemp..ksimport b on a.采购单编号=b.采购单号 and a.料号=b.料号 where b.flag =0 "
''AddSql2 (strSql)
strsql = "UPDATE t1 SET t1.清关日期=t2.入场日期,t1.清关总数=isnull(t2.清关总数量,0),t1.五天前清关总数=isnull(t2.五天前清关数量,0) from erpbase..OPENPO_WAFER t1 left join  " & _
" (SELECT ISNULL(采购单号,'') AS 采购单号 ,ISNULL(料号,'') AS 料号 ,sum( CASE WHEN DATEDIFF(day, 入场日期, getdate())>5 THEN 订单数量 ELSE 0 END )AS 五天前清关数量 , " & _
" sum( 订单数量)AS 清关总数量  ,max(isnull(入场日期,0)) AS 入场日期 FROM erptemp..ksimport WHERE flag=0 GROUP  BY 采购单号,料号) as t2  on t1.采购单编号=t2.采购单号 and t1.料号=t2.料号 "

AddSql2 (strsql)



'同步到货数量
strsql = "UPDATE a SET a.到货数量=isnull(t1.到货数量,0),a.未到货数量=a.清关总数-isnull(t1.到货数量,0) from erpbase..OPENPO_WAFER a  left JOIN ( SELECT b.采购单编号 ,b.物料编号 ,sum(b.到货数量) AS 到货数量 FROM erpbase..tblToRecEntry b  GROUP BY b.采购单编号,b.物料编号  ) AS t1 ON  a.采购单编号 =t1.采购单编号 AND a.物料编号=t1.物料编号"
AddSql2 (strsql)

'同步入库数量
strsql = "UPDATE t1 SET  t1.已入库数量=Isnull(t2.已入库数量,0),t1. 已入98仓数量= isnull(t2.已入库数量98,0),t1.已入52仓数量=isnull(t2. 已入库数量52,0) FROM     erpbase..OPENPO_WAFER  t1 left JOIN " & _
" (SELECT aa.采购单编号,aa.物料编号,sum( bb.实入数量*cc.入库类型 )AS 已入库数量,sum( CASE cc.仓库编号 WHEN '52' then  bb.实入数量*cc.入库类型 ELSE 0 END )AS 已入库数量52" & _
" ,sum( CASE cc.仓库编号 WHEN '98' then  bb.实入数量*cc.入库类型 ELSE 0 END )AS 已入库数量98  FROM erpbase..tblToRecEntry aa " & _
" LEFT JOIN  erpbase..TblToInSub bb ON aa.到货单编号=bb.到货单编号 AND aa.分录号=bb.分录号 " & _
" INNER JOIN erpbase..TblToInrec cc  ON bb.入库单编号=cc.入库单编号" & _
" GROUP BY aa.采购单编号,aa.物料编号) AS t2 ON t1.采购单编号=t2.采购单编号 AND  t1.物料编号=t2.物料编号"
AddSql2 (strsql)



strsql = "SELECT row_number() over (order by t1.采购单编号,t1.料号) as 序号,t1.* FROM ( " & _
" select distinct a.采购单编号,a.料号,a.PO数量 as PO总量,a.五天前清关总数,a.清关总数,a.PO数量-a.清关总数 as 'PO总量-清关总数' , a.五天前清关总数-a.清关总数 as '五天前清关总数-清关总数'  , a.到货数量 as  已到货数量, " & _
" a.清关总数- a.到货数量  as '清关总数- 已到货数量' ,a.已入库数量,a.已入98仓数量,a.已入52仓数量,a.到货数量-a.已入库数量 as '已到货数量-已入库数量' , " & _
" CASE WHEN isnull(a.已入库数量,0)=0 AND isnull(a.已入库数量,0)<a.清关总数  THEN '未入' WHEN isnull(a.已入库数量,0)<a.清关总数 THEN '未入满' WHEN isnull(a.已入库数量,0)>a.清关总数 THEN '入超'   ELSE '' END AS 是否入库, " & _
" e.客户代码,d.FName as 供应商名称,convert(VARCHAR(10),b.审核日期,112) as PO日期,a.清关日期  " & _
" from erpbase..OPENPO_WAFER a " & _
" inner join erpbase..tblcpurdata b on a.采购单编号=b.采购单编号 " & _
" inner join AIS20141114094336.dbo.t_Supplier  d on b.供应商编号=d.FNumber " & _
" inner join dbo.tblXCustomer  e on d.FName=e.客户名称 " & _
" left join erptemp..tbltsvnpiproduct  f on a.料号=f.MARKETLASTUPDATE_BY " & _
" where  a.清关总数-a.已入库数量 <>0 "

If Trim(txtCust.text) <> "" Then
  strsql = strsql & " and   e.客户代码='" & Trim(txtCust.text) & "'"
End If

If Trim(TxtPN.text) <> "" Then
  strsql = strsql & " and   a.料号='" & Trim(TxtPN.text) & "'"
End If


If Optpatial.Value = True Then '关务未维护
    strsql = strsql & " and isnull(a.清关日期,'')=''"
End If
If Optpatial2.Value = True Then '关务已维护
    strsql = strsql & " and isnull(a.清关日期,'')<>''"
End If
strsql = strsql & " ) as t1"

Set rs = Get_SqlserveRs(strsql)
Call ListDataType(rs, fpS_Clear)
Err_Query:
If Err.number <> 0 Then
   MsgBox "QueryData遇到错误,错误原因:" & Err.DESCRIPTION
End If

End Sub







Private Sub ExportExcel(fpS As fpSpread)
    Dim xlsApp      As Excel.Application
    Dim xlsBook     As Excel.Workbook
    Dim xlsSheet    As Excel.Worksheet
    Dim i           As Long
    Dim j           As Long
    
    On Error GoTo Ert
    
    Set xlsApp = CreateObject("Excel.Application")
    Set xlsBook = xlsApp.Workbooks.Add
    Set xlsSheet = xlsBook.Worksheets(1)

    With xlsApp
        .Rows(1).Font.Bold = True

    End With
   
    With fpS

        For i = 0 To .MaxRows
            For j = 1 To .MaxCols
                .Col = j
                .Row = i
                If j <= 14 Then
                    xlsSheet.Cells(i + 1, j) = .text
                Else
                
                    xlsSheet.Cells(i + 1, j) = Trim$(("'" & .text))
                End If

            Next j
       
        Next i

    End With
    xlsApp.Visible = True
    
    With xlsSheet.Range("2:" & i + 1)
        .horizontalAlignment = xlLeft
    End With
    xlsSheet.Range("A1").Select
    xlsApp.Columns.AutoFit

    
    
    Set xlsApp = Nothing
    Set xlsSheet = Nothing
    Set xlsBook = Nothing
    Exit Sub
    
Ert:
    MsgBox Err.DESCRIPTION
    
    If Not (xlsApp Is Nothing) Then
        
        Set xlsApp = Nothing
        Set xlsSheet = Nothing
        Set xlsBook = Nothing

    End If
    

End Sub


Private Function fpscopy()
    '入场日期  发票号  报关单号  AWB   货代  手册编号
    Dim RCRQ As String
    Dim QCRQ As String
    Dim FPH As String
    Dim BGDH As String
    Dim AWB As String
    Dim HD As String
    Dim SCBH As String
    RCRQ = ""
    FPH = ""
    BGDH = ""
    AWB = ""
    HD = ""
    SCBH = ""
    With fpS(0)
        .Row = 1
        .Col = 10
            SCBH = .text
        .Col = 15
            RCRQ = .text
        .Col = 16
            FPH = .text
        .Col = 26
            BGDH = .text
        .Col = 27
            AWB = .text
        .Col = 28
            HD = .text
        For i = 2 To .MaxRows
               .Row = i
               .Col = 10
                    .text = SCBH
               .Row = i
               .Col = 15
                    .text = RCRQ
               .Row = i
               .Col = 16
                    .text = FPH
               .Row = i
               .Col = 26
                    .text = BGDH
               .Row = i
               .Col = 27
                    .text = AWB
               .Row = i
               .Col = 28
                    .text = HD
        Next
    End With
      

End Function











