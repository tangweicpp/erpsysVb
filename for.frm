VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_GWZLWH 
   Caption         =   "����ά��"
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
      TabCaption(0)   =   "����ά��"
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
      TabCaption(1)   =   "�ֲ��ά��"
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
      TabCaption(2)   =   "��ԲOPEN PO��ѯ"
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
         Caption         =   "������ά��"
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
         Caption         =   "����"
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
         Caption         =   "Open PO��ѯ"
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
         Caption         =   "��ʾ����"
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
         Caption         =   "����δά��"
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
         Caption         =   "ɾ ��"
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
         Caption         =   "���һ��"
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
         Caption         =   "��ѯ"
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
         Caption         =   "�ϴ�"
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
               Caption         =   "��ѯ"
               Key             =   "QUE"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "ADD"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "MOD"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "DEL"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "EXIT"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
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
         Caption         =   "��       ��"
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
         Caption         =   "�ͻ�����"
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
         Caption         =   "������"
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
         Caption         =   "�ܽ��"
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
         Caption         =   "��������"
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
         Caption         =   "��ʼ����"
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
         Caption         =   "�ֲ���"
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
         Caption         =   "ó�׷�ʽ"
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
         Caption         =   "�ֲ����"
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
         Caption         =   "��ѡ���ϴ�������:"
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
         Caption         =   "ps:�������ɹ�����ʱ�� - Ϊ����� ʾ��:C120905027-C120905028"
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
         Caption         =   "��ѯ����"
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
         Caption         =   "ά������"
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
    F_no                    '����
    F_purchaseno            '�ɹ����� null
    F_partno                '�Ϻ�
    F_modelno               '�ͺ�
    F_modetrade             '���
    F_orderqty              '�������� null
    F_die                   '��׼die
    F_totaldie              '��die��
    F_manualno              '�ֲ���
    F_itemno                '���
    F_name                  'Ʒ��
    F_baoguanqty            '������
    F_unit                  '������λ
    F_indate                '�볡����
    F_invoice               '��Ʊ��
    F_caseqty               '����
    F_currency              '�ұ�
    F_unitprice             '�ɹ�����
    F_baoguanvalue          '���ؽ��
    F_rate                  '����
    F_tariffrate            '��˰��
    F_tariff                '��˰
    F_addtaxrate            '��ֵ˰��
    F_addtax                '��ֵ˰
    F_declarationno         '���ص���
    F_awb                   'AWB#
    F_freight               '����
    F_chargebackdate        '�˵�����
    F_mark                  '��ע
    F_id                    'id
     
End Enum


'export
Private Enum E_FPS
        
    E_gx = 1
    e_NO                    '����
    E_exportno              '��������
    E_partno                '�Ϻ�
    E_modetrade             '���
    e_Invoice               '��Ʊ��
    E_exportdate            '��������
    E_exportquantity        '��������
    E_declarationno         '���ص���
    E_manualno              '�ֲ���
    E_itemno                '�ֲ����
    E_name                  'Ʒ��
    E_UNIT                  '������λ
    E_currency              '�ұ�
    E_totalprice            '�ܼ�
    E_unitprice             '����
    E_AWB                   'AWB#
    E_destination           'Ŀ�ĵ�
    E_freight               '����
    E_chargebackdate        '�˵�����
    E_mark                  '��ע
    e_ID                    'id
    E_flienum               '�ļ����
End Enum

Private Enum E_FPS0          '
   ' E_CHOOSE = 1
    e_ID = 1
    E_CGDBH = 2 '�ɹ������
    E_CGDITEM '�ɹ������
    E_PODATE 'PO��Ч����
    E_CUSTOMSCLEARDATE 'PO�������
  'E_CreateDate 'ά������
  ' E_Createby 'ά����Ա
    E_cust '�ͻ�����
    E_PN '�Ϻ�
    E_SUPPLIERNAME '��Ӧ������
    E_SUPPLIERCODE '��Ӧ�̱��
    e_device 'Device
    E_POqty '����
    E_Entryqty '����
    E_Lastqty '����
    E_END
    
End Enum



Private Sub cmdExport_Click()
Call ExportExcel(fpS_Clear)
End Sub

Private Sub Combo1_Click()

    Select Case Combo1.text
            
        Case "������ϸ��"
        
            Command4.Visible = False
            
            Command5.Visible = False
            
            lb2.Visible = True
            
            Text1.Visible = True
        
            lb2 = "��Ʊ����"
            
            lb3.Visible = True
            
            lb3 = "ps:��������Ʊ�ű��ʱ��/Ϊ����� ʾ��:S1902200012/S1902200013"
            
            lb4.Visible = True
            
            comBo2.Visible = True
            
            comBo2.Clear
            
            comBo2.AddItem ("���϶Կ�")
            comBo2.AddItem ("һ��ó��")
            comBo2.AddItem ("�������������")
            comBo2.AddItem ("�����ϼ�����")
            comBo2.AddItem ("���ϳ�Ʒ�˻�")
            comBo2.AddItem ("������Ʒ")
            comBo2.AddItem ("�豸����")
            comBo2.AddItem ("����")
              
            lb5.Visible = False
            
            Combo3.Visible = False
            
            fpss(0).Visible = False
            
            fpS(0).MaxRows = 0
            fpS(0).MaxCols = 0
            fpss(0).MaxRows = 0
            fpss(0).MaxCols = 0
            
        Case "������ϸ��(����)"
        
            Command4.Visible = False
            
            Command5.Visible = False
            
            lb2.Visible = True
            
            Text1.Visible = True
            
            lb2 = "��Ʊ����"
            
            lb3.Visible = False
            
            lb4.Visible = True
            
            comBo2.Visible = True
            
            comBo2.Clear
            
            comBo2.AddItem ("���϶Կ�")
            comBo2.AddItem ("һ��ó��")
            comBo2.AddItem ("�������������")
            comBo2.AddItem ("�����ϼ�����")
            comBo2.AddItem ("���ϳ�Ʒ�˻�")
            comBo2.AddItem ("������Ʒ")
            comBo2.AddItem ("�豸����")
            comBo2.AddItem ("����")
              
            lb5.Visible = False
            
            Combo3.Visible = False
            
            fpss(0).Visible = False
            
            fpS(0).MaxRows = 0
            fpS(0).MaxCols = 0
            fpss(0).MaxRows = 0
            fpss(0).MaxCols = 0
            
        Case "������ϸ��"
            
            Command4.Visible = False
            
            Command5.Visible = False
            
            lb2.Visible = True
            
            Text1.Visible = True
        
            lb2 = "�ɹ�����"
            
            lb3.Visible = True
            
            lb3 = "ps:�������ɹ�����ʱ��/Ϊ����� ʾ��:C120905027/C120905028"
            
            lb4.Visible = False
            
            comBo2.Visible = False
            
            lb5.Visible = False
            
            Combo3.Visible = False
            
            fpS(0).MaxRows = 0
            fpS(0).MaxCols = 0
            fpss(0).MaxRows = 0
            fpss(0).MaxCols = 0
            
        Case "������ϸ��(����)"
            
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
    
        MsgBox "��ѡ��ó�׷�ʽ", vbInformation, "��ʾ"
        Exit Sub
    
    End If
    
    If comBo2.text = "���϶Կ�" Or comBo2.text = "���ϳ�Ʒ�˻�" Or comBo2.text = "�����ϼ�����" Then
    
        lb5.Visible = True
            
        Combo3.Visible = True
            
        Combo3.Clear
            
        strsql = "select distinct �ֲ��� from erptemp.dbo.ksmanual where 1 = 1 and flag = '2'"

        If rs.State = 1 Then rs.Close
        rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

        rs.MoveFirst

        For j = 1 To rs.RecordCount

            Combo3.AddItem (rs("�ֲ���"))
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

    CommonDialog1.Filter = "EXCEL�ļ�(*.xlsx)|*.xlsx|EXCEL�ļ�(*.xls)|*.xls"
    CommonDialog1.ShowOpen
    '�õ��ļ���
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
        MsgBox "��ѡ����ϴ����ļ�"
        Exit Sub

    End If
    
    SumCount = 0
    SumDelCount = 0

    'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text2.text)    '���ļ�

    Set xlSheet = xlBook.Worksheets("sheet1")        '��sheet�еı�

    '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 6 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        Exit Sub

    End If

    For i = 1 To xlSheet.Range("A1").CurrentRegion.Rows.count
   
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
        
            strChar = Chr(96 + j)
            tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ
        
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
        
        If Get_SqlserverCnt("select * from erptemp.dbo.ksmanual where �ֲ��� = '" & temp1 & "' and flag = '" & temp2 & "' and ��� = '" & temp3 & "' ") <> 0 Then

            AddSql2 ("DELETE FROM erptemp.dbo.ksmanual where �ֲ��� = '" & temp1 & "' and flag = '" & temp2 & "' and ��� = '" & temp3 & "'")

            SumDelCount = SumDelCount + 1

        End If

        AddSql2 ("insert into erptemp.dbo.ksmanual values('" & temp1 & "','" & temp2 & "','" & temp3 & "','" & temp4 & "','" & temp5 & "','" & temp6 & "')")
    
        
        SumCount = SumCount + 1
        
    Next i

    xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing
    
    If SumCount > 0 Then
    
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�", , "��������"
    Else

        MsgBox "�������ϴ��ɹ�", vbInformation, "��ʾ"
    
    End If
    
    If SumDelCount > 0 Then

        MsgBox "�������ݿ�����" & SumDelCount & "�ʣ�", , "��������"

    End If

End Sub

Private Sub Command3_Click()

    Dim aflag     As String
    
    Dim strmanual As String
    
    Dim strsql    As String
    
    Dim rs        As New ADODB.Recordset
    
    strmanual = Trim$(Text3.text)
    
    If Text3.text <> "" Then
    
        strsql = "select �ֲ���,case flag when '1' then '�ϼ���' when '2' then '��Ʒ��' end as ����,���,��Ʒ����,������λ,�걨���� from erptemp.dbo.ksmanual where �ֲ��� = '" & strmanual & "' order by �ֲ���,flag,���"
    
    Else
    
        strsql = "select �ֲ���,case flag when '1' then '�ϼ���' when '2' then '��Ʒ��' end as ����,���,��Ʒ����,������λ,�걨���� from erptemp.dbo.ksmanual where 1 = 1 order by �ֲ���,flag,��� "

    End If

    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType2(rs)
    Else
        
        MsgBox "��ѯ�������ֲ���Ϣ", vbInformation, "��ʾ"
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
    
        Case "������ϸ��(����)"
    
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
                       
                .TypeComboBoxList = .TypeComboBoxList & "���϶Կ�"
            
                .TypeComboBoxList = .TypeComboBoxList & "һ��ó��"
            
                .TypeComboBoxList = .TypeComboBoxList & "�������������"
            
                .TypeComboBoxList = .TypeComboBoxList & "�����ϼ�����"
            
                .TypeComboBoxList = .TypeComboBoxList & "���ϳ�Ʒ�˻�"
            
                .TypeComboBoxList = .TypeComboBoxList & "������Ʒ"
            
                .TypeComboBoxList = .TypeComboBoxList & "�豸����"
            
                .TypeComboBoxList = .TypeComboBoxList & "����"
                
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
            
        Case "������ϸ��(����)"

            With fpS(0)

                .MaxRows = .MaxRows + 1
                i = .MaxRows
        
                .Row = i
                .Col = F_fp.F_no
                .text = stridd
                
                .Row = i
                .Col = F_fp.F_modetrade
                .CellType = CellTypeComboBox
                       
                .TypeComboBoxList = .TypeComboBoxList & "���϶Կ�"
            
                .TypeComboBoxList = .TypeComboBoxList & "һ��ó��"
            
                .TypeComboBoxList = .TypeComboBoxList & "�������������"
            
                .TypeComboBoxList = .TypeComboBoxList & "��Ʒ����"
            
                .TypeComboBoxList = .TypeComboBoxList & "ά����Ʒ"
            
                .TypeComboBoxList = .TypeComboBoxList & "�ϼ�����"
            
                .TypeComboBoxList = .TypeComboBoxList & "���ϳ�Ʒ�˻�"
            
                .TypeComboBoxList = .TypeComboBoxList & "����"
                
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
                
                strsql = "select distinct �ֲ��� from erptemp.dbo.ksmanual where 1 = 1 and flag = '1'"

                If rs.State = 1 Then rs.Close
                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                .Col = F_fp.F_manualno
                .ColWidth(F_fp.F_manualno) = 12
                .CellType = CellTypeComboBox

                rs.MoveFirst

                For i = 1 To rs.RecordCount

                    .TypeComboBoxList = .TypeComboBoxList & rs("�ֲ���")
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
            
                MsgBox "û����Ҫɾ������", vbInformation, "��ʾ"
            
                Exit Sub
            
            End If
            
            .MaxRows = strsum1 - strsum
            '            .MaxRows = .MaxRows - 1

            For j = 1 To .MaxRows
            
                Select Case Combo1.text
            
                    Case "������ϸ��(����)"
                    
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
                
                    Case "������ϸ��(����)"
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
            
            MsgBox "�Ѿ������Ͽ�ɾ������ȷ��", vbInformation, "��ʾ"
            
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
    
    If Combo1.text = "������ϸ��" Or Combo1.text = "������ϸ��(����)" Then
        If Text1.text <> "" Then
            fpscopy '�Զ������¼��к��� ZYF 20200331
        End If
    End If
    
    Select Case Combo1.text
   
        Case "������ϸ��"
            
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
                        '�����ֻ�н��϶Կ�&��Ʒ������������ĲŻ������ֲ�ţ�������������ѡ���ֲ�������
                        .Col = 3
                        
                        strcool1 = Trim$(.text)
                        
                        .Col = 6

                        If .text = "" Then
                
                            MsgBox "���������", vbInformation, "��ʾ"
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
        
                                If strsty = "���϶Կ�" Or strsty = "��Ʒ����" Then
                
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
                
                        strNo1 = Get_SqlStr("SELECT isnull(SUM(a.��׼�ɹ�����),0) FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b WHERE a.�ɹ������ = '" & strInv1 & "' and a.���ϱ�� = b.���ϱ�� and b.�Ϻ� = '" & strInv2 & "' ")
                
                        strNo2 = Get_SqlStr("SELECT isnull(SUM(��������),0) FROM erptemp.dbo.ksimport where �ɹ����� = '" & strInv1 & "' and �Ϻ� = '" & strInv2 & "' and flag = '0'")
                
                        strNo3 = Val(strNo1) - Val(strNo2)
                
                        If Val(strInv5) > Val(strNo3) Then
                        
                            MsgBox "�ñ��Ϻ�" & strInv2 & "��׼�ɹ�����: " & strNo1 & ",�Ѿ�ά������������" & strNo2 & ",�������ֻ��ά����" & strNo3 & "", vbInformation, "��ʾ"
                            Exit Sub

                        End If
                
                        If Val(strInv5) <= 0 Then
                
                            MsgBox "������������С�ڵ���0", vbInformation, "��ʾ"
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
                
                        'ֻ��ѡ�����ֲ�Ųſ��Գ���ѡ����ŵĹ��ܣ����ܴ���Ʒ�����������Ҫ�ֹ�����Ʒ���뵥λ
                
                        If strsty = "���϶Կ�" Or strsty = "��Ʒ����" Then
                
                            .Col = 10
                            stritem = Trim(.text)

                            If stritem = "" Then
                
                                MsgBox "�������ֲ���", vbInformation, "��ʾ"
                                Exit Sub

                            End If

                            .Col = 11

                            If Trim$(.text) <> "" Then
                            
                                If Get_SqlserverCnt("SELECT ��Ʒ���� as Ʒ��,������λ " & " FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & stritem & "' and flag = '" & strflag & "' and ���= '" & Trim$(.text) & "'") = 0 Then
                                    
                                    MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"
                                    
                                    .SetText 12, Row, Trim$("")
                                    .SetText 14, Row, Trim$("")
                                    
                                    Exit Sub
                                
                                End If

                                strsql = "SELECT ��Ʒ���� as Ʒ��,������λ " & " FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & stritem & "' and flag = '" & strflag & "' and ���= '" & Trim$(.text) & "'"
                                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                                If Not rs.EOF Then

                                    With fpS(0)
                                        .SetText 12, Row, Trim$("" & rs!Ʒ��)
                                        .SetText 14, Row, Trim$("" & rs!������λ)

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
    
        Case "������ϸ��"
        
            strflag = 2
                
            '���϶Կ�/���ϳ�Ʒ�˻�/�����ϼ�����
            
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

                            MsgBox "���������", vbInformation, "��ʾ"
                            Exit Sub

                        End If

                        strsty = Trim(.text)

                        If strsty = "���϶Կ�" Or strsty = "���ϳ�Ʒ�˻�" Or strsty = "�����ϼ�����" Then

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
                
                        'ֻ��ѡ�����ֲ�Ųſ��Գ���ѡ����ŵĹ��ܣ����ܴ���Ʒ�����������Ҫ�ֹ�����Ʒ���뵥λ
                
                        If strsty = "���϶Կ�" Or strsty = "���ϳ�Ʒ�˻�" Or strsty = "�����ϼ�����" Then
                
                            .Col = 10
                            stritem = Trim(.text)

                            .Col = 11

                            If Trim$(.text) <> "" Then
                                    
                                If strsty = "���϶Կ�" Or strsty = "���ϳ�Ʒ�˻�" Then
                                
                                    If Get_SqlserverCnt("SELECT ��Ʒ���� as Ʒ��,������λ " & " FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & stritem & "' and flag = '" & strflag & "' and ���= '" & Trim$(.text) & "'") = 0 Then
                                    
                                        MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"
                                    
                                        .SetText 12, Row, Trim$("")
                                        .SetText 13, Row, Trim$("")
                                    
                                        Exit Sub
                                
                                    End If

                                    strsql = "SELECT ��Ʒ���� as Ʒ��,������λ " & " FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & stritem & "' and flag = '" & strflag & "' and ���= '" & Trim$(.text) & "'"
                                Else

                                    If Get_SqlserverCnt("SELECT ��Ʒ���� as Ʒ��,������λ " & " FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & stritem & "' and flag = '1' and ���= '" & Trim$(.text) & "'") = 0 Then
                                    
                                        MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"
                                    
                                        .SetText 12, Row, Trim$("")
                                        .SetText 13, Row, Trim$("")
                                    
                                        Exit Sub
                                
                                    End If

                                    strsql = "SELECT ��Ʒ���� as Ʒ��,������λ " & " FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & stritem & "' and flag = '1' and ���= '" & Trim$(.text) & "'"

                                End If

                                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                                If Not rs.EOF Then

                                    With fpS(0)
                                    
                                        .SetText 12, Row, Trim$("" & rs!Ʒ��)
                                        .SetText 13, Row, Trim$("" & rs!������λ)

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
                            MsgBox "�������ܼ�", vbInformation, "��ʾ"
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
            
        Case "������ϸ��(����)"
            
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
                
                            MsgBox "���������", vbInformation, "��ʾ"
                            Exit Sub

                        End If

                        strsty = Trim$(.text)
        
                        If strsty = "���϶Կ�" Or strsty = "��Ʒ����" Then
                
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
                
                        If strsty = "���϶Կ�" Or strsty = "��Ʒ����" Then
                
                            .Col = F_fp.F_manualno
                            stritem = Trim(.text)

                            If stritem = "" Then
                
                                MsgBox "�������ֲ���", vbInformation, "��ʾ"
                                Exit Sub

                            End If

                            .Col = F_fp.F_itemno

                            If Trim$(.text) <> "" Then
                            
                                If Get_SqlserverCnt("SELECT ��Ʒ���� as Ʒ��,������λ " & " FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & stritem & "' and flag = '" & strflag & "' and ���= '" & Trim$(.text) & "'") = 0 Then
                                    
                                    MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"
                                    
                                    .SetText F_fp.F_name, Row, Trim$("")
                                    .SetText F_fp.F_unit, Row, Trim$("")
                                    
                                    Exit Sub
                                
                                End If

                                strsql = "SELECT ��Ʒ���� as Ʒ��,������λ " & " FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & stritem & "' and flag = '" & strflag & "' and ���= '" & Trim$(.text) & "'"
                                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                                If Not rs.EOF Then

                                    With fpS(0)
                                        .SetText F_fp.F_name, Row, Trim$("" & rs!Ʒ��)
                                        .SetText F_fp.F_unit, Row, Trim$("" & rs!������λ)

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
                        
                            MsgBox "�����뱨������", vbInformation, "��ʾ"
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
            
        Case "������ϸ��(����)"
        
            strflag = 2
                
            '���϶Կ�/���ϳ�Ʒ�˻�/�����ϼ�����
            
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

                            MsgBox "���������", vbInformation, "��ʾ"
                            Exit Sub

                        End If

                        strsty = Trim(.text)

                        If strsty = "���϶Կ�" Or strsty = "���ϳ�Ʒ�˻�" Or strsty = "�����ϼ�����" Then

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
                
                        'ֻ��ѡ�����ֲ�Ųſ��Գ���ѡ����ŵĹ��ܣ����ܴ���Ʒ�����������Ҫ�ֹ�����Ʒ���뵥λ
                
                        If strsty = "���϶Կ�" Or strsty = "���ϳ�Ʒ�˻�" Or strsty = "�����ϼ�����" Then
                
                            .Col = E_FPS.E_manualno
                            stritem = Trim(.text)

                            .Col = E_FPS.E_itemno
                            
                            If Trim$(.text) <> "" Then
                                
                                If strsty = "���϶Կ�" Or strsty = "���ϳ�Ʒ�˻�" Then
                                
                                    If Get_SqlserverCnt("SELECT ��Ʒ���� as Ʒ��,������λ " & " FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & stritem & "' and flag = '" & strflag & "' and ���= '" & Trim$(.text) & "'") = 0 Then
                                    
                                        MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"
                                    
                                        .SetText E_FPS.E_name, Row, Trim$("")
                                        .SetText E_FPS.E_UNIT, Row, Trim$("")
                                    
                                        Exit Sub
                                
                                    End If

                                    strsql = "SELECT ��Ʒ���� as Ʒ��,������λ " & " FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & stritem & "' and flag = '" & strflag & "' and ���= '" & Trim$(.text) & "'"
                                
                                Else
                                
                                    If Get_SqlserverCnt("SELECT ��Ʒ���� as Ʒ��,������λ " & " FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & stritem & "' and flag = '1' and ���= '" & Trim$(.text) & "'") = 0 Then
                                    
                                        MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"
                                    
                                        .SetText E_FPS.E_name, Row, Trim$("")
                                        .SetText E_FPS.E_UNIT, Row, Trim$("")
                                    
                                        Exit Sub
                                
                                    End If

                                    strsql = "SELECT ��Ʒ���� as Ʒ��,������λ " & " FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & stritem & "' and flag = '1' and ���= '" & Trim$(.text) & "'"
                                
                                End If

                                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                                If Not rs.EOF Then

                                    With fpS(0)
                                    
                                        .SetText E_FPS.E_name, Row, Trim$("" & rs!Ʒ��)
                                        .SetText E_FPS.E_UNIT, Row, Trim$("" & rs!������λ)

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
                        
                            MsgBox "�������������", vbInformation, "��ʾ"
                            Exit Sub
                            
                        End If
                    
                        .Col = E_FPS.E_totalprice

                        If Trim$(.text) = "" Then
                        
                            MsgBox "�������ܼ�", vbInformation, "��ʾ"
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
    
  ' enter ��
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
                    
                Case "������ϸ��"
                    ForMod5
                
                Case "������ϸ��"
                    ForMod6
                    
                Case "������ϸ��(����)"
                    ForMod2
                
                Case "������ϸ��(����)"
                    ForMod1

            End Select
        
        Case "DEL"
            
            ForDel

        Case "EXIT"
            Unload Me
            
        Case "RET"
            Toolbar1.Buttons(1).Enabled = True
            Toolbar1.Buttons(1).Caption = "��ѯ"
            Toolbar1.Buttons(1).Image = 1
            Toolbar1.Buttons(3).Enabled = True
            Toolbar1.Buttons(3).Caption = "����"
            Toolbar1.Buttons(3).Image = 3
            Toolbar1.Buttons(5).Enabled = True
            Toolbar1.Buttons(5).Caption = "�޸�"
            Toolbar1.Buttons(5).Image = 4
            Toolbar1.Buttons(7).Enabled = True
            Toolbar1.Buttons(7).Caption = "ɾ��"
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
                    
                Case "������ϸ��(����)"
                
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
        MsgBox "��ѡ��ά������", vbInformation, "��ʾ"
        Exit Sub

    End If
          
    lb6.Visible = True
            
    lb7.Visible = True
            
    Text4.Visible = True
            
    Text5.Visible = True

    Select Case Combo1.text
        
        Case "������ϸ��"
        
            QueType5
                
        Case "������ϸ��"
        
            QueType6
            
        Case "������ϸ��(����)"
        
            QueType5
                
        Case "������ϸ��(����)"
        
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
    '        MsgBox "�����뷢Ʊ��", vbInformation, "��ʾ"
    '        Exit Sub
    '
    '    End If
    strInv = Trim$(Text1.text)

    a = Split(strInv, "/")
    
    leni = UBound(a) - LBound(a) + 1
    
    AddSql2 ("delete from erptemp.dbo.ksexport_temp where 1 = 1")

    For i = 0 To leni - 1
        
        If Get_SqlserverCnt("SELECT * FROM erpdata..tblsale A WHERE A.���۵���� = '" & a(i) & "'") = 0 Then
            
            strflag1 = 1
            
        End If
        
        strsssql = "select delivery from erpbase.dbo.tblCustomerShippingUp where delivery = '" & a(i) & "'"
        
        If Get_SqlserverCnt(strsssql) = 0 Then
            
            strflag2 = 1
            
        End If
        
        Select Case Combo1.text
                
            Case "������ϸ��"

                If strflag1 = 1 And strflag2 = 1 Then
        
                    MsgBox "û�д˷�Ʊ��" & a(i) & ",����������", vbInformation, "��ʾ"
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
    
        MsgBox "��ʼ���ڲ���ѡ����ڽ�������", vbInformation, "��ʾ"
            
        Exit Sub
    
    End If

    If Text1.text = "" Then
    
        If strstate = True Then
        
            strsql = "select '' as '��',����,��������,�Ϻ�,���,��Ʊ��,��������,����,���ص���,�ֲ���,�ֲ����,Ʒ��,������λ,�ұ�,�ܼ�,����,AWB#,Ŀ�ĵ�,����,�˵�����,��ע,id from erptemp.dbo.ksexport where flag = '0' and CONVERT(varchar(100),����ʱ��, 23) >= '" & strstart & "' and CONVERT(varchar(100),����ʱ��, 23) <= '" & strend & "' and ���� = '" & stridid & "' order by ����,id"
        
        Else
            Select Case Combo1.text
                
                Case "������ϸ��"
                
                    strsql = "select '' as '��',����,��������,�Ϻ�,���,��Ʊ��,��������,����,���ص���,�ֲ���,�ֲ����,Ʒ��,������λ,�ұ�,�ܼ�,����,AWB#,Ŀ�ĵ�,����,�˵�����,��ע,id from erptemp.dbo.ksexport where flag = '0' and CONVERT(varchar(100),����ʱ��, 23) >= '" & strstart & "' and CONVERT(varchar(100),����ʱ��, 23) <= '" & strend & "' and �������� <> '' order by ����,id"
            
                Case "������ϸ��(����)"
                
                    strsql = "select '' as '��',����,��������,�Ϻ�,���,��Ʊ��,��������,����,���ص���,�ֲ���,�ֲ����,Ʒ��,������λ,�ұ�,�ܼ�,����,AWB#,Ŀ�ĵ�,����,�˵�����,��ע,id from erptemp.dbo.ksexport where flag = '0' and CONVERT(varchar(100),����ʱ��, 23) >= '" & strstart & "' and CONVERT(varchar(100),����ʱ��, 23) <= '" & strend & "' and �������� = ''  order by ����,id"
                
            End Select
            
        
        End If
    
    Else
        
        strsql = "select '' as '��',����,��������,�Ϻ�,���,��Ʊ��,��������,����,���ص���,�ֲ���,�ֲ����,Ʒ��,������λ,�ұ�,�ܼ�,����,AWB#,Ŀ�ĵ�,����,�˵�����,��ע,id from erptemp.dbo.ksexport where ��Ʊ��  in (select distinct purchase from erptemp.dbo.ksexport_temp where 1 = 1)  and flag = '0' and CONVERT(varchar(100),����ʱ��, 23) >= '" & strstart & "' and CONVERT(varchar(100),����ʱ��, 23) <= '" & strend & "' order by ����,id "

    End If
    
    fpS(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType5(rs)
    Else
            
        strtet = 0
        
        strval = 0
        
        lb7 = "��������"
        
        strtet = Format(Trim$(strtet), "0.000")
        
        lb6 = "�����ܶ�"
        
        strval = Format(Trim$(strval), "0.0000")
        
        Text4.text = strval
    
        Text5.text = strtet
        
        MsgBox "��ѯ�����ó��ڵ�����Ϣ", vbInformation, "��ʾ"
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
                
            Case "������ϸ��"
        
                If Get_SqlserverCnt("SELECT * FROM erpbase..tblCPurDataSub WHERE �ɹ������ = '" & a(i) & "'") = 0 Then
                    MsgBox "û�д˲ɹ�����" & a(i) & ",����������", vbInformation, "��ʾ"
                    Exit Sub
    
                End If

        End Select
        
        AddSql2 (" insert into erptemp.dbo.ksimport_temp values('" & a(i) & "') ")

    Next
    
    strstart = Format(DTPicker1.Value, "yyyy-MM-dd")
    
    strend = Format(DTPicker2.Value, "yyyy-MM-dd")
    
    If strstart > strend Then
    
        MsgBox "��ʼ���ڲ���ѡ����ڽ�������", vbInformation, "��ʾ"
            
        Exit Sub
    
    End If
    
    If Text1.text = "" Then
    
        
        If strstate1 = True Then
    
            strsql = "select '' as '��',����,�ɹ�����,�Ϻ�,�ͺ�,���,��������,��׼die,��die��,�ֲ���,���,Ʒ��,������,������λ,�볡����,��Ʊ��,����,�ұ�,�ɹ�����,���ؽ��,����,��˰��,��˰,��ֵ˰��,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id from erptemp.dbo.ksimport where  flag = '0' and CONVERT(varchar(100),����ʱ��, 23) >= '" & strstart & "' and CONVERT(varchar(100),����ʱ��, 23) <= '" & strend & "' and ���� = '" & stridid1 & "' order by ����,id"
        
        Else
            
            Select Case Combo1.text
                
                Case "������ϸ��"
            
                    strsql = "select '' as '��',����,�ɹ�����,�Ϻ�,�ͺ�,���,��������,��׼die,��die��,�ֲ���,���,Ʒ��,������,������λ,�볡����,��Ʊ��,����,�ұ�,�ɹ�����,���ؽ��,����,��˰��,��˰,��ֵ˰��,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id from erptemp.dbo.ksimport where  flag = '0' and CONVERT(varchar(100),����ʱ��, 23) >= '" & strstart & "' and CONVERT(varchar(100),����ʱ��, 23) <= '" & strend & "' and �ɹ����� <> '' order by ����,id"
                
                Case "������ϸ��(����)"
                
                    strsql = "select '' as '��',����,�ɹ�����,�Ϻ�,�ͺ�,���,��������,��׼die,��die��,�ֲ���,���,Ʒ��,������,������λ,�볡����,��Ʊ��,����,�ұ�,�ɹ�����,���ؽ��,����,��˰��,��˰,��ֵ˰��,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id from erptemp.dbo.ksimport where  flag = '0' and CONVERT(varchar(100),����ʱ��, 23) >= '" & strstart & "' and CONVERT(varchar(100),����ʱ��, 23) <= '" & strend & "' and �ɹ����� = ''  order by ����,id"
            
            End Select
        End If
    
    Else
        
        strsql = "select '' as '��',����,�ɹ�����,�Ϻ�,�ͺ�,���,��������,��׼die,��die��,�ֲ���,���,Ʒ��,������,������λ,�볡����,��Ʊ��,����,�ұ�,�ɹ�����,���ؽ��,����,��˰��,��˰,��ֵ˰��,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id from erptemp.dbo.ksimport where �ɹ����� in (select distinct purchase from erptemp.dbo.ksimport_temp where 1 = 1) and flag = '0' and CONVERT(varchar(100),����ʱ��, 23) >= '" & strstart & "' and CONVERT(varchar(100),����ʱ��, 23) <= '" & strend & "' order by ����,id "
        
    End If
    
    fpS(0).MaxRows = 0
    fpS(0).MaxCols = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
       
        Call ListDataType6(rs)
       
    Else
            
        strtet = 0
        
        strval = 0
        
        lb7 = "��������"
        
        strtet = Format(Trim$(strtet), "0.000")
        
        lb6 = "�����ܶ�"
        
        strval = Format(Trim$(strval), "0.0000")
        
        Text4.text = strval
    
        Text5.text = strtet

        MsgBox "��ѯ�����òɹ�������Ϣ", vbInformation, "��ʾ"
        Exit Sub

    End If
  
    Select Case Combo1.text
                
        Case "������ϸ��"
        
            fpss(0).Visible = True
                
        Case "������ϸ��(����)"
        
            fpss(0).Visible = False

    End Select
    
    If Text1.text = "" Then
    
        strsql = "SELECT  �ɹ�����,�Ϻ�,isnull(sum(��������),0) as �������� from erptemp.dbo.ksimport where 1 = 1 and flag = '0' and CONVERT(varchar(100),����ʱ��, 23) >= '" & strstart & "' and CONVERT(varchar(100),����ʱ��, 23) <= '" & strend & "' AND �ɹ����� <> '' group by �ɹ�����,�Ϻ� order by �ɹ�����,�Ϻ�"
    
    Else
        
        strsql = "SELECT  �ɹ�����,�Ϻ�,isnull(sum(��������),0) as �������� from erptemp.dbo.ksimport where 1 = 1 and flag = '0' and �ɹ����� in (select distinct purchase from erptemp.dbo.ksimport_temp where 1 = 1) group by �ɹ�����,�Ϻ� order by �ɹ�����,�Ϻ� "
    
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
        
        lb7 = "��������"
        strtet = Format(Trim$(strtet), "0.000")
        
        lb6 = "�����ܶ�"
        
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
            '                Case "������ϸ��(����)"
                    
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
                       
            .TypeComboBoxList = .TypeComboBoxList & "���϶Կ�"
            
            .TypeComboBoxList = .TypeComboBoxList & "һ��ó��"
            
            .TypeComboBoxList = .TypeComboBoxList & "�������������"
            
            .TypeComboBoxList = .TypeComboBoxList & "��Ʒ����"
            
            .TypeComboBoxList = .TypeComboBoxList & "ά����Ʒ"
            
            .TypeComboBoxList = .TypeComboBoxList & "�ϼ�����"
            
            .TypeComboBoxList = .TypeComboBoxList & "���ϳ�Ʒ�˻�"
            
            .TypeComboBoxList = .TypeComboBoxList & "����"

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
        
        lb7 = "��������"
        
        strtet = Format(Trim$(strtet), "0.000")
        
        lb6 = "�����ܶ�"
        
        strval = Format(Trim$(strval), "0.0000")
        
        Text4.text = strval
    
        Text5.text = strtet

    End With
     
End Sub

Private Sub ForAdd()

    If Toolbar1.Buttons(3).Caption = "�ύ" Then
        
        Select Case Combo1.text
             
            Case "������ϸ��"
                ForCommit1
                
            Case "������ϸ��"
                ForCommit2

            Case "������ϸ��(����)"
                ForCommit1

            Case "������ϸ��(����)"
                ForCommit2

        End Select
        
        Exit Sub

    End If

    If Combo1.text = "" Then
        MsgBox "��ѡ��ά������", vbInformation, "��ʾ"
        Exit Sub

    End If

    Select Case Combo1.text

        Case "������ϸ��"
            AddType5

        Case "������ϸ��"
            AddType6
            
        Case "������ϸ��(����)"
            AddType1
        
        Case "������ϸ��(����)"
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

        Case "������ϸ��(����)"
        
            stridd = Get_SqlStr("select isnull(max(����),0) from erptemp.dbo.ksexport where flag = '0'")

        Case "������ϸ��(����)"
        
            stridd = Get_SqlStr("select isnull(max(����),0) from erptemp.dbo.ksimport where flag = '0'")
            
        Case "������ϸ��"
        
            stridd = Get_SqlStr("select isnull(max(����),0) from erptemp.dbo.ksexport where flag = '0'")

        Case "������ϸ��"
        
            stridd = Get_SqlStr("select isnull(max(����),0) from erptemp.dbo.ksimport where flag = '0'")

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
        
        .SetText E_FPS.E_gx, 0, "��"
        .SetText E_FPS.e_NO, 0, "����"
        .SetText E_FPS.E_exportno, 0, "��������"
        .SetText E_FPS.E_partno, 0, "�Ϻ�"
        .SetText E_FPS.E_modetrade, 0, "���"
        .SetText E_FPS.e_Invoice, 0, "��Ʊ��"
        .SetText E_FPS.E_exportdate, 0, "��������"
        .SetText E_FPS.E_exportquantity, 0, "��������"
        .SetText E_FPS.E_declarationno, 0, "���ص���"
        .SetText E_FPS.E_manualno, 0, "�ֲ���"
        .SetText E_FPS.E_itemno, 0, "�ֲ����"
        .SetText E_FPS.E_name, 0, "Ʒ��"
        .SetText E_FPS.E_UNIT, 0, "������λ"
        .SetText E_FPS.E_currency, 0, "�ұ�"
        .SetText E_FPS.E_totalprice, 0, "�ܼ�"
        .SetText E_FPS.E_unitprice, 0, "����"
        .SetText E_FPS.E_AWB, 0, "AWB#"
        .SetText E_FPS.E_destination, 0, "Ŀ�ĵ�"
        .SetText E_FPS.E_freight, 0, "����"
        .SetText E_FPS.E_chargebackdate, 0, "�˵�����"
        .SetText E_FPS.E_mark, 0, "��ע"
        .SetText E_FPS.e_ID, 0, "id"
        
        '        .RowHeight(0) = 22
        '        .RowHeight(-1) = 22

        .Col = E_FPS.E_gx    'ѡ��
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
                       
        .TypeComboBoxList = .TypeComboBoxList & "���϶Կ�"
            
        .TypeComboBoxList = .TypeComboBoxList & "һ��ó��"
            
        .TypeComboBoxList = .TypeComboBoxList & "�������������"
            
        .TypeComboBoxList = .TypeComboBoxList & "�����ϼ�����"
            
        .TypeComboBoxList = .TypeComboBoxList & "���ϳ�Ʒ�˻�"
            
        .TypeComboBoxList = .TypeComboBoxList & "������Ʒ"
            
        .TypeComboBoxList = .TypeComboBoxList & "�豸����"
            
        .TypeComboBoxList = .TypeComboBoxList & "����"
                
        .Col = E_FPS.E_currency
        
        .CellType = CellTypeComboBox
            
        .TypeComboBoxList = "USD"
            
        .TypeComboBoxList = .TypeComboBoxList & "JPY"

        .TypeComboBoxList = .TypeComboBoxList & "EUR"

        .TypeComboBoxList = .TypeComboBoxList & "RMB"
        
        strsql = "select distinct �ֲ��� from erptemp.dbo.ksmanual where 1 = 1 and flag = '2'"

        If rs.State = 1 Then rs.Close
        rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

        .Col = E_FPS.E_manualno
        .ColWidth(E_FPS.E_manualno) = 12
        .CellType = CellTypeComboBox

        rs.MoveFirst

        For i = 1 To rs.RecordCount

            .TypeComboBoxList = .TypeComboBoxList & rs("�ֲ���")
            rs.MoveNext
        Next
        
        rs.Clone
        
        Set rs = Nothing
        
        '        .ReDraw = True
        
    End With
    
    Toolbar1.Buttons(3).Caption = "�ύ"
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
        
        If Trim$(comBo2.text) = "���϶Կ�" Or Trim$(comBo2.text) = "���ϳ�Ʒ�˻�" Or Trim$(comBo2.text) = "�����ϼ�����" Then
                
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
    
    lb7 = "��������"
    
    lb6.Visible = True
    
    lb6 = "�����ܶ�"
    
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
        
        .SetText F_fp.F_gx, 0, "��"
        .SetText F_fp.F_no, 0, "����"
        .SetText F_fp.F_purchaseno, 0, "�ɹ�����"
        .SetText F_fp.F_partno, 0, "�Ϻ�"
        .SetText F_fp.F_modelno, 0, "�ͺ�"
        .SetText F_fp.F_modetrade, 0, "���"
        .SetText F_fp.F_orderqty, 0, "��������"
        .SetText F_fp.F_die, 0, "��׼die"
        .SetText F_fp.F_totaldie, 0, "��die��"
        .SetText F_fp.F_manualno, 0, "�ֲ���"
        .SetText F_fp.F_itemno, 0, "���"
        .SetText F_fp.F_name, 0, "Ʒ��"
        .SetText F_fp.F_baoguanqty, 0, "������"
        .SetText F_fp.F_unit, 0, "������λ"
        .SetText F_fp.F_indate, 0, "�볡����"
        .SetText F_fp.F_invoice, 0, "��Ʊ��"
        .SetText F_fp.F_caseqty, 0, "����"
        .SetText F_fp.F_currency, 0, "�ұ�"
        .SetText F_fp.F_unitprice, 0, "����"
        .SetText F_fp.F_baoguanvalue, 0, "���ؽ��"
        .SetText F_fp.F_rate, 0, "����"
        .SetText F_fp.F_tariffrate, 0, "��˰��"
        .SetText F_fp.F_tariff, 0, "��˰"
        .SetText F_fp.F_addtaxrate, 0, "��ֵ˰��"
        .SetText F_fp.F_addtax, 0, "��ֵ˰"
        .SetText F_fp.F_declarationno, 0, "���ص���"
        .SetText F_fp.F_awb, 0, "AWB#"
        .SetText F_fp.F_freight, 0, "����"
        .SetText F_fp.F_chargebackdate, 0, "�˵�����"
        .SetText F_fp.F_mark, 0, "��ע"
        .SetText F_fp.F_id, 0, "id"

        .Col = F_fp.F_gx    'ѡ��
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
                       
        .TypeComboBoxList = .TypeComboBoxList & "���϶Կ�"
            
        .TypeComboBoxList = .TypeComboBoxList & "һ��ó��"
            
        .TypeComboBoxList = .TypeComboBoxList & "�������������"
            
        .TypeComboBoxList = .TypeComboBoxList & "��Ʒ����"
            
        .TypeComboBoxList = .TypeComboBoxList & "ά����Ʒ"
            
        .TypeComboBoxList = .TypeComboBoxList & "�ϼ�����"
            
        .TypeComboBoxList = .TypeComboBoxList & "���ϳ�Ʒ�˻�"
            
        .TypeComboBoxList = .TypeComboBoxList & "����"

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
        
        strsql = "select distinct �ֲ��� from erptemp.dbo.ksmanual where 1 = 1 and flag = '1'"

        If rs.State = 1 Then rs.Close
        rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

        .Col = F_fp.F_manualno
        .ColWidth(F_fp.F_manualno) = 12
        .CellType = CellTypeComboBox

        rs.MoveFirst

        For i = 1 To rs.RecordCount

            .TypeComboBoxList = .TypeComboBoxList & rs("�ֲ���")
            rs.MoveNext
        Next
        
        rs.Clone
        
        Set rs = Nothing

    End With
    
    Toolbar1.Buttons(3).Caption = "�ύ"
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

    lb7 = "��������"
    
    lb6.Visible = True
    
    lb6 = "�����ܶ�"
    
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
        MsgBox "����дҪά���ķ�Ʊ��", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    If comBo2.text = "" Then
          
        MsgBox "��ѡ��ó�׷�ʽ", vbInformation, "��ʾ"
        Exit Sub
    
    End If
    
    
    If comBo2.text = "���϶Կ�" Or comBo2.text = "���ϳ�Ʒ�˻�" Or comBo2.text = "�����ϼ�����" Then
        
        If Combo3.text = "" Then
    
            MsgBox "��ѡ���ֲ����", vbInformation, "��ʾ"
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

        If Get_SqlserverCnt("SELECT * FROM erpdata..tblsale A WHERE A.���۵���� = '" & a(i) & "'") = 0 Then
            
            strflag1 = 1
            
        End If
        
        strsssql = "select DN from erpdata..tblStockNumTree where DN = '" & a(i) & "'"
        
        If Get_SqlserverCnt(strsssql) = 0 Then
            
            strflag2 = 1
            
        End If
        
        If strflag1 = 1 And strflag2 = 1 Then
        
            MsgBox "û�д˷�Ʊ��" & a(i) & ",����������", vbInformation, "��ʾ"
            Exit Sub
        
        End If
        
        strflag1 = 0
    
        strflag2 = 0
        
        AddSql2 (" insert into erptemp.dbo.ksexport_temp values('" & a(i) & "') ")
        
    Next

    strsql = " select '' as '��','" & stridd & "' as ���� ,b.���ݱ�� as ��������,c.�Ϻ�,'" & strcom1 & "' as ���,e.delivery as ��Ʊ��,  " & _
    " CONVERT(varchar(100), b.��������, 23) as ��������,CONVERT(decimal(19,3),(SUM(b.ʵ����Ʒ��+b.ʵ��������+b.ʵ���Ƴ̲�����)/1000.00)) as ����,'' as ���ص���, " & _
    " '" & strcom2 & "' as �ֲ���,'' as �ֲ����,'' as Ʒ��,'' as ������λ,'USD' as �ұ�, " & _
    " CONVERT(decimal(19,2),e.�ܼ�) as �ܼ�, " & _
    " CONVERT(decimal(19,6),e.�ܼ�/CONVERT(decimal(19,3),(SUM(b.ʵ����Ʒ��+b.ʵ��������+b.ʵ���Ƴ̲�����)/1000.00))) as ����,'' as AWB#,'' as Ŀ�ĵ�,'' as ����,'' as �˵�����,'' as ��ע,'' as id " & _
    " from  erpdata..tblSmainM2 c,erpdata..tblStockMove b " & _
    " LEFT JOIN (SELECT distinct p1.���ݱ��,d.DN as delivery,sum((ISNULL(p1.����, 0) + ISNULL(p1.�͹����ϵ���, 0)) * p1.����) AS �ܼ�,p1.�Ϻ� FROM erpdata..tblSaleRec p1  " & _
    " LEFT JOIN (select distinct p3.DN,p1.���ݱ�� from erpdata..tblSaleRec p1  inner join erpdata..tblStocksqfhsub p2 on p1.���ݱ�� = p2.���ݱ�� and p1.������� = p2.������� " & _
    " inner join erpdata..tblStockNumTree p3 on p3.��� = p2.���  where p3.DN in (select distinct purchase from erptemp.dbo.ksexport_temp where 1 = 1)) d on d.���ݱ�� = p1.���ݱ�� " & _
    " where p1.���ݱ�� in( select distinct p1.���ݱ�� from erpdata..tblSaleRec p1 inner join erpdata..tblStocksqfhsub p2 on p1.���ݱ�� = p2.���ݱ�� and p1.������� = p2.������� " & _
    " inner join erpdata..tblStockNumTree p3 on p3.��� = p2.���  where p3.DN in (select distinct purchase from erptemp.dbo.ksexport_temp where 1 = 1)) " & _
    " group by p1.���ݱ��,p1.�Ϻ�,d.DN UNION ALL " & _
    " SELECT distinct RTRIM(b.���ݱ��) as ���ݱ��,a.���۵���� as delivery,sum(b.���� * (b.�͹����ϵ��� + b.����)) AS �ܼ�,b.�Ϻ� FROM erpdata..tblsale a INNER JOIN erpdata..tblSaleRec b ON a.���۵���� = b.���۵���� " & _
    " where a.���۵���� in (select distinct purchase from erptemp.dbo.ksexport_temp where 1 = 1) group by b.���ݱ��,a.���۵����,b.�Ϻ�) e " & _
    " ON  e.���ݱ�� = b.���ݱ�� AND  e.delivery in (select distinct purchase from erptemp.dbo.ksexport_temp where 1 = 1) " & _
    " where e.���ݱ�� = b.���ݱ�� AND e.delivery in (select distinct purchase from erptemp.dbo.ksexport_temp where 1 = 1) AND c.���ϱ�� = b.���ϱ�� and e.�Ϻ� = c.�Ϻ� and c.�Ϻ� not in (select distinct �Ϻ� from erptemp.dbo.ksexport where �������� = e.���ݱ�� and flag = '0' )" & _
    " group by b.���ݱ��,c.�Ϻ�,CONVERT(varchar(100), b.��������, 23),e.delivery,e.�ܼ� "
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType5(rs)
    Else
            
        If Get_SqlserverCnt("select ��������  from erptemp.dbo.ksexport where ��Ʊ�� in (select distinct purchase from erptemp.dbo.ksexport_temp where 1 = 1)") <> 0 Then
            
                    
             MsgBox "�˱��Ѿ�������,���ʵ��", vbInformation, "��ʾ"
             Exit Sub
                    
         Else
         
         
             MsgBox "��ѯ�����ó��ڵ�����Ϣ", vbInformation, "��ʾ"
             Exit Sub
        
                    
        End If
        
       
    End If
    
    Toolbar1.Buttons(3).Caption = "�ύ"
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
            If .text = "���϶Կ�" Or .text = "���ϳ�Ʒ�˻�" Or .text = "�����ϼ�����" Then
                
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
'            strSql = "select distinct �ֲ��� from erptemp.dbo.ksmanual where 1 = 1 and flag = '2'"
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
'                .TypeComboBoxList = .TypeComboBoxList & Rs("�ֲ���")
'                Rs.MoveNext
'            Next
'
'            Rs.Clone
'
'            Set Rs = Nothing
'
        Next

    End With
    
    lb7 = "��������"
    
    lb6.Visible = True
    
    lb6 = "�����ܶ�"
    
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
    
        MsgBox "����дҪά���Ĳɹ�����", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    fpS(0).MaxRows = 0
    
    stridd = Createid

    strInv = Trim$(Text1.text)
    
    a = Split(strInv, "/")
    
    leni = UBound(a) - LBound(a) + 1
    
    AddSql2 ("delete from erptemp.dbo.ksimport_temp where 1 = 1")

    For i = 0 To leni - 1
    
        If Get_SqlserverCnt("SELECT * FROM erpbase..tblCPurDataSub WHERE �ɹ������ = '" & a(i) & "'") = 0 Then
            MsgBox "û�д˲ɹ�����" & a(i) & ",����������", vbInformation, "��ʾ"
            Exit Sub
    
        End If
        
        AddSql2 (" insert into erptemp.dbo.ksimport_temp values('" & a(i) & "') ")

    Next
    
    strsql = "SELECT '' as '��','" & stridd & "' as ����,a.�ɹ������,b.�Ϻ�,b.����ͺ� as �ͺ�,'' AS ���,ceiling(sum(a.��׼�ɹ�����) - isnull(c.��������,0)) as ��������,t7.qty1 as ��׼die," & _
    "ceiling(sum(a.��׼�ɹ�����) - isnull(c.��������,0))* t7.qty1 as ��die��,'' as �ֲ���,'' as ���,'' as Ʒ��,'0' as ������,'' as ������λ,'' as �볡����,'' as ��Ʊ��,'' as ����,'USD' as �ұ�,a.���� as �ɹ�����,((sum(a.��׼�ɹ�����) - isnull(c.��������,0))) * a.���� as ���ؽ��," & _
    " '' as ����,'' as ��˰��,'' as ��˰,'' as ��ֵ˰��,'' as ��ֵ˰,'' as ���ص���,'' as AWB#,'' as ����,'' as �˵�����,'' as ��ע," & _
    " '' as id FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b  left join (select �ɹ�����,�Ϻ�,isnull(sum(��������),0) as ��������, " & _
    " flag from erptemp.dbo.ksimport where flag = '0' group by �ɹ�����,�Ϻ�,flag) c on c.�Ϻ� = b.�Ϻ� and flag = '0' and c.�ɹ����� in (select distinct purchase from erptemp.dbo.ksimport_temp " & _
    " where 1 = 1) left join (select t1.�ɹ������,t6.�Ϻ�,isnull(t8.qty,0) as qty1 ,t1.�빺�����,t1.�빺����� ,t1.�ɹ������  from  erpbase..tblCPurDataSub t1  inner join  erpdata..tblSmainM2 t6  on t1.���ϱ�� = t6.���ϱ��  " & _
    " left join  (select m2.�Ϻ�,max(m1.QTECHDIEQTY) as qty  from erptemp..TBLTSVNPIPRODUCT m1,erpdata..TSVtblMRuleData m2 where 1=1  " & _
    " and m1.QTECHPTNO2 = m2.����� group by m2.�Ϻ�) t8 on t8.�Ϻ� = t6.�Ϻ� where 1=1 ) t7  on t7.�Ϻ� = b.�Ϻ�  and  " & _
    " t7.�ɹ������ in (select distinct purchase from erptemp.dbo.ksimport_temp where 1 = 1) " & _
    " WHERE a.�Ƿ���� = '0' and a.�ɹ������ in (select distinct purchase from erptemp.dbo.ksimport_temp where 1 = 1) and t7.�ɹ������ = a.�ɹ������  AND t7.�빺����� = a.�빺�����  AND t7.�빺����� = a.�빺�����  AND t7.�ɹ������ = a.�ɹ������ " & _
    " and a.���ϱ�� = b.���ϱ�� GROUP by a.����,b.����ͺ�,a.�ɹ������,b.�Ϻ�,c.��������,t7.qty1,a.�ɹ������ order by a.�ɹ������ "
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
    Call ListDataType6(rs)
    
    fpss(0).Visible = True
    
    strsql = "SELECT  �ɹ�����,�Ϻ�,isnull(sum(��������),0) as �������� from erptemp.dbo.ksimport where 1 = 1 and flag = '0' and �ɹ����� in (select distinct purchase from erptemp.dbo.ksimport_temp where 1 = 1) group by �ɹ�����,�Ϻ�"
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
    Call ListDataType1(rs)
    
    Toolbar1.Buttons(3).Caption = "�ύ"
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
            
            strsql = "select distinct �ֲ��� from erptemp.dbo.ksmanual where 1 = 1 and flag = '1'"

            If rs.State = 1 Then rs.Close
            rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

            .Col = F_fp.F_manualno

            .CellType = CellTypeComboBox

            rs.MoveFirst

            For j = 1 To rs.RecordCount

                .TypeComboBoxList = .TypeComboBoxList & rs("�ֲ���")
                rs.MoveNext
            Next
        
            rs.Clone
        
            Set rs = Nothing
            
        Next

    End With
    
    lb7 = "��������"
    
    lb6.Visible = True
    
    lb6 = "�����ܶ�"
    
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
            MsgBox "û������", vbInformation, "��ʾ"
        
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
                '����
                strInv21 = Trim$(.text)
                
                If Trim$(stridd) <> Trim$(strInv21) Then
                
                    MsgBox "�����б䶯,����Ϊ " & strInv21 & "", vbInformation, "��ʾ"
                    
'                    strInv21 = stridd
                
                End If
                
                .Col = E_FPS.E_exportno
                
                Select Case Combo1.text
                
                
                    Case "������ϸ��"
                        
                        If .text = "" Then
                            
                            MsgBox "�������������", vbInformation, "��ʾ"
                            Exit Sub

                        End If
                    
                    Case "������ϸ��(����)"
                            
                        .text = ""
                    
                End Select
    
                strInv1 = Trim$(.text)
    
                .Col = E_FPS.E_partno

                If .text = "" Then
                    MsgBox "�������Ϻ�", vbInformation, "��ʾ"
                    Exit Sub

                End If
    
                strInv2 = Trim$(.text)
                
                Select Case Combo1.text
                
                
                    Case "������ϸ��"
                        
                       If Get_SqlserverCnt("select * from erptemp.dbo.ksexport where �������� = '" & strInv1 & "' and �Ϻ� = '" & strInv2 & "' and flag = '0'") > 0 Then
                        
                            MsgBox "�ñ������Ѿ�������", vbInformation, "��ʾ"
                            Exit Sub

                       End If
                    
                End Select
                
                .Col = E_FPS.E_modetrade
                
                If .text = "" Then
                
                    MsgBox "��ѡ�����", vbInformation, "��ʾ"
                    Exit Sub

                End If
                
                strInv3 = Trim$(.text)
                
                .Col = E_FPS.e_Invoice
                
                strInv4 = Trim$(.text)
        
                .Col = E_FPS.E_exportdate
                
                If .text = "" Then
                
                    MsgBox "�������������", vbInformation, "��ʾ"
                    Exit Sub

                End If
                
                strInv5 = Trim$(.text)
        
                .Col = E_FPS.E_exportquantity
                
                If .text = "" Then
                    
                    MsgBox "�������������", vbInformation, "��ʾ"
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
                
                If strInv3 = "���϶Կ�" Or strInv3 = "���ϳ�Ʒ�˻�" Or strInv3 = "�����ϼ�����" Then
                'Ʒ��
                    If strInv3 = "���϶Կ�" Or strInv3 = "���ϳ�Ʒ�˻�" Then
                    
                        If Get_SqlserverCnt("SELECT ��Ʒ���� FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & strInv8 & "' and flag = '2' and ���= '" & strInv9 & "'") = 0 Then
                                    
                            MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"
                                    
                            Exit Sub
                                
                        End If
                    
                        stritemname = Get_SqlStr("select distinct ��Ʒ���� from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '2' and  ��� = '" & strInv9 & "'")

                    Else
                    
                        If Get_SqlserverCnt("SELECT ��Ʒ���� FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & strInv8 & "' and flag = '1' and ���= '" & strInv9 & "'") = 0 Then
                                    
                            MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"
                                    
                            Exit Sub
                                
                        End If
                    
                        stritemname = Get_SqlStr("select distinct ��Ʒ���� from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '1' and  ��� = '" & strInv9 & "'")
                    
                    End If
                    
                    .text = stritemname


                End If
                
                strInv10 = Trim$(.text)
        
                .Col = E_FPS.E_UNIT
                
                '������λ
                If strInv3 = "���϶Կ�" Or strInv3 = "���ϳ�Ʒ�˻�" Or strInv3 = "�����ϼ�����" Then
                    
                    If strInv3 = "���϶Կ�" Or strInv3 = "���ϳ�Ʒ�˻�" Then
                
                        strunit = Get_SqlStr("select distinct ������λ from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '2' and  ��� = '" & strInv9 & "'")

                    Else
                    
                        strunit = Get_SqlStr("select distinct ������λ from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '1' and  ��� = '" & strInv9 & "'")
                        
                    End If
                    
                    .text = strunit
                    
                End If

                strInv11 = Trim$(.text)
        
                .Col = E_FPS.E_currency
                strInv12 = Trim$(.text)
        
                .Col = E_FPS.E_totalprice
                '�ܼ�
                If .text = "" Then
                    MsgBox "�������ܼ�", vbInformation, "��ʾ"
                    Exit Sub

                End If
    
                strInv13 = Trim$(.text)
        
                .Col = E_FPS.E_unitprice
                
                If strInv13 <> "" And .text = "" Then
                    
                    .text = Format(Val(strInv13) / Val(strInv6), "0.000000")
                
                End If
                '����
                
                strInv14 = Trim$(.text)
        
                .Col = E_FPS.E_AWB
                strInv15 = Trim$(.text)
        
                .Col = E_FPS.E_destination
                strInv16 = Trim$(.text)
                
                .Col = E_FPS.E_freight
                strInv17 = Trim$(.text)
                
                .Col = E_FPS.E_chargebackdate
                
                Select Case Combo1.text
                
                
                    Case "������ϸ��(����)"
                        
                        .text = ""
                        
                End Select
                
                strInv18 = Trim$(.text)
                
                .Col = E_FPS.E_mark
                
                Select Case Combo1.text
                
                
                    Case "������ϸ��(����)"
                    
                        .text = ""
                    
                End Select
                
                strInv19 = Trim$(.text)
                
                .Col = E_FPS.e_ID
                strInv20 = Trim$(i)
                
                
                AddSql2 ("insert into erptemp.dbo.ksexport( ����,��������,�Ϻ�,���,��Ʊ��,��������,����,���ص���,�ֲ���,�ֲ����,Ʒ��,������λ,�ұ�,�ܼ�,����,AWB#,Ŀ�ĵ�,����,�˵�����,��ע,����ʱ��,�޸�״̬,�޸�ʱ��,ɾ��ʱ��,flag,id) values('" & strInv21 & "','" & strInv1 & "','" & strInv2 & "','" & strInv3 & "','" & strInv4 & "','" & strInv5 & "','" & strInv6 & "','" & strInv7 & "','" & strInv8 & "','" & strInv9 & "','" & strInv10 & "','" & strInv11 & "','" & strInv12 & "','" & strInv13 & "','" & strInv14 & "','" & strInv15 & "','" & strInv16 & "','" & strInv17 & "','" & strInv18 & "','" & strInv19 & "',GetDate(),NULL,NULL,NULL,'0','" & strInv20 & "')")

            End If

        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "��ѡ��Ҫ��������", vbInformation, "��ʾ"
            Exit Sub
            
        End If

    End With
    
    MsgBox "�����ɹ�", vbInformation, "��ʾ"
    Toolbar1.Buttons(3).Caption = "����"
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
            MsgBox "û������", vbInformation, "��ʾ"
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
                
                    MsgBox "�����б䶯,����Ϊ " & strInv29 & "", vbInformation, "��ʾ"
                    
'                    strInv29 = stridd

                End If
                
                .Col = F_fp.F_purchaseno
                
                Select Case Combo1.text
                
                    Case "������ϸ��(����)"
                    
                        .text = ""
                        
                    Case "������ϸ��"
                    
                        If .text = "" Then
                            
                            MsgBox "������ɹ�����", vbInformation, "��ʾ"
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
                    MsgBox "���������", vbInformation, "��ʾ"
                    Exit Sub

                End If
                
                strInv4 = Trim$(.text)
                              
                .Col = F_fp.F_orderqty
                
                Select Case Combo1.text
                
                    Case "������ϸ��(����)"
                    
                        .text = 0
                        
                    Case "������ϸ��"
                        
                        strInv5 = Trim$(.text)
                
                        strNo1 = Get_SqlStr("SELECT isnull(SUM(a.��׼�ɹ�����),0) FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b WHERE a.�ɹ������ = '" & strInv1 & "' and a.���ϱ�� = b.���ϱ�� and b.�Ϻ� = '" & strInv2 & "' ")
                    
                        strNo2 = Get_SqlStr("SELECT isnull(SUM(��������),0) FROM erptemp.dbo.ksimport where �ɹ����� = '" & strInv1 & "' and �Ϻ� = '" & strInv2 & "' and flag = '0'")
                    
                        strNo3 = Val(strNo1) - Val(strNo2)
                    
                        If Val(strInv5) > Val(strNo3) Then
                            
                            MsgBox "�ñ��Ϻ�" & strInv2 & "��׼�ɹ�����: " & strNo1 & ",�Ѿ�ά������������" & strNo2 & ",�������ֻ��ά����" & strNo3 & "", vbInformation, "��ʾ"
                            Exit Sub

                        End If
                    
                        If Val(strInv5) <= 0 Then
                    
                            MsgBox "������������С�ڵ���0", vbInformation, "��ʾ"
                            Exit Sub
                        
                        End If
            
                End Select
                
                strInv5 = Format(Trim$(.text), "0.00")
                
                .Col = F_fp.F_die
                '��׼die
                
                Select Case Combo1.text
                
                    Case "������ϸ��(����)"
                    
                        If .text = "" Then
                    
                            .text = 0
                    
                        End If

                End Select
                
                strInv6 = Format(Trim$(.text), "0.00")
        
                .Col = F_fp.F_totaldie
                '��die����
                
                Select Case Combo1.text
                
                    Case "������ϸ��(����)"
                    
                        If .text = "" Then
                    
                            .text = 0
                    
                        End If
                         
                        strInv7 = Format(Trim$(.text), "0.00")
                    
                    Case "������ϸ��"
                        
                        strInv7 = Val(strInv5) * Val(strInv6)

                End Select
        
                .Col = F_fp.F_manualno
                '�ֲ��
                strInv8 = Trim$(.text)
                
                .Col = F_fp.F_itemno
                '���
                
                strInv9 = Trim$(.text)
        
                .Col = F_fp.F_name

                'Ʒ��
                If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                
                    If Get_SqlserverCnt("SELECT ��Ʒ���� FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & strInv8 & "' and flag = '1' and ���= '" & strInv9 & "'") = 0 Then
                                    
                        MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"

                        Exit Sub
                    
                    End If
                                
                    stritemname = Get_SqlStr("select distinct ��Ʒ���� from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '1' and  ��� = '" & strInv9 & "'")

                    .text = stritemname

                End If
                
                strInv10 = Trim$(.text)
                
                .Col = F_fp.F_baoguanqty
                
                Select Case Combo1.text
                
                    Case "������ϸ��(����)"
                    
                        If .text = "" Then
                    
                            MsgBox "�����뱨������", vbInformation, "��ʾ"
                            Exit Sub
                    
                        End If
                    
                        If Val(.text) <= 0 Then
                        
                            MsgBox "������������С�ڵ���0", vbInformation, "��ʾ"
                            Exit Sub
                    
                        End If
                
                        strInv11 = Format(Trim$(.text), "0.0000")
                
                    Case "������ϸ��"

                        '��������
                        If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                
                            strbaono1 = Get_SqlStr("select isnull(�걨����,0) from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '1' and  ��� = '" & strInv9 & "'")
                
                            strbaono2 = Get_SqlStr("select isnull(sum(������),0) from erptemp.dbo.ksimport where flag = '0' and  �ɹ����� = ' " & strInv1 & "' and �Ϻ� = '" & strInv2 & "'")
                
                            strbaono3 = strbaono1 - strbaono2
                
                            If .text = "" Then
                
                                MsgBox "�����뱨������", vbInformation, "��ʾ"
                                Exit Sub
                
                            End If
                
                            If Val(.text) <= 0 Then
                
                                MsgBox "������������С�ڵ���0", vbInformation, "��ʾ"
                                Exit Sub
                
                            End If
                
                            strInv11 = Format(Trim$(.text), "0.000")
                
                            If Val(strInv11) > Val(strbaono3) Then
                
                                MsgBox "����ı���������������ķ�Χ,�걨����Ϊ" & strbaono1 & ",Ŀǰϵͳ��¼������ " & strbaono2 & "", vbInformation, "��ʾ"
                
                            End If
                
                        Else
                
                            If .text = "" Then
                    
                                MsgBox "�����뱨������", vbInformation, "��ʾ"
                                Exit Sub
                    
                            End If
                    
                            If Val(.text) <= 0 Then
                        
                                MsgBox "������������С�ڵ���0", vbInformation, "��ʾ"
                                Exit Sub
                    
                            End If
                
                            strInv11 = Format(Trim$(.text), "0.000")
                    
                        End If

                End Select
                
                .Col = F_fp.F_unit

                '������λ
                If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                
                    strunit = Get_SqlStr("select distinct ������λ from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '1' and  ��� = '" & strInv9 & "'")

                    .text = strunit

                End If
                
                strInv12 = Trim$(.text)
            
                .Col = F_fp.F_indate
                If Trim$(.text) <> "" And Len(Trim$(.text)) <> 8 Then
                    MsgBox "������������YYYYMMDD��ʽ��д,��20200501", vbInformation, "��ʾ"
                    Exit Sub
                End If
                '�볡����
                strInv13 = Trim$(.text)
              
                .Col = F_fp.F_invoice
                '��Ʊ��
                strInv14 = Trim$(.text)
                
                .Col = F_fp.F_caseqty
                '����
                strInv15 = Trim$(.text)
                
                .Col = F_fp.F_currency
                '�ұ�
                strInv16 = Trim$(.text)
                
                .Col = F_fp.F_unitprice

                '�ɹ�����
                If .text = "" Then
                    
                    .text = 0
                    
                End If

                strInv30 = Format(Trim$(.text), "0.000")
                
                .Col = F_fp.F_baoguanvalue

                '���ؽ��
                If .text = "" Then
                    
                    .text = 0
                    
                End If
                
                strInv17 = Format(Trim$(.text), "0.0000")
                
                Select Case Combo1.text
                
                    Case "������ϸ��(����)"
    
                            strInv30 = Format(Trim$(Val(strInv17) / Val(strInv11)), "0.000")

                End Select
                
                .Col = F_fp.F_rate

                '����
                If .text = "" Then
                    
                    .text = 0
                    
                End If

                strInv18 = Format(Trim$(.text), "0.0000")
    
                .Col = F_fp.F_tariffrate

                '��˰��
                If .text = "" Then
                    
                    .text = 0
                    
                End If

                strInv19 = Format(Trim$(.text), "0.0000")
                    
                .Col = F_fp.F_tariff
                '��˰
                
                .text = Val(strInv18) * Val(strInv17) * Val(strInv19)
                    
                strInv20 = Format(Trim$(.text), "0.00")
                    
                .Col = F_fp.F_addtaxrate

                '��ֵ˰��
                If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                
                    .text = 0
                    strInv21 = Format(Trim$(.text), "0.0000")
                    
                Else
                    .text = 0.13
                    strInv21 = Format(Trim$(.text), "0.0000")
                    
                End If

                .Col = F_fp.F_addtax
                '��ֵ˰=����˰+��ֵ*���ʣ�*0.16
                    
                .text = Val(strInv20) * Val(strInv21) + Val(strInv17) * Val(strInv21) * Val(strInv18)
                    
                strInv22 = Format(Trim$(.text), "0.00")
                            
                .Col = F_fp.F_declarationno
                '���ص���
                
                strInv23 = Trim$(.text)
                
                .Col = F_fp.F_awb
                'AWB#
                
                strInv24 = Trim$(.text)
                
                .Col = F_fp.F_freight
                '����
        
                strInv25 = Trim$(.text)
                
                .Col = F_fp.F_chargebackdate
                '�˵�����
                
                strInv26 = Trim$(.text)
                
                .Col = F_fp.F_mark
                '��ע
                strInv27 = Trim$(.text)
                
                .Col = F_fp.F_id
                
                strInv28 = Trim$(i)
                
                AddSql2 ("insert into erptemp.dbo.ksimport( ����,�ɹ�����,�Ϻ�,�ͺ�,���,��������,��׼die,��die��,�ֲ���,���,Ʒ��,������,������λ,�볡����,��Ʊ��,����,�ұ�,�ɹ�����,���ؽ��,����,��˰��,��˰,��ֵ˰��,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id,����ʱ��,�޸�״̬,�޸�ʱ��,ɾ��ʱ��,flag) values('" & strInv29 & "','" & strInv1 & "','" & strInv2 & "','" & strInv3 & "','" & strInv4 & "','" & strInv5 & "','" & strInv6 & "','" & strInv7 & "','" & strInv8 & "','" & strInv9 & "','" & strInv10 & "','" & strInv11 & "','" & strInv12 & "','" & strInv13 & "','" & strInv14 & "','" & strInv15 & "','" & strInv16 & "','" & strInv30 & "','" & strInv17 & "','" & strInv18 & "','" & strInv19 & "','" & strInv20 & "','" & strInv21 & "','" & strInv22 & "','" & strInv23 & "','" & strInv24 & "','" & strInv25 & "','" & strInv26 & "','" & strInv27 & "','" & strInv28 & "',GetDate(),NULL,NULL,NULL,'0')")
            
            End If
            
        Next
        
        'j = 0 ��ȡ�����û���Ҫ���������
        If bFlag = False And j = 0 Then
            MsgBox "��ѡ��Ҫ��������", vbInformation, "��ʾ"
            Exit Sub
            
        End If

    End With
    
    MsgBox "�����ɹ�", vbInformation, "��ʾ"
    Toolbar1.Buttons(3).Caption = "����"
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

    If Toolbar1.Buttons(5).Caption <> "�ύ" Then

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
                If .text = "���϶Կ�" Or .text = "���ϳ�Ʒ�˻�" Or .text = "�����ϼ�����" Then
                
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

                .TypeComboBoxList = .TypeComboBoxList & "���϶Կ�"
    
                .TypeComboBoxList = .TypeComboBoxList & "һ��ó��"

                .TypeComboBoxList = .TypeComboBoxList & "�������������"

                .TypeComboBoxList = .TypeComboBoxList & "�����ϼ�����"
    
                .TypeComboBoxList = .TypeComboBoxList & "���ϳ�Ʒ�˻�"

                .TypeComboBoxList = .TypeComboBoxList & "������Ʒ"

                .TypeComboBoxList = .TypeComboBoxList & "�豸����"

                .TypeComboBoxList = .TypeComboBoxList & "����"
    
                 
                .Col = E_FPS.E_currency
                .CellType = CellTypeComboBox
            
                .TypeComboBoxList = .TypeComboBoxList & "USD"
            
                .TypeComboBoxList = .TypeComboBoxList & "JPY"

                .TypeComboBoxList = .TypeComboBoxList & "EUR"

                .TypeComboBoxList = .TypeComboBoxList & "RMB"
                                
                .Col = E_FPS.E_exportquantity
            
                .text = Format(Trim$(.text), "0.000")
                
                .LockBackColor = vbYellow
                
                strsql = "select distinct �ֲ��� from erptemp.dbo.ksmanual where 1 = 1 and flag = '2'"

                If rs.State = 1 Then rs.Close
                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                .Col = E_FPS.E_manualno

                .CellType = CellTypeComboBox

     '           .TypeComboBoxList = ""

                rs.MoveFirst

                For j = 1 To rs.RecordCount

                    .TypeComboBoxList = .TypeComboBoxList & rs("�ֲ���")
                    rs.MoveNext
                Next
        
                rs.Clone
        
                Set rs = Nothing
    
            Next
        
        End With
    
        Toolbar1.Buttons(5).Caption = "�ύ"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    bFlag = False
    
    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "û������", vbInformation, "��ʾ"
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
                    MsgBox "���������", vbInformation, "��ʾ"
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
                
                If strInv3 = "���϶Կ�" Or strInv3 = "���ϳ�Ʒ�˻�" Or strInv3 = "�����ϼ�����" Then
                'Ʒ��
                    If strInv3 = "���϶Կ�" Or strInv3 = "���ϳ�Ʒ�˻�" Then
                        
                        If Get_SqlserverCnt("SELECT ��Ʒ���� FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & strInv8 & "' and flag = '2' and ���= '" & strInv9 & "'") = 0 Then
                                    
                            MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"
                                    
                            Exit Sub
                                
                        End If
                    
                        stritemname = Get_SqlStr("select distinct ��Ʒ���� from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '2' and  ��� = '" & strInv9 & "'")
                    
                    Else
                    
                        If Get_SqlserverCnt("SELECT ��Ʒ���� FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & strInv8 & "' and flag = '1' and ���= '" & strInv9 & "'") = 0 Then
                                    
                            MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"
                                    
                            Exit Sub
                                
                        End If
                    
                        stritemname = Get_SqlStr("select distinct ��Ʒ���� from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '1' and  ��� = '" & strInv9 & "'")
                    
                    End If
                    
                    .text = stritemname

                End If
                
                strInv10 = Trim$(.text)
        
                .Col = E_FPS.E_UNIT
                
                '������λ
                If strInv3 = "���϶Կ�" Or strInv3 = "���ϳ�Ʒ�˻�" Or strInv3 = "�����ϼ�����" Then
                
                    If strInv3 = "���϶Կ�" Or strInv3 = "���ϳ�Ʒ�˻�" Then
                    
                        strunit = Get_SqlStr("select distinct ������λ from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '2' and  ��� = '" & strInv9 & "'")
                    
                    Else
                        
                        strunit = Get_SqlStr("select distinct ������λ from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '1' and  ��� = '" & strInv9 & "'")
                    
                    End If
                    
                    .text = strunit
                
                End If
                
                strInv11 = Trim$(.text)
        
                .Col = E_FPS.E_currency
                strInv12 = Trim$(.text)
        
                .Col = E_FPS.E_totalprice
                '�ܼ�
                
                If .text = "" Then
                
                    MsgBox "�������ܼ�", vbInformation, "��ʾ"
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
                
                AddSql2 ("insert into erptemp.dbo.ksexport (����,��������,�Ϻ�,���,��Ʊ��,��������,����,���ص���,�ֲ���,�ֲ����,Ʒ��,������λ,�ұ�,�ܼ�,����,AWB#,Ŀ�ĵ�,����,�˵�����,��ע,����ʱ��,�޸�״̬,�޸�ʱ��,ɾ��ʱ��,flag,id) SELECT ����,��������,�Ϻ�,���,��Ʊ��,��������,����,���ص���,�ֲ���,�ֲ����,Ʒ��,������λ,�ұ�,�ܼ�,����,AWB#,Ŀ�ĵ�,����,�˵�����,��ע,����ʱ��,'�޸�ǰ',�޸�ʱ��,ɾ��ʱ��,'2',id FROM erptemp.dbo.ksexport WHERE ���� = '" & strInv21 & "' and id = '" & strInv20 & "' AND  flag = '0'")
                
                AddSql2 ("update erptemp.dbo.ksexport set �Ϻ� =  '" & strInv2 & "', ��� =  '" & strInv3 & "',���ص��� =  '" & strInv7 & "',�ֲ��� =  '" & strInv8 & "',�ֲ���� =  '" & strInv9 & "',Ʒ�� =  '" & strInv10 & "',������λ =  '" & strInv11 & "',�ұ� =  '" & strInv12 & "',�ܼ� =  '" & strInv13 & "',���� = '" & strInv14 & "',AWB# =  '" & strInv15 & "',Ŀ�ĵ� =  '" & strInv16 & "',���� =  '" & strInv17 & "',�˵����� =  '" & strInv18 & "',��ע =  '" & strInv19 & "',�޸�״̬ = '�޸ĺ�',�޸�ʱ�� = '" & strtime & "' where ���� = '" & strInv21 & "' and id = '" & strInv20 & "' and flag = '0'")
            
            End If
            
        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "��ѡ��Ҫ�޸ĵ���", vbInformation, "��ʾ"
            Exit Sub
            
        End If
    
    End With
    
    MsgBox "�޸ĳɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(5).Caption = "�޸�"
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

    If Toolbar1.Buttons(5).Caption <> "�ύ" Then

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
                If .text = "���϶Կ�" Or .text = "��Ʒ����" Then
                
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
                    
                    '��˰
                    .Col = F_fp.F_tariff
                    .Lock = False
                    
                    '��ֵ˰
                    .Col = F_fp.F_addtax
                    .Lock = False
                        
                End If

                .LockBackColor = vbYellow
                
                strsql = "select distinct �ֲ��� from erptemp.dbo.ksmanual where 1 = 1 and flag = '1'"

                If rs.State = 1 Then rs.Close
                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                .Col = F_fp.F_manualno

                .CellType = CellTypeComboBox

                rs.MoveFirst

                For j = 1 To rs.RecordCount

                    .TypeComboBoxList = .TypeComboBoxList & rs("�ֲ���")
                    rs.MoveNext
                    
                Next
        
                rs.Clone
        
                Set rs = Nothing
    
            Next
        
        End With
    
        Toolbar1.Buttons(5).Caption = "�ύ"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    bFlag = False
    
    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "û������", vbInformation, "��ʾ"
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
                    MsgBox "���������", vbInformation, "��ʾ"
                    Exit Sub

                End If
                
                strInv4 = Trim$(.text)
                
                .Col = F_fp.F_orderqty

                strInv5 = Trim$(.text)
    
                .Col = F_fp.F_die

                '��׼die
                
                strInv6 = Trim$(.text)
                
                .Col = F_fp.F_totaldie
                '��die����
                
                strInv7 = Trim$(.text)
        
                .Col = F_fp.F_manualno
                strInv8 = Trim$(.text)
        
                .Col = F_fp.F_itemno
                strInv9 = Trim$(.text)
                
                .Col = F_fp.F_name
                
                If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                
                    If Get_SqlserverCnt("SELECT ��Ʒ���� FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & strInv8 & "' and flag = '1' and ���= '" & strInv9 & "'") = 0 Then
                                    
                        MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"

                        Exit Sub
                    
                    End If
                                
                    stritemname = Get_SqlStr("select distinct ��Ʒ���� from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '1' and  ��� = '" & strInv9 & "'")

                    .text = stritemname

                End If
                
                strInv10 = Trim$(.text)
        
                .Col = F_fp.F_baoguanqty
                '��������
             
                If .text = "" Then
                    
                    MsgBox "�����뱨������", vbInformation, "��ʾ"
                    Exit Sub
                    
                End If
                    
                If Val(.text) <= 0 Then
                        
                    MsgBox "������������С�ڵ���0", vbInformation, "��ʾ"
                    Exit Sub
                
                End If
                
                strInv11 = Format(Trim$(.text), "0.000")
                    
                .Col = F_fp.F_unit
                
                If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                
                    strunit = Get_SqlStr("select distinct ������λ from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '1' and  ��� = '" & strInv9 & "'")
                
                    .text = strunit
                                
                End If
                
                strInv12 = Trim$(.text)
        
                .Col = F_fp.F_indate
                If Trim$(.text) <> "" And Len(Trim$(.text)) <> 8 Then
                    MsgBox "������������YYYYMMDD��ʽ��д,��20200501", vbInformation, "��ʾ"
                    Exit Sub
                End If
                
                '�볡����
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
                '����
                If .text = "" Then
                    
                    .text = 0
                    
                End If
                strInv18 = Format(Trim$(.text), "0.0000")
                
                .Col = F_fp.F_tariffrate
                '��˰��
                If .text = "" Then
                    
                    .text = 0
                    
                End If
                
                strInv19 = Format(Trim$(.text), "0.0000")
                
                .Col = F_fp.F_id
                strInv28 = Trim$(.text)
                
                strbao1 = Get_SqlStr("select distinct ���ؽ�� from erptemp.dbo.ksimport where ���� = '" & strInv29 & "' and id = '" & strInv28 & "' and flag = '0' ")
                
                strbao2 = Get_SqlStr("select distinct ���� from erptemp.dbo.ksimport where ���� = '" & strInv29 & "' and id = '" & strInv28 & "' and flag = '0' ")
                
                strbao3 = Get_SqlStr("select distinct ��˰�� from erptemp.dbo.ksimport where ���� = '" & strInv29 & "' and id = '" & strInv28 & "' and flag = '0' ")
                
                .Col = F_fp.F_tariff
                
                '��˰
                
                If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                    
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
                
                '��ֵ˰��
                If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                
                    .text = 0
                    strInv21 = Format(Trim$(.text), "0.0000")
                        
                Else
                
                    .text = 0.13
                    strInv21 = Format(Trim$(.text), "0.0000")
                    
                End If
                
                .Col = F_fp.F_addtax
                
                '��ֵ˰
                
                If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                
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
                
                AddSql2 ("insert into erptemp.dbo.ksimport(����,�ɹ�����,�Ϻ�,�ͺ�,���,��������,��׼die,��die��,�ֲ���,���,Ʒ��,������,������λ,�볡����,��Ʊ��,����,�ұ�,�ɹ�����,���ؽ��,����,��˰��,��˰,��ֵ˰��,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id,����ʱ��,�޸�״̬,�޸�ʱ��,ɾ��ʱ��,flag) SELECT ����,�ɹ�����,�Ϻ�,�ͺ�,���,��������,��׼die,��die��,�ֲ���,���,Ʒ��,������,������λ,�볡����,��Ʊ��,����,�ұ�,�ɹ�����,���ؽ��,����,��˰��,��˰,��ֵ˰��,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id,����ʱ��,'�޸�ǰ',�޸�ʱ��,ɾ��ʱ��,'2' FROM erptemp.dbo.ksimport WHERE ���� =  '" & strInv29 & "' AND id =  '" & strInv28 & "'  AND  flag = '0'")
                
                AddSql2 ("update erptemp.dbo.ksimport set �Ϻ� = '" & strInv2 & "', �ͺ� = '" & strInv3 & "',��� = '" & strInv4 & "',�������� = '" & strInv5 & "',��׼die =  '" & strInv6 & "',��die�� =  '" & strInv7 & "',�ֲ��� = '" & strInv8 & "',��� = '" & strInv9 & "',Ʒ�� =  '" & strInv10 & "',������ =  '" & strInv11 & "', " & " ������λ  =  '" & strInv12 & "',�볡���� =  '" & strInv13 & "',��Ʊ�� =  '" & strInv14 & "',���� =  '" & strInv15 & "',�ұ� =  '" & strInv16 & "',���ؽ�� =  '" & strInv17 & "',���� =  '" & strInv18 & "',��˰�� =  '" & strInv19 & "',��˰ =  '" & strInv20 & "',��ֵ˰�� =  '" & strInv21 & "',��ֵ˰ =  '" & strInv22 & "',���ص��� =  '" & strInv23 & "',AWB#  =  '" & strInv24 & "',���� =  '" & strInv25 & "',�˵����� =  '" & strInv26 & "',��ע =  '" & strInv27 & "',�޸�״̬ = '�޸ĺ�',�޸�ʱ�� = '" & strtime & "' where ���� =  '" & strInv29 & "'  and flag = '0'  and id =  '" & strInv28 & "' ")
            
            End If
            
        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "��ѡ��Ҫ�޸ĵ���", vbInformation, "��ʾ"
            Exit Sub
            
        End If
    
    End With
    
    MsgBox "�޸ĳɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(5).Caption = "�޸�"
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

    If Toolbar1.Buttons(5).Caption <> "�ύ" Then

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
                If .text = "���϶Կ�" Or .text = "���ϳ�Ʒ�˻�" Or .text = "�����ϼ�����" Then
                
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

                .TypeComboBoxList = .TypeComboBoxList & "���϶Կ�"
    
                .TypeComboBoxList = .TypeComboBoxList & "һ��ó��"

                .TypeComboBoxList = .TypeComboBoxList & "�������������"

                .TypeComboBoxList = .TypeComboBoxList & "�����ϼ�����"
    
                .TypeComboBoxList = .TypeComboBoxList & "���ϳ�Ʒ�˻�"

                .TypeComboBoxList = .TypeComboBoxList & "������Ʒ"

                .TypeComboBoxList = .TypeComboBoxList & "�豸����"

                .TypeComboBoxList = .TypeComboBoxList & "����"
    
                 
                .Col = 14
                .CellType = CellTypeComboBox
            
                .TypeComboBoxList = .TypeComboBoxList & "USD"
            
                .TypeComboBoxList = .TypeComboBoxList & "JPY"

                .TypeComboBoxList = .TypeComboBoxList & "EUR"

                .TypeComboBoxList = .TypeComboBoxList & "RMB"
                                
                .Col = 8
            
                .text = Format(Trim$(.text), "0.000")
                
                .LockBackColor = vbYellow
                
                strsql = "select distinct �ֲ��� from erptemp.dbo.ksmanual where 1 = 1 and flag = '2'"

                If rs.State = 1 Then rs.Close
                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                .Col = 10

                .CellType = CellTypeComboBox

     '           .TypeComboBoxList = ""

                rs.MoveFirst

                For j = 1 To rs.RecordCount

                    .TypeComboBoxList = .TypeComboBoxList & rs("�ֲ���")
                    rs.MoveNext
                Next
        
                rs.Clone
        
                Set rs = Nothing
    
            Next
        
        End With
    
        Toolbar1.Buttons(5).Caption = "�ύ"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    bFlag = False
    
    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "û������", vbInformation, "��ʾ"
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
'                    MsgBox "�������������", vbInformation, "��ʾ"
'                    Exit Sub
'
'                End If
                
                strInv1 = Trim$(.text)
    
                .Col = 4
                strInv2 = Trim$(.text)
    
                .Col = 5
                
                If .text = "" Then
                    MsgBox "���������", vbInformation, "��ʾ"
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
                
                If strInv3 = "���϶Կ�" Or strInv3 = "���ϳ�Ʒ�˻�" Or strInv3 = "�����ϼ�����" Then
                'Ʒ��
                    If strInv3 = "���϶Կ�" Or strInv3 = "���ϳ�Ʒ�˻�" Then
                    
                        If Get_SqlserverCnt("SELECT ��Ʒ���� FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & strInv8 & "' and flag = '2' and ���= '" & strInv9 & "'") = 0 Then
                                    
                            MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"
                                    
                            Exit Sub
                                    
                        End If
                    
                        stritemname = Get_SqlStr("select distinct ��Ʒ���� from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '2' and  ��� = '" & strInv9 & "'")
                    
                    Else
                    
                        If Get_SqlserverCnt("SELECT ��Ʒ���� FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & strInv8 & "' and flag = '1' and ���= '" & strInv9 & "'") = 0 Then
                                    
                            MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"
                                    
                            Exit Sub
                                
                        End If
                    
                        stritemname = Get_SqlStr("select distinct ��Ʒ���� from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '1' and  ��� = '" & strInv9 & "'")
                    
                    End If
                    
                    .text = stritemname

                End If
                
                strInv10 = Trim$(.text)
        
                .Col = 13
                
                '������λ
                If strInv3 = "���϶Կ�" Or strInv3 = "���ϳ�Ʒ�˻�" Or strInv3 = "�����ϼ�����" Then
                
                    If strInv3 = "���϶Կ�" Or strInv3 = "���ϳ�Ʒ�˻�" Then
                    
                        strunit = Get_SqlStr("select distinct ������λ from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '2' and  ��� = '" & strInv9 & "'")

                    Else
                    
                        strunit = Get_SqlStr("select distinct ������λ from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '1' and  ��� = '" & strInv9 & "'")
                    
                    End If
                    
                    .text = strunit
                
                End If
                
                strInv11 = Trim$(.text)
        
                .Col = 14
                strInv12 = Trim$(.text)
        
                .Col = 15
                '�ܼ�
                
                If .text = "" Then
                
                    MsgBox "�������ܼ�", vbInformation, "��ʾ"
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
                
                AddSql2 ("insert into erptemp.dbo.ksexport (����,��������,�Ϻ�,���,��Ʊ��,��������,����,���ص���,�ֲ���,�ֲ����,Ʒ��,������λ,�ұ�,�ܼ�,����,AWB#,Ŀ�ĵ�,����,�˵�����,��ע,����ʱ��,�޸�״̬,�޸�ʱ��,ɾ��ʱ��,flag,id) SELECT ����,��������,�Ϻ�,���,��Ʊ��,��������,����,���ص���,�ֲ���,�ֲ����,Ʒ��,������λ,�ұ�,�ܼ�,����,AWB#,Ŀ�ĵ�,����,�˵�����,��ע,����ʱ��,'�޸�ǰ',�޸�ʱ��,ɾ��ʱ��,'2',id FROM erptemp.dbo.ksexport WHERE ���� = '" & strInv21 & "' and �������� = '" & strInv1 & "'  AND �Ϻ� =  '" & strInv2 & "' and id = '" & strInv20 & "' AND  flag = '0'")
                
                AddSql2 ("update erptemp.dbo.ksexport set ��� =  '" & strInv3 & "',���ص��� =  '" & strInv7 & "',�ֲ��� =  '" & strInv8 & "',�ֲ���� =  '" & strInv9 & "',Ʒ�� =  '" & strInv10 & "',������λ =  '" & strInv11 & "',�ұ� =  '" & strInv12 & "',�ܼ� =  '" & strInv13 & "',���� = '" & strInv14 & "',AWB# =  '" & strInv15 & "',Ŀ�ĵ� =  '" & strInv16 & "',���� =  '" & strInv17 & "',�˵����� =  '" & strInv18 & "',��ע =  '" & strInv19 & "',�޸�״̬ = '�޸ĺ�',�޸�ʱ�� = '" & strtime & "' where ���� = '" & strInv21 & "' and �������� = '" & strInv1 & "'  and �Ϻ�  = '" & strInv2 & "' and id = '" & strInv20 & "' and flag = '0'")
            
            End If
            
        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "��ѡ��Ҫ�޸ĵ���", vbInformation, "��ʾ"
            Exit Sub
            
        End If
    
    End With
    
    MsgBox "�޸ĳɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(5).Caption = "�޸�"
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

    If Toolbar1.Buttons(5).Caption <> "�ύ" Then

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
                If .text = "���϶Կ�" Or .text = "��Ʒ����" Then
                
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
                    
                    '��˰
                    .Col = 23
                    .Lock = False
                    
                    '��ֵ˰
                    .Col = 25
                    .Lock = False
                        
                End If

                
                .LockBackColor = vbYellow
                
                strsql = "select distinct �ֲ��� from erptemp.dbo.ksmanual where 1 = 1 and flag = '1'"

                If rs.State = 1 Then rs.Close
                rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                .Col = 10

                .CellType = CellTypeComboBox

              '  .TypeComboBoxList = ""

                rs.MoveFirst

                For j = 1 To rs.RecordCount

                    .TypeComboBoxList = .TypeComboBoxList & rs("�ֲ���")
                    rs.MoveNext
                Next
        
                rs.Clone
        
                Set rs = Nothing
    
            Next
        
        End With
    
        Toolbar1.Buttons(5).Caption = "�ύ"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    bFlag = False
    
    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "û������", vbInformation, "��ʾ"
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
                    MsgBox "���������", vbInformation, "��ʾ"
                    Exit Sub

                End If
                
                strInv4 = Trim$(.text)

                '����ǽ��϶Կ�&��Ʒ�����������������û�л��ʡ���˰����ֵ˰�ģ���Ϊ�Ǳ�˰��
                
                .Col = 7

                strInv5 = Trim$(.text)
                
                If Val(strInv5) <= 0 Then
                
                     MsgBox "������������С�ڵ���0", vbInformation, "��ʾ"
                     Exit Sub
                     
                End If
           
                .Col = 8

                '��׼die
'                If .Text = "" Then
'
'                    strssql = "select isnull(t8.qty,0) from  erpbase..tblCPurDataSub t1  " & " inner join  erpdata..tblSmainM2 t6 " & " on t1.���ϱ�� = t6.���ϱ��  " & " left join  (select m2.�Ϻ�,max(m1.QTECHDIEQTY) as qty  from erptemp..TBLTSVNPIPRODUCT m1,erpdata..TSVtblMRuleData m2 where 1=1 " & " and m1.QTECHPTNO2 = m2.����� group by m2.�Ϻ�) t8 " & " on t8.�Ϻ� = t6.�Ϻ� where 1=1 and t1.�ɹ������  = '" & strInv1 & "' and t6.�Ϻ� = '" & strInv2 & "' "
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
                '��die����
                
                strInv7 = Val(strInv5) * Val(strInv6)
                
                'strInv7 = Trim$(.Text)
        
                .Col = 10
                strInv8 = Trim$(.text)
        
                .Col = 11
                strInv9 = Trim$(.text)
                
                .Col = 12
                
                If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                
                    If Get_SqlserverCnt("SELECT ��Ʒ���� FROM erptemp.dbo.ksmanual WHERE �ֲ��� = '" & strInv8 & "' and flag = '1' and ���= '" & strInv9 & "'") = 0 Then
                                    
                        MsgBox "������ֲ�� + ��� �޶�Ӧ��Ʒ����������λ,��ȷ��", vbInformation, "��ʾ"

                        Exit Sub
                    
                    End If
                                
                    stritemname = Get_SqlStr("select distinct ��Ʒ���� from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '1' and  ��� = '" & strInv9 & "'")

                    .text = stritemname

                End If
                
                strInv10 = Trim$(.text)
        
                .Col = 13
                '��������
                If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                
                    strbaono1 = Get_SqlStr("select distinct isnull(�걨����,0) from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '1' and  ��� = '" & strInv9 & "'")
                
                    strbaono2 = Get_SqlStr("select isnull(sum(������),0) from erptemp.dbo.ksimport where flag = '0' and  �ɹ����� = ' " & strInv1 & "' and �Ϻ� = '" & strInv2 & "'")
                
                    strbaono3 = strbaono1 - strbaono2
                
                    If .text = "" Then
                    
                        MsgBox "�����뱨������", vbInformation, "��ʾ"
                        Exit Sub
                    
                    End If
                    
                    If Val(.text) <= 0 Then
                        
                        MsgBox "������������С�ڵ���0", vbInformation, "��ʾ"
                        Exit Sub
                    
                    End If
                
                    strInv11 = Format(Trim$(.text), "0.000")
                
                    If Val(strInv11) > Val(strbaono3) Then
                    
                        MsgBox "����ı���������������ķ�Χ,�걨����Ϊ" & strbaono1 & ",Ŀǰϵͳ��¼������ " & strbaono2 & "", vbInformation, "��ʾ"
                        
                        Exit Sub
                    
                    End If
                    
                Else
                         
                    If .text = "" Then
                    
                        MsgBox "�����뱨������", vbInformation, "��ʾ"
                        Exit Sub
                    
                    End If
                    
                    If Val(.text) <= 0 Then
                        
                        MsgBox "������������С�ڵ���0", vbInformation, "��ʾ"
                        Exit Sub
                    
                    End If
                    strInv11 = Format(Trim$(.text), "0.000")
                    
                End If
        
                .Col = 14
                
                If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                
                    strunit = Get_SqlStr("select distinct ������λ from erptemp.dbo.ksmanual where �ֲ��� = '" & strInv8 & "' and flag = '1' and  ��� = '" & strInv9 & "'")
                
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
                '����
                If .text = "" Then
                    
                    .text = 0
                    
                End If
                strInv18 = Format(Trim$(.text), "0.0000")
                
                .Col = 22
                '��˰��
                If .text = "" Then
                    
                    .text = 0
                    
                End If
                
                strInv19 = Format(Trim$(.text), "0.0000")
                
                .Col = 31
                strInv28 = Trim$(.text)
                
                strbao1 = Get_SqlStr("select distinct ���ؽ�� from erptemp.dbo.ksimport where �ɹ����� = '" & strInv1 & "' and �Ϻ� = '" & strInv2 & "' and id = '" & strInv28 & "' and flag = '0' ")
                
                strbao2 = Get_SqlStr("select distinct ���� from erptemp.dbo.ksimport where �ɹ����� = '" & strInv1 & "' and �Ϻ� = '" & strInv2 & "' and id = '" & strInv28 & "' and flag = '0' ")
                
                strbao3 = Get_SqlStr("select distinct ��˰�� from erptemp.dbo.ksimport where �ɹ����� = '" & strInv1 & "' and �Ϻ� = '" & strInv2 & "' and id = '" & strInv28 & "' and flag = '0' ")
                
                .Col = 23
                
                '��˰
                
                If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                    
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
                
                '��ֵ˰��
                If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                
                    .text = 0
                    strInv21 = Format(Trim$(.text), "0.0000")
                        
                Else
                
                    .text = 0.13
                    strInv21 = Format(Trim$(.text), "0.0000")
                    
                End If
                
                .Col = 25
                
                '��ֵ˰
                
                If strInv4 = "���϶Կ�" Or strInv4 = "��Ʒ����" Then
                
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
                
                     
                strNo1 = Get_SqlStr("SELECT isnull(SUM(a.��׼�ɹ�����),0) FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b WHERE a.�ɹ������ = '" & strInv1 & "' and a.���ϱ�� = b.���ϱ�� and b.�Ϻ� = '" & strInv2 & "' ")
                
                strNo2 = Get_SqlStr("SELECT isnull(SUM(��������),0) FROM erptemp.dbo.ksimport where �ɹ����� = '" & strInv1 & "' and �Ϻ� = '" & strInv2 & "' and id <> '" & strInv28 & "' and flag = '0'")
                
                strNo3 = Val(strNo1) - Val(strNo2)
                
                If Val(strInv5) > Val(strNo3) Then
                
                    MsgBox "�ñ��Ϻ�" & strInv2 & "��׼�ɹ�����: " & strNo1 & ",�Ѿ�ά������������" & strNo2 & ",�������ֻ��ά����" & strNo3 & "", vbInformation, "��ʾ"
                    Exit Sub

                End If
                
                strtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
                
                AddSql2 ("insert into erptemp.dbo.ksimport(����,�ɹ�����,�Ϻ�,�ͺ�,���,��������,��׼die,��die��,�ֲ���,���,Ʒ��,������,������λ,�볡����,��Ʊ��,����,�ұ�,�ɹ�����,���ؽ��,����,��˰��,��˰,��ֵ˰��,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id,����ʱ��,�޸�״̬,�޸�ʱ��,ɾ��ʱ��,flag) SELECT ����,�ɹ�����,�Ϻ�,�ͺ�,���,��������,��׼die,��die��,�ֲ���,���,Ʒ��,������,������λ,�볡����,��Ʊ��,����,�ұ�,�ɹ�����,���ؽ��,����,��˰��,��˰,��ֵ˰��,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id,����ʱ��,'�޸�ǰ',�޸�ʱ��,ɾ��ʱ��,'2' FROM erptemp.dbo.ksimport WHERE ���� =  '" & strInv29 & "' and �ɹ����� = '" & strInv1 & "'  AND �Ϻ� =  '" & strInv2 & "' AND id =  '" & strInv28 & "'  AND  flag = '0'")
                
                AddSql2 ("update erptemp.dbo.ksimport set �ͺ� = '" & strInv3 & "',��� = '" & strInv4 & "',�������� = '" & strInv5 & "',��׼die =  '" & strInv6 & "',��die�� =  '" & strInv7 & "',�ֲ��� = '" & strInv8 & "',��� = '" & strInv9 & "',Ʒ�� =  '" & strInv10 & "',������ =  '" & strInv11 & "', " & " ������λ  =  '" & strInv12 & "',�볡���� =  '" & strInv13 & "',��Ʊ�� =  '" & strInv14 & "',���� =  '" & strInv15 & "',�ұ� =  '" & strInv16 & "',���ؽ�� =  '" & strInv17 & "',���� =  '" & strInv18 & "',��˰�� =  '" & strInv19 & "',��˰ =  '" & strInv20 & "',��ֵ˰�� =  '" & strInv21 & "',��ֵ˰ =  '" & strInv22 & "',���ص��� =  '" & strInv23 & "',AWB#  =  '" & strInv24 & "',���� =  '" & strInv25 & "',�˵����� =  '" & strInv26 & "',��ע =  '" & strInv27 & "',�޸�״̬ = '�޸ĺ�',�޸�ʱ�� = '" & strtime & "' where ���� =  '" & strInv29 & "' and �ɹ����� = '" & strInv1 & "' and flag = '0' and �Ϻ�  = '" & strInv2 & "' and id =  '" & strInv28 & "' ")
            
            End If
            
        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "��ѡ��Ҫ�޸ĵ���", vbInformation, "��ʾ"
            Exit Sub
            
        End If
    
    End With
    
    MsgBox "�޸ĳɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(5).Caption = "�޸�"
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

    If Toolbar1.Buttons(7).Caption <> "�ύ" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
                
                .Col = E_FPS.E_gx
                .Lock = False
              
            Next
        
        End With
    
        Toolbar1.Buttons(7).Caption = "�ύ"
        Toolbar1.Buttons(7).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        Exit Sub

    End If

    bFlag = False

    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "û������", vbInformation, "��ʾ"
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
                
                    Case "������ϸ��"
                        .Col = F_fp.F_id
                        strInv2 = Trim$(.text)
                        AddSql2 ("update erptemp.dbo.ksimport set flag = '1',ɾ��ʱ��  = '" & strtime & "' where ���� = '" & strInv1 & "'  and id = '" & strInv2 & "' and flag = '0'")
                 
                    Case "������ϸ��(����)"
                        .Col = F_fp.F_id
                        strInv2 = Trim$(.text)
                        AddSql2 ("update erptemp.dbo.ksimport set flag = '1',ɾ��ʱ��  = '" & strtime & "' where ���� = '" & strInv1 & "'  and id = '" & strInv2 & "' and flag = '0'")
                     
                    Case "������ϸ��"
                        .Col = E_FPS.e_ID
                        strInv2 = Trim$(.text)
                        AddSql2 ("update erptemp.dbo.ksexport set flag = '1',ɾ��ʱ��  = '" & strtime & "' where ���� = '" & strInv1 & "' and id = '" & strInv2 & "' and flag = '0'")
                 
                    Case "������ϸ��(����)"
                
                        .Col = E_FPS.e_ID
                        strInv2 = Trim$(.text)
                        AddSql2 ("update erptemp.dbo.ksexport set flag = '1',ɾ��ʱ��  = '" & strtime & "' where ���� = '" & strInv1 & "' and id = '" & strInv2 & "' and flag = '0'")
                
                End Select
               
            End If

        Next

    End With

    If bFlag = False And j = 0 Then
        MsgBox "��ѡ��Ҫɾ������", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    MsgBox "ɾ���ɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(7).Caption = "ɾ��"
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
        '�趨������
       ' .Col = E_FPS0.E_CHOOSE   'ѡ��
       ' .CellType = CellTypeCheckBox
       ' .TypeHAlign = TypeVAlignCenter
       ' .TypeVAlign = TypeVAlignCenter
        
        '�趨�п�
        .ColWidth(-1) = 10
      '  .ColWidth(E_FPS0.E_CHOOSE) = 4
        .ColWidth(E_FPS0.e_ID) = 4
        .ColWidth(E_FPS0.E_CGDITEM) = 4
         .ColWidth(E_FPS0.E_PN) = 14
        .ColWidth(E_FPS0.E_SUPPLIERNAME) = 20
        .RowHeight(-1) = 10
        '�趨�Ƿ�����
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
'ͬ���ɹ���
strsql = "INSERT INTO erpbase..OPENPO_WAFER(�ɹ������,���ϱ��,PO����,��������,δ��������,�Ϻ�) SELECT a.�ɹ������, a.���ϱ��,sum(a.��׼�ɹ�����),0,0,c.F_101 FROM erpbase..tblCPurDataSub a, erpbase..tblCPurData b  ,AIS20141114094336..t_ICItem c WHERE a.�ɹ������=b.�ɹ������ AND a.���ϱ��=c.FNumber  and  a.�ɹ������ like 'c%' and a.���ϱ�� LIKE '01.01%' and  b.��˰���=1 AND a.�Ƿ����=0 group by a.�ɹ������, a.���ϱ�� ,c.F_101"
AddSql2 (strsql)


'ͬ���������,�������
'strSql = "UPDATE a SET a.�������=max(b.�볡����) from erpbase..OPENPO_WAFER a  left JOIN erptemp..ksimport b on a.�ɹ������=b.�ɹ����� and a.�Ϻ�=b.�Ϻ� where b.flag =0 "
''AddSql2 (strSql)
strsql = "UPDATE t1 SET t1.�������=t2.�볡����,t1.�������=isnull(t2.���������,0),t1.����ǰ�������=isnull(t2.����ǰ�������,0) from erpbase..OPENPO_WAFER t1 left join  " & _
" (SELECT ISNULL(�ɹ�����,'') AS �ɹ����� ,ISNULL(�Ϻ�,'') AS �Ϻ� ,sum( CASE WHEN DATEDIFF(day, �볡����, getdate())>5 THEN �������� ELSE 0 END )AS ����ǰ������� , " & _
" sum( ��������)AS ���������  ,max(isnull(�볡����,0)) AS �볡���� FROM erptemp..ksimport WHERE flag=0 GROUP  BY �ɹ�����,�Ϻ�) as t2  on t1.�ɹ������=t2.�ɹ����� and t1.�Ϻ�=t2.�Ϻ� "

AddSql2 (strsql)



'ͬ����������
strsql = "UPDATE a SET a.��������=isnull(t1.��������,0),a.δ��������=a.�������-isnull(t1.��������,0) from erpbase..OPENPO_WAFER a  left JOIN ( SELECT b.�ɹ������ ,b.���ϱ�� ,sum(b.��������) AS �������� FROM erpbase..tblToRecEntry b  GROUP BY b.�ɹ������,b.���ϱ��  ) AS t1 ON  a.�ɹ������ =t1.�ɹ������ AND a.���ϱ��=t1.���ϱ��"
AddSql2 (strsql)

'ͬ���������
strsql = "UPDATE t1 SET  t1.���������=Isnull(t2.���������,0),t1. ����98������= isnull(t2.���������98,0),t1.����52������=isnull(t2. ���������52,0) FROM     erpbase..OPENPO_WAFER  t1 left JOIN " & _
" (SELECT aa.�ɹ������,aa.���ϱ��,sum( bb.ʵ������*cc.������� )AS ���������,sum( CASE cc.�ֿ��� WHEN '52' then  bb.ʵ������*cc.������� ELSE 0 END )AS ���������52" & _
" ,sum( CASE cc.�ֿ��� WHEN '98' then  bb.ʵ������*cc.������� ELSE 0 END )AS ���������98  FROM erpbase..tblToRecEntry aa " & _
" LEFT JOIN  erpbase..TblToInSub bb ON aa.���������=bb.��������� AND aa.��¼��=bb.��¼�� " & _
" INNER JOIN erpbase..TblToInrec cc  ON bb.��ⵥ���=cc.��ⵥ���" & _
" GROUP BY aa.�ɹ������,aa.���ϱ��) AS t2 ON t1.�ɹ������=t2.�ɹ������ AND  t1.���ϱ��=t2.���ϱ��"
AddSql2 (strsql)



strsql = "SELECT row_number() over (order by t1.�ɹ������,t1.�Ϻ�) as ���,t1.* FROM ( " & _
" select distinct a.�ɹ������,a.�Ϻ�,a.PO���� as PO����,a.����ǰ�������,a.�������,a.PO����-a.������� as 'PO����-�������' , a.����ǰ�������-a.������� as '����ǰ�������-�������'  , a.�������� as  �ѵ�������, " & _
" a.�������- a.��������  as '�������- �ѵ�������' ,a.���������,a.����98������,a.����52������,a.��������-a.��������� as '�ѵ�������-���������' , " & _
" CASE WHEN isnull(a.���������,0)=0 AND isnull(a.���������,0)<a.�������  THEN 'δ��' WHEN isnull(a.���������,0)<a.������� THEN 'δ����' WHEN isnull(a.���������,0)>a.������� THEN '�볬'   ELSE '' END AS �Ƿ����, " & _
" e.�ͻ�����,d.FName as ��Ӧ������,convert(VARCHAR(10),b.�������,112) as PO����,a.�������  " & _
" from erpbase..OPENPO_WAFER a " & _
" inner join erpbase..tblcpurdata b on a.�ɹ������=b.�ɹ������ " & _
" inner join AIS20141114094336.dbo.t_Supplier  d on b.��Ӧ�̱��=d.FNumber " & _
" inner join dbo.tblXCustomer  e on d.FName=e.�ͻ����� " & _
" left join erptemp..tbltsvnpiproduct  f on a.�Ϻ�=f.MARKETLASTUPDATE_BY " & _
" where  a.�������-a.��������� <>0 "

If Trim(txtCust.text) <> "" Then
  strsql = strsql & " and   e.�ͻ�����='" & Trim(txtCust.text) & "'"
End If

If Trim(TxtPN.text) <> "" Then
  strsql = strsql & " and   a.�Ϻ�='" & Trim(TxtPN.text) & "'"
End If


If Optpatial.Value = True Then '����δά��
    strsql = strsql & " and isnull(a.�������,'')=''"
End If
If Optpatial2.Value = True Then '������ά��
    strsql = strsql & " and isnull(a.�������,'')<>''"
End If
strsql = strsql & " ) as t1"

Set rs = Get_SqlserveRs(strsql)
Call ListDataType(rs, fpS_Clear)
Err_Query:
If Err.number <> 0 Then
   MsgBox "QueryData��������,����ԭ��:" & Err.DESCRIPTION
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
    '�볡����  ��Ʊ��  ���ص���  AWB   ����  �ֲ���
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











