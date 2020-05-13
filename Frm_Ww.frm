VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_ww 
   Caption         =   "委外"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   18720
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
   ScaleHeight     =   10935
   ScaleWidth      =   18720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin TabDlg.SSTab SSTab2 
      Height          =   10815
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   19076
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Frm_Ww.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSTab1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Toolbar1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameC"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame FrameC 
         Caption         =   "查询"
         Height          =   2100
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   18135
         Begin VB.TextBox Txt_sqdh 
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
            Height          =   330
            Left            =   1080
            TabIndex        =   53
            Top             =   1680
            Width           =   2175
         End
         Begin VB.CheckBox Chk_NG 
            Caption         =   "不良品"
            Height          =   375
            Left            =   9120
            TabIndex        =   52
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtPN 
            Height          =   375
            Left            =   4800
            TabIndex        =   47
            Top             =   1200
            Width           =   2055
         End
         Begin VB.CheckBox Chk_Keepdata 
            Caption         =   "查询时保留已选择的Lot/机种/料号"
            Height          =   435
            Left            =   7080
            TabIndex        =   17
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox TxtLot 
            Height          =   405
            Left            =   4800
            TabIndex        =   16
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox TxtCustpn 
            Height          =   375
            Left            =   4800
            TabIndex        =   15
            Top             =   720
            Width           =   2055
         End
         Begin VB.ComboBox Cob_Shipto 
            BackColor       =   &H00FFC0FF&
            Height          =   315
            Left            =   1080
            TabIndex        =   13
            Top             =   720
            Width           =   2175
         End
         Begin VB.ComboBox Cob_kf_dest 
            BackColor       =   &H00FFC0FF&
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
            Left            =   8040
            TabIndex        =   5
            Top             =   840
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ComboBox Cob_kf_former 
            BackColor       =   &H00FFC0FF&
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
            Left            =   8040
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Left            =   1080
            TabIndex        =   3
            Top             =   1200
            Width           =   2175
         End
         Begin VB.ComboBox Cmbcust 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
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
            Left            =   1080
            TabIndex        =   2
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请单号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   54
            Top             =   1680
            Width           =   840
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
            Left            =   11640
            TabIndex        =   51
            Top             =   600
            Width           =   165
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
            Left            =   11640
            TabIndex        =   50
            Top             =   1320
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
            Left            =   10440
            TabIndex        =   49
            Top             =   1080
            Width           =   2100
         End
         Begin VB.Line Line2 
            X1              =   10320
            X2              =   13440
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label lblShippingQty 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "当前累计DIE数(DIE &PCS):"
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
            Left            =   10440
            TabIndex        =   48
            Top             =   360
            Width           =   2415
         End
         Begin VB.Shape Shape1 
            Height          =   1455
            Left            =   10320
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "料    号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   3480
            TabIndex        =   46
            Top             =   1320
            Width           =   840
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "客户机种"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   3480
            TabIndex        =   14
            Top             =   840
            Width           =   840
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "调 拨 人"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   840
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出货地址"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   840
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "目标库房"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   7080
            TabIndex        =   9
            Top             =   840
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "业务库房"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   7080
            TabIndex        =   8
            Top             =   360
            Visible         =   0   'False
            Width           =   1080
         End
         Begin MSForms.Label Label3 
            Height          =   210
            Left            =   120
            TabIndex        =   7
            Top             =   270
            Width           =   855
            ForeColor       =   0
            VariousPropertyBits=   276824091
            Caption         =   "客户代码"
            Size            =   "1508;370"
            FontName        =   "宋体"
            FontHeight      =   210
            FontCharSet     =   134
            FontPitchAndFamily=   34
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "工单(LOT)号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   3480
            TabIndex        =   6
            Top             =   240
            Width           =   1155
         End
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   870
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   18180
         _ExtentX        =   32068
         _ExtentY        =   1535
         ButtonWidth     =   1773
         ButtonHeight    =   1482
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   24
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "  查  询"
               Key             =   "Query"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "回货查询"
               Key             =   "Query_VT"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "  委外申请 "
               Key             =   "Request"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "回货申请"
               Key             =   "Backrequest"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "我的申请"
               Key             =   "ViewMyRequest"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               ImageIndex      =   8
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   " 撤销申请"
               Key             =   "CancerRequest"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Caption         =   "委外回货"
               Key             =   "BackRequest"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "A004"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "待调拨"
               Key             =   "WaitMove"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Caption         =   "委外调拨 "
               Key             =   "move"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Caption         =   "回货接收"
               Key             =   "A10"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "调拨"
               Key             =   "stockmove"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "委外撤销"
               Key             =   "CancerStockMove"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "  退   出  "
               Key             =   "A11"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         BorderStyle     =   1
         MousePointer    =   99
         MouseIcon       =   "Frm_Ww.frx":001C
         Begin MSComDlg.CommonDialog ee 
            Left            =   10200
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
            DialogTitle     =   "Excel导入"
            Filter          =   "*xls"
            InitDir         =   "D:\"
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   9600
            Top             =   0
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
                  Picture         =   "Frm_Ww.frx":017E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_Ww.frx":22B8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_Ww.frx":5142
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_Ww.frx":78F4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_Ww.frx":9A2E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_Ww.frx":C1E0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_Ww.frx":E992
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_Ww.frx":11A14
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_Ww.frx":141C6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_Ww.frx":144E0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_Ww.frx":151BA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_Ww.frx":1823C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm_Ww.frx":1A9EE
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   7095
         Left            =   240
         TabIndex        =   18
         Top             =   3360
         Width           =   18375
         _ExtentX        =   32411
         _ExtentY        =   12515
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "申请"
         TabPicture(0)   =   "Frm_Ww.frx":1B2C8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label9"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "fpS_Box"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "fpS_wafer"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "ChkAll2"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "调拨"
         TabPicture(1)   =   "Frm_Ww.frx":1B2E4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2"
         Tab(1).Control(1)=   "fpS_stockview"
         Tab(1).Control(2)=   "Command2"
         Tab(1).Control(3)=   "ChkAll"
         Tab(1).Control(4)=   "Option4"
         Tab(1).Control(5)=   "TxtRequestNo"
         Tab(1).Control(6)=   "Option3"
         Tab(1).Control(7)=   "Option2"
         Tab(1).Control(8)=   "Option1"
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "撤销"
         TabPicture(2)   =   "Frm_Ww.frx":1B300
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "TxtdbNo"
         Tab(2).Control(1)=   "Command1"
         Tab(2).Control(2)=   "Option5"
         Tab(2).Control(3)=   "Option6"
         Tab(2).Control(4)=   "fpS_Cancerview"
         Tab(2).Control(5)=   "Label1"
         Tab(2).Control(6)=   "Label4"
         Tab(2).ControlCount=   7
         TabCaption(3)   =   "回货"
         TabPicture(3)   =   "Frm_Ww.frx":1B31C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "fpS_vt_lot"
         Tab(3).Control(1)=   "Cmd_VT"
         Tab(3).Control(2)=   "DTP1"
         Tab(3).Control(3)=   "DTP2"
         Tab(3).Control(4)=   "Label8"
         Tab(3).Control(5)=   "Label7"
         Tab(3).Control(6)=   "Label5"
         Tab(3).ControlCount=   7
         TabCaption(4)   =   "Tab 4"
         TabPicture(4)   =   "Frm_Ww.frx":1B338
         Tab(4).ControlEnabled=   0   'False
         Tab(4).ControlCount=   0
         Begin FPSpreadADO.fpSpread fpS_vt_lot 
            Height          =   5055
            Left            =   -75000
            TabIndex        =   40
            Top             =   840
            Width           =   14895
            _Version        =   524288
            _ExtentX        =   26273
            _ExtentY        =   8916
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
            MaxCols         =   2
            MaxRows         =   0
            SpreadDesigner  =   "Frm_Ww.frx":1B354
            AppearanceStyle =   0
         End
         Begin VB.OptionButton Option1 
            Caption         =   "厂内调拨"
            Height          =   195
            Left            =   -64080
            TabIndex        =   32
            Top             =   450
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "委外"
            Height          =   195
            Left            =   -62880
            TabIndex        =   31
            Top             =   450
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            Caption         =   "回货"
            Height          =   195
            Left            =   -62040
            TabIndex        =   30
            Top             =   450
            Width           =   855
         End
         Begin VB.TextBox TxtRequestNo 
            Height          =   405
            Left            =   -73200
            TabIndex        =   29
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            Caption         =   "所有"
            Height          =   195
            Left            =   -61200
            TabIndex        =   28
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox ChkAll 
            Caption         =   "全选"
            Height          =   375
            Left            =   -75000
            TabIndex        =   27
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox ChkAll2 
            Caption         =   "全选"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox TxtdbNo 
            Height          =   375
            Left            =   -73920
            TabIndex        =   24
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton Command1 
            Caption         =   "查询"
            Height          =   375
            Left            =   -72000
            TabIndex        =   23
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "查询"
            Height          =   375
            Left            =   -71520
            TabIndex        =   22
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            Caption         =   "撤销申请单"
            Height          =   255
            Left            =   -74760
            TabIndex        =   21
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Option6 
            Caption         =   "撤销调拨单"
            Height          =   255
            Left            =   -73440
            TabIndex        =   20
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton Cmd_VT 
            Caption         =   "查询回货资料"
            Height          =   375
            Left            =   -68400
            TabIndex        =   19
            Top             =   360
            Width           =   2055
         End
         Begin FPSpreadADO.fpSpread fpS_Cancerview 
            Height          =   4695
            Left            =   -75000
            TabIndex        =   25
            Top             =   1440
            Width           =   14895
            _Version        =   524288
            _ExtentX        =   26273
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
            SpreadDesigner  =   "Frm_Ww.frx":1B74C
         End
         Begin FPSpreadADO.fpSpread fpS_stockview 
            Height          =   4935
            Left            =   -74880
            TabIndex        =   33
            Top             =   960
            Width           =   14775
            _Version        =   524288
            _ExtentX        =   26061
            _ExtentY        =   8705
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
            SpreadDesigner  =   "Frm_Ww.frx":1BB36
            AppearanceStyle =   0
         End
         Begin FPSpreadADO.fpSpread fpS_wafer 
            Height          =   5895
            Left            =   7080
            TabIndex        =   34
            Top             =   720
            Width           =   11175
            _Version        =   524288
            _ExtentX        =   19711
            _ExtentY        =   10398
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
            SpreadDesigner  =   "Frm_Ww.frx":1BF2E
            AppearanceStyle =   0
         End
         Begin FPSpreadADO.fpSpread fpS_Box 
            Height          =   5775
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   6855
            _Version        =   524288
            _ExtentX        =   12091
            _ExtentY        =   10186
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
            SpreadDesigner  =   "Frm_Ww.frx":1C326
            AppearanceStyle =   0
         End
         Begin MSComCtl2.DTPicker DTP1 
            Height          =   375
            Left            =   -72600
            TabIndex        =   42
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   102367233
            CurrentDate     =   41424
         End
         Begin MSComCtl2.DTPicker DTP2 
            Height          =   375
            Left            =   -70080
            TabIndex        =   43
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   102367233
            CurrentDate     =   41424
         End
         Begin VB.Label Label9 
            Caption         =   "标记灰色的表示已委外且已回货,不可再次委外"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1080
            TabIndex        =   55
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结束时间："
            Height          =   195
            Left            =   -71040
            TabIndex        =   45
            Top             =   480
            Width           =   900
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "开始时间："
            Height          =   195
            Left            =   -73440
            TabIndex        =   44
            Top             =   480
            Width           =   900
         End
         Begin VB.Label Label6 
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   960
            TabIndex        =   41
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "申请单号"
            Height          =   255
            Left            =   -74040
            TabIndex        =   39
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "调拨单号"
            Height          =   255
            Left            =   -74880
            TabIndex        =   38
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "（只能查询自己申请或调拨的单号）"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   -72120
            TabIndex        =   37
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label Label5 
            Caption         =   "请先输入客户代码"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   -75000
            TabIndex        =   36
            Top             =   480
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "Frm_ww"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Enum E_BOX

    E_CHOOSE = 1

    E_CUSTCODE     '客户代码
    E_CUSTPN       'KEHUJIZHONG
    E_qtechPTNo    'changneijizhong
    E_LOTID
    E_BOXID
    E_Matcode    '物料编号
    E_partno '料号
    E_Matspec    '规格
    E_Mattype    '型号
    E_UNIT    '单位
    E_Passqty  '合格数
    E_Ngqty1            '来料不良数
    E_Ngqty2             '制程不良数
    e_ID      '序号
    E_StockID      '仓库代码
    E_END

End Enum

Enum E_WAFERID

    E_CHOOSE = 1
    E_LOTID
    E_BigBoxID
    E_BOXID
    E_WAFERID
    E_PN
    E_QTY  '数量
    E_Passqty  '合格数
    E_Ngqty1            '来料不良数
    E_Ngqty2             '制程不良数
    e_ID
    E_END

End Enum

Enum E_StockView


    E_CHOOSE = 1
    e_order_num
    e_Item
    E_CUSTCODE
    E_CUSTPN
  '  E_qtechPTNo
    E_KF_FORMER
    E_KF_DEST
    E_LOT
    E_REMARK1
    e_Qbox
    E_Wafer    '
    E_GOOD_DIE '合格数
    E_BAD1_DIE '来料不良数
    E_BAD2_DIE '制程不良数
    e_ID
    E_END

End Enum



Enum E_VT
    E_CHOOSE = 1
    E_Result
  '  E_ShipDate
  '  E_CUSTPN
    E_LOTID
    E_Wafer    '
    E_GOOD_DIE '合格数
    E_BAD_DIE '来料不良数
   ' E_CUSTCODE
  '  E_ID
    e_waferlist_vt
    e_waferlist_stock
    E_BOX_STOCK
    E_Passqty_STOCK
    E_Ngqty1_STOCK
    E_Ngqty2_STOCK
    E_KF_FORMER
    E_ID_STOCK
    E_END

End Enum




Dim adorst2     As New ADODB.Recordset
Dim strXH As String
Dim strgdh As String
Dim strLCK As String
Dim strlps As String
Dim strbls As String
Dim strzcbls As String
Dim strmbkf As String

Dim strid As Long
Dim strxh_big As String
Dim strnewbox As String
Dim newboxid As String
Dim strWholeName As String
Dim strdepartment As String


Private Sub Check1_Click()

End Sub

Private Sub ChkAll2_Click()
    Dim i As Integer
    
    With fpS_Box
        If ChkAll2.Value = 1 Then
            For i = 1 To .MaxRows
      
                .Row = i
                .Col = 1
                .text = 1
                 Call Fps_Box_Click(0, i)
                
            Next i
        ElseIf ChkAll2.Value = 0 Then
        
            For i = 1 To .MaxRows
                  
                .Row = i
                .Col = 1
                .text = 0
                Call Fps_Box_Click(1, i)
                
            Next i
            
        End If
    End With
End Sub

Private Sub Cmbcust_DropDown()
Dim i As Integer
    Set adorst2 = New ADODB.Recordset
    Set adorst2.ActiveConnection = INIadoCon2
    adorst2.Source = "select distinct 客户代码  from tblXCustomer "
    adorst2.Open , , , , adCmdText
    Cmbcust.Clear
    If adorst2.RecordCount > 0 Then
      For i = 1 To adorst2.RecordCount
        Cmbcust.AddItem Trim(adorst2("客户代码"))
        adorst2.MoveNext
      Next i
    Else
       Cmbcust.text = ""
       Exit Sub
    End If
End Sub





Private Sub Cmd_VT_Click()
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer

    If Trim(Cmbcust.text) = "" Then
        MsgBox "请先选择客户代码", vbInformation, "提示"
        Exit Sub
    End If
    
    If SMR.State = adStateOpen Then SMR.Close

    
    strSql = "SELECT 0 AS 选择,SHIPDATE ,DELIVERYNO,CUSTLOT, GOODDIEQTY,NGDIEQTY,TTL,BATCH,REMARK,CUSTOMERSHORTNAME from TSV_VT_History WHERE CREATED_DATE<'" & Format(DTP2.Value + 1, "yyyy/mm/dd") & "' and  CREATED_DATE>'" & Format(DTP1.Value, "yyyy/mm/dd") & "' and CUSTOMERSHORTNAME='" & Trim(UCase(Cmbcust.text)) & "'"
    SMR.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        With fpS_vt_lot
           .MaxRows = 0
           Set .DataSource = SMR
          
        End With
        
    Else
    
        With fpS_vt_lot
           .MaxRows = 0

          
        End With

    End If
   
End Sub

Private Sub Cob_Shipto_DropDown()
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Source = "SELECT DISTINCT SHIP_TO  FROM erptemp..customer_information a WHERE a.CUSTOMER = '" & Trim(Cmbcust.text) & "'"
    SMR.Open , INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            Cob_Shipto.AddItem (Trim(SMR("SHIP_TO")))
            SMR.MoveNext
        Next
    End If
End Sub

Private Sub Cob_kf_dest_Click()
    Dim SMR        As New ADODB.Recordset
    Dim i As Integer
    
    Dim Kf_former As String
    Dim Kf_dest As String
    
    Kf_former = ""
    Kf_dest = ""
    
    If Trim(Cob_kf_former.text) <> "" Then
        Kf_former = Left(Trim(Cob_kf_former.text), InStr(Trim(Cob_kf_former.text), " ") - 1)
    End If
    If Trim(Cob_kf_dest.text) <> "" Then
        Kf_dest = Left(Trim(Cob_kf_dest.text), InStr(Trim(Cob_kf_dest.text), " ") - 1)
    End If
    If Kf_dest = "72" Then
        
        Toolbar1.Buttons("Request").Caption = "委外申请"
        Toolbar1.Buttons("Request").Enabled = True
        Toolbar1.Buttons("Backrequest").Enabled = False
    Else
        If Kf_former <> "72" Then
            Toolbar1.Buttons("Request").Caption = "调拨申请"
            Toolbar1.Buttons("Request").Enabled = True
            Toolbar1.Buttons("Backrequest").Enabled = False
        Else
            Toolbar1.Buttons("Request").Enabled = False
            Toolbar1.Buttons("Backrequest").Enabled = True
        End If
    End If
       
End Sub



Private Sub Cob_kf_former_Click()
    Dim Kf_former As String
    Dim Kf_dest As String
    
    Kf_former = ""
    Kf_dest = ""
    
    If Trim(Cob_kf_former.text) <> "" Then
        Kf_former = Left(Trim(Cob_kf_former.text), InStr(Trim(Cob_kf_former.text), " ") - 1)
    End If
    If Trim(Cob_kf_dest.text) <> "" Then
        Kf_dest = Left(Trim(Cob_kf_dest.text), InStr(Trim(Cob_kf_dest.text), " ") - 1)
    End If
    
    If Kf_dest = "72" Then
        Toolbar1.Buttons("Request").Caption = "委外申请"
        Toolbar1.Buttons("Request").Enabled = True
        Toolbar1.Buttons("Backrequest").Enabled = False
    Else
        If Kf_former <> "72" Then
            Toolbar1.Buttons("Request").Caption = "调拨申请"
            Toolbar1.Buttons("Request").Enabled = True
            Toolbar1.Buttons("Backrequest").Enabled = False
        Else
            Toolbar1.Buttons("Request").Enabled = False
            Toolbar1.Buttons("Backrequest").Enabled = True
        End If
    End If
           
    
End Sub

Private Sub Cob_kf_former_DropDown()
Dim intNext As Integer
Dim adoRstStocEntry As New ADODB.Recordset
   Set adoRstStocEntry = New ADODB.Recordset
   adoRstStocEntry.ActiveConnection = INIadoCon2
   adoRstStocEntry.Source = "select 库房代码,库房名称 from erpbase..tblstock  where 仓库属性='成品仓'  order by 库房代码"
   adoRstStocEntry.Open , , , , adCmdText
   If adoRstStocEntry.RecordCount > 0 Then
      Cob_kf_former.Clear
      adoRstStocEntry.MoveFirst
      For intNext = 1 To adoRstStocEntry.RecordCount
          Cob_kf_former.AddItem Trim(adoRstStocEntry("库房代码")) & Space(1) & Trim(adoRstStocEntry("库房名称"))
          adoRstStocEntry.MoveNext
      Next intNext
   Else
   End If
  adoRstStocEntry.Close
  Set adoRstStocEntry = Nothing
  
End Sub







Private Sub Cob_kf_dest_DropDown()
Dim adorst11 As New ADODB.Recordset
Dim intSubN As Integer
 Set adorst11 = New ADODB.Recordset
  adorst11.ActiveConnection = INIadoCon2
  adorst11.Source = "SELECT 库房代码+' '+库房名称 仓库名称 FROM erpbase..tblstock WHERE 仓库属性='成品仓'"
  adorst11.Open , , , , adCmdText
  Cob_kf_dest.Clear
  If adorst11.RecordCount > 0 Then
    adorst11.MoveFirst
    For intSubN = 1 To adorst11.RecordCount
      Cob_kf_dest.AddItem Trim(adorst11.Fields(0))
    adorst11.MoveNext
    Next intSubN
  End If
  adorst11.Close
  Set adorst11 = Nothing
End Sub



Private Sub Command1_Click()
If Option5.Value = False And Option6.Value = False Then
    MsgBox "请选择撤销方式", vbInformation, "提示"
    Exit Sub
End If


ListCancerView

End Sub

Private Sub Command2_Click()
ListStockView
Toolbar1.Buttons("stockmove").Enabled = True
End Sub


Private Sub Form_Load()
   

    inictrl
    SSTab1.Tab = 0
    Call SSTab1_Click(1)
    SSTab1.TabVisible(4) = False

    strWholeName = gUserName
    Select Case gUserName

    Case "17363"
        strWholeName = gUserName & " 秦智"
        strdepartment = "07"
    Case "19809"
        strWholeName = gUserName & " 申无疆"
        strdepartment = "07"
    Case "19536"
        strWholeName = gUserName & " 李月"
        strdepartment = "07"
    Case "07952"
        strWholeName = gUserName & " 韩海燕"
        strdepartment = "07"
    Case "10222"
        strWholeName = gUserName & " 薛振江"
        strdepartment = "07"
    Case "12825"
        strWholeName = gUserName & " 杨静锋"
        strdepartment = "07"

    End Select
    Text1.text = Trim(strWholeName)
    'Cob_kf_former.Text ="07 保税成品仓"
    'Cob_kf_dest.Text="72 WLA委外仓"

End Sub






Private Sub Fps_Box_Click(ByVal Col As Long, ByVal Row As Long)

Dim i       As Long
Dim j       As Integer
Dim strid As String
Dim strLotID As String
Dim strLotID_temp As String
Dim strBoxID As String
Dim strStockID As String
Dim strchoose As String
If Col <> 1 Then Exit Sub

With fpS_Box
    
    .Row = Row
    .Col = E_BOX.E_LOTID
    strLotID = Trim(.text)
    
    .Row = Row
    .Col = E_BOX.E_CHOOSE
    If .Value = 0 Then
        If .BackColor = 8421504 Then
           MsgBox "此笔已回货，不能委外", vbInformation, "提示"

           Exit Sub
        End If
        If UCase(Trim(Cmbcust.text)) = "GC" Then
            '同一个lot分存于两个大箱，一个单号只能出一个大箱号
            For i = 1 To .MaxRows
                If i <> Row Then
                    .Row = i
                    .Col = E_BOX.E_CHOOSE
                    strchoose = Trim(.text)
                    .Row = i
                    .Col = E_BOX.E_LOTID
                    strLotID_temp = Trim(.text)
                    If strchoose = "1" And strLotID_temp = strLotID Then
                         MsgBox "同一个lot分存于两个大箱，一个单号只能出一个大箱号", vbInformation, "提示"
                         Exit Sub
                    End If
                End If
            
            Next
        End If
    End If
    
    .Row = Row
    .Col = E_BOX.E_CHOOSE
    .Value = Abs(Val(.Value) - 1)
    
    If .Value = 1 Then
        .Col = -1
        .ForeColor = &HFF8080
        .Col = E_BOX.E_LOTID
        strLotID = "$" & Trim$(.text)

        .Col = E_BOX.E_BOXID
        strBoxID = Trim$(.text) & "★"

        .Col = E_BOX.E_StockID
        strStockID = Trim$(.text)



        
        
       If Get_SqlserverCnt(" SELECT * FROM erptemp..tblstockdbsub_temp a,  erptemp..tblstockdb_temp b where a.remark1='" & Replace(strBoxID, "★", "") & "' and b.flag=1 and a.ORDER_NUM=b.ORDER_NUM and a.ITEM=B.ITEM") > 0 Then
           MsgBox "箱号" & strBoxID & "已提过委外申请，不可重复申请", vbInformation, "提示"
           .Col = E_BOX.E_CHOOSE
           .Value = 0
       Else
           Call SearchWaferID_ByBoxID(strStockID, strLotID, strBoxID, 1)
        End If
       
    ElseIf .Value = 0 Then
        .Col = -1
        .ForeColor = vbBlack
        .Col = E_BOX.E_LOTID
        strLotID = Trim$(.text)

        .Col = E_BOX.E_BOXID
        strBoxID = Trim$(.text)
        
        .Col = E_BOX.E_StockID
        strStockID = Trim$(.text)
               
        
        Call SearchWaferID_ByBoxID(strStockID, strLotID, strBoxID, 2)

    End If

    
End With
End Sub



Private Sub fpS_stockview_Click(ByVal Col As Long, ByVal Row As Long)
Dim i           As Long
Dim j           As Integer
Dim strnoTmp      As String
Dim stritemTmp      As String
Dim strno_select      As String
Dim stritem_select      As String

    '点击把选择的单号都选上
    If Row < 1 Then Exit Sub
    If Col <> 1 Then Exit Sub
    With fpS_stockview

        .Col = 1
        .Row = Row
        .Value = Abs(Val(.Value) - 1)
'        strDJBH = ""
        If Val(.Value) = 1 Then
            '将所有一样的单号+序号的选择上
            .Col = 2
            .Row = Row
            strno_select = Trim$(.text)
            .Col = 3
            .Row = Row
            stritem_select = Trim$(.text)
            For i = 1 To .MaxRows
                .Col = 2
                .Row = i
                strnoTmp = Trim$(.text)
                .Col = 3
                .Row = i
                stritemTmp = Trim$(.text)
                
                
                If strno_select = strnoTmp And stritem_select = stritemTmp Then
                    .Col = 1
                    .Value = 1
                    .Col = -1
                    .ForeColor = &HFF8080
                End If
            Next
        Else
            '将所有一样的单号+序号的选择上
            .Col = 2
            .Row = Row
            strno_select = Trim$(.text)
            .Col = 3
            .Row = Row
            stritem_select = Trim$(.text)
            For i = 1 To .MaxRows
            .Col = 2
            .Row = i
            strnoTmp = Trim$(.text)
            .Col = 3
            .Row = i
            stritemTmp = Trim$(.text)
            
            
            If strno_select = strnoTmp And stritem_select = stritemTmp Then
                .Col = 1
                .Value = 0
                .Col = -1
                .ForeColor = vbBlack
            End If
            Next


        End If
        
    End With
End Sub




Private Sub Option1_Click()
    ListStockView
End Sub

Private Sub Option2_Click()
    ListStockView
End Sub

Private Sub Option3_Click()
    ListStockView
End Sub

Private Sub Option4_Click()
    ListStockView
End Sub



Private Sub Option5_Click()
Label1.Caption = "申请单号"
End Sub




Private Sub Option6_Click()
Label1.Caption = "调拨单号"
End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab

    Case 0
           Toolbar1.Buttons("Query").Enabled = True
           Toolbar1.Buttons("Request").Enabled = True
           Toolbar1.Buttons("Query_VT").Enabled = True
           Toolbar1.Buttons("Backrequest").Enabled = True
           Toolbar1.Buttons("ViewMyRequest").Enabled = False
           Toolbar1.Buttons("CancerRequest").Enabled = False
           Toolbar1.Buttons("WaitMove").Enabled = False
           Toolbar1.Buttons("stockmove").Enabled = False
           Toolbar1.Buttons("CancerStockMove").Enabled = False
    Case 1
           Toolbar1.Buttons("Query").Enabled = False
           Toolbar1.Buttons("Query_VT").Enabled = False
           Toolbar1.Buttons("Request").Enabled = False
           
           Toolbar1.Buttons("Backrequest").Enabled = False
           Toolbar1.Buttons("ViewMyRequest").Enabled = True
           Toolbar1.Buttons("CancerRequest").Enabled = False
           Toolbar1.Buttons("WaitMove").Enabled = True
           Toolbar1.Buttons("stockmove").Enabled = False
           Toolbar1.Buttons("CancerStockMove").Enabled = False
    Case 2
           Toolbar1.Buttons("Query").Enabled = False
           Toolbar1.Buttons("Request").Enabled = False
           Toolbar1.Buttons("Query_VT").Enabled = False
           Toolbar1.Buttons("Backrequest").Enabled = False
           Toolbar1.Buttons("ViewMyRequest").Enabled = False
           Toolbar1.Buttons("CancerRequest").Enabled = False
           Toolbar1.Buttons("WaitMove").Enabled = False
           Toolbar1.Buttons("stockmove").Enabled = False
           Toolbar1.Buttons("CancerStockMove").Enabled = False
     Case 3
           Toolbar1.Buttons("Query").Enabled = False
           Toolbar1.Buttons("Request").Enabled = False
           Toolbar1.Buttons("Query_VT").Enabled = False
           Toolbar1.Buttons("Backrequest").Enabled = False
           Toolbar1.Buttons("ViewMyRequest").Enabled = False
           Toolbar1.Buttons("CancerRequest").Enabled = False
           Toolbar1.Buttons("WaitMove").Enabled = False
           Toolbar1.Buttons("stockmove").Enabled = False
           Toolbar1.Buttons("CancerStockMove").Enabled = False
           
           DTP1.Value = Now

           DTP2.Value = Now
           
End Select



End Sub





Sub ListStockView()
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    If SMR.State = adStateOpen Then SMR.Close
    strSql = "select 0 as '选择',a.ORDER_NUM as 申请单号,a.ITEM as 序号,d.CUSTOMERSHORTNAME as 客户代码,d.MPN_DESC as 客户机种,a.FORMER as 原仓库,a.DESTINATION as 目标仓库,b.LOT as 工单号,b.REMARK1 as 大箱号,b.QBOX as 小箱号, " & _
             " b.WAFER as 流程卡编号,b.GOOD_DIE as 合格数,b.BAD1_DIE as 来料不良数 ,b.BAD2_DIE as 制程不良数,b.id as id from erptemp..tblstockdb_temp a " & _
             " left join erptemp..tblstockdbsub_temp b  on a.ORDER_NUM=b.ORDER_NUM and a.item=b.item " & _
             " left join erpbase..tblmappingdata c on c.SUBSTRATEID=b.WAFER and c.LOTID=b.LOT " & _
             " left join erpbase..tblcustomeroi d on convert(varchar(20)  ,d.id)=c.FILENAME and c.LOTID= d.SOURCE_BATCH_ID " & _
             " where a.flag=1 "
             
    If gUserName <> "07885" Then strSql = strSql & " and a.APPLICANT<>'" & strWholeName & " '"
    If Trim(TxtRequestNo.text) <> "" Then
        strSql = strSql & " and a.ORDER_NUM='" & Trim(TxtRequestNo.text) & "'"
    End If

    If Option1.Value = True Then
        strSql = strSql & " and a.FORMER<>'72'  and a.DESTINATION<>'72'  "
    ElseIf Option2.Value = True Then
        strSql = strSql & " and a.FORMER<>'72'  and a.DESTINATION='72'  "
    ElseIf Option3.Value = True Then
        strSql = strSql & " and a.FORMER='72'  and a.DESTINATION<>'72'  "
    End If
    strSql = strSql & "order by a.ORDER_NUM,a.FORMER,a.DESTINATION,a.item,b.WAFER"
        
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    With fpS_stockview
       .MaxRows = 0
       Set .DataSource = SMR
    End With

End Sub

Sub ListCancerView()
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    If SMR.State = adStateOpen Then SMR.Close
    If Label1.Caption = "调拨单号" Then
       If Trim(TxtdbNo.text) = "" Then
            MsgBox "请输入调拨单号", vbInformation, "提示"
            Exit Sub
       End If
       
    
        strSql = "select 1 as '选择',a.调拨编号,a.序号,d.CUSTOMERSHORTNAME as 客户代码,d.MPN_DESC as 客户机种,a.原仓库,a.目标仓库,b.工单号,'' as remark,b.箱号, " & _
                 " b.流程卡编号,b.合格数,b.来料不良数 ,b.制程不良数,b.id as id from erpdata..tblstockdb a " & _
                 " left join erpdata..tblstockdbsub b  on a.调拨编号=b.调拨编号 and a.序号=b.序号 " & _
                 " left join erpbase..tblmappingdata c on c.SUBSTRATEID=b.流程卡编号 and c.LOTID=b.工单号 " & _
                 " left join erpbase..tblcustomeroi d on convert(varchar(20)  ,d.id)=c.FILENAME and c.LOTID= d.SOURCE_BATCH_ID  "

        If Trim(TxtdbNo.text) <> "" Then
            strSql = strSql & " where  a.调拨编号='" & Trim(TxtdbNo.text) & "'"
            If gUserName <> "07885" Then strSql = strSql & " and  a.申请人员='" & strWholeName & " '"
        Else
            If gUserName <> "07885" Then strSql = strSql & " where a.申请人员='" & strWholeName & " '"
        End If
        Toolbar1.Buttons("CancerStockMove").Enabled = True
    ElseIf Label1.Caption = "申请单号" Then
        strSql = "select 1 as '选择',a.ORDER_NUM as 申请单号,a.ITEM as 序号,d.CUSTOMERSHORTNAME as 客户代码,d.MPN_DESC as 客户机种,a.FORMER as 原仓库,a.DESTINATION as 目标仓库,b.LOT as 工单号,'' as remark,b.QBOX as 小箱号, " & _
                 " b.WAFER as 流程卡编号,b.GOOD_DIE as 合格数,b.BAD1_DIE as 来料不良数 ,b.BAD2_DIE as 制程不良数,b.id as id from erptemp..tblstockdb_temp a " & _
                 " left join erptemp..tblstockdbsub_temp b  on a.ORDER_NUM=b.ORDER_NUM and a.item=b.item " & _
                 " left join erpbase..tblmappingdata c on c.SUBSTRATEID=b.WAFER and c.LOTID=b.LOT " & _
                 " left join erpbase..tblcustomeroi d on convert(varchar(20)  ,d.id)=c.FILENAME and c.LOTID= d.SOURCE_BATCH_ID " & _
                 " where a.flag=1 and APPLICANT='" & strWholeName & " '"
        If Trim(TxtdbNo.text) <> "" Then
            strSql = strSql & " and a.ORDER_NUM='" & Trim(TxtdbNo.text) & "'"
        End If
        Toolbar1.Buttons("CancerRequest").Enabled = True
    End If

    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        With fpS_Cancerview
           .MaxRows = 0
           Set .DataSource = SMR
           
        End With

    End If

End Sub





Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
strXH = ""
strgdh = ""
strLCK = ""
strlps = ""
strbls = ""
strzcbls = ""
strnewbox = ""
strmbkf = ""

Dim i As Integer
Dim WAFER_QTY As Integer
Dim boxtemp As String
Dim boxtemp1 As String

boxtemp = ""
boxtemp1 = ""
newboxid = ""
WAFER_QTY = 1
i = 1

   Select Case Button.Key
  
    Case "Request"  '申请
         If Trim(Cmbcust.text) = "" Then
             MsgBox "请选择客户代码!", vbInformation, "提示"
             Exit Sub
         End If
         If Trim(Cob_Shipto.text) = "" Then
             MsgBox "请选择出货地址!", vbInformation, "提示"
             Exit Sub
         End If
        Toolbar1.Buttons("Request").Enabled = False
        CreateApplication ("WW")
        Toolbar1.Buttons("Request").Enabled = True
    Case "Backrequest"  '回货申请
          
        CreateApplication ("VT")
        
    Case "ViewMyRequest"  '
         Option1.Visible = False
         Option2.Visible = False
         Option3.Visible = False
         Option4.Visible = False
        
        ListMyRequest
        
         
    Case "CancerRequest"  '撤销申请
        Toolbar1.Buttons("CancerRequest").Enabled = False
        CancerRequest
           
    Case "WaitMove"  '待调拨
         Option1.Visible = True
         Option2.Visible = True
         Option3.Visible = True
         Option4.Visible = True
         
         ListStockView
         
    Case "stockmove"   '调拨
        Toolbar1.Buttons("stockmove").Enabled = False
        stockmove
        
    Case "Query" '查询
    
         If Trim(Cmbcust.text) = "" Then
             MsgBox "请选择客户代码!", vbInformation, "提示"
             Exit Sub
         End If
        
         Call ListView1Data("WW")
    Case "Query_VT" '查询
    
         Cob_kf_former.text = "72 WLA委外仓"

         Call ListView1Data("VT")
         
     Case "CancerStockMove"   '调拨撤销
        Toolbar1.Buttons("CancerStockMove").Enabled = False
        cancerstockmove
         

     Case "A11"
     
         Unload Me
            
  End Select
End Sub

Sub ListMyRequest()

   Dim SMR        As New ADODB.Recordset
   Dim strSql     As String

             
    SSTab1.Tab = 1

    If SMR.State = adStateOpen Then SMR.Close
    strSql = "select 0 as '选择',a.ORDER_NUM as 申请单号,a.ITEM as 序号,d.CUSTOMERSHORTNAME as 客户代码,d.MPN_DESC as 客户机种,a.FORMER as 原仓库,a.DESTINATION as 目标仓库,b.LOT as 工单号,b.REMARK1 as 大箱号,b.QBOX as 小箱号, " & _
             " b.WAFER as 流程卡编号,b.GOOD_DIE as 合格数,b.BAD1_DIE as 来料不良数 ,b.BAD2_DIE as 制程不良数,b.id as id from erptemp..tblstockdb_temp a " & _
             " left join erptemp..tblstockdbsub_temp b  on a.ORDER_NUM=b.ORDER_NUM and a.item=b.item " & _
             " left join erpbase..tblmappingdata c on c.SUBSTRATEID=b.WAFER and c.LOTID=b.LOT " & _
             " left join erpbase..tblcustomeroi d on convert(varchar(20)  ,d.id)=c.FILENAME and c.LOTID= d.SOURCE_BATCH_ID " & _
             " where a.flag=1 and APPLICANT='" & strWholeName & " '"
    If Trim(TxtRequestNo.text) <> "" Then
        strSql = strSql & " and a.ORDER_NUM='" & Trim(TxtRequestNo.text) & "'"
    End If

        
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    With fpS_stockview
       .MaxRows = 0
       Set .DataSource = SMR
    End With


End Sub

Sub CancerRequest()
 Dim strrequestno As String
 Dim strrequestitem As String
 Dim strSql As String
 Dim SumCount As Integer
 Dim i As Integer
SumCount = 0
If Label1.Caption <> "申请单号" Then
    Exit Sub
End If
If TxtdbNo.text = "" Then
   MsgBox "请按申请单号查询", vbInformation, "提示"
   Exit Sub
End If

 With fpS_Cancerview
    If .MaxRows <= 0 Then
        MsgBox "请先查询", vbInformation, "提示"
        Exit Sub
    End If
    For i = 1 To .MaxRows
   
       .Row = i
       .Col = E_StockView.E_CHOOSE
       
       If .text = 1 Then
           SumCount = SumCount + 1
       End If
    Next i
    If MsgBox("你确认要取消申请单" & TxtRequestNo & ",共" & SumCount & "笔记录吗?", vbOKCancel, "提示") = vbCancel Then
        Exit Sub
    End If
    SumCount = 0
    For i = 1 To .MaxRows
   
       .Row = i
       .Col = E_StockView.E_CHOOSE
       
       If .text = 1 Then
           .Row = i
           .Col = E_StockView.e_order_num
           strrequestno = Trim(.text)
           .Row = i
           .Col = E_StockView.e_Item
           strrequestitem = Trim(.text)
           strSql = "update erptemp..tblstockdb_temp set flag=0 where ORDER_NUM='" & strrequestno & "' and ITEM=" & strrequestitem & ""
           AddSql2 (strSql)
           SumCount = SumCount + 1

       End If
    Next i
   End With
   MsgBox SumCount & "笔申请记录撤销成功"

End Sub

    
    

Private Sub CreateApplication(Apptype As String)
    Dim RequestNo As String
    Dim strSql As String
    Dim SMR        As New ADODB.Recordset
    Dim i          As Integer
    Dim j          As Integer
    Dim intnum          As Integer
    
    Dim strid  As String
    Dim Kf_former  As String
    Dim Kf_dest   As String
    Dim strKF As String
    Dim strmatcode  As String
    Dim strCustCode  As String
    Dim Douprice  As Double
    Dim intqty   As Long
    Dim intitem As Integer
    Dim strbond  As String
    Dim SumCount As Integer
    Dim strchoose  As String
    Dim strbigbox_sel  As String
    Dim strlot_db  As String
    Dim strlot_sel  As String
    ' If Trim(Cob_kf_former.Text) = "" Then
       ' MsgBox "请选择业务库房！", vbInformation, Me.Caption
       ' Exit Sub
    ' End If
        
    'Kf_former = Left(Trim(Cob_kf_former.Text), InStr(Trim(Cob_kf_former.Text), " ") - 1)
'创建申请
    ' If Trim(Cob_kf_dest.Text) = Trim(Cob_kf_former.Text) Then
       ' MsgBox "业务库房和目标库房相同，申请失败！", vbInformation, Me.Caption
       ' Exit Sub
    ' End If
    ' If Apptype <> "VT" Then
        ' If Trim(Cob_kf_dest.Text) = "" Then
           ' MsgBox "请选择目标库房！", vbInformation, Me.Caption
           ' Exit Sub
        ' End If
         ' Kf_dest = Left(Trim(Cob_kf_dest.Text), InStr(Trim(Cob_kf_dest.Text), " ") - 1)
    
        ' If Left(Trim(Cob_kf_dest.Text), InStr(Trim(Cob_kf_dest.Text), " ") - 1) <> "72" And Left(Trim(Cob_kf_former.Text), InStr(Trim(Cob_kf_former.Text), " ") - 1) <> "72" Then
            ' strbond = "SELECT COUNT(*) FROM erpdata..tblstock a,erpdata..tblstock b WHERE a.库房代码 = '" & Left(Trim(Cob_kf_dest.Text), InStr(Trim(Cob_kf_dest.Text), " ") - 1) & "' AND b.库房代码 = '" & Left(Trim(Cob_kf_former.Text), InStr(Trim(Cob_kf_former.Text), " ") - 1) & "' AND b.库房类型 = a.库房类型"
        
             ' If SMR.State = adStateOpen Then SMR.Close
             ' SMR.Open strbond, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
             ' If SMR.Fields(0).Value = 0 Then
                ' MsgBox "业务库房和目标库房类型不同！", vbInformation, Me.Caption
                ' SMR.Close
                ' Exit Sub
             ' End If
             
         ' End If
        ' If Kf_former = "72" Then
           ' MsgBox "非委外回货，不可从72仓调拨！", vbInformation, Me.Caption
           ' Exit Sub
        ' End If
        ' If Kf_dest = "72" Then
           ' If Trim(Cob_Shipto.Text) = "" Then
               ' MsgBox "委外调拨，请选择发货地址！", vbInformation, Me.Caption
               ' Exit Sub
           ' End If
        ' End If
         

    ' Else
        ' If Kf_former <> "72" Then
               ' MsgBox "委回调拨，请选择72仓！", vbInformation, Me.Caption
               ' Exit Sub
        ' End If
    ' End If


   Txt_sqdh.text = ""
    
    
    
    With fpS_wafer
        If .MaxRows <= 0 Then
            MsgBox "请选择要操作的记录!", vbInformation, "提示"
            Exit Sub
        End If
    End With
    
    strbigbox_sel = ""
    strlot_sel = "" '大箱号#lot
    'merry20200202同一大箱中有多个lot,不能分开出货
    With fpS_Box
        For i = 1 To .MaxRows
            
            .Row = i
            .Col = E_BOX.E_CHOOSE
            strchoose = Trim(.text)
            If strchoose = "1" Then
                .Col = E_BOX.E_BOXID
                If InStr(strbigbox_sel, Trim(.text)) = 0 Then
                    If strbigbox_sel = "" Then
                        strbigbox_sel = Trim(.text)
                    Else
                        strbigbox_sel = strbigbox_sel & "," & Trim(.text)
                    End If
                End If
                strlot_sel = strlot_sel & "," & Trim(.text)
                .Col = E_BOX.E_LOTID
                strlot_sel = strlot_sel & "#" & Trim(.text)
            End If
        Next
        For i = 0 To UBound(Split(strbigbox_sel, ","))
            strSql = " select distinct rtrim(c.箱号) + '#' + rtrim(工单号) as 工单号 from erpdata..tblPackMainInfsub  a " & _
                   " inner join erpdata..tblPacktreeinf b on a.箱号=b.箱号 " & _
                   " inner join erpdata..tblPacktreeinf c on b.上级序号=c.序号  " & _
                   " where c.箱号 = '" & Split(strbigbox_sel, ",")(i) & "' "
            
            If SMR.State = adStateOpen Then SMR.Close
            SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            If SMR.RecordCount > 0 Then
                SMR.MoveFirst
                For j = 1 To SMR.RecordCount
                    strlot_db = SMR("工单号")
                    If InStr(strlot_sel, strlot_db) = 0 Then
                        MsgBox "大箱号" & Split(strbigbox_sel, ",")(i) & " 中还有其他lot,不可拆箱出货", vbInformation, "提示"
                        Exit Sub
                    End If
                    SMR.MoveNext
                Next
            End If
            SMR.Close
            Set SMR = Nothing
        Next

    End With

    
     '生成申请单号

    RequestNo = GetID()
    intitem = 0
    SumCount = 0

    With fpS_wafer
        If .MaxRows <= 0 Then
            MsgBox "请选择要操作的记录!", vbInformation, "提示"
            Exit Sub
        End If
    
        For intnum = 1 To .MaxRows
            .Row = intnum
            .Col = E_WAFERID.E_CHOOSE
            If .text <> "" Then
                If .text = 1 Then
                    .Col = E_WAFERID.E_BOXID
                    strXH = Trim(.text)      '箱号
                    .Col = E_WAFERID.E_WAFERID
                    strLCK = Trim(.text)     '流程卡编号
                    If Get_SqlserverCnt(" SELECT * FROM erptemp..tblstockdbsub_temp a ,erptemp..tblstockdb_temp b  where b.flag=1 and a.ORDER_NUM=b.ORDER_NUM and a.ITEM=b.ITEM  and a.qbox='" & strXH & "' and RTRIM(a.WAFER)='" & strLCK & "'") > 0 Then
                        MsgBox "箱号" & strXH & "，流程卡编号" & strLCK & " 已申请过，不可重复申请", vbInformation, "提示"
                        Exit Sub
                    End If
                    If Apptype = "WW" Then
                        If Get_SqlserverCnt(" SELECT * FROM erpdata..tblstockdbsub a ,erpdata..tblstockdb b  where  a.调拨编号=b.调拨编号 and a.序号=b.序号  and a.箱号='" & strXH & "' and RTRIM(a.流程卡编号)='" & strLCK & "' and a.调拨编号 not in (select 调拨编号 from erptemp..invalidstockdb) and a.调拨编号  not in (select 关联调拨编号 from erptemp..invalidstockdb)") > 0 Then
                            MsgBox "箱号" & strXH & "，流程卡编号" & strLCK & " 已回货，不可再次申请委外", vbInformation, "提示"
                            Exit Sub
                        End If
                    End If
                    
                End If
            End If
        Next
    
        
        
        For intnum = 1 To .MaxRows
            .Row = intnum
            .Col = E_WAFERID.E_CHOOSE
            If .text <> "" Then
                If .text = 1 Then
                    .Col = E_WAFERID.E_BOXID
                    strXH = Trim(.text)      '箱号
                    .Col = E_WAFERID.E_BigBoxID
                    strxh_big = Trim(.text)      '大箱号
                    
                    .Col = E_WAFERID.E_LOTID
                    strgdh = Trim(.text)      '工单号
                    
                    .Col = E_WAFERID.E_WAFERID
                    strLCK = Trim(.text)     '流程卡编号
                    
                    .Col = E_WAFERID.E_QTY
                    intqty = Val(.text)        '数量
                    .Col = E_WAFERID.E_Passqty
                    strlps = Trim(.text)        '良品数
                    .Col = E_WAFERID.E_Ngqty1
                    strbls = Trim(.text)        '不良品数
                    .Col = E_WAFERID.E_Ngqty2
                    strzcbls = Trim(.text)      '制程不良数
                    .Col = E_WAFERID.e_ID
                    strid = Trim(.text)
                    If Get_SqlserverCnt("select * from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'and  id=" & strid) > 0 Then
                        strSql = "select ITEM from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'and  id=" & strid
                        intitem = GetSqlServerStr(strSql)
                    Else
                        strSql = "select isnull(max(ITEM),0) from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'"
                        intitem = GetSqlServerStr(strSql) + 1
   
                        If SMR.State = adStateOpen Then SMR.Close
                        strSql = " select 库房编号,物料编号,客户代码,isnull(单价,0) from erpdata..tblStockNum where id=" & strid
                        SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
                        If SMR.RecordCount = 1 Then
                            SMR.MoveFirst
                            strKF = Trim(SMR("库房编号"))
                            strmatcode = Trim(SMR("物料编号"))
                            strCustCode = Trim(SMR("客户代码"))
                        
                        End If
                        If Apptype = "VT" Then
                            strSql = " select top 1 rtrim(a.原仓库) from erpdata..tblStockdb a,erpdata..tblStockdbsub b   where rtrim(b.流程卡编号)='" & strLCK & "' and a.调拨编号=b.调拨编号 and a.序号=b.序号 and a.目标仓库='72'"
                            Kf_dest = GetSqlServerStr(strSql)
                          
                            If Chk_NG.Value = 1 Then
                            
                                 Select Case Kf_dest
                                 Case "07", "20"
                                     Kf_dest = "30"
                                 Case "16", "19"
                                     Kf_dest = "28"
                                 Case Else
            
                                 End Select
                          
                            End If
                            
                            
                        ElseIf Apptype = "WW" Then
                            Kf_dest = "72"
                        Else
                            Kf_dest = Left(Trim(Cob_kf_dest.text), InStr(Trim(Cob_kf_dest.text), " ") - 1)
                        End If
                        
                      '上传主表

                     '调拨编号,序号,物料编号, 调拨数量,原仓库,目标仓库,申请人员,申请时间,审核人员,审核时间, 申请部门,状态,REMARK1,REMARK2,REMARK3,REMARK4,REMARK5,ID
                        strSql = "insert into erptemp..tblstockdb_temp(ORDER_NUM,ITEM, MATERIALS,QTY,FORMER, DESTINATION, APPLICANT, APPLICATION_TIME, AUDITOR, AUDIT_TIME, DEPT, FLAG,ID,REMARK1) values( " & _
                        "'" & RequestNo & "'," & intitem & ",'" & strmatcode & "'," & 0 & ",'" & strKF & "','" & Kf_dest & "','" & strWholeName & "',sysdatetime(),'','','',1," & strid & ",'" & Trim(Cob_Shipto.text) & "')"
                    
                        AddSql2 (strSql)
                       
                        
                    End If
                    
                    '上传子表
                    
                    '调拨编号, 序号, 箱号, 流程卡编号, 工单号, 合格数, 制程不良数, 来料不良数, ID
                     strSql = "insert into erptemp..tblstockdbsub_temp(ORDER_NUM,ITEM,WAFER,LOT,GOOD_DIE,BAD1_DIE,BAD2_DIE,ID,REMARK1,QBOX) values( " & _
                    "'" & RequestNo & "'," & intitem & ",'" & strLCK & "','" & strgdh & "'," & strlps & "," & strbls & "," & strzcbls & "," & strid & ",'" & strxh_big & "','" & strXH & "')"
                  
                    AddSql2 (strSql)
                    
                    'update主表数量
                    strSql = "Update erptemp..tblstockdb_temp set QTY =QTY+" & Val(strlps) + Val(strbls) + Val(strzcbls) & " where ORDER_NUM='" & RequestNo & "' and ITEM=" & intitem
                   
                    AddSql2 (strSql)
                    SumCount = SumCount + 1
                    
                    
                    
                End If
            End If
        Next intnum

    End With
    If SumCount > 0 Then
        MsgBox SumCount & "笔记录申请成功", vbInformation, "提示"
        Txt_sqdh.text = RequestNo
    End If
     
End Sub


     
Function GetID()
'FWW1911140011
'生成方式：FWW+YYMMDD +4位流水码
Dim CODE       As String
Dim strSql     As String
Dim YearStr    As String
Dim MonthStr   As String
Dim DayStr     As String
Dim SMR        As New ADODB.Recordset


YearStr = Right(Year(Now()), 2)
If Len(Month(Now())) = 1 Then
    MonthStr = "0" & Month(Now())
Else
    MonthStr = Month(Now())
End If
If Len(Day(Now())) = 1 Then
    DayStr = "0" & Day(Now())
Else
    DayStr = Day(Now())
End If
CODE = YearStr & MonthStr & DayStr

strSql = "Select Isnull(max(RIGHT(ORDER_NUM,LEN(ORDER_NUM)-3)),0) as ORDER_NUM from erptemp..tblStockdb_temp where left(ORDER_NUM,9)='FWW" & CODE & "'"


If SMR.State = adStateOpen Then SMR.Close
SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If SMR("ORDER_NUM") = 0 Then

    GetID = "FWW" & CODE & "0001"
Else
    GetID = "FWW" & Val(SMR("ORDER_NUM")) + 1
End If
SMR.Close
Set SMR = Nothing

End Function

Sub ListView1Data(searchtype As String)
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim gdh As String
    Dim Kf_former  As String
    Dim i As Integer
    

    
    If Chk_Keepdata.Value = 1 Then
        With fpS_Box
        '先判断Lot号是否已经存在
        For i = 1 To .MaxRows
            .Row = i
            .Col = E_BOX.E_LOTID
            If Trim(TxtLot.text) = Trim(.text) Then
                MsgBox "此Lot已经查询过，不要重复查询", vbInformation, "提示"
                Exit Sub
            End If
        Next
        End With
    Else
        With fpS_Box
        .DataSource = Nothing
        .MaxRows = 0
        End With
        With fpS_wafer
            .DataSource = Nothing
            .MaxRows = 0
        End With
    End If

    

    
   '

    strSql = "SELECT distinct 0 as '√',a.客户代码,i.MPN_DESC as 客户机种 , g.QTECHPTNO as 厂内机种,a.工单号,dbo.f_getparent(f.箱号)  as 大箱号 ,a.物料编号,a.料号, b.规格,b.型号,b.计量单位名称,a.合格数,a.不良数 AS 不良数,a.制程不良数 , a.id,c.库房代码" & _
    " FROM  erpdata..tblStockNum AS a " & _
    " INNER JOIN  erpbase..tblSmainM2 AS b ON a.物料编号 = b.物料编号  " & _
    " INNER JOIN  erpbase..tblstock AS c ON a.库房编号 = c.库房代码  " & _
    " INNER JOIN  erpdata..tblbase d on a.产线标记=d.名称 and d.说明2='产线标记'  " & _
    " LEFT JOIN erpdata..tblWithWork e ON a.订单编号=e.订单编号 AND a.订单项次=e.订单项次    " & _
    " LEFT JOIN  erpdata..tblStockNumsub f on  f.id=a.id  " & _
    " LEFT JOIN  erptemp..tbltsvnpiproduct g ON g.QTECHPTNO2=f.料号   " & _
     " left join erpbase..tblmappingdata h on h.SUBSTRATEID=f.流程卡编号 and h.LOTID=f.工单号 " & _
     " left join erpbase..tblcustomeroi i on convert(varchar(20)  ,i.id)=h.FILENAME and h.LOTID= i.SOURCE_BATCH_ID " & _
    " where a.合格数+a.不良数+a.制程不良数>0 "

    'If Kf_former <> "" Then strSql = strSql & " and  a.库房编号='" & Kf_former & "'"

    If Trim(Cmbcust.text) <> "" Then strSql = strSql & " and  a.客户代码='" & Trim(Cmbcust.text) & "'"
    If Trim(TxtCustpn.text) <> "" Then strSql = strSql & " and  i.MPN_DESC='" & Trim(TxtCustpn.text) & "'"
    If Trim(TxtPN.text) <> "" Then strSql = strSql & " and  a.料号='" & Trim(TxtPN.text) & "'"
    If Trim(TxtLot.text) <> "" Then strSql = strSql & " and  a.工单号='" & Trim(TxtLot.text) & "'"
    If searchtype = "VT" Then
     strSql = strSql & " AND a.库房编号 IN ('72')"
     'strSql = strSql & "  and  a.工单号 in (select distinct CUSTLOT from erptemp..TSV_VT_History_sub  where  flag=1  and CUSTOMERSHORTNAME='" & Trim(Cmbcust.Text) & "') "
     
    ElseIf searchtype = "WW" Then
        If Chk_NG.Value = 1 Then
            strSql = strSql & " AND a.库房编号 IN ('28','30')"
        Else
            strSql = strSql & " AND a.库房编号 IN ('07','16','19','20')"
        End If
        'If Trim(TxtLot.Text) <> "" Then strSql = strSql & " and  a.工单号='" & Trim(TxtLot.Text) & "'"
    Else
        If Trim(Cob_kf_former.text) <> "" Then
            Kf_former = Left(Cob_kf_former.text, InStr(Cob_kf_former.text, " ") - 1)
            If Kf_former <> "" Then strSql = strSql & " and  a.库房编号='" & Kf_former & "'"
        End If
    End If

    
    If gdh = "" Then

    Else
        strSql = strSql & " GROUP BY  a.id, a.客户代码, a.工单号, a.物料编号, a.料号, b.规格,  b.型号,  b.计量单位名称 ,dbo.f_getparent(d.箱号)"
    End If
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
 
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        If Chk_Keepdata.Value = 1 Then
        '查询时保留已选择的Lot号
            With fpS_Box
               For i = 1 To SMR.RecordCount
                   .MaxRows = .MaxRows + 1
                   .SetText E_BOX.E_CHOOSE, .MaxRows, 0
                   .SetText E_BOX.E_CUSTCODE, .MaxRows, SMR("客户代码")
                   .SetText E_BOX.E_CUSTPN, .MaxRows, SMR("客户机种")
                   .SetText E_BOX.E_qtechPTNo, .MaxRows, SMR("厂内机种")
                   .SetText E_BOX.E_LOTID, .MaxRows, SMR("工单号")
                   .SetText E_BOX.E_BOXID, .MaxRows, SMR("大箱号")
                   .SetText E_BOX.E_Matcode, .MaxRows, SMR("物料编号")
                   .SetText E_BOX.E_partno, .MaxRows, SMR("料号")
                   .SetText E_BOX.E_Matspec, .MaxRows, SMR("规格")
                   .SetText E_BOX.E_Mattype, .MaxRows, SMR("型号")
                   .SetText E_BOX.E_UNIT, .MaxRows, SMR("计量单位名称")
                   .SetText E_BOX.E_Passqty, .MaxRows, SMR("合格数")
                   .SetText E_BOX.E_Ngqty1, .MaxRows, SMR("不良数")
                   .SetText E_BOX.E_Ngqty2, .MaxRows, SMR("制程不良数")
                   .SetText E_BOX.e_ID, .MaxRows, SMR("ID")
                   .SetText E_BOX.E_StockID, .MaxRows, SMR("库房代码")

                   SMR.MoveNext
               Next
            End With
        Else
            With fpS_Box
            .MaxRows = 0
            Set .DataSource = SMR
            
            End With
        End If
    End If
   If searchtype = "WW" Then
        With fpS_Box
            For i = 1 To .MaxRows
               .Row = i
               .Col = E_BOX.E_BOXID
               
                strSql = " select d.箱号 from erpdata..tblStockdbsub a " & _
                    " inner join erpdata..tblStockdb  b on a.调拨编号=b.调拨编号 and a.序号=b.序号 " & _
                    " inner join erpdata..tblStockNumTree  c on a.箱号=c.箱号 " & _
                    " inner join erpdata..tblStockNumTree  d on c.上级序号=d.序号 " & _
                    " where b.目标仓库='72' and exists( select * from erpdata..tblStockNumSub where 流程卡编号=a.流程卡编号) " & _
                    " and not exists( select * from erptemp..InvalidStockDb  where 关联调拨编号=a.调拨编号) " & _
                    " and rtrim(d.箱号)='" & Trim(.text) & "'"
                    
                If Get_SqlserverCnt(strSql) > 0 Then
                    .Row = i
                    .Col = -1
                    .BackColor = &H808080
                Else
                    .BackColor = &HFFFFFF
                End If
            Next
        End With
    End If
                    
    
End Sub



Private Sub SearchWaferID_ByBoxID(kf As String, strLotID As String, strBoxID As String, intBJ As Integer)
    Dim i          As Integer
    Dim j          As Integer
    Dim strSql     As String
    Dim rs         As New ADODB.Recordset
    Dim Lot_temp   As String
    Dim Box_temp   As String
    Dim Stock_temp As String
   ' Dim kf As String
    Dim BoxIdExist As Boolean
    Dim adorst1         As New ADODB.Recordset
 
   ' kf = Left(Trim(Cob_kf_former.Text), InStr(Trim(Cob_kf_former.Text), " ") - 1)

    If intBJ = 1 Then '勾选

        With fpS_wafer
           If .MaxRows = 0 Then
                   '查询资料
                Set adorst1 = New ADODB.Recordset
                Set adorst1.ActiveConnection = INIadoCon2
                
            adorst1.Source = "SELECT a.id,a.箱号, a.工单号,a.流程卡编号,a.料号, a.物料编号, sum(a.数量) as 数量, case when a.合格标记=0 then sum(a.数量) else 0 end as 合格品, case when a.合格标记=2 then sum(a.数量) else 0 end  as 不良品,case when a.合格标记=1 then sum(a.数量) else 0 end  as 制程不良品,a.发货标记 ,'' 新箱号 FROM  dbo.tblStockNumSub AS a INNER JOIN dbo.f_kcdb('" & strBoxID & "') AS b ON a.箱号 = b.箱号 INNER JOIN erpdata..tblstocknum c ON c.ID = a.ID " & _
              "  where charindex(rtrim(a.工单号),'" & strLotID & " ')>0 and  a.数量>0 and c.库房编号 = '" & kf & "' group by  a.id,a.箱号,a.料号,a.物料编号,a.合格标记,a.发货标记 ,a.工单号,a.流程卡编号  " & _
              " union " & _
              "  SELECT a.id,a.箱号, a.工单号,a.流程卡编号, a.料号, a.物料编号, sum(a.数量) as 数量, case when a.合格标记=0 then sum(a.数量) else 0 end as 合格品, case when a.合格标记=2 then sum(a.数量) else 0 end as 不良品,case when a.合格标记=1 then sum(a.数量) else 0 end  as 制程不良品,a.发货标记 ,'' 新箱号  FROM  dbo.tblStockNumSub AS a INNER JOIN tblStockNumtree AS b ON a.箱号 = b.箱号 INNER JOIN erpdata..tblstocknum c ON c.ID = a.ID " & _
              "  where  rtrim(a.工单号)='" & Replace(strLotID, "$", "") & "' and  a.数量>0 and rtrim(a.箱号)='" & Replace(Trim(strBoxID), "★", "") & "' and c.库房编号 = '" & kf & "'  group by  a.id,a.箱号,a.料号,a.物料编号,a.合格标记,a.发货标记, a.工单号,a.流程卡编号 order by a.流程卡编号 "
                            
                adorst1.Open , , adOpenStatic, adLockReadOnly, adCmdText
               
            
                If adorst1.RecordCount > 0 Then
                    adorst1.MoveFirst

                    For j = 1 To adorst1.RecordCount

                        .MaxRows = .MaxRows + 1
                        
                        .SetText E_WAFERID.E_CHOOSE, .MaxRows, 1
                        .SetText E_WAFERID.E_LOTID, .MaxRows, Trim$("" & adorst1!工单号)
                        .SetText E_WAFERID.E_BigBoxID, .MaxRows, Replace(strBoxID, "★", "")
                        .SetText E_WAFERID.E_BOXID, .MaxRows, Trim$("" & adorst1!箱号)
                        .SetText E_WAFERID.E_WAFERID, .MaxRows, Trim$("" & adorst1!流程卡编号)
                        .SetText E_WAFERID.E_PN, .MaxRows, Trim$("" & adorst1!料号)
                        .SetText E_WAFERID.E_QTY, .MaxRows, Trim$("" & adorst1!数量)
                        
                        .SetText E_WAFERID.E_Passqty, .MaxRows, Trim$("" & adorst1!合格品)
                        .SetText E_WAFERID.E_Passqty, .MaxRows, Trim$("" & adorst1!合格品)
                        .SetText E_WAFERID.E_Ngqty1, .MaxRows, Trim$("" & adorst1!不良品)
                        .SetText E_WAFERID.E_Ngqty2, .MaxRows, Trim$("" & adorst1!制程不良品)
                        
                        .SetText E_WAFERID.e_ID, .MaxRows, Trim$("" & adorst1!id)

                        adorst1.MoveNext
                    Next
        
                End If
            Else

                For i = 1 To .MaxRows
                    .Row = i
                    .Col = E_WAFERID.E_BigBoxID
                    Box_temp = Trim$(.text)
                    .Row = i
                    .Col = E_WAFERID.E_LOTID
                    Lot_temp = Trim$(.text)

                    If Replace(strBoxID, "★", "") = Trim(Box_temp) And Replace(strLotID, "$", "") = Lot_temp Then
                        Exit Sub
                    End If

                Next

                   '查询资料
                Set adorst1 = New ADODB.Recordset
                Set adorst1.ActiveConnection = INIadoCon2
       ' adorst1.Source = "SELECT a.id,a.箱号, a.工单号,a.流程卡编号,a.料号, a.物料编号, sum(a.数量) as 数量,case when a.合格标记=0 then sum(a.数量) else 0 end as 合格品, case when a.合格标记=2 then sum(a.数量) else 0 end  as 不良品,case when a.合格标记=1 then sum(a.数量) else 0 end  as 制程不良品,a.发货标记 ,'' 新箱号 FROM  dbo.tblStockNumSub AS a INNER JOIN dbo.f_kcdb('" & strBoxID & "') AS b ON a.箱号 = b.箱号 INNER JOIN erpdata..tblstocknum c ON c.ID = a.ID " & _
        '  "  where charindex(rtrim(a.工单号),'" & strLotID & " ')>0 and  a.数量>0 and c.库房编号 = '" & kf & "' group by  a.id,a.箱号,a.料号,a.物料编号,a.合格标记,a.发货标记 ,a.工单号,a.流程卡编号 " & _
        '  " union " & _
        '  "  SELECT a.id,a.箱号, a.工单号,a.流程卡编号, a.料号, a.物料编号,  sum(a.数量) as 数量, case when a.合格标记=0 then sum(a.数量) else 0 end as 合格品, case when a.合格标记=2 then sum(a.数量) else 0 end as 不良品,case when a.合格标记=1 then sum(a.数量) else 0 end  as 制程不良品,a.发货标记 ,'' 新箱号  FROM  dbo.tblStockNumSub AS a INNER JOIN tblStockNumtree AS b ON a.箱号 = b.箱号 INNER JOIN erpdata..tblstocknum c ON c.ID = a.ID " & _
        '  "  where  charindex(rtrim(a.工单号),'" & strLotID & " ')>0 and  a.数量>0 and charindex(rtrim(a.箱号),'" & Trim(strBoxID) & " ')>0 and c.库房编号 = '" & kf & "'  group by  a.id,a.箱号,a.料号,a.物料编号,a.合格标记,a.发货标记, a.工单号,a.流程卡编号"
            
            adorst1.Source = "SELECT a.id,a.箱号, a.工单号,a.流程卡编号,a.料号, a.物料编号, sum(a.数量) as 数量, case when a.合格标记=0 then sum(a.数量) else 0 end as 合格品, case when a.合格标记=2 then sum(a.数量) else 0 end  as 不良品,case when a.合格标记=1 then sum(a.数量) else 0 end  as 制程不良品,a.发货标记 ,'' 新箱号 FROM  dbo.tblStockNumSub AS a INNER JOIN dbo.f_kcdb('" & strBoxID & "') AS b ON a.箱号 = b.箱号 INNER JOIN erpdata..tblstocknum c ON c.ID = a.ID " & _
              "  where charindex(rtrim(a.工单号),'" & strLotID & " ')>0 and  a.数量>0 and c.库房编号 = '" & kf & "' group by  a.id,a.箱号,a.料号,a.物料编号,a.合格标记,a.发货标记 ,a.工单号,a.流程卡编号  " & _
              " union " & _
              "  SELECT a.id,a.箱号, a.工单号,a.流程卡编号, a.料号, a.物料编号, sum(a.数量) as 数量, case when a.合格标记=0 then sum(a.数量) else 0 end as 合格品, case when a.合格标记=2 then sum(a.数量) else 0 end as 不良品,case when a.合格标记=1 then sum(a.数量) else 0 end  as 制程不良品,a.发货标记 ,'' 新箱号  FROM  dbo.tblStockNumSub AS a INNER JOIN tblStockNumtree AS b ON a.箱号 = b.箱号 INNER JOIN erpdata..tblstocknum c ON c.ID = a.ID " & _
              "  where  rtrim(a.工单号)='" & Replace(strLotID, "$", "") & "' and  a.数量>0 and rtrim(a.箱号)='" & Replace(Trim(strBoxID), "★", "") & "' and c.库房编号 = '" & kf & "'  group by  a.id,a.箱号,a.料号,a.物料编号,a.合格标记,a.发货标记, a.工单号,a.流程卡编号 order by a.流程卡编号 "
                                          
            adorst1.Open , , adOpenStatic, adLockReadOnly, adCmdText
          
                If adorst1.RecordCount > 0 Then
                    adorst1.MoveFirst

                    For j = 1 To adorst1.RecordCount

                        .MaxRows = .MaxRows + 1
                        .SetText E_WAFERID.E_CHOOSE, .MaxRows, 1
                        .SetText E_WAFERID.E_LOTID, .MaxRows, Trim$("" & adorst1!工单号)
                        .SetText E_WAFERID.E_BigBoxID, .MaxRows, Replace(strBoxID, "★", "")
                        .SetText E_WAFERID.E_BOXID, .MaxRows, Trim$("" & adorst1!箱号)
                        .SetText E_WAFERID.E_WAFERID, .MaxRows, Trim$("" & adorst1!流程卡编号)
                        .SetText E_WAFERID.E_PN, .MaxRows, Trim$("" & adorst1!料号)
                        .SetText E_WAFERID.E_QTY, .MaxRows, Trim$("" & adorst1!数量)
                        .SetText E_WAFERID.E_Passqty, .MaxRows, Trim$("" & adorst1!合格品)
                        .SetText E_WAFERID.E_Ngqty1, .MaxRows, Trim$("" & adorst1!不良品)
                        .SetText E_WAFERID.E_Ngqty2, .MaxRows, Trim$("" & adorst1!制程不良品)
                        .SetText E_WAFERID.e_ID, .MaxRows, Trim$("" & adorst1!id)

                        adorst1.MoveNext
                    Next
        
                End If



            End If

        End With

    End If

    If intBJ = 2 Then '取消勾选

        With fpS_wafer

            For i = .MaxRows To 1 Step -1
                    .Row = i
                    .Col = E_WAFERID.E_BigBoxID
                    Box_temp = Trim$(.text)
                    .Row = i
                    .Col = E_WAFERID.E_LOTID
                    Lot_temp = Trim$(.text)

                If Replace(strBoxID, "★", "") = Trim(Box_temp) And Replace(strLotID, "$", "") = Lot_temp Then
                    .DeleteRows i, 1
                    .MaxRows = .MaxRows - 1

                End If

            Next

        End With

    End If
reflashQty
End Sub




Sub inictrl()
Dim i As Integer

    
    With fpS_Box
        .MaxCols = E_BOX.E_END - 1
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
        .Col = E_BOX.E_CHOOSE   '选择
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        '设定列宽
        .ColWidth(-1) = 10
        .ColWidth(E_BOX.E_CHOOSE) = 4
        .ColWidth(E_BOX.E_CUSTCODE) = 6
        .RowHeight(-1) = 10
        '设定是否排序
        .UserColAction = UserColActionSort

        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
  

         .SetText E_BOX.E_CHOOSE, 0, "√"
         .SetText E_BOX.E_CUSTCODE, 0, "客户"
         .SetText E_BOX.E_qtechPTNo, 0, "厂内机种"
         .SetText E_BOX.E_LOTID, 0, "工单号"
         .SetText E_BOX.E_BOXID, 0, "大箱号"
         .SetText E_BOX.E_Matcode, 0, "物料编号"
         .SetText E_BOX.E_partno, 0, "料号"
         .SetText E_BOX.E_Matspec, 0, "规格"
         .SetText E_BOX.E_Mattype, 0, "型号"
         .SetText E_BOX.E_UNIT, 0, "计量单位名称"
         .SetText E_BOX.E_Passqty, 0, "合格数"
         .SetText E_BOX.E_Ngqty1, 0, "不良数"
         .SetText E_BOX.E_Ngqty2, 0, "制程不良数"
         .SetText E_BOX.e_ID, 0, "id"
         
        
        .ZOrder
        .ReDraw = True
    End With
    
    With fpS_wafer
        .MaxCols = E_WAFERID.E_END - 1
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
        .Col = E_WAFERID.E_CHOOSE   '选择
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        '设定列宽
        .ColWidth(-1) = 10
        .ColWidth(E_WAFERID.E_CHOOSE) = 4
        .RowHeight(-1) = 10
        '设定是否排序
        .UserColAction = UserColActionSort

        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next

    
        .SetText E_WAFERID.E_CHOOSE, 0, "√"
        .SetText E_WAFERID.E_LOTID, 0, "工单号"
        .SetText E_WAFERID.E_BigBoxID, 0, "大箱号"
        .SetText E_WAFERID.E_BOXID, 0, "箱号"
        .SetText E_WAFERID.E_WAFERID, 0, "流程卡编号"
        .SetText E_WAFERID.E_PN, 0, "料号"
        .SetText E_WAFERID.E_QTY, 0, "数量"
        .SetText E_WAFERID.E_Passqty, 0, "合格数"
        .SetText E_WAFERID.E_Ngqty1, 0, "来料不良数"
        .SetText E_WAFERID.E_Ngqty2, 0, "制程不良数"
        .SetText E_WAFERID.e_ID, 0, "ID"
        .ZOrder
        .ReDraw = True
    End With
    

       
     With fpS_stockview
    
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
        .Col = 1 '选择
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
        .ReDraw = False
    End With
        
     With fpS_Cancerview
    
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
        .Col = 1 '选择
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
        .ReDraw = False
    End With
    With fpS_vt_lot
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
        .Col = 1 '选择
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        
        '设定列宽
        .ColWidth(-1) = 10
        .ColWidth(E_VT.E_CHOOSE) = 4

        
        .RowHeight(-1) = 10
        '设定是否排序
        .UserColAction = UserColActionSort
        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next

        .ZOrder
        .ReDraw = False
    End With

End Sub

Private Sub fpS_Wafer_Click(ByVal Col As Long, ByVal Row As Long)
Exit Sub
Dim i           As Long
Dim j           As Integer
Dim strTmp      As String

    '点击把选择的单号都选上
    If Row < 1 Then Exit Sub
    If Col <> 1 Then Exit Sub
    With fpS_wafer

        .Col = E_WAFERID.E_CHOOSE
        .Row = Row
        .Value = Abs(Val(.Value) - 1)
        If Val(.Value) = 1 Then   '将所有一样的大箱号选择上
            .Col = E_WAFERID.E_BigBoxID
            .Row = Row
            strTmp = Trim$(.text)
            For i = 1 To .MaxRows
                .Row = i
                .Col = E_WAFERID.E_BigBoxID
                If Trim$(.text) = strTmp Then
                    .Col = E_WAFERID.E_CHOOSE
                    .Value = 1
                    .Col = -1
                    .ForeColor = &HFF8080
                End If
            Next
        Else
            '将所有一样的单号选择上
            .Col = E_WAFERID.E_BigBoxID
            .Row = Row
            strTmp = Trim$(.text)
'            strDJBH = Trim$(.Text) '共用的单据编号，在导出打印时会用到
            For i = 1 To .MaxRows
                .Row = i
                .Col = E_WAFERID.E_BigBoxID
                If Trim$(.text) = strTmp Then
                    .Col = E_WAFERID.E_CHOOSE
                    .Value = 0
                    .Col = -1
                    .ForeColor = vbBlack
                End If
            Next
        End If
        
    End With
    
End Sub








Private Sub stockmove()
 Dim i As Integer
 Dim strrequestno As String
 Dim strrequestitem As String
 Dim strSql As String
 Dim MsgRly As String
 Dim dbno_cancer As String
 Dim strid As String
 Dim dbno_item As String
 Dim dbno As String
 Dim dbitem As String
 Dim strykf As String
 Dim selcnt As Integer
 '@XH CHAR(8000),--小箱号
 '@lck   CHAR(8000) ,---流程卡编号
 '@gdh  CHAR(8000),--工单号
 '@lps   CHAR(8000) ,---良品数
 '@blS   CHAR(8000) ,---不良品数
 '@zcbls  CHAR(8000),--制程不良数
 '@FDCStock CHAR(50),--目标库房
 '@dbry CHAR(50), --调拨人员
 '@sqbm CHAR(20)='07', --申请部门 默认计划部
 '@NEWBOX CHAR(50) = '

    strrequestno = ""
    If strWholeName = gUserName Then
        MsgBox "您没有权限执行此动作", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(TxtRequestNo.text) = "" Then
        MsgBox "请输入申请单号，查询", vbInformation, "提示"
        Exit Sub
    End If
    selcnt = 0
    
 With fpS_stockview
    '先check
    For i = 1 To .MaxRows
        .Row = i
        .Col = E_StockView.E_CHOOSE
        
        If .text = 1 Then
            selcnt = i
           .Row = i
           .Col = E_StockView.e_order_num
           
            If Trim(.text) <> Trim(TxtRequestNo.text) Then
                MsgBox "不同申请单号不可一起调拨！", vbInformation, "提示"
                Exit Sub
            End If
    
            .Row = i
            .Col = E_StockView.E_KF_FORMER
            strykf = Trim(.text) ' 原库房

            .Row = i
            .Col = E_StockView.E_Wafer
            strLCK = Trim(.text)  '流程卡编号
            
            .Row = i
            .Col = E_StockView.e_Qbox
                        
            strSql = "select distinct rtrim(库房编号) from erpdata..tblstocknumsub where 箱号='" & Trim(.text) & "' and 流程卡编号='" & strLCK & "'"
            If GetSqlServerStr(strSql) <> strykf Then
               MsgBox "箱号" & Trim(.text) & "不在" & strykf & "库房,无法调拨", vbInformation, "提示"
               Exit Sub
            End If
        End If
    
            
    Next
    
    If selcnt = 0 Then
        MsgBox "没有需要调拨的单号,请先查询", vbInformation, "提示"
        Exit Sub
    End If
    strXH = ""
    strgdh = ""
    strLCK = ""
    strlps = ""
    strbls = ""
    strzcbls = ""
    strmbkf = ""
    For i = 1 To .MaxRows
    
        .Row = i
        .Col = E_StockView.E_CHOOSE
        
        If .text = 1 Then
            .Row = i
            .Col = E_StockView.E_KF_DEST
            '目标库房不同，需拆分成不同的调拨单号
            If strmbkf = "" Then
                strmbkf = Trim(.text) ' 目标库房
            Else
                If Trim(.text) <> strmbkf Then
                    Call DataOpt
                    strXH = ""
                    strgdh = ""
                    strLCK = ""
                    strlps = ""
                    strbls = ""
                    strzcbls = ""
                    strmbkf = Trim(.text) ' 目标库房不同
                End If
            End If
            '原库房不同，需拆分成不同的调拨单号
            .Row = i
            .Col = E_StockView.E_KF_FORMER
            If strykf = "" Then
                strykf = Trim(.text) ' 原库房
            Else
                If Trim(.text) <> strykf Then
                    Call DataOpt
                    strXH = ""
                    strgdh = ""
                    strLCK = ""
                    strlps = ""
                    strbls = ""
                    strzcbls = ""
                    strykf = Trim(.text) ' 原库房
                End If
            End If
            .Row = i
            .Col = E_StockView.e_Qbox
            strXH = strXH & Trim(.text) & "★" '小箱号
    
            .Row = i
            .Col = E_StockView.E_LOT
            strgdh = strgdh & Trim(.text) & "★" '工单号
           
            .Row = i
            .Col = E_StockView.E_Wafer
            strLCK = strLCK & Trim(.text) & "★" '流程卡编号
           
            .Row = i
            .Col = E_StockView.E_GOOD_DIE
            strlps = strlps & Trim(.text) & "★" '良品数
           
            .Row = i
            .Col = E_StockView.E_BAD1_DIE
            strbls = strbls & Trim(.text) & "★" '不良品数
        
           
            .Row = i
            .Col = E_StockView.E_BAD2_DIE
            strzcbls = strzcbls & Trim(.text) & "★" '制程不良数
        End If
    Next i
    
    
 End With
    strrequestno = ""
    strrequestitem = ""
    If DataOpt() = True Then

       'update erptemp..tblstockdb_temp的flag状态
        With fpS_stockview
            For i = 1 To .MaxRows
           
               .Row = i
               .Col = E_StockView.E_CHOOSE
               
               If .text = 1 Then
                   .Row = i
                   .Col = E_StockView.e_order_num
                   strrequestno = Trim(.text)
                   .Row = i
                   .Col = E_StockView.e_Item
                   strrequestitem = Trim(.text)
                   .Row = i
                   .Col = E_StockView.e_ID
                   strid = Trim(.text)
                   
                   strSql = "select top 1 rtrim(调拨编号) + '-' + convert(varchar(5),序号)  from erpdata..tblstockdb where rtrim(申请人员)='" & strWholeName & "' and id=" & strid & " and DATEDIFF(mi,申请时间,sysdatetime())<5 ORDER BY 申请时间 desc"
                   dbno_item = GetSqlServerStr(strSql)
                   If InStr(dbno_item, "-") > 0 Then
                       dbno = Split(dbno_item, "-")(0)
                       dbitem = Split(dbno_item, "-")(1)
                       strSql = "update erptemp..tblstockdb_temp set flag=2, remark2='" & dbno & "', remark3='" & dbitem & "',AUDITOR='" & strWholeName & "', AUDIT_TIME=sysdatetime()  where ORDER_NUM='" & strrequestno & "' and ITEM=" & strrequestitem & " and flag=1 "
                       AddSql2 (strSql)
                       strSql = "update erptemp..tblstockdb_temp set flag=4, remark2='" & dbno & "', remark3='" & dbitem & "',AUDITOR='" & strWholeName & "', AUDIT_TIME=sysdatetime()  where ORDER_NUM='" & strrequestno & "' and ITEM=" & strrequestitem & " and flag=3 "
                       AddSql2 (strSql)
                    Else
                       MsgBox strrequestno & strrequestitem & "调拨出现异常，请提出", vbInformation, "提示"
                       Exit Sub
                    End If
           
                  
               End If
           Next i
           
        End With

   Else
       Exit Sub
    End If
    
 End Sub



Private Sub cancerstockmove()
 Dim i As Integer
 Dim strrequestno As String
 Dim strrequestitem As String
 Dim strSql As String
 Dim MsgRly As String
 Dim dbno_cancer As String
 Dim strykf As String
 Dim selcnt As Integer
 
 
 '@XH CHAR(8000),--小箱号
 '@lck   CHAR(8000) ,---流程卡编号
 '@gdh  CHAR(8000),--工单号
 '@lps   CHAR(8000) ,---良品数
 '@blS   CHAR(8000) ,---不良品数
 '@zcbls  CHAR(8000),--制程不良数
 '@FDCStock CHAR(50),--目标库房
 '@dbry CHAR(50), --调拨人员
 '@sqbm CHAR(20)='07', --申请部门 默认计划部
 '@NEWBOX CHAR(50) = '
 
 
    If strWholeName = gUserName Then
        MsgBox "您没有权限执行此动作", vbInformation, "提示"
        'Exit Sub
    End If
    
    If Label1.Caption <> "调拨单号" Then
        MsgBox "请按调拨单号撤销", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(TxtdbNo.text) = "" Then
        MsgBox "请按调拨单号撤销", vbInformation, "提示"
        Exit Sub
    End If
    selcnt = 0
 With fpS_Cancerview
 
    For i = 1 To .MaxRows
    
        .Row = i
        .Col = E_StockView.E_CHOOSE
        If .text = 1 Then
            .Row = i
            .Col = E_StockView.e_Qbox
            strXH = Trim(.text)  '小箱号
            
            .Row = i
            .Col = E_StockView.E_Wafer
            strLCK = Trim(.text) '流程卡编号
        
            .Row = i
            .Col = E_StockView.E_KF_DEST
            strykf = Trim(.text) ' 原库房'撤销调拨 ，反向,将目标库房定为原库房
            If Trim(.text) <> "72" Then
                MsgBox "目标仓非72仓，不可撤销！", vbInformation, "提示"
                Exit Sub
            End If

            strSql = "select distinct Rtrim(库房编号) from erpdata..tblstocknumsub where 箱号='" & strXH & "' and 流程卡编号='" & strLCK & "'"
            If GetSqlServerStr(strSql) <> strykf Then
               MsgBox "箱号" & Trim(.text) & "已不在" & strykf & "库房,无法撤销", vbInformation, "提示"
               Exit Sub
               
            End If
        End If
    Next
    strXH = ""
    strgdh = ""
    strLCK = ""
    strlps = ""
    strbls = ""
    strzcbls = ""
    strmbkf = ""

    For i = 1 To .MaxRows
    
        .Row = i
        .Col = E_StockView.E_CHOOSE
        
        If .text = 1 Then
        
           .Row = i
           .Col = E_StockView.E_KF_FORMER
           strmbkf = Trim(.text) ' 目标库房'撤销调拨 ，反向,将原库房定为目标库房

           .Row = i
           .Col = E_StockView.E_KF_DEST
           strykf = Trim(.text) ' 原库房'撤销调拨 ，反向,将目标库房定为原库房
        
          .Row = i
          .Col = E_StockView.e_Qbox
           strXH = strXH & Trim(.text) & "★" '小箱号
           
          .Row = i
          .Col = E_StockView.E_LOT
           strgdh = strgdh & Trim(.text) & "★" '工单号
           
           .Row = i
          .Col = E_StockView.E_Wafer
           strLCK = strLCK & Trim(.text) & "★" '流程卡编号
           
          .Row = i
          .Col = E_StockView.E_GOOD_DIE
           strlps = strlps & Trim(.text) & "★" '良品数
           
          .Row = i
          .Col = E_StockView.E_BAD1_DIE
           strbls = strbls & Trim(.text) & "★" '不良品数
                   
          .Row = i
          .Col = E_StockView.E_BAD2_DIE
           strzcbls = strzcbls & Trim(.text) & "★" '制程不良数


           
      
        End If
    Next i
 End With
    If strXH = "" Then
        MsgBox "请选择要操作的记录!", vbCritical + vbOKCancel, "系统提示"
        Exit Sub
    End If

    If DataOpt() = True Then
       
        strSql = "select top 1 调拨编号 from erpdata..tblstockdb where 申请人员='" & strWholeName & "'  and DATEDIFF(mi,申请时间,sysdatetime())<5 ORDER BY 申请时间 desc"
        dbno_cancer = GetSqlServerStr(strSql)
        If dbno_cancer = "" Then
            MsgBox "撤销出现异常，请提出", vbInformation, "提示"
            Exit Sub
        Else
            strSql = "insert into  erptemp..invalidstockdb(调拨编号,关联调拨编号) values('" & dbno_cancer & "','" & Trim(TxtdbNo.text) & "')"
            AddSql2 (strSql)
        End If
    Else
        Exit Sub
    End If
    
 
 
 '
 End Sub
 
 
Function DataOpt()
 Err.Clear
 On Error GoTo there
  Dim num As Integer
  Dim strCon As String
  Dim TblName As String
  Dim strrkd As String
  
  Dim strbond As String
  Dim rs     As New ADODB.Recordset
  
  Dim i As Integer
  Dim j As Integer
  Dim adoPrmReturn As New ADODB.Parameter
  Dim adoprm1 As New ADODB.Parameter
  Dim adoprm2 As New ADODB.Parameter
  Dim adoPrm3 As New ADODB.Parameter
  Dim adoPrm4 As New ADODB.Parameter
  Dim adoPrm5 As New ADODB.Parameter
  Dim adoPrm6 As New ADODB.Parameter
  Dim adoPrm7 As New ADODB.Parameter
  Dim adoPrm8 As New ADODB.Parameter
  Dim adoPrm9 As New ADODB.Parameter
  Dim adoprm10 As New ADODB.Parameter


  
  DataOpt = False

     Set adoCmd = New ADODB.Command
     Set adoCmd.ActiveConnection = INIadoCon2
     adoCmd.CommandText = "uspcp_kcdb1"
     adoCmd.Parameters.Refresh
     adoCmd.CommandType = adCmdStoredProc
  
        Set adoPrmReturn = New ADODB.Parameter         '返回执行成功标记
        adoPrmReturn.type = adInteger
        adoPrmReturn.Direction = adParamReturnValue
        adoCmd.Parameters.Append adoPrmReturn

        
        Set adoprm1 = New ADODB.Parameter                 '4箱号
        adoprm1.type = adChar
        adoprm1.Size = 8000
        adoprm1.Direction = adParamInput
        adoprm1.Value = Trim(strXH)
        adoCmd.Parameters.Append adoprm1
        
        Set adoprm2 = New ADODB.Parameter                 '5lck
        adoprm2.type = adChar
        adoprm2.Size = 8000
        adoprm2.Direction = adParamInput
        adoprm2.Value = Trim(strLCK)
        adoCmd.Parameters.Append adoprm2
        
        Set adoPrm3 = New ADODB.Parameter                 '6gdh
        adoPrm3.type = adChar
        adoPrm3.Size = 8000
        adoPrm3.Direction = adParamInput
        adoPrm3.Value = Trim(strgdh)
        adoCmd.Parameters.Append adoPrm3
        
        Set adoPrm4 = New ADODB.Parameter                 '7lps
        adoPrm4.type = adChar
        adoPrm4.Size = 8000
        adoPrm4.Direction = adParamInput
        adoPrm4.Value = Trim(strlps)
        adoCmd.Parameters.Append adoPrm4
        
        Set adoPrm5 = New ADODB.Parameter               '8bls
        adoPrm5.type = adChar
        adoPrm5.Size = 8000
        adoPrm5.Direction = adParamInput
        adoPrm5.Value = Trim(strbls)                 '
        adoCmd.Parameters.Append adoPrm5
        
        Set adoPrm6 = New ADODB.Parameter               '9zcbls
        adoPrm6.type = adChar
        adoPrm6.Size = 8000
        adoPrm6.Direction = adParamInput
        adoPrm6.Value = Trim(strzcbls)
        adoCmd.Parameters.Append adoPrm6
        
        Set adoPrm7 = New ADODB.Parameter             '库房编号
        adoPrm7.type = adChar
        adoPrm7.Size = 20
        adoPrm7.Direction = adParamInput
        adoPrm7.Value = strmbkf
        adoCmd.Parameters.Append adoPrm7
        
        Set adoPrm8 = New ADODB.Parameter               '调拨人员
        adoPrm8.type = adChar
        adoPrm8.Size = 20
        adoPrm8.Direction = adParamInput
        adoPrm8.Value = Trim(strWholeName)
        adoCmd.Parameters.Append adoPrm8
        
        Set adoPrm9 = New ADODB.Parameter               '申请部门,存储过程中默认07
        adoPrm9.type = adChar
        adoPrm9.Size = 20
        adoPrm9.Direction = adParamInput
        adoPrm9.Value = strdepartment
        adoCmd.Parameters.Append adoPrm9
        
        
        
        Set adoprm10 = New ADODB.Parameter               '新箱号
        adoprm10.type = adChar
        adoprm10.Size = 8000
        adoprm10.Direction = adParamInput
        adoprm10.Value = Trim(newboxid)
        adoCmd.Parameters.Append adoprm10
        
     
     adoCmd.Execute
     Screen.MousePointer = 0
     If adoPrmReturn.Value = 0 Then
       MsgBox "已经成功执行您的任务！", vbInformation, Me.Caption
       DataOpt = True

     Else
        GoTo there
     End If
  Exit Function
there:
  Screen.MousePointer = 0
  MsgBox "执行失败！" + Chr(13) + "原因:" + Err.DESCRIPTION, vbExclamation, Me.Caption
End Function


Private Sub ChkAll_Click()

    Dim i As Integer
    
    With fpS_stockview
        If ChkAll.Value = 1 Then
            For i = 1 To .MaxRows
     
                .Row = i
                .Col = 1
                .text = 1
            Next i
        ElseIf ChkAll.Value = 0 Then
            For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                .text = 0
            Next i
        End If
    End With
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
        .Col = E_WAFERID.E_CHOOSE
        If .Value = 1 Then
            .Col = E_WAFERID.E_WAFERID
            If strWaferID = "" Then
                strWaferID = Trim$(.text)
                lQty2 = 1
            Else
                If strWaferID <> Trim$(.text) Then
                    strWaferID = Trim$(.text)
                    lQty2 = lQty2 + 1

                End If

            End If
            .Col = E_WAFERID.E_Passqty
            lQty = lQty + CLng(.text)

        End If

    Next

End With

lblQty.Caption = lQty
lblQtyPecs.Caption = lQty2

End Sub





