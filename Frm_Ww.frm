VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_ww 
   Caption         =   "ί��"
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
   StartUpPosition =   3  '����ȱʡ
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
         Caption         =   "��ѯ"
         Height          =   2100
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   18135
         Begin VB.TextBox Txt_sqdh 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "����Ʒ"
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
            Caption         =   "��ѯʱ������ѡ���Lot/����/�Ϻ�"
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "���뵥��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "�ۼ�Ƭ��(Wafer &PCS):"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��ǰ�ۼ�DIE��(DIE &PCS):"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��    ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "�ͻ�����"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "�� �� ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "������ַ"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "Ŀ��ⷿ"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "ҵ��ⷿ"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "�ͻ�����"
            Size            =   "1508;370"
            FontName        =   "����"
            FontHeight      =   210
            FontCharSet     =   134
            FontPitchAndFamily=   34
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(LOT)��"
            BeginProperty Font 
               Name            =   "����"
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
               Caption         =   "  ��  ѯ"
               Key             =   "Query"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ػ���ѯ"
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
               Caption         =   "  ί������ "
               Key             =   "Request"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ػ�����"
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
               Caption         =   "�ҵ�����"
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
               Caption         =   " ��������"
               Key             =   "CancerRequest"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Caption         =   "ί��ػ�"
               Key             =   "BackRequest"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "A004"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "������"
               Key             =   "WaitMove"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Caption         =   "ί����� "
               Key             =   "move"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Caption         =   "�ػ�����"
               Key             =   "A10"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "stockmove"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ί�⳷��"
               Key             =   "CancerStockMove"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "  ��   ��  "
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
            DialogTitle     =   "Excel����"
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
         TabCaption(0)   =   "����"
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
         TabCaption(1)   =   "����"
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
         TabCaption(2)   =   "����"
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
         TabCaption(3)   =   "�ػ�"
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
            Caption         =   "���ڵ���"
            Height          =   195
            Left            =   -64080
            TabIndex        =   32
            Top             =   450
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "ί��"
            Height          =   195
            Left            =   -62880
            TabIndex        =   31
            Top             =   450
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            Caption         =   "�ػ�"
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
            Caption         =   "����"
            Height          =   195
            Left            =   -61200
            TabIndex        =   28
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox ChkAll 
            Caption         =   "ȫѡ"
            Height          =   375
            Left            =   -75000
            TabIndex        =   27
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox ChkAll2 
            Caption         =   "ȫѡ"
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
            Caption         =   "��ѯ"
            Height          =   375
            Left            =   -72000
            TabIndex        =   23
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "��ѯ"
            Height          =   375
            Left            =   -71520
            TabIndex        =   22
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            Caption         =   "�������뵥"
            Height          =   255
            Left            =   -74760
            TabIndex        =   21
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Option6 
            Caption         =   "����������"
            Height          =   255
            Left            =   -73440
            TabIndex        =   20
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton Cmd_VT 
            Caption         =   "��ѯ�ػ�����"
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
            Caption         =   "��ǻ�ɫ�ı�ʾ��ί�����ѻػ�,�����ٴ�ί��"
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
            Caption         =   "����ʱ�䣺"
            Height          =   195
            Left            =   -71040
            TabIndex        =   45
            Top             =   480
            Width           =   900
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ʼʱ�䣺"
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
            Caption         =   "���뵥��"
            Height          =   255
            Left            =   -74040
            TabIndex        =   39
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "��������"
            Height          =   255
            Left            =   -74880
            TabIndex        =   38
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "��ֻ�ܲ�ѯ�Լ����������ĵ��ţ�"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   -72120
            TabIndex        =   37
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label Label5 
            Caption         =   "��������ͻ�����"
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

    E_CUSTCODE     '�ͻ�����
    E_CUSTPN       'KEHUJIZHONG
    E_qtechPTNo    'changneijizhong
    E_LOTID
    E_BOXID
    E_Matcode    '���ϱ��
    E_partno '�Ϻ�
    E_Matspec    '���
    E_Mattype    '�ͺ�
    E_UNIT    '��λ
    E_Passqty  '�ϸ���
    E_Ngqty1            '���ϲ�����
    E_Ngqty2             '�Ƴ̲�����
    e_ID      '���
    E_StockID      '�ֿ����
    E_END

End Enum

Enum E_WAFERID

    E_CHOOSE = 1
    E_LOTID
    E_BigBoxID
    E_BOXID
    E_WAFERID
    E_PN
    E_QTY  '����
    E_Passqty  '�ϸ���
    E_Ngqty1            '���ϲ�����
    E_Ngqty2             '�Ƴ̲�����
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
    E_GOOD_DIE '�ϸ���
    E_BAD1_DIE '���ϲ�����
    E_BAD2_DIE '�Ƴ̲�����
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
    E_GOOD_DIE '�ϸ���
    E_BAD_DIE '���ϲ�����
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
    adorst2.Source = "select distinct �ͻ�����  from tblXCustomer "
    adorst2.Open , , , , adCmdText
    Cmbcust.Clear
    If adorst2.RecordCount > 0 Then
      For i = 1 To adorst2.RecordCount
        Cmbcust.AddItem Trim(adorst2("�ͻ�����"))
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
        MsgBox "����ѡ��ͻ�����", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If SMR.State = adStateOpen Then SMR.Close

    
    strSql = "SELECT 0 AS ѡ��,SHIPDATE ,DELIVERYNO,CUSTLOT, GOODDIEQTY,NGDIEQTY,TTL,BATCH,REMARK,CUSTOMERSHORTNAME from TSV_VT_History WHERE CREATED_DATE<'" & Format(DTP2.Value + 1, "yyyy/mm/dd") & "' and  CREATED_DATE>'" & Format(DTP1.Value, "yyyy/mm/dd") & "' and CUSTOMERSHORTNAME='" & Trim(UCase(Cmbcust.text)) & "'"
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
        
        Toolbar1.Buttons("Request").Caption = "ί������"
        Toolbar1.Buttons("Request").Enabled = True
        Toolbar1.Buttons("Backrequest").Enabled = False
    Else
        If Kf_former <> "72" Then
            Toolbar1.Buttons("Request").Caption = "��������"
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
        Toolbar1.Buttons("Request").Caption = "ί������"
        Toolbar1.Buttons("Request").Enabled = True
        Toolbar1.Buttons("Backrequest").Enabled = False
    Else
        If Kf_former <> "72" Then
            Toolbar1.Buttons("Request").Caption = "��������"
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
   adoRstStocEntry.Source = "select �ⷿ����,�ⷿ���� from erpbase..tblstock  where �ֿ�����='��Ʒ��'  order by �ⷿ����"
   adoRstStocEntry.Open , , , , adCmdText
   If adoRstStocEntry.RecordCount > 0 Then
      Cob_kf_former.Clear
      adoRstStocEntry.MoveFirst
      For intNext = 1 To adoRstStocEntry.RecordCount
          Cob_kf_former.AddItem Trim(adoRstStocEntry("�ⷿ����")) & Space(1) & Trim(adoRstStocEntry("�ⷿ����"))
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
  adorst11.Source = "SELECT �ⷿ����+' '+�ⷿ���� �ֿ����� FROM erpbase..tblstock WHERE �ֿ�����='��Ʒ��'"
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
    MsgBox "��ѡ������ʽ", vbInformation, "��ʾ"
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
        strWholeName = gUserName & " ����"
        strdepartment = "07"
    Case "19809"
        strWholeName = gUserName & " ���޽�"
        strdepartment = "07"
    Case "19536"
        strWholeName = gUserName & " ����"
        strdepartment = "07"
    Case "07952"
        strWholeName = gUserName & " ������"
        strdepartment = "07"
    Case "10222"
        strWholeName = gUserName & " Ѧ��"
        strdepartment = "07"
    Case "12825"
        strWholeName = gUserName & " ���"
        strdepartment = "07"

    End Select
    Text1.text = Trim(strWholeName)
    'Cob_kf_former.Text ="07 ��˰��Ʒ��"
    'Cob_kf_dest.Text="72 WLAί���"

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
           MsgBox "�˱��ѻػ�������ί��", vbInformation, "��ʾ"

           Exit Sub
        End If
        If UCase(Trim(Cmbcust.text)) = "GC" Then
            'ͬһ��lot�ִ����������䣬һ������ֻ�ܳ�һ�������
            For i = 1 To .MaxRows
                If i <> Row Then
                    .Row = i
                    .Col = E_BOX.E_CHOOSE
                    strchoose = Trim(.text)
                    .Row = i
                    .Col = E_BOX.E_LOTID
                    strLotID_temp = Trim(.text)
                    If strchoose = "1" And strLotID_temp = strLotID Then
                         MsgBox "ͬһ��lot�ִ����������䣬һ������ֻ�ܳ�һ�������", vbInformation, "��ʾ"
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
        strBoxID = Trim$(.text) & "��"

        .Col = E_BOX.E_StockID
        strStockID = Trim$(.text)



        
        
       If Get_SqlserverCnt(" SELECT * FROM erptemp..tblstockdbsub_temp a,  erptemp..tblstockdb_temp b where a.remark1='" & Replace(strBoxID, "��", "") & "' and b.flag=1 and a.ORDER_NUM=b.ORDER_NUM and a.ITEM=B.ITEM") > 0 Then
           MsgBox "���" & strBoxID & "�����ί�����룬�����ظ�����", vbInformation, "��ʾ"
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

    '�����ѡ��ĵ��Ŷ�ѡ��
    If Row < 1 Then Exit Sub
    If Col <> 1 Then Exit Sub
    With fpS_stockview

        .Col = 1
        .Row = Row
        .Value = Abs(Val(.Value) - 1)
'        strDJBH = ""
        If Val(.Value) = 1 Then
            '������һ���ĵ���+��ŵ�ѡ����
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
            '������һ���ĵ���+��ŵ�ѡ����
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
Label1.Caption = "���뵥��"
End Sub




Private Sub Option6_Click()
Label1.Caption = "��������"
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
    strSql = "select 0 as 'ѡ��',a.ORDER_NUM as ���뵥��,a.ITEM as ���,d.CUSTOMERSHORTNAME as �ͻ�����,d.MPN_DESC as �ͻ�����,a.FORMER as ԭ�ֿ�,a.DESTINATION as Ŀ��ֿ�,b.LOT as ������,b.REMARK1 as �����,b.QBOX as С���, " & _
             " b.WAFER as ���̿����,b.GOOD_DIE as �ϸ���,b.BAD1_DIE as ���ϲ����� ,b.BAD2_DIE as �Ƴ̲�����,b.id as id from erptemp..tblstockdb_temp a " & _
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
    If Label1.Caption = "��������" Then
       If Trim(TxtdbNo.text) = "" Then
            MsgBox "�������������", vbInformation, "��ʾ"
            Exit Sub
       End If
       
    
        strSql = "select 1 as 'ѡ��',a.�������,a.���,d.CUSTOMERSHORTNAME as �ͻ�����,d.MPN_DESC as �ͻ�����,a.ԭ�ֿ�,a.Ŀ��ֿ�,b.������,'' as remark,b.���, " & _
                 " b.���̿����,b.�ϸ���,b.���ϲ����� ,b.�Ƴ̲�����,b.id as id from erpdata..tblstockdb a " & _
                 " left join erpdata..tblstockdbsub b  on a.�������=b.������� and a.���=b.��� " & _
                 " left join erpbase..tblmappingdata c on c.SUBSTRATEID=b.���̿���� and c.LOTID=b.������ " & _
                 " left join erpbase..tblcustomeroi d on convert(varchar(20)  ,d.id)=c.FILENAME and c.LOTID= d.SOURCE_BATCH_ID  "

        If Trim(TxtdbNo.text) <> "" Then
            strSql = strSql & " where  a.�������='" & Trim(TxtdbNo.text) & "'"
            If gUserName <> "07885" Then strSql = strSql & " and  a.������Ա='" & strWholeName & " '"
        Else
            If gUserName <> "07885" Then strSql = strSql & " where a.������Ա='" & strWholeName & " '"
        End If
        Toolbar1.Buttons("CancerStockMove").Enabled = True
    ElseIf Label1.Caption = "���뵥��" Then
        strSql = "select 1 as 'ѡ��',a.ORDER_NUM as ���뵥��,a.ITEM as ���,d.CUSTOMERSHORTNAME as �ͻ�����,d.MPN_DESC as �ͻ�����,a.FORMER as ԭ�ֿ�,a.DESTINATION as Ŀ��ֿ�,b.LOT as ������,'' as remark,b.QBOX as С���, " & _
                 " b.WAFER as ���̿����,b.GOOD_DIE as �ϸ���,b.BAD1_DIE as ���ϲ����� ,b.BAD2_DIE as �Ƴ̲�����,b.id as id from erptemp..tblstockdb_temp a " & _
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
  
    Case "Request"  '����
         If Trim(Cmbcust.text) = "" Then
             MsgBox "��ѡ��ͻ�����!", vbInformation, "��ʾ"
             Exit Sub
         End If
         If Trim(Cob_Shipto.text) = "" Then
             MsgBox "��ѡ�������ַ!", vbInformation, "��ʾ"
             Exit Sub
         End If
        Toolbar1.Buttons("Request").Enabled = False
        CreateApplication ("WW")
        Toolbar1.Buttons("Request").Enabled = True
    Case "Backrequest"  '�ػ�����
          
        CreateApplication ("VT")
        
    Case "ViewMyRequest"  '
         Option1.Visible = False
         Option2.Visible = False
         Option3.Visible = False
         Option4.Visible = False
        
        ListMyRequest
        
         
    Case "CancerRequest"  '��������
        Toolbar1.Buttons("CancerRequest").Enabled = False
        CancerRequest
           
    Case "WaitMove"  '������
         Option1.Visible = True
         Option2.Visible = True
         Option3.Visible = True
         Option4.Visible = True
         
         ListStockView
         
    Case "stockmove"   '����
        Toolbar1.Buttons("stockmove").Enabled = False
        stockmove
        
    Case "Query" '��ѯ
    
         If Trim(Cmbcust.text) = "" Then
             MsgBox "��ѡ��ͻ�����!", vbInformation, "��ʾ"
             Exit Sub
         End If
        
         Call ListView1Data("WW")
    Case "Query_VT" '��ѯ
    
         Cob_kf_former.text = "72 WLAί���"

         Call ListView1Data("VT")
         
     Case "CancerStockMove"   '��������
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
    strSql = "select 0 as 'ѡ��',a.ORDER_NUM as ���뵥��,a.ITEM as ���,d.CUSTOMERSHORTNAME as �ͻ�����,d.MPN_DESC as �ͻ�����,a.FORMER as ԭ�ֿ�,a.DESTINATION as Ŀ��ֿ�,b.LOT as ������,b.REMARK1 as �����,b.QBOX as С���, " & _
             " b.WAFER as ���̿����,b.GOOD_DIE as �ϸ���,b.BAD1_DIE as ���ϲ����� ,b.BAD2_DIE as �Ƴ̲�����,b.id as id from erptemp..tblstockdb_temp a " & _
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
If Label1.Caption <> "���뵥��" Then
    Exit Sub
End If
If TxtdbNo.text = "" Then
   MsgBox "�밴���뵥�Ų�ѯ", vbInformation, "��ʾ"
   Exit Sub
End If

 With fpS_Cancerview
    If .MaxRows <= 0 Then
        MsgBox "���Ȳ�ѯ", vbInformation, "��ʾ"
        Exit Sub
    End If
    For i = 1 To .MaxRows
   
       .Row = i
       .Col = E_StockView.E_CHOOSE
       
       If .text = 1 Then
           SumCount = SumCount + 1
       End If
    Next i
    If MsgBox("��ȷ��Ҫȡ�����뵥" & TxtRequestNo & ",��" & SumCount & "�ʼ�¼��?", vbOKCancel, "��ʾ") = vbCancel Then
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
   MsgBox SumCount & "�������¼�����ɹ�"

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
       ' MsgBox "��ѡ��ҵ��ⷿ��", vbInformation, Me.Caption
       ' Exit Sub
    ' End If
        
    'Kf_former = Left(Trim(Cob_kf_former.Text), InStr(Trim(Cob_kf_former.Text), " ") - 1)
'��������
    ' If Trim(Cob_kf_dest.Text) = Trim(Cob_kf_former.Text) Then
       ' MsgBox "ҵ��ⷿ��Ŀ��ⷿ��ͬ������ʧ�ܣ�", vbInformation, Me.Caption
       ' Exit Sub
    ' End If
    ' If Apptype <> "VT" Then
        ' If Trim(Cob_kf_dest.Text) = "" Then
           ' MsgBox "��ѡ��Ŀ��ⷿ��", vbInformation, Me.Caption
           ' Exit Sub
        ' End If
         ' Kf_dest = Left(Trim(Cob_kf_dest.Text), InStr(Trim(Cob_kf_dest.Text), " ") - 1)
    
        ' If Left(Trim(Cob_kf_dest.Text), InStr(Trim(Cob_kf_dest.Text), " ") - 1) <> "72" And Left(Trim(Cob_kf_former.Text), InStr(Trim(Cob_kf_former.Text), " ") - 1) <> "72" Then
            ' strbond = "SELECT COUNT(*) FROM erpdata..tblstock a,erpdata..tblstock b WHERE a.�ⷿ���� = '" & Left(Trim(Cob_kf_dest.Text), InStr(Trim(Cob_kf_dest.Text), " ") - 1) & "' AND b.�ⷿ���� = '" & Left(Trim(Cob_kf_former.Text), InStr(Trim(Cob_kf_former.Text), " ") - 1) & "' AND b.�ⷿ���� = a.�ⷿ����"
        
             ' If SMR.State = adStateOpen Then SMR.Close
             ' SMR.Open strbond, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
             ' If SMR.Fields(0).Value = 0 Then
                ' MsgBox "ҵ��ⷿ��Ŀ��ⷿ���Ͳ�ͬ��", vbInformation, Me.Caption
                ' SMR.Close
                ' Exit Sub
             ' End If
             
         ' End If
        ' If Kf_former = "72" Then
           ' MsgBox "��ί��ػ������ɴ�72�ֵ�����", vbInformation, Me.Caption
           ' Exit Sub
        ' End If
        ' If Kf_dest = "72" Then
           ' If Trim(Cob_Shipto.Text) = "" Then
               ' MsgBox "ί���������ѡ�񷢻���ַ��", vbInformation, Me.Caption
               ' Exit Sub
           ' End If
        ' End If
         

    ' Else
        ' If Kf_former <> "72" Then
               ' MsgBox "ί�ص�������ѡ��72�֣�", vbInformation, Me.Caption
               ' Exit Sub
        ' End If
    ' End If


   Txt_sqdh.text = ""
    
    
    
    With fpS_wafer
        If .MaxRows <= 0 Then
            MsgBox "��ѡ��Ҫ�����ļ�¼!", vbInformation, "��ʾ"
            Exit Sub
        End If
    End With
    
    strbigbox_sel = ""
    strlot_sel = "" '�����#lot
    'merry20200202ͬһ�������ж��lot,���ֿܷ�����
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
            strSql = " select distinct rtrim(c.���) + '#' + rtrim(������) as ������ from erpdata..tblPackMainInfsub  a " & _
                   " inner join erpdata..tblPacktreeinf b on a.���=b.��� " & _
                   " inner join erpdata..tblPacktreeinf c on b.�ϼ����=c.���  " & _
                   " where c.��� = '" & Split(strbigbox_sel, ",")(i) & "' "
            
            If SMR.State = adStateOpen Then SMR.Close
            SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            If SMR.RecordCount > 0 Then
                SMR.MoveFirst
                For j = 1 To SMR.RecordCount
                    strlot_db = SMR("������")
                    If InStr(strlot_sel, strlot_db) = 0 Then
                        MsgBox "�����" & Split(strbigbox_sel, ",")(i) & " �л�������lot,���ɲ������", vbInformation, "��ʾ"
                        Exit Sub
                    End If
                    SMR.MoveNext
                Next
            End If
            SMR.Close
            Set SMR = Nothing
        Next

    End With

    
     '�������뵥��

    RequestNo = GetID()
    intitem = 0
    SumCount = 0

    With fpS_wafer
        If .MaxRows <= 0 Then
            MsgBox "��ѡ��Ҫ�����ļ�¼!", vbInformation, "��ʾ"
            Exit Sub
        End If
    
        For intnum = 1 To .MaxRows
            .Row = intnum
            .Col = E_WAFERID.E_CHOOSE
            If .text <> "" Then
                If .text = 1 Then
                    .Col = E_WAFERID.E_BOXID
                    strXH = Trim(.text)      '���
                    .Col = E_WAFERID.E_WAFERID
                    strLCK = Trim(.text)     '���̿����
                    If Get_SqlserverCnt(" SELECT * FROM erptemp..tblstockdbsub_temp a ,erptemp..tblstockdb_temp b  where b.flag=1 and a.ORDER_NUM=b.ORDER_NUM and a.ITEM=b.ITEM  and a.qbox='" & strXH & "' and RTRIM(a.WAFER)='" & strLCK & "'") > 0 Then
                        MsgBox "���" & strXH & "�����̿����" & strLCK & " ��������������ظ�����", vbInformation, "��ʾ"
                        Exit Sub
                    End If
                    If Apptype = "WW" Then
                        If Get_SqlserverCnt(" SELECT * FROM erpdata..tblstockdbsub a ,erpdata..tblstockdb b  where  a.�������=b.������� and a.���=b.���  and a.���='" & strXH & "' and RTRIM(a.���̿����)='" & strLCK & "' and a.������� not in (select ������� from erptemp..invalidstockdb) and a.�������  not in (select ����������� from erptemp..invalidstockdb)") > 0 Then
                            MsgBox "���" & strXH & "�����̿����" & strLCK & " �ѻػ��������ٴ�����ί��", vbInformation, "��ʾ"
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
                    strXH = Trim(.text)      '���
                    .Col = E_WAFERID.E_BigBoxID
                    strxh_big = Trim(.text)      '�����
                    
                    .Col = E_WAFERID.E_LOTID
                    strgdh = Trim(.text)      '������
                    
                    .Col = E_WAFERID.E_WAFERID
                    strLCK = Trim(.text)     '���̿����
                    
                    .Col = E_WAFERID.E_QTY
                    intqty = Val(.text)        '����
                    .Col = E_WAFERID.E_Passqty
                    strlps = Trim(.text)        '��Ʒ��
                    .Col = E_WAFERID.E_Ngqty1
                    strbls = Trim(.text)        '����Ʒ��
                    .Col = E_WAFERID.E_Ngqty2
                    strzcbls = Trim(.text)      '�Ƴ̲�����
                    .Col = E_WAFERID.e_ID
                    strid = Trim(.text)
                    If Get_SqlserverCnt("select * from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'and  id=" & strid) > 0 Then
                        strSql = "select ITEM from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'and  id=" & strid
                        intitem = GetSqlServerStr(strSql)
                    Else
                        strSql = "select isnull(max(ITEM),0) from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'"
                        intitem = GetSqlServerStr(strSql) + 1
   
                        If SMR.State = adStateOpen Then SMR.Close
                        strSql = " select �ⷿ���,���ϱ��,�ͻ�����,isnull(����,0) from erpdata..tblStockNum where id=" & strid
                        SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
                        If SMR.RecordCount = 1 Then
                            SMR.MoveFirst
                            strKF = Trim(SMR("�ⷿ���"))
                            strmatcode = Trim(SMR("���ϱ��"))
                            strCustCode = Trim(SMR("�ͻ�����"))
                        
                        End If
                        If Apptype = "VT" Then
                            strSql = " select top 1 rtrim(a.ԭ�ֿ�) from erpdata..tblStockdb a,erpdata..tblStockdbsub b   where rtrim(b.���̿����)='" & strLCK & "' and a.�������=b.������� and a.���=b.��� and a.Ŀ��ֿ�='72'"
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
                        
                      '�ϴ�����

                     '�������,���,���ϱ��, ��������,ԭ�ֿ�,Ŀ��ֿ�,������Ա,����ʱ��,�����Ա,���ʱ��, ���벿��,״̬,REMARK1,REMARK2,REMARK3,REMARK4,REMARK5,ID
                        strSql = "insert into erptemp..tblstockdb_temp(ORDER_NUM,ITEM, MATERIALS,QTY,FORMER, DESTINATION, APPLICANT, APPLICATION_TIME, AUDITOR, AUDIT_TIME, DEPT, FLAG,ID,REMARK1) values( " & _
                        "'" & RequestNo & "'," & intitem & ",'" & strmatcode & "'," & 0 & ",'" & strKF & "','" & Kf_dest & "','" & strWholeName & "',sysdatetime(),'','','',1," & strid & ",'" & Trim(Cob_Shipto.text) & "')"
                    
                        AddSql2 (strSql)
                       
                        
                    End If
                    
                    '�ϴ��ӱ�
                    
                    '�������, ���, ���, ���̿����, ������, �ϸ���, �Ƴ̲�����, ���ϲ�����, ID
                     strSql = "insert into erptemp..tblstockdbsub_temp(ORDER_NUM,ITEM,WAFER,LOT,GOOD_DIE,BAD1_DIE,BAD2_DIE,ID,REMARK1,QBOX) values( " & _
                    "'" & RequestNo & "'," & intitem & ",'" & strLCK & "','" & strgdh & "'," & strlps & "," & strbls & "," & strzcbls & "," & strid & ",'" & strxh_big & "','" & strXH & "')"
                  
                    AddSql2 (strSql)
                    
                    'update��������
                    strSql = "Update erptemp..tblstockdb_temp set QTY =QTY+" & Val(strlps) + Val(strbls) + Val(strzcbls) & " where ORDER_NUM='" & RequestNo & "' and ITEM=" & intitem
                   
                    AddSql2 (strSql)
                    SumCount = SumCount + 1
                    
                    
                    
                End If
            End If
        Next intnum

    End With
    If SumCount > 0 Then
        MsgBox SumCount & "�ʼ�¼����ɹ�", vbInformation, "��ʾ"
        Txt_sqdh.text = RequestNo
    End If
     
End Sub


     
Function GetID()
'FWW1911140011
'���ɷ�ʽ��FWW+YYMMDD +4λ��ˮ��
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
        '���ж�Lot���Ƿ��Ѿ�����
        For i = 1 To .MaxRows
            .Row = i
            .Col = E_BOX.E_LOTID
            If Trim(TxtLot.text) = Trim(.text) Then
                MsgBox "��Lot�Ѿ���ѯ������Ҫ�ظ���ѯ", vbInformation, "��ʾ"
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

    strSql = "SELECT distinct 0 as '��',a.�ͻ�����,i.MPN_DESC as �ͻ����� , g.QTECHPTNO as ���ڻ���,a.������,dbo.f_getparent(f.���)  as ����� ,a.���ϱ��,a.�Ϻ�, b.���,b.�ͺ�,b.������λ����,a.�ϸ���,a.������ AS ������,a.�Ƴ̲����� , a.id,c.�ⷿ����" & _
    " FROM  erpdata..tblStockNum AS a " & _
    " INNER JOIN  erpbase..tblSmainM2 AS b ON a.���ϱ�� = b.���ϱ��  " & _
    " INNER JOIN  erpbase..tblstock AS c ON a.�ⷿ��� = c.�ⷿ����  " & _
    " INNER JOIN  erpdata..tblbase d on a.���߱��=d.���� and d.˵��2='���߱��'  " & _
    " LEFT JOIN erpdata..tblWithWork e ON a.�������=e.������� AND a.�������=e.�������    " & _
    " LEFT JOIN  erpdata..tblStockNumsub f on  f.id=a.id  " & _
    " LEFT JOIN  erptemp..tbltsvnpiproduct g ON g.QTECHPTNO2=f.�Ϻ�   " & _
     " left join erpbase..tblmappingdata h on h.SUBSTRATEID=f.���̿���� and h.LOTID=f.������ " & _
     " left join erpbase..tblcustomeroi i on convert(varchar(20)  ,i.id)=h.FILENAME and h.LOTID= i.SOURCE_BATCH_ID " & _
    " where a.�ϸ���+a.������+a.�Ƴ̲�����>0 "

    'If Kf_former <> "" Then strSql = strSql & " and  a.�ⷿ���='" & Kf_former & "'"

    If Trim(Cmbcust.text) <> "" Then strSql = strSql & " and  a.�ͻ�����='" & Trim(Cmbcust.text) & "'"
    If Trim(TxtCustpn.text) <> "" Then strSql = strSql & " and  i.MPN_DESC='" & Trim(TxtCustpn.text) & "'"
    If Trim(TxtPN.text) <> "" Then strSql = strSql & " and  a.�Ϻ�='" & Trim(TxtPN.text) & "'"
    If Trim(TxtLot.text) <> "" Then strSql = strSql & " and  a.������='" & Trim(TxtLot.text) & "'"
    If searchtype = "VT" Then
     strSql = strSql & " AND a.�ⷿ��� IN ('72')"
     'strSql = strSql & "  and  a.������ in (select distinct CUSTLOT from erptemp..TSV_VT_History_sub  where  flag=1  and CUSTOMERSHORTNAME='" & Trim(Cmbcust.Text) & "') "
     
    ElseIf searchtype = "WW" Then
        If Chk_NG.Value = 1 Then
            strSql = strSql & " AND a.�ⷿ��� IN ('28','30')"
        Else
            strSql = strSql & " AND a.�ⷿ��� IN ('07','16','19','20')"
        End If
        'If Trim(TxtLot.Text) <> "" Then strSql = strSql & " and  a.������='" & Trim(TxtLot.Text) & "'"
    Else
        If Trim(Cob_kf_former.text) <> "" Then
            Kf_former = Left(Cob_kf_former.text, InStr(Cob_kf_former.text, " ") - 1)
            If Kf_former <> "" Then strSql = strSql & " and  a.�ⷿ���='" & Kf_former & "'"
        End If
    End If

    
    If gdh = "" Then

    Else
        strSql = strSql & " GROUP BY  a.id, a.�ͻ�����, a.������, a.���ϱ��, a.�Ϻ�, b.���,  b.�ͺ�,  b.������λ���� ,dbo.f_getparent(d.���)"
    End If
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
 
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        If Chk_Keepdata.Value = 1 Then
        '��ѯʱ������ѡ���Lot��
            With fpS_Box
               For i = 1 To SMR.RecordCount
                   .MaxRows = .MaxRows + 1
                   .SetText E_BOX.E_CHOOSE, .MaxRows, 0
                   .SetText E_BOX.E_CUSTCODE, .MaxRows, SMR("�ͻ�����")
                   .SetText E_BOX.E_CUSTPN, .MaxRows, SMR("�ͻ�����")
                   .SetText E_BOX.E_qtechPTNo, .MaxRows, SMR("���ڻ���")
                   .SetText E_BOX.E_LOTID, .MaxRows, SMR("������")
                   .SetText E_BOX.E_BOXID, .MaxRows, SMR("�����")
                   .SetText E_BOX.E_Matcode, .MaxRows, SMR("���ϱ��")
                   .SetText E_BOX.E_partno, .MaxRows, SMR("�Ϻ�")
                   .SetText E_BOX.E_Matspec, .MaxRows, SMR("���")
                   .SetText E_BOX.E_Mattype, .MaxRows, SMR("�ͺ�")
                   .SetText E_BOX.E_UNIT, .MaxRows, SMR("������λ����")
                   .SetText E_BOX.E_Passqty, .MaxRows, SMR("�ϸ���")
                   .SetText E_BOX.E_Ngqty1, .MaxRows, SMR("������")
                   .SetText E_BOX.E_Ngqty2, .MaxRows, SMR("�Ƴ̲�����")
                   .SetText E_BOX.e_ID, .MaxRows, SMR("ID")
                   .SetText E_BOX.E_StockID, .MaxRows, SMR("�ⷿ����")

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
               
                strSql = " select d.��� from erpdata..tblStockdbsub a " & _
                    " inner join erpdata..tblStockdb  b on a.�������=b.������� and a.���=b.��� " & _
                    " inner join erpdata..tblStockNumTree  c on a.���=c.��� " & _
                    " inner join erpdata..tblStockNumTree  d on c.�ϼ����=d.��� " & _
                    " where b.Ŀ��ֿ�='72' and exists( select * from erpdata..tblStockNumSub where ���̿����=a.���̿����) " & _
                    " and not exists( select * from erptemp..InvalidStockDb  where �����������=a.�������) " & _
                    " and rtrim(d.���)='" & Trim(.text) & "'"
                    
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

    If intBJ = 1 Then '��ѡ

        With fpS_wafer
           If .MaxRows = 0 Then
                   '��ѯ����
                Set adorst1 = New ADODB.Recordset
                Set adorst1.ActiveConnection = INIadoCon2
                
            adorst1.Source = "SELECT a.id,a.���, a.������,a.���̿����,a.�Ϻ�, a.���ϱ��, sum(a.����) as ����, case when a.�ϸ���=0 then sum(a.����) else 0 end as �ϸ�Ʒ, case when a.�ϸ���=2 then sum(a.����) else 0 end  as ����Ʒ,case when a.�ϸ���=1 then sum(a.����) else 0 end  as �Ƴ̲���Ʒ,a.������� ,'' ����� FROM  dbo.tblStockNumSub AS a INNER JOIN dbo.f_kcdb('" & strBoxID & "') AS b ON a.��� = b.��� INNER JOIN erpdata..tblstocknum c ON c.ID = a.ID " & _
              "  where charindex(rtrim(a.������),'" & strLotID & " ')>0 and  a.����>0 and c.�ⷿ��� = '" & kf & "' group by  a.id,a.���,a.�Ϻ�,a.���ϱ��,a.�ϸ���,a.������� ,a.������,a.���̿����  " & _
              " union " & _
              "  SELECT a.id,a.���, a.������,a.���̿����, a.�Ϻ�, a.���ϱ��, sum(a.����) as ����, case when a.�ϸ���=0 then sum(a.����) else 0 end as �ϸ�Ʒ, case when a.�ϸ���=2 then sum(a.����) else 0 end as ����Ʒ,case when a.�ϸ���=1 then sum(a.����) else 0 end  as �Ƴ̲���Ʒ,a.������� ,'' �����  FROM  dbo.tblStockNumSub AS a INNER JOIN tblStockNumtree AS b ON a.��� = b.��� INNER JOIN erpdata..tblstocknum c ON c.ID = a.ID " & _
              "  where  rtrim(a.������)='" & Replace(strLotID, "$", "") & "' and  a.����>0 and rtrim(a.���)='" & Replace(Trim(strBoxID), "��", "") & "' and c.�ⷿ��� = '" & kf & "'  group by  a.id,a.���,a.�Ϻ�,a.���ϱ��,a.�ϸ���,a.�������, a.������,a.���̿���� order by a.���̿���� "
                            
                adorst1.Open , , adOpenStatic, adLockReadOnly, adCmdText
               
            
                If adorst1.RecordCount > 0 Then
                    adorst1.MoveFirst

                    For j = 1 To adorst1.RecordCount

                        .MaxRows = .MaxRows + 1
                        
                        .SetText E_WAFERID.E_CHOOSE, .MaxRows, 1
                        .SetText E_WAFERID.E_LOTID, .MaxRows, Trim$("" & adorst1!������)
                        .SetText E_WAFERID.E_BigBoxID, .MaxRows, Replace(strBoxID, "��", "")
                        .SetText E_WAFERID.E_BOXID, .MaxRows, Trim$("" & adorst1!���)
                        .SetText E_WAFERID.E_WAFERID, .MaxRows, Trim$("" & adorst1!���̿����)
                        .SetText E_WAFERID.E_PN, .MaxRows, Trim$("" & adorst1!�Ϻ�)
                        .SetText E_WAFERID.E_QTY, .MaxRows, Trim$("" & adorst1!����)
                        
                        .SetText E_WAFERID.E_Passqty, .MaxRows, Trim$("" & adorst1!�ϸ�Ʒ)
                        .SetText E_WAFERID.E_Passqty, .MaxRows, Trim$("" & adorst1!�ϸ�Ʒ)
                        .SetText E_WAFERID.E_Ngqty1, .MaxRows, Trim$("" & adorst1!����Ʒ)
                        .SetText E_WAFERID.E_Ngqty2, .MaxRows, Trim$("" & adorst1!�Ƴ̲���Ʒ)
                        
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

                    If Replace(strBoxID, "��", "") = Trim(Box_temp) And Replace(strLotID, "$", "") = Lot_temp Then
                        Exit Sub
                    End If

                Next

                   '��ѯ����
                Set adorst1 = New ADODB.Recordset
                Set adorst1.ActiveConnection = INIadoCon2
       ' adorst1.Source = "SELECT a.id,a.���, a.������,a.���̿����,a.�Ϻ�, a.���ϱ��, sum(a.����) as ����,case when a.�ϸ���=0 then sum(a.����) else 0 end as �ϸ�Ʒ, case when a.�ϸ���=2 then sum(a.����) else 0 end  as ����Ʒ,case when a.�ϸ���=1 then sum(a.����) else 0 end  as �Ƴ̲���Ʒ,a.������� ,'' ����� FROM  dbo.tblStockNumSub AS a INNER JOIN dbo.f_kcdb('" & strBoxID & "') AS b ON a.��� = b.��� INNER JOIN erpdata..tblstocknum c ON c.ID = a.ID " & _
        '  "  where charindex(rtrim(a.������),'" & strLotID & " ')>0 and  a.����>0 and c.�ⷿ��� = '" & kf & "' group by  a.id,a.���,a.�Ϻ�,a.���ϱ��,a.�ϸ���,a.������� ,a.������,a.���̿���� " & _
        '  " union " & _
        '  "  SELECT a.id,a.���, a.������,a.���̿����, a.�Ϻ�, a.���ϱ��,  sum(a.����) as ����, case when a.�ϸ���=0 then sum(a.����) else 0 end as �ϸ�Ʒ, case when a.�ϸ���=2 then sum(a.����) else 0 end as ����Ʒ,case when a.�ϸ���=1 then sum(a.����) else 0 end  as �Ƴ̲���Ʒ,a.������� ,'' �����  FROM  dbo.tblStockNumSub AS a INNER JOIN tblStockNumtree AS b ON a.��� = b.��� INNER JOIN erpdata..tblstocknum c ON c.ID = a.ID " & _
        '  "  where  charindex(rtrim(a.������),'" & strLotID & " ')>0 and  a.����>0 and charindex(rtrim(a.���),'" & Trim(strBoxID) & " ')>0 and c.�ⷿ��� = '" & kf & "'  group by  a.id,a.���,a.�Ϻ�,a.���ϱ��,a.�ϸ���,a.�������, a.������,a.���̿����"
            
            adorst1.Source = "SELECT a.id,a.���, a.������,a.���̿����,a.�Ϻ�, a.���ϱ��, sum(a.����) as ����, case when a.�ϸ���=0 then sum(a.����) else 0 end as �ϸ�Ʒ, case when a.�ϸ���=2 then sum(a.����) else 0 end  as ����Ʒ,case when a.�ϸ���=1 then sum(a.����) else 0 end  as �Ƴ̲���Ʒ,a.������� ,'' ����� FROM  dbo.tblStockNumSub AS a INNER JOIN dbo.f_kcdb('" & strBoxID & "') AS b ON a.��� = b.��� INNER JOIN erpdata..tblstocknum c ON c.ID = a.ID " & _
              "  where charindex(rtrim(a.������),'" & strLotID & " ')>0 and  a.����>0 and c.�ⷿ��� = '" & kf & "' group by  a.id,a.���,a.�Ϻ�,a.���ϱ��,a.�ϸ���,a.������� ,a.������,a.���̿����  " & _
              " union " & _
              "  SELECT a.id,a.���, a.������,a.���̿����, a.�Ϻ�, a.���ϱ��, sum(a.����) as ����, case when a.�ϸ���=0 then sum(a.����) else 0 end as �ϸ�Ʒ, case when a.�ϸ���=2 then sum(a.����) else 0 end as ����Ʒ,case when a.�ϸ���=1 then sum(a.����) else 0 end  as �Ƴ̲���Ʒ,a.������� ,'' �����  FROM  dbo.tblStockNumSub AS a INNER JOIN tblStockNumtree AS b ON a.��� = b.��� INNER JOIN erpdata..tblstocknum c ON c.ID = a.ID " & _
              "  where  rtrim(a.������)='" & Replace(strLotID, "$", "") & "' and  a.����>0 and rtrim(a.���)='" & Replace(Trim(strBoxID), "��", "") & "' and c.�ⷿ��� = '" & kf & "'  group by  a.id,a.���,a.�Ϻ�,a.���ϱ��,a.�ϸ���,a.�������, a.������,a.���̿���� order by a.���̿���� "
                                          
            adorst1.Open , , adOpenStatic, adLockReadOnly, adCmdText
          
                If adorst1.RecordCount > 0 Then
                    adorst1.MoveFirst

                    For j = 1 To adorst1.RecordCount

                        .MaxRows = .MaxRows + 1
                        .SetText E_WAFERID.E_CHOOSE, .MaxRows, 1
                        .SetText E_WAFERID.E_LOTID, .MaxRows, Trim$("" & adorst1!������)
                        .SetText E_WAFERID.E_BigBoxID, .MaxRows, Replace(strBoxID, "��", "")
                        .SetText E_WAFERID.E_BOXID, .MaxRows, Trim$("" & adorst1!���)
                        .SetText E_WAFERID.E_WAFERID, .MaxRows, Trim$("" & adorst1!���̿����)
                        .SetText E_WAFERID.E_PN, .MaxRows, Trim$("" & adorst1!�Ϻ�)
                        .SetText E_WAFERID.E_QTY, .MaxRows, Trim$("" & adorst1!����)
                        .SetText E_WAFERID.E_Passqty, .MaxRows, Trim$("" & adorst1!�ϸ�Ʒ)
                        .SetText E_WAFERID.E_Ngqty1, .MaxRows, Trim$("" & adorst1!����Ʒ)
                        .SetText E_WAFERID.E_Ngqty2, .MaxRows, Trim$("" & adorst1!�Ƴ̲���Ʒ)
                        .SetText E_WAFERID.e_ID, .MaxRows, Trim$("" & adorst1!id)

                        adorst1.MoveNext
                    Next
        
                End If



            End If

        End With

    End If

    If intBJ = 2 Then 'ȡ����ѡ

        With fpS_wafer

            For i = .MaxRows To 1 Step -1
                    .Row = i
                    .Col = E_WAFERID.E_BigBoxID
                    Box_temp = Trim$(.text)
                    .Row = i
                    .Col = E_WAFERID.E_LOTID
                    Lot_temp = Trim$(.text)

                If Replace(strBoxID, "��", "") = Trim(Box_temp) And Replace(strLotID, "$", "") = Lot_temp Then
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
        '�趨������
        .Col = E_BOX.E_CHOOSE   'ѡ��
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        '�趨�п�
        .ColWidth(-1) = 10
        .ColWidth(E_BOX.E_CHOOSE) = 4
        .ColWidth(E_BOX.E_CUSTCODE) = 6
        .RowHeight(-1) = 10
        '�趨�Ƿ�����
        .UserColAction = UserColActionSort

        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
  

         .SetText E_BOX.E_CHOOSE, 0, "��"
         .SetText E_BOX.E_CUSTCODE, 0, "�ͻ�"
         .SetText E_BOX.E_qtechPTNo, 0, "���ڻ���"
         .SetText E_BOX.E_LOTID, 0, "������"
         .SetText E_BOX.E_BOXID, 0, "�����"
         .SetText E_BOX.E_Matcode, 0, "���ϱ��"
         .SetText E_BOX.E_partno, 0, "�Ϻ�"
         .SetText E_BOX.E_Matspec, 0, "���"
         .SetText E_BOX.E_Mattype, 0, "�ͺ�"
         .SetText E_BOX.E_UNIT, 0, "������λ����"
         .SetText E_BOX.E_Passqty, 0, "�ϸ���"
         .SetText E_BOX.E_Ngqty1, 0, "������"
         .SetText E_BOX.E_Ngqty2, 0, "�Ƴ̲�����"
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
        '�趨������
        .Col = E_WAFERID.E_CHOOSE   'ѡ��
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        '�趨�п�
        .ColWidth(-1) = 10
        .ColWidth(E_WAFERID.E_CHOOSE) = 4
        .RowHeight(-1) = 10
        '�趨�Ƿ�����
        .UserColAction = UserColActionSort

        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next

    
        .SetText E_WAFERID.E_CHOOSE, 0, "��"
        .SetText E_WAFERID.E_LOTID, 0, "������"
        .SetText E_WAFERID.E_BigBoxID, 0, "�����"
        .SetText E_WAFERID.E_BOXID, 0, "���"
        .SetText E_WAFERID.E_WAFERID, 0, "���̿����"
        .SetText E_WAFERID.E_PN, 0, "�Ϻ�"
        .SetText E_WAFERID.E_QTY, 0, "����"
        .SetText E_WAFERID.E_Passqty, 0, "�ϸ���"
        .SetText E_WAFERID.E_Ngqty1, 0, "���ϲ�����"
        .SetText E_WAFERID.E_Ngqty2, 0, "�Ƴ̲�����"
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
        '�趨������
        .Col = 1 'ѡ��
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        
        '�趨�п�
        .ColWidth(-1) = 10
        .ColWidth(1) = 4
        .RowHeight(-1) = 10
        '�趨�Ƿ�����
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
        '�趨������
        .Col = 1 'ѡ��
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        
        '�趨�п�
        .ColWidth(-1) = 10
        .ColWidth(1) = 4
        .RowHeight(-1) = 10
        '�趨�Ƿ�����
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
        '�趨������
        .Col = 1 'ѡ��
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        
        '�趨�п�
        .ColWidth(-1) = 10
        .ColWidth(E_VT.E_CHOOSE) = 4

        
        .RowHeight(-1) = 10
        '�趨�Ƿ�����
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

    '�����ѡ��ĵ��Ŷ�ѡ��
    If Row < 1 Then Exit Sub
    If Col <> 1 Then Exit Sub
    With fpS_wafer

        .Col = E_WAFERID.E_CHOOSE
        .Row = Row
        .Value = Abs(Val(.Value) - 1)
        If Val(.Value) = 1 Then   '������һ���Ĵ����ѡ����
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
            '������һ���ĵ���ѡ����
            .Col = E_WAFERID.E_BigBoxID
            .Row = Row
            strTmp = Trim$(.text)
'            strDJBH = Trim$(.Text) '���õĵ��ݱ�ţ��ڵ�����ӡʱ���õ�
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
 '@XH CHAR(8000),--С���
 '@lck   CHAR(8000) ,---���̿����
 '@gdh  CHAR(8000),--������
 '@lps   CHAR(8000) ,---��Ʒ��
 '@blS   CHAR(8000) ,---����Ʒ��
 '@zcbls  CHAR(8000),--�Ƴ̲�����
 '@FDCStock CHAR(50),--Ŀ��ⷿ
 '@dbry CHAR(50), --������Ա
 '@sqbm CHAR(20)='07', --���벿�� Ĭ�ϼƻ���
 '@NEWBOX CHAR(50) = '

    strrequestno = ""
    If strWholeName = gUserName Then
        MsgBox "��û��Ȩ��ִ�д˶���", vbInformation, "��ʾ"
        Exit Sub
    End If
    If Trim(TxtRequestNo.text) = "" Then
        MsgBox "���������뵥�ţ���ѯ", vbInformation, "��ʾ"
        Exit Sub
    End If
    selcnt = 0
    
 With fpS_stockview
    '��check
    For i = 1 To .MaxRows
        .Row = i
        .Col = E_StockView.E_CHOOSE
        
        If .text = 1 Then
            selcnt = i
           .Row = i
           .Col = E_StockView.e_order_num
           
            If Trim(.text) <> Trim(TxtRequestNo.text) Then
                MsgBox "��ͬ���뵥�Ų���һ�������", vbInformation, "��ʾ"
                Exit Sub
            End If
    
            .Row = i
            .Col = E_StockView.E_KF_FORMER
            strykf = Trim(.text) ' ԭ�ⷿ

            .Row = i
            .Col = E_StockView.E_Wafer
            strLCK = Trim(.text)  '���̿����
            
            .Row = i
            .Col = E_StockView.e_Qbox
                        
            strSql = "select distinct rtrim(�ⷿ���) from erpdata..tblstocknumsub where ���='" & Trim(.text) & "' and ���̿����='" & strLCK & "'"
            If GetSqlServerStr(strSql) <> strykf Then
               MsgBox "���" & Trim(.text) & "����" & strykf & "�ⷿ,�޷�����", vbInformation, "��ʾ"
               Exit Sub
            End If
        End If
    
            
    Next
    
    If selcnt = 0 Then
        MsgBox "û����Ҫ�����ĵ���,���Ȳ�ѯ", vbInformation, "��ʾ"
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
            'Ŀ��ⷿ��ͬ�����ֳɲ�ͬ�ĵ�������
            If strmbkf = "" Then
                strmbkf = Trim(.text) ' Ŀ��ⷿ
            Else
                If Trim(.text) <> strmbkf Then
                    Call DataOpt
                    strXH = ""
                    strgdh = ""
                    strLCK = ""
                    strlps = ""
                    strbls = ""
                    strzcbls = ""
                    strmbkf = Trim(.text) ' Ŀ��ⷿ��ͬ
                End If
            End If
            'ԭ�ⷿ��ͬ�����ֳɲ�ͬ�ĵ�������
            .Row = i
            .Col = E_StockView.E_KF_FORMER
            If strykf = "" Then
                strykf = Trim(.text) ' ԭ�ⷿ
            Else
                If Trim(.text) <> strykf Then
                    Call DataOpt
                    strXH = ""
                    strgdh = ""
                    strLCK = ""
                    strlps = ""
                    strbls = ""
                    strzcbls = ""
                    strykf = Trim(.text) ' ԭ�ⷿ
                End If
            End If
            .Row = i
            .Col = E_StockView.e_Qbox
            strXH = strXH & Trim(.text) & "��" 'С���
    
            .Row = i
            .Col = E_StockView.E_LOT
            strgdh = strgdh & Trim(.text) & "��" '������
           
            .Row = i
            .Col = E_StockView.E_Wafer
            strLCK = strLCK & Trim(.text) & "��" '���̿����
           
            .Row = i
            .Col = E_StockView.E_GOOD_DIE
            strlps = strlps & Trim(.text) & "��" '��Ʒ��
           
            .Row = i
            .Col = E_StockView.E_BAD1_DIE
            strbls = strbls & Trim(.text) & "��" '����Ʒ��
        
           
            .Row = i
            .Col = E_StockView.E_BAD2_DIE
            strzcbls = strzcbls & Trim(.text) & "��" '�Ƴ̲�����
        End If
    Next i
    
    
 End With
    strrequestno = ""
    strrequestitem = ""
    If DataOpt() = True Then

       'update erptemp..tblstockdb_temp��flag״̬
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
                   
                   strSql = "select top 1 rtrim(�������) + '-' + convert(varchar(5),���)  from erpdata..tblstockdb where rtrim(������Ա)='" & strWholeName & "' and id=" & strid & " and DATEDIFF(mi,����ʱ��,sysdatetime())<5 ORDER BY ����ʱ�� desc"
                   dbno_item = GetSqlServerStr(strSql)
                   If InStr(dbno_item, "-") > 0 Then
                       dbno = Split(dbno_item, "-")(0)
                       dbitem = Split(dbno_item, "-")(1)
                       strSql = "update erptemp..tblstockdb_temp set flag=2, remark2='" & dbno & "', remark3='" & dbitem & "',AUDITOR='" & strWholeName & "', AUDIT_TIME=sysdatetime()  where ORDER_NUM='" & strrequestno & "' and ITEM=" & strrequestitem & " and flag=1 "
                       AddSql2 (strSql)
                       strSql = "update erptemp..tblstockdb_temp set flag=4, remark2='" & dbno & "', remark3='" & dbitem & "',AUDITOR='" & strWholeName & "', AUDIT_TIME=sysdatetime()  where ORDER_NUM='" & strrequestno & "' and ITEM=" & strrequestitem & " and flag=3 "
                       AddSql2 (strSql)
                    Else
                       MsgBox strrequestno & strrequestitem & "���������쳣�������", vbInformation, "��ʾ"
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
 
 
 '@XH CHAR(8000),--С���
 '@lck   CHAR(8000) ,---���̿����
 '@gdh  CHAR(8000),--������
 '@lps   CHAR(8000) ,---��Ʒ��
 '@blS   CHAR(8000) ,---����Ʒ��
 '@zcbls  CHAR(8000),--�Ƴ̲�����
 '@FDCStock CHAR(50),--Ŀ��ⷿ
 '@dbry CHAR(50), --������Ա
 '@sqbm CHAR(20)='07', --���벿�� Ĭ�ϼƻ���
 '@NEWBOX CHAR(50) = '
 
 
    If strWholeName = gUserName Then
        MsgBox "��û��Ȩ��ִ�д˶���", vbInformation, "��ʾ"
        'Exit Sub
    End If
    
    If Label1.Caption <> "��������" Then
        MsgBox "�밴�������ų���", vbInformation, "��ʾ"
        Exit Sub
    End If
    If Trim(TxtdbNo.text) = "" Then
        MsgBox "�밴�������ų���", vbInformation, "��ʾ"
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
            strXH = Trim(.text)  'С���
            
            .Row = i
            .Col = E_StockView.E_Wafer
            strLCK = Trim(.text) '���̿����
        
            .Row = i
            .Col = E_StockView.E_KF_DEST
            strykf = Trim(.text) ' ԭ�ⷿ'�������� ������,��Ŀ��ⷿ��Ϊԭ�ⷿ
            If Trim(.text) <> "72" Then
                MsgBox "Ŀ��ַ�72�֣����ɳ�����", vbInformation, "��ʾ"
                Exit Sub
            End If

            strSql = "select distinct Rtrim(�ⷿ���) from erpdata..tblstocknumsub where ���='" & strXH & "' and ���̿����='" & strLCK & "'"
            If GetSqlServerStr(strSql) <> strykf Then
               MsgBox "���" & Trim(.text) & "�Ѳ���" & strykf & "�ⷿ,�޷�����", vbInformation, "��ʾ"
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
           strmbkf = Trim(.text) ' Ŀ��ⷿ'�������� ������,��ԭ�ⷿ��ΪĿ��ⷿ

           .Row = i
           .Col = E_StockView.E_KF_DEST
           strykf = Trim(.text) ' ԭ�ⷿ'�������� ������,��Ŀ��ⷿ��Ϊԭ�ⷿ
        
          .Row = i
          .Col = E_StockView.e_Qbox
           strXH = strXH & Trim(.text) & "��" 'С���
           
          .Row = i
          .Col = E_StockView.E_LOT
           strgdh = strgdh & Trim(.text) & "��" '������
           
           .Row = i
          .Col = E_StockView.E_Wafer
           strLCK = strLCK & Trim(.text) & "��" '���̿����
           
          .Row = i
          .Col = E_StockView.E_GOOD_DIE
           strlps = strlps & Trim(.text) & "��" '��Ʒ��
           
          .Row = i
          .Col = E_StockView.E_BAD1_DIE
           strbls = strbls & Trim(.text) & "��" '����Ʒ��
                   
          .Row = i
          .Col = E_StockView.E_BAD2_DIE
           strzcbls = strzcbls & Trim(.text) & "��" '�Ƴ̲�����


           
      
        End If
    Next i
 End With
    If strXH = "" Then
        MsgBox "��ѡ��Ҫ�����ļ�¼!", vbCritical + vbOKCancel, "ϵͳ��ʾ"
        Exit Sub
    End If

    If DataOpt() = True Then
       
        strSql = "select top 1 ������� from erpdata..tblstockdb where ������Ա='" & strWholeName & "'  and DATEDIFF(mi,����ʱ��,sysdatetime())<5 ORDER BY ����ʱ�� desc"
        dbno_cancer = GetSqlServerStr(strSql)
        If dbno_cancer = "" Then
            MsgBox "���������쳣�������", vbInformation, "��ʾ"
            Exit Sub
        Else
            strSql = "insert into  erptemp..invalidstockdb(�������,�����������) values('" & dbno_cancer & "','" & Trim(TxtdbNo.text) & "')"
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
  
        Set adoPrmReturn = New ADODB.Parameter         '����ִ�гɹ����
        adoPrmReturn.type = adInteger
        adoPrmReturn.Direction = adParamReturnValue
        adoCmd.Parameters.Append adoPrmReturn

        
        Set adoprm1 = New ADODB.Parameter                 '4���
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
        
        Set adoPrm7 = New ADODB.Parameter             '�ⷿ���
        adoPrm7.type = adChar
        adoPrm7.Size = 20
        adoPrm7.Direction = adParamInput
        adoPrm7.Value = strmbkf
        adoCmd.Parameters.Append adoPrm7
        
        Set adoPrm8 = New ADODB.Parameter               '������Ա
        adoPrm8.type = adChar
        adoPrm8.Size = 20
        adoPrm8.Direction = adParamInput
        adoPrm8.Value = Trim(strWholeName)
        adoCmd.Parameters.Append adoPrm8
        
        Set adoPrm9 = New ADODB.Parameter               '���벿��,�洢������Ĭ��07
        adoPrm9.type = adChar
        adoPrm9.Size = 20
        adoPrm9.Direction = adParamInput
        adoPrm9.Value = strdepartment
        adoCmd.Parameters.Append adoPrm9
        
        
        
        Set adoprm10 = New ADODB.Parameter               '�����
        adoprm10.type = adChar
        adoprm10.Size = 8000
        adoprm10.Direction = adParamInput
        adoprm10.Value = Trim(newboxid)
        adoCmd.Parameters.Append adoprm10
        
     
     adoCmd.Execute
     Screen.MousePointer = 0
     If adoPrmReturn.Value = 0 Then
       MsgBox "�Ѿ��ɹ�ִ����������", vbInformation, Me.Caption
       DataOpt = True

     Else
        GoTo there
     End If
  Exit Function
there:
  Screen.MousePointer = 0
  MsgBox "ִ��ʧ�ܣ�" + Chr(13) + "ԭ��:" + Err.DESCRIPTION, vbExclamation, Me.Caption
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





