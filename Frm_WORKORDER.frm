VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_WORKORDER 
   Caption         =   "��������ά��2.0"
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
      Caption         =   "������ϸ"
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
      Caption         =   "����ѡ��"
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
         Caption         =   "����"
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
         Caption         =   "ȫѡ/��ѡ"
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
         Caption         =   "��������"
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
         Caption         =   "������Ա"
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
         Caption         =   "��������"
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
         Caption         =   "������;"
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
         Caption         =   "�ͻ�����"
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
         Caption         =   "������"
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
         Caption         =   "NPI������"
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
         Caption         =   "Ԥ���깤"
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
         Caption         =   "Ԥ�ƿ���"
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
         Caption         =   "��������"
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
         Caption         =   "����ǰ׺"
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
         Caption         =   "��Ʒ�Ϻ�"
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
         Caption         =   "���ڻ���"
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
         Caption         =   "��������"
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
         Caption         =   "�ͻ�����"
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
         Caption         =   "�ͻ�����"
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
            Caption         =   " ��ӡ "
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "���"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "  ��  �� "
            Key             =   "INSERT"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ѯ����"
            Key             =   "READ"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "���ɹ���"
            Key             =   "CREATE"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�޸Ĺ���"
            Key             =   "UPDATE"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ɾ������"
            Key             =   "DELETE"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��������"
            Key             =   "EXPORT"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "  �� ��"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "A004"
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "�����"
            Key             =   "WAIT_PASS"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "  ��  ��"
            Key             =   "PASS"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "�����"
            Key             =   "CANCEL_PASS"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  ��  ��"
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
'    Project    : ��ʽ����1
'
'    Description: PMC����������ά��
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
' Project    :       ��ʽ����1
' Procedure  :       Form_Load
' Description:       �������
' Created by :       ����
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
' Project    :       ��ʽ����1
' Procedure  :       initCtrls
' Description:       ��ʼ���ؼ�״̬
' Created by :       ����
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
' Project    :       ��ʽ����1
' Procedure  :       InitFps
' Description:       ��ʼ��FPS
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/11-16:40:25
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitFps()

'������ϸ
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
    .SetText 0, 0, "���"
    .ColWidth(0) = 4
    .SetText E_WO_DETAIL.E_CHOOSE, 0, "��"
    .SetText E_WO_DETAIL.E_LOTID, 0, "������"
    .SetText E_WO_DETAIL.E_WAFERNO, 0, "��Բ���"
    .SetText E_WO_DETAIL.E_WAFERID, 0, "��ԲID"
    .SetText E_WO_DETAIL.E_GROSSDIES, 0, "��DIES"
    .SetText E_WO_DETAIL.E_GOODDIES, 0, "��ƷDIES"
    .SetText E_WO_DETAIL.E_NGDIES, 0, "����ƷDIES"
    .SetText E_WO_DETAIL.E_MARKINGCODE, 0, "�����"
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
' Project    :       ��ʽ����1
' Procedure  :       initDT
' Description:       ��ʼ������
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
' Project    :       ��ʽ����1
' Procedure  :       initCB_WOType
' Description:       ��ʼ�����������б�
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
' Project    :       ��ʽ����1
' Procedure  :       initCB_CustCode
' Description:       ��ʼ���ͻ������б�
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-9:44:13
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub initCB_CustCode()
Dim rs     As New ADODB.Recordset
Dim strSql As String

strSql = "select distinct �ͻ����� from erpdata..tblxcustomer where �ͻ����� is not null"
Set rs = Get_SqlserveRs(strSql)
cbCustCode.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbCustCode.AddItem Trim("" & rs!�ͻ�����)
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       initCB_WOName
' Description:       ��ʼ������ǰ׺�б�
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
' Project    :       ��ʽ����1
' Procedure  :       initCB_37PRI
' Description:       ��ʼ��37PRI�б�
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
' Project    :       ��ʽ����1
' Procedure  :       InitWOCreater
' Description:       ��ʼ������������Ա
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
' Project    :       ��ʽ����1
' Procedure  :       initData
' Description:       ��ʼ������
' Created by :       ����
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-9:40:23
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitData()

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       cbCustCode_LostFocus
' Description:       �ͻ�����ת��д
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
' Project    :       ��ʽ����1
' Procedure  :       cbCustCode_Change
' Description:       �ͻ�����ı�����ͻ�����/���ڻ����б�,���lstLotID
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
' Project    :       ��ʽ����1
' Procedure  :       cbCustCode_DropDown
' Description:       �ͻ������������ͻ�����/���ڻ����б�,���lstLotID
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
' Project    :       ��ʽ����1
' Procedure  :       cbHTPN_Change
' Description:       ���ڻ��ֱ������Ψһ�Ŀͻ�����/��Ʒ�Ϻ�
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
' Project    :       ��ʽ����1
' Procedure  :       cbHTPN_DropDown
' Description:       ���ڻ��ֱ������Ψһ�Ŀͻ�����/��Ʒ�Ϻ�
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
' Project    :       ��ʽ����1
' Procedure  :       cbCustPN_Change
' Description:       ���ڻ��ֱ�������ͻ����ֱ��,���lstLotID
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
' Project    :       ��ʽ����1
' Procedure  :       cbCustPN_Click
' Description:       �ͻ����ֱ��,���lstLotID
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
' Project    :       ��ʽ����1
' Procedure  :       cbProductNO_Change
' Description:       ��Ʒ�Ϻű��������������/NPI������
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

' ��������
strProductNO = Trim(cbProduct.text)
strSql = "select Get_Product_Dept('" & strProductNO & "') qtechpt from dual"
strPart1 = Get_OracleStr(strSql)
strSql = "select FNumber from AIS20141114094336.dbo.t_Department where FName='" & strPart1 & "' "
strPart2 = Get_SqlStr(strSql)
txtWODept.text = strPart1 & strPart2
' NPI������
strSql = "select residual from tbltsvnpiproduct where qtechptno2 = '" & strProductNO & "'"
txtNPIOwner.text = Get_OracleStr(strSql)

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       cbProductNO_Click
' Description:       ��Ʒ�Ϻű��������������/NPI������
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

' ��������
strProductNO = Trim(cbProduct.text)
strSql = "select Get_Product_Dept('" & strProductNO & "') qtechpt from dual"
strPart1 = Get_OracleStr(strSql)
strSql = "select FNumber from AIS20141114094336.dbo.t_Department where FName='" & strPart1 & "' "
strPart2 = Get_SqlStr(strSql)
txtWODept.text = strPart1 & strPart2
' NPI������
strSql = "select residual from tbltsvnpiproduct where qtechptno2 = '" & strProductNO & "'"
txtNPIOwner.text = Get_OracleStr(strSql)

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       cbWOType_Change
' Description:       NPI������-������(E)+�ͻ�ʵ��(Q)
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
' Project    :       ��ʽ����1
' Procedure  :       cbWOType_Click
' Description:       PI������-������(E)+�ͻ�ʵ��(Q)
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
' Project    :       ��ʽ����1
' Procedure  :       txtNPIOwner_Change
' Description:       �����˹��Ŵ�������
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
' Project    :       ��ʽ����1
' Procedure  :       cbWOName_Change
' Description:       �����Ŵ���������;
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
' Project    :       ��ʽ����1
' Procedure  :       chkLotSelect_Click
' Description:       LOTIDȫѡ/��ѡ
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
' Project    :       ��ʽ����1
' Procedure  :       cmdQuery_Click
' Description:       ����LOTID
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
    MsgBox "������LOT ID", vbInformation, "��ʾ:"
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
    MsgBox "��ѯ������LOTID", vbInformation, "��ʾ"

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
' Project    :       ��ʽ����1
' Procedure  :       ReadData
' Description:       ��ȡ����
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
    MsgBox "������ͻ�����", vbInformation, "��ʾ"
    Exit Sub

End If

If strCustPN = "" Then
    MsgBox "������ͻ�����", vbInformation, "��ʾ"
    Exit Sub

End If

Call ShowLotList(strCustCode, strCustPN)

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       ShowLotList
' Description:       ��ѯLOTID
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

    Case "��ͨ����"
        strSqlPart_Flag = " and aa.flag = 'Y'"

    Case "�ع�����"
        strSqlPart_WaferID = " and instr(aa.substrateid,'+') > 0"

    Case "Dummy����"
        strSqlPart_Flag = " and aa.flag = 'T'"
        strSqlPart_LotID = " and (aa.lotid like 'D%' or aa.lotid like 'SI%')"

    Case "��������"
        strSqlPart_Flag = " and aa.flag = 'T'"
        strSqlPart_LotID = " and aa.lotid like 'G%' "

    Case "�������"
        strSqlPart_Flag = " and aa.flag = 'T'"
        strSqlPart_LotID = " and aa.lotid like 'SI%' "
        strSqlPart_WaferID = " and instr(aa.substrateid,'+') = 0"

    Case "FO_CSP����"
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

        Case "��ͨ����"
            MsgBox "��ѯ�����ÿͻ����ֵ��ϴ�WO,�����ϴ���WO�Ѿ����˹���" & vbCrLf & "���ѯLotID�Ƿ����,�Լ��������������", vbInformation, "��ʾ"

        Case "�ع�����"
            MsgBox "��ѯ�����ÿͻ����ֵ��ع�WO" & vbCrLf & "���ֶ�ά��", vbInformation, "��ʾ"

        Case "Dummy����"
            MsgBox "��ѯ�����ÿͻ����ֵ�DummyWO" & vbCrLf & "���ֶ�ά��", vbInformation, "��ʾ"

        Case "��������"
            MsgBox "��ѯ�����ÿͻ����ֵĲ���WO" & vbCrLf & "���ֶ�ά��", vbInformation, "��ʾ"

        Case "�������"
            MsgBox "��ѯ�����ÿͻ����ֵĹ��WO" & vbCrLf & "���ֶ�ά��", vbInformation, "��ʾ"

        Case "FO_CSP����"
            MsgBox "��ѯ�����ÿͻ����ֵ�FO_CSP WO" & vbCrLf & "���ֶ�ά��", vbInformation, "��ʾ"

    End Select

End If

rs.Close
Set rs = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       lstLotID_Click
' Description:       ����LOTIDչ��Wafer�б�
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
' Project    :       ��ʽ����1
' Procedure  :       ShowWaferList
' Description:       ��ʾ������ϸ
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
        '��ѯ����
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

'ˢ������
End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       CreateWorkOrder
' Description:       ��������
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
If Not CheckByWO Then Exit Function     '�����㼶���ݼ��
If Not CheckByLot Then Exit Function    'Lot�㼶���ݼ��
If Not CheckByWafer Then Exit Function  'Wafer�㼶���ݼ��
CheckHandler = True

End Function

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       CheckByWO
' Description:       ��鹤���㼶����
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
    MsgBox "�����빤������", vbCritical, "��ʾ"
    Exit Function

End If

If cbWOType(1).text = "" Then
    MsgBox "�����빤����;", vbCritical, "��ʾ"
    Exit Function

End If

If cbCustCode.text = "" Then
    MsgBox "������ͻ�����", vbCritical, "��ʾ"
    Exit Function

End If

If cbCustPN.text = "" Then
    MsgBox "������ͻ�����", vbCritical, "��ʾ"
    Exit Function

End If

If cbHTPN.text = "" Then
    MsgBox "�����볧�ڻ���", vbCritical, "��ʾ"
    Exit Function

End If

If cbProduct.text = "" Then
    MsgBox "�������Ʒ�Ϻ�", vbCritical, "��ʾ"
    Exit Function

End If

If txtWODept.text = "" Then
    MsgBox "�����빤������", vbCritical, "��ʾ"
    Exit Function

End If

If cbWOName.text = "" Then
    MsgBox "�����빤��ǰ׺", vbCritical, "��ʾ"
    Exit Function

End If

If cb37Pri(0).text = "" Then
    MsgBox "������PRI", vbCritical, "��ʾ"
    Exit Function

End If

strWOType = cbWOType(0).text
strCustCode = Trim$(cbCustCode.text)
strCustPN = Trim$(cbCustPN.text)
strHTPN = Trim$(cbHTPN.text)
strproduct = Trim$(cbProduct.text)
'NPI���ֶ��ձ���
strSql = "select * from tbltsvnpiproduct where customershortname = '" & strCustCode & "' and customerptno1 = '" & strCustPN & "' and qtechptno = '" & strHTPN & "' and qtechptno2 = '" & strproduct & "'"
If Get_OracleCnt(strSql) = 0 Then
    MsgBox "NPIδά����ض��ռ�¼,����ϵNPIά��", vbCritical, "����"
    Exit Function

End If

'�������ڿ���
If Not (dTEnd.Value > dTBegin.Value) Then
    MsgBox "�깤���ڱ�����ڿ�������", vbCritical, "����"
    Exit Function

End If

'��Ʒ�Ϻſ���
If strWOType = "��ͨ����" Then
    strSql = "SELECT b.�Ϻ� FROM [erpdata].[dbo].[TSVtblSetMRule] a inner join [erpdata].[dbo].[TSVtblMRuleData] b on a.���Ϲ淶��� = b.���Ϲ淶��� where a.���ϱ��='" & strproduct & "'  and a.������� is not null "
    If Get_SqlserverCnt(strSql) = 0 Then
        MsgBox "ϵͳ�и��Ϻŵ�BOM�����ڻ�δ���,����ϵ��ص���,��ά�������BOM", vbCritical, "����"
        Exit Function

    End If

End If

'������������
If strWOType = "��������" Then
    strSql = "select * from tbltsvnpiproduct where customershortname = '" & strCustCode & "' and customerptno1 = '" & strCustPN & "' and qtechptno = '" & strHTPN & "' and  customerptno3 is not null and customerptno4 is not null and customerptno5 is not null and customerptno6 is not null"
    If Get_OracleCnt(strSql) = 0 Then
        MsgBox "��������û��ά���ض�����Ϣ(��ϴ����,CV�߶�,��ϴ����,�������)" & vbCrLf & "����ϵNPIά����Ӧ���ֵ���Ϣ", vbCritical, "����"
        Exit Function

    End If

End If

CheckByWO = True

End Function

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       CheckByLot
' Description:       ���Lot�㼶����
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
' Project    :       ��ʽ����1
' Procedure  :       CheckByWafer
' Description:       ���Wafer�㼶����
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
    MsgBox "NPI���ձ�δά����ȷ��GROSSDIES,����ϵNPI����ά��", vbCritical, "����"
    Exit Function

End If

With fpSDetail

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_WO_DETAIL.E_CHOOSE
        If .Value = 1 Then
            bChoose = True
            'GrossDies����
            If strWOType = "��ͨ����" Then
                .Col = E_WO_DETAIL.E_GROSSDIES
                If CLng(.text) <> lNpiGrossDies Then
                    MsgBox "NPIά����GROSSDIESΪ: " & lNpiGrossDies & vbCrLf & "WOά����GROSSDIESΪ: " & .text & vbCrLf & "���߲�һ��,����ϵ˫��ȷ��", vbCritical, "����"
                    Exit Function

                End If

            End If

        End If

    Next i

End With

If Not bChoose Then
    MsgBox "��ѡ����Ҫ����������WaferID", vbCritical, "����"
    Exit Function

End If

CheckByWafer = True

End Function

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       SaveHandler
' Description:       ��������
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
' Project    :       ��ʽ����1
' Procedure  :       GetWOData
' Description:       ��ȡ�����㼶����
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/12-8:46:39
'
' Parameters :       tWOData (T_WO_HEADER)
'--------------------------------------------------------------------------------
Private Sub GetWOData(ByRef tWOData As T_WO_HEADER)



End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       SaveWOData
' Description:       ���湤���㼶����
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
' Project    :       ��ʽ����1
' Procedure  :       UpdateData
' Description:       �޸Ĺ���
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-25-9:53:06
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub UpdateWorkOrder()

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       DeleteData
' Description:       ɾ������
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-25-9:53:12
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub DeleteWorkOrder()

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       ExportData
' Description:       ��������
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-25-9:53:17
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub ExportWOData()

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       GetNewWOID
' Description:       ��ȡ�¹�����
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
