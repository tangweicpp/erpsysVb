VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmGW_Report 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "���񱨱�"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   17265
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Fra 
      Caption         =   "��ѯ����"
      ForeColor       =   &H00FF0000&
      Height          =   7455
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   3495
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1080
         TabIndex        =   34
         Top             =   5640
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   ".."
         Height          =   360
         Left            =   2880
         TabIndex        =   33
         Top             =   5640
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "GC�����ϴ�"
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
         Left            =   1080
         TabIndex        =   32
         Top             =   6360
         Width           =   1710
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmGW_Report.frx":0000
         Left            =   1080
         List            =   "FrmGW_Report.frx":0028
         TabIndex        =   31
         Top             =   4320
         Width           =   2295
      End
      Begin MSComDlg.CommonDialog com 
         Left            =   2640
         Top             =   4800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "UPLOAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1080
         TabIndex        =   30
         Top             =   4800
         Width           =   1455
      End
      Begin VB.CheckBox chk 
         Caption         =   "�ǳɱ���"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   1
         Left            =   1080
         TabIndex        =   18
         Top             =   1800
         Width           =   2355
      End
      Begin VB.ComboBox Cob 
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
         Index           =   2
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   2355
      End
      Begin VB.ComboBox Cob 
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
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Text            =   "Cob"
         Top             =   600
         Width           =   2355
      End
      Begin VB.ComboBox Cob 
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
         Index           =   0
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2355
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   2280
         Width           =   2355
      End
      Begin VB.ComboBox Cob 
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
         Index           =   3
         ItemData        =   "FrmGW_Report.frx":0077
         Left            =   1080
         List            =   "FrmGW_Report.frx":0079
         TabIndex        =   3
         Top             =   1320
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Top             =   2760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "YYYY-MM-DD"
         Format          =   107806721
         CurrentDate     =   41387
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   3360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "YYYY-MM-DD"
         Format          =   107806721
         CurrentDate     =   41387
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·��"
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
         Left            =   360
         TabIndex        =   38
         Top             =   5640
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϴ�ģ��"
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
         Left            =   120
         TabIndex        =   37
         Top             =   4920
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ģ��"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Top             =   4320
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ģ���ϴ�"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   1.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   30
         Left            =   120
         TabIndex        =   35
         Top             =   4920
         Width           =   120
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���������"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���۵����"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������ĩ"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   15
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ������"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���߱��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   660
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��        ��"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   720
      End
   End
   Begin VB.Frame Fra 
      ForeColor       =   &H000000FF&
      Height          =   7455
      Index           =   1
      Left            =   3480
      TabIndex        =   0
      Top             =   720
      Width           =   9615
      Begin VB.Frame frameBGFP 
         Caption         =   "���ط�Ʊ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   1080
         TabIndex        =   20
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox txtSalesNo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            TabIndex        =   28
            Top             =   960
            Width           =   2500
         End
         Begin VB.CommandButton btnCannel 
            Caption         =   "ȡ    ��"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2640
            TabIndex        =   24
            Top             =   2300
            Width           =   1005
         End
         Begin VB.TextBox txtFHDH 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            TabIndex        =   23
            Top             =   480
            Width           =   2500
         End
         Begin VB.CommandButton btnConfirm 
            Caption         =   "ȷ    ��"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1440
            TabIndex        =   22
            Top             =   2300
            Width           =   1005
         End
         Begin VB.TextBox txtBGFPNo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            TabIndex        =   21
            Top             =   1480
            Width           =   2500
         End
         Begin VB.Label lblSalesNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���۵����:"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   27
            Top             =   960
            Width           =   1035
         End
         Begin VB.Label lblFHDH 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������:"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Width           =   840
         End
         Begin VB.Label lblBGFPNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ط�Ʊ��:"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   25
            Top             =   1485
            Width           =   1035
         End
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   3255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6015
         _Version        =   524288
         _ExtentX        =   10610
         _ExtentY        =   5741
         _StockProps     =   64
         EditEnterAction =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   5
         SpreadDesigner  =   "FrmGW_Report.frx":007B
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin MSComctlLib.Toolbar TlBar 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   17265
      _ExtentX        =   30454
      _ExtentY        =   1535
      ButtonWidth     =   1773
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "������ӡ  "
            Key             =   "A01"
            Object.ToolTipText     =   "����Excel���ӡ"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��������"
            Key             =   "A02"
            Object.ToolTipText     =   "��������ѯ��������"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "L11"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�����˵���"
            Key             =   "KeySet"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "   �� ��   "
            Key             =   "A03"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "ɾ ��"
            Key             =   "A04"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "�� ��"
            Key             =   "A05"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��  ѯ"
            Key             =   "A06"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "A07"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "ȡ ��"
            Key             =   "A08"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "ȷ  ��"
            Key             =   "A09"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "A004"
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "��  ��"
            Key             =   "A10"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��  ��"
            Key             =   "A11"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10080
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGW_Report.frx":055C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGW_Report.frx":2696
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGW_Report.frx":5520
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGW_Report.frx":7CD2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGW_Report.frx":9E0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGW_Report.frx":C5BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGW_Report.frx":ED70
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGW_Report.frx":11DF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGW_Report.frx":145A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGW_Report.frx":148BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGW_Report.frx":15598
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGW_Report.frx":1861A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmGW_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_PATH = 260

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


Public strSysPath As String
Public strUserName As String
Dim strdjbh         As String
Dim rs              As New ADODB.Recordset
Dim Flag_Exceute    As Integer
Private Enum fpSDetail
    E_CHOOSE = 1
    e_DJBH = 2
    E_cust = 3
    e_YDH = 7
End Enum

'''''''''''''''''''''''''''''
Private Sub Command1_Click()
    
    Dim strFilePath         As String
    Dim strFileName         As String
    Dim strSql              As String
    Dim image_Data()        As Byte         'ͼƬ������
    Dim rs                  As New ADODB.Recordset
    '��ͼƬ
    If gUserName <> "07885" Then
        
         MsgBox "��Ȩ�޲���", vbInformation, "��ʾ"
         
         Exit Sub
        
    End If
    com.Filter = "�ϴ��ļ�(*.xls,*.xlsx)|*.xls;*.xlsx"
    com.ShowOpen '�򿪶Ի���
    strFilePath = Trim(com.filename)  '����·��
    
    If com.filename = "" Then
    
        MsgBox "��ѡ���ļ�", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    strFileName = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1) '�ļ���
    '��ʼ���浽���Ͽ�
    '����ת��Ϊ��
    Open strFilePath For Binary As #1
    ReDim image_Data(LOF(1) - 1)
    Get #1, , image_Data()
    Close #1
    '��ѯ�Ƿ񱣴����ͼƬ
    strSql = "SELECT * FROM erpdata..tblSystemTemplet Where TEMPLETNAME = '" & Trim$(strFileName) & "' "
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenKeyset, adLockOptimistic
    If Not rs.EOF Then
        rs("TEMPLETNAME") = strFileName
        rs("FILECONTENT") = image_Data()
        rs.Update
    Else
        rs.AddNew
        rs("SYS_NAME") = "����"
        rs("TEMPLETNAME") = strFileName
        rs("create_date") = Now
        rs("FILECONTENT") = image_Data()
        rs("Memo") = Null
        rs.Update
    End If
    rs.Close
    
    MsgBox "�ϴ��ɹ�", vbInformation, "��ʾ"
  
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub btnCannel_Click()
    
    txtBGFPNo.text = ""
    txtFHDH.text = ""
    txtSalesNo.text = ""
    frameBGFP.Visible = False
    Fps(0).Visible = True
    
End Sub

Private Sub btnConfirm_Click()
    Dim Execute_SQL As String
    If Flag_Exceute = 0 Then
        Execute_SQL = "insert into erpdata..tblGW_TranSportNo(��Ʊ����,����ʱ��,��������,���ط�Ʊ��) values('" & txtSalesNo.text & "',getdate(),'" & txtFHDH.text & "','" & txtBGFPNo.text & "')"
    Else
        Execute_SQL = "update erpdata..tblGW_TranSportNo set ���ط�Ʊ��='" & txtBGFPNo.text & "',����ʱ��=getdate() where ��Ʊ����='" & txtSalesNo.text & "' and ��������='" & txtFHDH.text & "' "
    End If
    
    If rs.State = 1 Then rs.Close
    rs.Open Execute_SQL, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    txtBGFPNo.text = ""
    txtFHDH.text = ""
    txtSalesNo.text = ""
    frameBGFP.Visible = False
    Fps(0).Visible = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Fra(0).Move 60, Fra(0).Top, Fra(0).Width, Me.ScaleHeight - TlBar.Height
    Fra(1).Move Fra(1).Left, Fra(1).Top, Me.ScaleWidth - Fra(0).Width - 120, Me.ScaleHeight - TlBar.Height
    Fps(0).Move 60, Fps(0).Top, Fra(1).Width - 120, Me.ScaleHeight - TlBar.Height - 3 * 120
End Sub
Private Sub Form_Load()
    '��ʼ���ؼ�
    InitCtrl
    Application.DisplayAlerts = False
End Sub

Public Function FileExist(filename As String) As Boolean
'ȷ���ļ��Ƿ����
On Error GoTo ErrHandle
Dim FileData As WIN32_FIND_DATA
Dim Re As Long
Re = FindFirstFile(filename, FileData)
If Re = -1 Then
    FileExist = False
Else
    FileExist = True
    FindClose Re
End If

Exit Function
ErrHandle:
    FileExist = False
End Function


Public Sub GetExcelTempletInfo(ByVal StrSys_Name As String, Strxls As String)
On Error GoTo ErrHandle

Dim strSql          As String, strFileName As String
Dim iStm As ADODB.Stream
Dim iRe As ADODB.Recordset

  strSysPath = App.Path

  If Dir(strSysPath & "\TEMPLET", vbDirectory) = "" Then MkDir strSysPath & "\TEMPLET"
  strFileName = strSysPath & "\TEMPLET\" & Strxls

  If FileExist(strFileName) Then
      Kill strFileName
  End If

  Set iRe = New ADODB.Recordset
  strSql = "SELECT * FROM tblSystemTemplet WHERE SYS_NAME='" & StrSys_Name & "' AND UPPER(TEMPLETNAME)='" & UCase(Strxls) & "'"
  iRe.Open strSql, INIadoCon2, adOpenKeyset, adLockReadOnly
  If iRe.RecordCount > 0 Then
    Set iStm = New ADODB.Stream
    With iStm
        .Mode = adModeReadWrite
        .type = adTypeBinary
        .Open
        .Write iRe("FILECONTENT")
        .SaveToFile strFileName
    End With
    iStm.Close
  End If
  iRe.Close

  Exit Sub

ErrHandle:
  MsgBox "����:" & Err.DESCRIPTION & vbCrLf & " ���������Ѱ���йذ�����", vbExclamation, "ϵͳ"
  Exit Sub
End Sub

Private Sub CheckXls()
    '����ģ��
    
    Select Case UCase(Combo1.text)
        
        Case "68"
        
            Call GetExcelTempletInfo("����", "68_Invoice.xls")
            Call GetExcelTempletInfo("����", "68_Packing_list.xls")

        Case "76", "US026"
        
            Call GetExcelTempletInfo("����", "76_Invoice.xls")
            Call GetExcelTempletInfo("����", "76_Packing_list.xls")
        Case "SG005"
        
            Call GetExcelTempletInfo("����", "SG005_Invoice.xls")
            Call GetExcelTempletInfo("����", "SG005_Packing_list.xls")
        Case "SG005_SO"
        
            Call GetExcelTempletInfo("����", "SG005_SO_Invoice.xls")
            Call GetExcelTempletInfo("����", "SG005_SO_Packing_list.xls")
        Case "TW067", "ͨ��ģ��"
        
            Call GetExcelTempletInfo("����", "TW067_Invoice.xls")
            Call GetExcelTempletInfo("����", "TW067_Packing_list.xls")
            
        Case "BD", "EQ"
           
            Call GetExcelTempletInfo("����", "BD_Invoice.xls")
            Call GetExcelTempletInfo("����", "BD_Packing_list.xls")
        
        Case "HK005"
        
            Call GetExcelTempletInfo("����", "HK005_Invoice.xls")
            Call GetExcelTempletInfo("����", "HK005_Packing_list.xls")
            
        Case "HK080"
        
            Call GetExcelTempletInfo("����", "HK080_Invoice.xls")
            Call GetExcelTempletInfo("����", "HK080_Packing_list.xls")
            
        Case "GC"

            Call GetExcelTempletInfo("����", "GC_Invoice.xls")
            Call GetExcelTempletInfo("����", "GC_Packing_list.xls")
        Case "HK075"

            Call GetExcelTempletInfo("����", "HK075_Invoice.xls")
            Call GetExcelTempletInfo("����", "HK075_Packing_list.xls")
            
            
            
    End Select

End Sub

'��ʼ���ؼ�
Private Sub InitCtrl()
Dim i                   As Integer
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
    
    strdjbh = ""
    '���ص�������
    strSql = "SELECT ˵�� FROM dbo.tblbase WHERE ����='���񱨱�' AND ˵��2='0' ORDER BY ����  "
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Cob(0).Clear
    If Not rs.EOF Then
        Do While Not rs.EOF
            Cob(0).AddItem Trim$("" & rs!˵��)
            rs.MoveNext
        Loop
        Cob(0).ListIndex = 0
    End If
    rs.Close
    '���ؿͻ�����
    strSql = "SELECT DISTINCT �ͻ����� FROM dbo.tblXCustomer  "
    If rs.State = 1 Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Cob(1).Clear
    If Not rs.EOF Then
        With Cob(1)
            .AddItem "����"
            Do While Not rs.EOF
                .AddItem Trim$("" & rs!�ͻ�����)
                rs.MoveNext
            Loop
            .ListIndex = 0
        End With
    End If
    rs.Close
    
    '���ز��߱��
    strSql = "SELECT RTRIM(�����)+' '+RTRIM(������) ���߱�� FROM tblareadata  "
    If rs.State = 1 Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Cob(2).Clear
    If Not rs.EOF Then
        With Cob(2)
            Do While Not rs.EOF
                .AddItem Trim$("" & rs!���߱��)
                rs.MoveNext
            Loop
            .ListIndex = 0
        End With
    End If
    rs.Close
    
    'Fps��ʼ��
    With Fps(0)
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
        .Col = fpSDetail.E_CHOOSE   'ѡ��
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        '�趨�п�
        .ColWidth(-1) = 10
        .ColWidth(fpSDetail.E_CHOOSE) = 4
        .RowHeight(-1) = 10
        '�趨�Ƿ�����
        .UserColAction = UserColActionSort
        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
        .ZOrder
        .ReDraw = True
    End With
    
   DTP(0).Value = Format(Now() - 1, "YYYY/MM/DD")
   DTP(1).Value = Format(Now(), "YYYY/MM/DD")
   '���ģ��
   CheckXls
   
End Sub

Private Sub fps_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i           As Long
Dim j           As Integer
Dim strTmp      As String

    '�����ѡ��ĵ��Ŷ�ѡ��
    If Row < 1 Then Exit Sub
    If Col <> 1 Then Exit Sub
    With Fps(0)
'        .Col = FpsDetail.e_Choose
'        For i = 1 To .MaxRows
'            .Row = i
'            If i <> Row Then
'                .Col = FpsDetail.e_Choose
'                If Val(.Value) = 1 Then
''                    .Value = 0
'                    .Col = -1
'                    .ForeColor = vbBlack
'                End If
'            End If
'        Next

        .Col = fpSDetail.E_CHOOSE
        .Row = Row
        .Value = Abs(Val(.Value) - 1)
'        strDJBH = ""
        If Val(.Value) = 1 Then
            '������һ���ĵ���ѡ����
            .Col = fpSDetail.e_DJBH
            .Row = Row
            strTmp = Trim$(.text)
'            strDJBH = Trim$(.Text) '���õĵ��ݱ�ţ��ڵ�����ӡʱ���õ�
            For i = 1 To .MaxRows
                .Row = i
                .Col = fpSDetail.e_DJBH
                If Trim$(.text) = strTmp Then
                    .Col = fpSDetail.E_CHOOSE
                    .Value = 1
                    .Col = -1
                    .ForeColor = &HFF8080
                End If
            Next
        Else
            '������һ���ĵ���ѡ����
            .Col = fpSDetail.e_DJBH
            .Row = Row
            strTmp = Trim$(.text)
'            strDJBH = Trim$(.Text) '���õĵ��ݱ�ţ��ڵ�����ӡʱ���õ�
            For i = 1 To .MaxRows
                .Row = i
                .Col = fpSDetail.e_DJBH
                If Trim$(.text) = strTmp Then
                    .Col = fpSDetail.E_CHOOSE
                    .Value = 0
                    .Col = -1
                    .ForeColor = vbBlack
                End If
            Next
        End If
        
    End With
    
End Sub

Private Sub Fps_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    Flag_Exceute = 0
    Dim fhdhvalue As String
    Dim salesvalue As String
    Dim strSql As String
    If Row < 1 Then Exit Sub
    With Fps(0)
     
       .Col = 8
       .Row = Row
       fhdhvalue = .text
       .Col = 2
       salesvalue = .text
       
       strSql = "SELECT * FROM erpdata..tblGW_TranSportNo" & _
                " WHERE ��Ʊ����='" & salesvalue & "' and ��������='" & fhdhvalue & "' "
       If rs.State = 1 Then rs.Close
       rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
       If rs.RecordCount > 0 Then
          Flag_Exceute = 1
          txtBGFPNo.text = Trim$(rs!���ط�Ʊ��)
       End If
       
       Fps(0).Visible = False
       frameBGFP.Visible = True
       txtFHDH.text = fhdhvalue
       txtSalesNo.text = salesvalue
       
       txtBGFPNo.SetFocus
       txtBGFPNo.SelStart = Len(txtBGFPNo.text)
       
       
    End With
End Sub

Private Sub TlBar_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrHandle
'Dim m           As New ClsMouse
Dim strTmp      As String
Dim strTemp()   As String
Dim strydh      As String '�˵���
Dim i           As Integer

'    m.SetPointer vbHourglass
    Select Case Button.Key
        Case "A01"      '������ӡ
        
            If Combo1.text = "" Then
                
                MsgBox "��ѡ����Ҫ��ӡģ��ĸ�ʽ��", vbInformation, "��ʾ"
                
                Exit Sub
                
            End If
            
            CheckXls
                      
            If MsgBox("ȷ��Ҫ������", vbInformation + vbYesNo, "��ʾ") = vbNo Then Exit Sub
            'У������
            If Not CheckData Then Exit Sub
            '�����ݱ�����²�ѯһ�£��õ����ݼ����д�ӡ
            Call Search(strdjbh)
            '��ֵ����ȥ
            Cob(0).text = Fra(1).Caption
            '������ӡ����
            'Invoice
            If Cob(0).ListIndex = 0 Then
                strTmp = "Invoice"
                Call InvoiceExportPrintExcel(rs, strdjbh)
            End If
            'Packing list
            If Cob(0).ListIndex = 1 Then
                strTmp = "Packing_list"
                'rs.Sort = "PO_NUM Asc"
                Call PackinglistExportPrintExcel(rs, strdjbh)
            End If
            
        Case "A02"      '������ѯ�����ı���
            'У������
            If Fps(0).MaxRows <= 0 Then Exit Sub
            '��������
            If Not ExportFpspreadToExcel(Fps(0), Trim(Fra(1).Caption), Trim(Fra(1).Caption)) Then Exit Sub
            
        Case "KeySet"   '�����˵��� 2014.12.31 modify ���ݹ��������޸� ����������˵������ڵ�������֮������ģ�������԰�Ļ���û���˵���
            'У������
            If Not CheckData Then Exit Sub
            If strdjbh = "" Then
                MsgBox "û��ѡ��Ҫ���õ��ڲ����۵��ţ�", vbInformation, "��ʾ"
                Exit Sub
            End If
            strydh = InputBox("���������۵���" & strdjbh & "���˵��ţ�", "�������˵���", "")
            If Trim(strydh) = "" Then Exit Sub
            '��������
            If InStr(strdjbh, ",") > 0 Then
                strTemp = Split(strdjbh, ",")
                For i = 0 To UBound(strTemp)
                    Call SaveYDH(strTemp(i), strydh)
                Next
            Else
                Call SaveYDH(strdjbh, strydh)
            End If
            '�����ݱ�����²�ѯһ�£��õ����ݼ����д�ӡ
            Call Search(strdjbh)
        Case "A06"      '��ѯ
            strdjbh = ""
            Call Search(strdjbh)
        Case "A11"      '�˳�
            Unload Me
    End Select

    Exit Sub
    
ErrHandle:
    Screen.MousePointer = 0
    MsgBox "ִ��ʧ�ܣ�" + Chr(13) + "ԭ��:" + Err.DESCRIPTION, vbInformation, Me.Caption
End Sub
Private Sub SaveYDH(strDJH As String, strydh As String)
Dim strSql      As String
Dim RsNew          As New ADODB.Recordset
    
    '����ʱ���м�鿴�Ƿ��������˵��ţ�û������ʾ����
    strSql = "SELECT * FROM erpdata..tblGW_TranSportNo Where ��Ʊ����='" & Trim(strDJH) & "'  "
    If RsNew.State = 1 Then RsNew.Close
    RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If RsNew.EOF Then
        '�������ݿ�
        strSql = "Insert Into erpdata..tblGW_TranSportNo(��Ʊ����,�˵���) Values('" & Trim(strDJH) & "','" & Trim(strydh) & "')"
        INIadoCon.Execute strSql
    Else '�в�ѯ�����ţ�ֻ��Ҫ�޸�
        '�޸����ݿ�
        strSql = "Update erpdata..tblGW_TranSportNo Set �˵���='" & Trim(strydh) & "' Where ��Ʊ����='" & Trim(strDJH) & "'"
        INIadoCon.Execute strSql
    End If
    RsNew.Close
            
End Sub
'У������
Private Function CheckData() As Boolean
Dim i               As Integer
Dim intCount        As Integer
Dim strCust         As String

    CheckData = False
    
    strdjbh = ""     '--���ݱ�ż�¼
    strCust = ""
    intCount = 0
    
    With Fps(0)
        If .MaxRows <= 0 Then
            MsgBox "û���κ�����,���Ȳ�ѯ��", vbInformation, "��ʾ"
            Exit Function
        End If
        '���Ƿ���ѡ��
        For i = 1 To .MaxRows
            .Row = i
            .Col = fpSDetail.E_CHOOSE  'ѡ��
            If .Value = 1 Then
                intCount = intCount + 1
                .Col = fpSDetail.e_DJBH '���ݱ��
                If InStr(strdjbh, Trim$(.text)) <= 0 Then
                    strdjbh = strdjbh + Trim$(.text) + ","
                End If
               .Col = fpSDetail.E_cust '�ͻ�����
               If UCase(Trim(Cob(1).text)) = "HK075" Then
                    .Col = 5
               Else
                    .Col = 3
               End If
                If strCust = "" Then
                    strCust = Trim$(.text)
                Else
                    If strCust <> Trim(.text) Then
                        MsgBox "��ͬ�ͻ����ϲ���ͬʱ������ӡ�������˵��ţ�", vbInformation, "��ʾ"
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
    'ȥ�����ݱ�����һ������
    strdjbh = Left$(strdjbh, Len(strdjbh) - 1)
    '--------------------------
    If intCount <= 0 Then
        MsgBox "û��ѡ���κ����ϣ�", vbInformation, "��ʾ"
        Exit Function
    End If
    
    CheckData = True
End Function

'��ѯ����
Public Sub Search(strdj As String)

    On Error GoTo ErrHandle

    Dim i            As Long

    Dim j            As Integer

    Dim strTDJBH     As String

    Dim adoprm1      As ADODB.Parameter

    Dim adoprm2      As ADODB.Parameter

    Dim adoPrm3      As ADODB.Parameter

    Dim adoPrm4      As ADODB.Parameter

    Dim adoPrm5      As ADODB.Parameter

    Dim adoPrm6      As ADODB.Parameter

    Dim adoPrm7      As ADODB.Parameter

    Dim adoPrm8      As ADODB.Parameter

    Dim adoPrm9      As ADODB.Parameter

    Dim adoprm10     As ADODB.Parameter

    Dim adoPrm11     As ADODB.Parameter

    Dim adoprmFG     As ADODB.Parameter

    Dim adoPrmReturn As ADODB.Parameter
    
    
    Dim strsql1  As String
    Dim strSql2 As String
    Dim strTDJBH1 As String
    
    Dim rs1          As New ADODB.Recordset
    Dim Rs2          As New ADODB.Recordset
    
       
    If strdj = "" Then
        
    strsql1 = " SELECT  DISTINCT  RTRIM(a.���۵����)  FROM erpdata..tblSale a  Where  a.�ͻ����� = '" & Trim(Cob(1).text) & "'  AND CONVERT(VARCHAR(20), a.��������,23) >= '" & Format(Trim(DTP(0).Value), "YYYY-MM-DD") & "' AND CONVERT(VARCHAR(20), a.��������,23) <= '" & Format(Trim$(DTP(1).Value), "YYYY-MM-DD") & "'"
    
    If rs1.State = adStateOpen Then rs1.Close
    rs1.Open strsql1, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
   If Not rs1.EOF Then
    Do While Not rs1.EOF
    
    strTDJBH = strTDJBH + rs1.Fields(0).Value + "','"
    rs1.MoveNext
  Loop
  Else
  
  MsgBox "û�г�����Ϣ", vbCritical, Me.Caption
  Exit Sub
  
  End If
        
  strTDJBH = Mid(strTDJBH, 1, Len(strTDJBH) - 3)
        
    Else
        strTDJBH = Replace(strdj, ",", "','")

    End If

    
  If Cob(0).ListIndex = "0" Then
  
     If UCase(Trim(Combo1.text)) = "HK075" Then           'HK075�Ĳ�ѯSQL ����by�˳̳� 2019-12-24
       strSql2 = "SELECT 0 AS ѡ��, x.���۵����,x.reticle_level_72 as ""Line_item"",x.reticle_level_73 as ""NC"",x.�ͻ�����,x.�ͻ�����,x.������,x.��������,x.�Ƶ�Ա,x.�˵���,x.��������, ROW_NUMBER() OVER(ORDER BY x.���۵����,x.���۵����) AS Item ,x.PO_NUM,x.PO_ITEM ,x.MPN_DESC" & _
    "  ,x.�Ϻ�,x.�����,x.С���,x.������ ,'Integrated Circuit chip' Specification,SUM(x.����) AS ����,����,SUM(x.���) AS ���,x.����,x.�ӹ��� ,�͹��ϵ���,SUM(x.�ӹ��ѽ��) AS �ӹ��ѽ�� " & _
    "  ,SUM(�͹��Ͻ��) AS �͹��Ͻ�� FROM ( SELECT RTRIM(a.���۵����) ���۵����,e.reticle_level_72,e.reticle_level_73, a.�ͻ�����,c.�ͻ�����,SUBSTRING(yy.SHIP_TO_AD,1,CHARINDEX('@',yy.SHIP_TO_AD) - 1 ) as  ������ ,CONVERT(VARCHAR(20), a.��������,23) �������� ,a.�Ƶ�Ա,ISNULL(f.�˵���,'') �˵��� " & _
    "  ,RTRIM(b.���ݱ��) �������� , b.���۵����,e.PO_NUM,e.PO_ITEM ,CASE WHEN e.CUSTOMERSHORTNAME in ('US026','SG005') THEN e.MPN_DESC + REPLACE(REPLACE(x.�ϸ���,'1','-D'),'0','') " & _
    "  WHEN A.�ͻ����� in ('76','AA') THEN  XX.MPN ELSE e.MPN_DESC end MPN_DESC ,b.�Ϻ�,b.С���,ISNULL(ad.remark2,replace(d.������,' ','')) as ������ " & _
    " ,x.���� * (CASE WHEN  SUBSTRING( d.���ݱ��,1,1) IN ('T','R') THEN -1 ELSE 1 END ) as ���� ,(b.����+b.�͹����ϵ���) ����,CONVERT(DECIMAL(18,2),x.����*(b.����+b.�͹����ϵ���)) ��� " & _
    "  ,case when d.���ݱ�� like 'T%' THEN 'Y' ELSE ISNULL(AA.FLAG,'N') END ���� ,b.���� AS �ӹ��� ,b.�͹����ϵ��� AS  �͹��ϵ��� ,CONVERT(DECIMAL(18,2), x.���� * b.����,2) as �ӹ��ѽ�� " & _
    "  ,CONVERT(DECIMAL(18,2), x.����*(b.����+b.�͹����ϵ���)) - CONVERT(DECIMAL(18,2),x.���� * b.����)  as �͹��Ͻ��,ISNULL(cc.���,bb.���)  �����  FROM erpdata..tblSale a  INNER JOIN erpdata..tblSaleRec b " & _
    "  ON a.���۵����=b.���۵���� INNER JOIN erpdata.dbo.tblXCustomer c ON a.�ͻ�����=c.�ͻ�����  INNER JOIN erpdata..tblStockSQfh d ON b.���ݱ��=d.���ݱ�� AND b.�������=d.��� " & _
    "  LEFT JOIN  erpdata.. tblStocksqfhsub x ON x.���ݱ�� = b.���ݱ��  AND x.������� = b.������� and x.��� = b.С��� INNER JOIN erpdata..tblStockNumTree bb ON bb.��� = x.��� AND bb.������ = 0  " & _
    " LEFT JOIN erpdata..tblStockNumTree cc ON cc.��� = bb.�ϼ���� AND cc.������ = 1  LEFT JOIN ERPBASE..tblmappingData dd " & _
    "  ON dd.SUBSTRATEID = x.���̿����  LEFT JOIN erpbase .. tblCustomerOI e ON CONVERT(VARCHAR(20), CONVERT(int,e.ID))  = dd.FILENAME AND e.SOURCE_BATCH_ID = dd.LOTID " & _
    "  Left JOIN erpdata..tblGW_TranSportNo f ON a.���۵����=f.��Ʊ����  LEFT JOIN ERPDATA..MDZD_TBL AA ON AA.SENT_ID = b.���ݱ��  LEFT join erptemp..mps_mark ad " & _
    "  ON ad.wafer_id = x.���̿����  LEFT JOIN  erpdata .. tblTSVworkorder xx ON xx.ORDERNAME = x.�󹤵�   LEFT JOIN erptemp..customer_information yy  ON yy.CUSTOMER = d.�ͻ����� AND yy.SHIP_TO = d.������ַ   Where   a.���۵���� IN ('" & strTDJBH & "') ) x " & _
    "   GROUP BY x.���۵����,x.���۵����,x.�ͻ�����,x.�ͻ�����,x.��������,x.�Ƶ�Ա ,x.�˵���,x.��������,x.PO_NUM,x.�Ϻ�,x.�����,x.С���,x.������ ,����,x.����,x.�ӹ���,�͹��ϵ���,x.PO_ITEM ,x.MPN_DESC,x.������,x.reticle_level_72,x.reticle_level_73 order by x.���۵����,x.reticle_level_72,x.�����,x.С��� "
   Else
        strSql2 = "SELECT 0 AS ѡ��, x.���۵����,x.�ͻ�����,x.�ͻ�����,x.������,x.��������,x.�Ƶ�Ա,x.�˵���,x.��������, ROW_NUMBER() OVER(ORDER BY x.���۵����,x.���۵����) AS Item ,x.PO_NUM,x.PO_ITEM ,x.MPN_DESC" & _
    "  ,x.�Ϻ�,x.�����,x.С���,x.������ ,'Integrated Circuit chip' Specification,SUM(x.����) AS ����,����,SUM(x.���) AS ���,x.����,x.�ӹ��� ,�͹��ϵ���,SUM(x.�ӹ��ѽ��) AS �ӹ��ѽ�� " & _
    "  ,SUM(�͹��Ͻ��) AS �͹��Ͻ�� FROM ( SELECT     RTRIM(a.���۵����) ���۵����,a.�ͻ�����,c.�ͻ�����,SUBSTRING(yy.SHIP_TO_AD,1,CHARINDEX('@',yy.SHIP_TO_AD) - 1 ) as  ������ ,CONVERT(VARCHAR(20), a.��������,23) �������� ,a.�Ƶ�Ա,ISNULL(f.�˵���,'') �˵��� " & _
    "  ,RTRIM(b.���ݱ��) �������� , b.���۵����,e.PO_NUM,e.PO_ITEM ,CASE WHEN e.CUSTOMERSHORTNAME in ('US026','SG005') THEN e.MPN_DESC + REPLACE(REPLACE(x.�ϸ���,'1','-D'),'0','') " & _
    "  WHEN A.�ͻ����� in ('76','AA') THEN  XX.MPN ELSE e.MPN_DESC end MPN_DESC ,b.�Ϻ�,b.С���,ISNULL(ad.remark2,replace(d.������,' ','')) as ������ " & _
    " ,x.���� * (CASE WHEN  SUBSTRING( d.���ݱ��,1,1) IN ('T','R') THEN -1 ELSE 1 END ) as ���� ,(b.����+b.�͹����ϵ���) ����,CONVERT(DECIMAL(18,2),x.����*(b.����+b.�͹����ϵ���)) ��� " & _
    "  ,case when d.���ݱ�� like 'T%' THEN 'Y' ELSE ISNULL(AA.FLAG,'N') END ���� ,b.���� AS �ӹ��� ,b.�͹����ϵ��� AS  �͹��ϵ��� ,CONVERT(DECIMAL(18,2), x.���� * b.����,2) as �ӹ��ѽ�� " & _
    "  ,CONVERT(DECIMAL(18,2), x.����*(b.����+b.�͹����ϵ���)) - CONVERT(DECIMAL(18,2),x.���� * b.����)  as �͹��Ͻ��,ISNULL(cc.���,bb.���)  �����  FROM erpdata..tblSale a  INNER JOIN erpdata..tblSaleRec b " & _
    "  ON a.���۵����=b.���۵���� INNER JOIN erpdata.dbo.tblXCustomer c ON a.�ͻ�����=c.�ͻ�����  INNER JOIN erpdata..tblStockSQfh d ON b.���ݱ��=d.���ݱ�� AND b.�������=d.��� " & _
    "  LEFT JOIN  erpdata.. tblStocksqfhsub x ON x.���ݱ�� = b.���ݱ��  AND x.������� = b.������� and x.��� = b.С��� INNER JOIN erpdata..tblStockNumTree bb ON bb.��� = x.��� AND bb.������ = 0  " & _
    " LEFT JOIN erpdata..tblStockNumTree cc ON cc.��� = bb.�ϼ���� AND cc.������ = 1  LEFT JOIN ERPBASE..tblmappingData dd " & _
    "  ON dd.SUBSTRATEID = x.���̿����  LEFT JOIN erpbase .. tblCustomerOI e ON CONVERT(VARCHAR(20), CONVERT(int,e.ID))  = dd.FILENAME AND e.SOURCE_BATCH_ID = dd.LOTID " & _
    "  Left JOIN erpdata..tblGW_TranSportNo f ON a.���۵����=f.��Ʊ����  LEFT JOIN ERPDATA..MDZD_TBL AA ON AA.SENT_ID = b.���ݱ��  LEFT join erptemp..mps_mark ad " & _
    "  ON ad.wafer_id = x.���̿����  LEFT JOIN  erpdata .. tblTSVworkorder xx ON xx.ORDERNAME = x.�󹤵�    LEFT JOIN erptemp..customer_information yy  ON yy.CUSTOMER = d.�ͻ����� AND yy.SHIP_TO = d.������ַ    Where   a.���۵���� IN ('" & strTDJBH & "') ) x " & _
    "   GROUP BY x.���۵����,x.���۵����,x.�ͻ�����,x.�ͻ�����,x.��������,x.�Ƶ�Ա ,x.�˵���,x.��������,x.PO_NUM,x.�Ϻ�,x.�����,x.С���,x.������ ,����,x.����,x.�ӹ���,�͹��ϵ���,x.PO_ITEM ,x.MPN_DESC ,x.������  "
    
      If UCase(Trim(Combo1.text)) = "ͨ��ģ��" Then
    
         strSql2 = strSql2 & " order by x.���۵����,x.�������� ,x.������ "
    
      Else
         strSql2 = strSql2 & " order by x.���۵����,x.��������,x.�����,x.С��� "
    
      End If
    End If
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql2, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
        With Fps(0)
            .MaxRows = 0
            Set .DataSource = rs
            .MaxRows = rs.RecordCount
        End With
        Fra(1).Caption = Trim$(Cob(0).text)
   Else
     If UCase(Trim(Combo1.text)) = "HK075" Then           'HK075�Ĳ�ѯSQL
         strSql2 = " SELECT 0 ѡ��,RTRIM(a.���۵����) ���۵����,a.�ͻ�����,c.�ͻ�����,SUBSTRING(yy.SHIP_TO_AD,1,CHARINDEX('@',yy.SHIP_TO_AD) - 1 ) as  ������ ,CONVERT(VARCHAR(100),a.��������,23)  ��������,a.�Ƶ�Ա,ISNULL(f.�˵���,'') �˵���,RTRIM(b.���ݱ��) ��������,e.PO_NUM,e.PO_ITEM  AS Item " & _
       " ,CASE WHEN e.CUSTOMERSHORTNAME in ('US026','SG005') THEN e.MPN_DESC + REPLACE(REPLACE(x.�ϸ���,'1','-D'),'0','')  " & _
       " WHEN A.�ͻ����� in ('76','AA') THEN  XX.MPN  ELSE e.MPN_DESC end MPN_DESC ,b.�Ϻ�,b.С���,ISNULL(ad.remark2,replace(d.������,' ','')) ������ " & _
       " ,'Integrated Circuit chip' Specification,SUM(x.����) ����,ISNULL(cc.���,bb.���)  �����    " & _
       " ,isnull(cc.����,0) as ����,CASE WHEN a.�ͻ����� in ('GC') THEN round(SUM(x.����) * 0.1/6000,2) ELSE  cast (isnull(cc.����,0) as NUMERIC) * 0.25 END as ����,isnull(cc.�ߴ�,'') MEAS " & _
       ",e.RETICLE_LEVEL_72 AS 'LineItem',e.RETICLE_LEVEL_73 AS 'NC12' " & _
       "  FROM erpdata..tblSale a  INNER JOIN erpdata..tblSaleRec b ON a.���۵����=b.���۵����  " & _
       " INNER JOIN erpdata.dbo.tblXCustomer c ON a.�ͻ�����=c.�ͻ�����  INNER JOIN erpdata..tblStockSQfh d ON b.���ݱ��=d.���ݱ�� AND b.�������=d.��� " & _
       " inner JOIN  erpdata..tblStocksqfhsub x ON x.���ݱ�� = b.���ݱ��  AND x.������� = b.������� AND x.������ = b.������  AND x.��� = b.С���   " & _
       " INNER JOIN erpdata..tblStockNumTree bb ON bb.��� = x.��� AND bb.������ = 0 LEFT JOIN erpdata..tblStockNumTree cc ON cc.��� = bb.�ϼ���� AND cc.������ = 1   " & _
       " LEFT JOIN ERPBASE..tblmappingData dd ON dd.SUBSTRATEID = x.���̿���� LEFT JOIN erpbase..tblCustomerOI e ON CONVERT(VARCHAR(20), CONVERT(int,e.ID))  = dd.FILENAME AND e.SOURCE_BATCH_ID = dd.LOTID   " & _
       "  Left JOIN erpdata..tblGW_TranSportNo f ON a.���۵����=f.��Ʊ���� left join erptemp..mps_mark ad on ad.wafer_id = x.���̿����   " & _
       "  LEFT JOIN  erpdata .. tblTSVworkorder xx ON xx.ORDERNAME = x.�󹤵�   LEFT JOIN erptemp..customer_information yy  ON yy.CUSTOMER = d.�ͻ����� AND yy.SHIP_TO = d.������ַ  " & _
       "  Where  a.���۵���� IN  ('" & strTDJBH & "') GROUP BY a.���۵����,a.�ͻ�����,c.�ͻ�����,a.��������,a.�Ƶ�Ա,ISNULL(f.�˵���,''),RTRIM(b.���ݱ��)   " & _
       " ,e.PO_NUM,e.PO_ITEM, e.CUSTOMERSHORTNAME , e.MPN_DESC,b.�Ϻ�,b.С���,b.�����,d.������ ,ISNULL(cc.���,bb.���),x.�ϸ���,XX.MPN,ad.remark2 ,cc.����,cc.�ߴ�,yy.SHIP_TO_AD ,e.RETICLE_LEVEL_72,e.RETICLE_LEVEL_73 " & _
       " Order by a.���۵����,b.�����,b.С���"
        

    ElseIf UCase(Trim(Combo1.text)) = "68" Then
           
           strSql2 = "SELECT 0 ѡ��,RTRIM(a.���۵����) ���۵����,a.�ͻ�����,c.�ͻ�����,SUBSTRING(yy.SHIP_TO_AD,1,CHARINDEX('@',yy.SHIP_TO_AD) - 1 ) as  ������ ,CONVERT(VARCHAR(100),a.��������,23)  ��������,a.�Ƶ�Ա,ISNULL(f.�˵���,'') �˵���,RTRIM(b.���ݱ��) ��������,e.PO_NUM,e.PO_ITEM  AS Item " & _
    "   ,CASE WHEN e.CUSTOMERSHORTNAME in ('US026','SG005') THEN e.MPN_DESC + REPLACE(REPLACE(x.�ϸ���,'1','-D'),'0','') " & _
    "   WHEN A.�ͻ����� in ('76','AA') THEN  XX.MPN  ELSE e.MPN_DESC end MPN_DESC ,b.�Ϻ�,b.С���,ISNULL(ad.remark2,replace(d.������,' ','')) ������ " & _
    "   ,'Integrated Circuit chip' Specification,SUM(x.����) ����,ISNULL(cc.���,bb.���)  �����  " & _
    "  ,isnull(cc.����,0) as ����,CASE WHEN a.�ͻ����� in ('GC') THEN round(SUM(x.����) * 0.1/6000,2) ELSE  cast (isnull(cc.����,0) as NUMERIC) * 0.25 END as ����,isnull(cc.�ߴ�,'') MEAS " & _
    " ,right(year(xx.erpcreatedate),2 )  +'' + convert(VARCHAR(2),RIGHT(100+DATEPART(week,xx.erpcreatedate),2) ) as DC  " & _
    "   FROM erpdata..tblSale a  INNER JOIN erpdata..tblSaleRec b ON a.���۵����=b.���۵���� " & _
    "   INNER JOIN erpdata.dbo.tblXCustomer c ON a.�ͻ�����=c.�ͻ�����  INNER JOIN erpdata..tblStockSQfh d ON b.���ݱ��=d.���ݱ�� AND b.�������=d.��� " & _
    "   inner JOIN  erpdata..tblStocksqfhsub x ON x.���ݱ�� = b.���ݱ��  AND x.������� = b.������� AND x.������ = b.������  AND x.��� = b.С��� " & _
    "   INNER JOIN erpdata..tblStockNumTree bb ON bb.��� = x.��� AND bb.������ = 0 LEFT JOIN erpdata..tblStockNumTree cc ON cc.��� = bb.�ϼ���� AND cc.������ = 1 " & _
    "   LEFT JOIN ERPBASE..tblmappingData dd ON dd.SUBSTRATEID = x.���̿���� LEFT JOIN erpbase..tblCustomerOI e ON CONVERT(VARCHAR(20), CONVERT(int,e.ID))  = dd.FILENAME AND e.SOURCE_BATCH_ID = dd.LOTID " & _
    "   Left JOIN erpdata..tblGW_TranSportNo f ON a.���۵����=f.��Ʊ���� left join erptemp..mps_mark ad on ad.wafer_id = x.���̿���� " & _
    "   LEFT JOIN  erpdata .. tblTSVworkorder xx ON xx.ORDERNAME = x.�󹤵�   LEFT JOIN erptemp..customer_information yy  ON yy.CUSTOMER = d.�ͻ����� AND yy.SHIP_TO = d.������ַ   Where  a.���۵���� IN  ('" & strTDJBH & "') GROUP BY a.���۵����,a.�ͻ�����,c.�ͻ�����,a.��������,a.�Ƶ�Ա,ISNULL(f.�˵���,''),RTRIM(b.���ݱ��) " & _
    "   ,e.PO_NUM,e.PO_ITEM, e.CUSTOMERSHORTNAME , e.MPN_DESC,b.�Ϻ�,b.С���,b.�����,d.������ ,ISNULL(cc.���,bb.���),x.�ϸ���,XX.MPN,ad.remark2 ,cc.����,cc.�ߴ�,yy.SHIP_TO_AD ,right(year(xx.erpcreatedate),2 )  +'' + convert(VARCHAR(2),RIGHT(100+DATEPART(week,xx.erpcreatedate),2) )  order by a.���۵����,RTRIM(b.���ݱ��),b.�����,b.С���"

    
    
    Else
   
       
           strSql2 = "SELECT 0 ѡ��,RTRIM(a.���۵����) ���۵����,a.�ͻ�����,c.�ͻ�����,SUBSTRING(yy.SHIP_TO_AD,1,CHARINDEX('@',yy.SHIP_TO_AD) - 1 ) as  ������ ,CONVERT(VARCHAR(100),a.��������,23)  ��������,a.�Ƶ�Ա,ISNULL(f.�˵���,'') �˵���,RTRIM(b.���ݱ��) ��������,e.PO_NUM,e.PO_ITEM  AS Item " & _
    "   ,CASE WHEN e.CUSTOMERSHORTNAME in ('US026','SG005') THEN e.MPN_DESC + REPLACE(REPLACE(x.�ϸ���,'1','-D'),'0','') " & _
    "   WHEN A.�ͻ����� in ('76','AA') THEN  XX.MPN  ELSE e.MPN_DESC end MPN_DESC ,b.�Ϻ�,b.С���,ISNULL(ad.remark2,replace(d.������,' ','')) ������ " & _
    "   ,'Integrated Circuit chip' Specification,SUM(x.����) ����,ISNULL(cc.���,bb.���)  �����  " & _
    "  ,isnull(cc.����,0) as ����,CASE WHEN a.�ͻ����� in ('GC') THEN round(SUM(x.����) * 0.1/6000,2) ELSE  cast (isnull(cc.����,0) as NUMERIC) * 0.25 END as ����,isnull(cc.�ߴ�,'') MEAS " & _
    "   FROM erpdata..tblSale a  INNER JOIN erpdata..tblSaleRec b ON a.���۵����=b.���۵���� " & _
    "   INNER JOIN erpdata.dbo.tblXCustomer c ON a.�ͻ�����=c.�ͻ�����  INNER JOIN erpdata..tblStockSQfh d ON b.���ݱ��=d.���ݱ�� AND b.�������=d.��� " & _
    "   inner JOIN  erpdata..tblStocksqfhsub x ON x.���ݱ�� = b.���ݱ��  AND x.������� = b.������� AND x.������ = b.������  AND x.��� = b.С��� " & _
    "   INNER JOIN erpdata..tblStockNumTree bb ON bb.��� = x.��� AND bb.������ = 0 LEFT JOIN erpdata..tblStockNumTree cc ON cc.��� = bb.�ϼ���� AND cc.������ = 1 " & _
    "   LEFT JOIN ERPBASE..tblmappingData dd ON dd.SUBSTRATEID = x.���̿���� LEFT JOIN erpbase..tblCustomerOI e ON CONVERT(VARCHAR(20), CONVERT(int,e.ID))  = dd.FILENAME AND e.SOURCE_BATCH_ID = dd.LOTID " & _
    "   Left JOIN erpdata..tblGW_TranSportNo f ON a.���۵����=f.��Ʊ���� left join erptemp..mps_mark ad on ad.wafer_id = x.���̿���� " & _
    "   LEFT JOIN  erpdata .. tblTSVworkorder xx ON xx.ORDERNAME = x.�󹤵�   LEFT JOIN erptemp..customer_information yy  ON yy.CUSTOMER = d.�ͻ����� AND yy.SHIP_TO = d.������ַ   Where  a.���۵���� IN  ('" & strTDJBH & "') GROUP BY a.���۵����,a.�ͻ�����,c.�ͻ�����,a.��������,a.�Ƶ�Ա,ISNULL(f.�˵���,''),RTRIM(b.���ݱ��) " & _
    "   ,e.PO_NUM,e.PO_ITEM, e.CUSTOMERSHORTNAME , e.MPN_DESC,b.�Ϻ�,b.С���,b.�����,d.������ ,ISNULL(cc.���,bb.���),x.�ϸ���,XX.MPN,ad.remark2 ,cc.����,cc.�ߴ�,yy.SHIP_TO_AD  order by a.���۵����,RTRIM(b.���ݱ��),b.�����,b.С���"
    End If
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql2, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
     With Fps(0)
            .MaxRows = 0
            Set .DataSource = rs
            .MaxRows = rs.RecordCount
        End With
        Fra(1).Caption = Trim$(Cob(0).text)
    
    
    
    
       
'    Set adoCmd = New ADODB.Command
'    Set adoCmd.ActiveConnection = INIadoCon
'    adoCmd.CommandText = "erptemp..usGW_ReportSearch"
'    '       adoCmd.CommandText = "erpdata..uspGW_ReportSearch"
'    adoCmd.Parameters.Refresh
'    adoCmd.CommandType = adCmdStoredProc
'
'    Set adoPrmReturn = New ADODB.Parameter         '����ִ�гɹ����
'    adoPrmReturn.Type = adInteger
'    adoPrmReturn.Direction = adParamReturnValue
'    adoCmd.Parameters.Append adoPrmReturn
'
'    Set adoprmFG = New ADODB.Parameter             '��ѯ���
'    adoprmFG.Type = adInteger
'    adoprmFG.Direction = adParamInput
'    adoprmFG.Value = Cob(0).ListIndex
'    adoCmd.Parameters.Append adoprmFG
'
'    Set adoprm2 = New ADODB.Parameter             '�ͻ�����
'    adoprm2.Type = adVarChar
'    adoprm2.Size = 20
'    adoprm2.Direction = adParamInput
'    adoprm2.Value = Trim(Cob(1).Text)
'    adoCmd.Parameters.Append adoprm2
'
'    Set adoPrm3 = New ADODB.Parameter             '���߱��
'    adoPrm3.Type = adInteger
'    adoPrm3.Direction = adParamInput
'    adoPrm3.Value = Val(Trim(Cob(2).Text))
'    adoCmd.Parameters.Append adoPrm3
'
'    Set adoPrm7 = New ADODB.Parameter              '���۵����
'    adoPrm7.Type = adVarChar
'    adoPrm7.Size = 1000
'    adoPrm7.Direction = adParamInput
'    adoPrm7.Value = strTDJBH 'IIf(strdj = "", Trim(Cob(3).Text), Replace(strdj, ",", "','")) 'ע�ⵥ�ݱ�ŵ��ж�
'    adoCmd.Parameters.Append adoPrm7
'
'    Set adoPrm6 = New ADODB.Parameter             '�Ϻ�
'    adoPrm6.Type = adVarChar
'    adoPrm6.Size = 50
'    adoPrm6.Direction = adParamInput
'    adoPrm6.Value = Trim(txt(0).Text)
'    adoCmd.Parameters.Append adoPrm6
'
'    Set adoPrm8 = New ADODB.Parameter             '��ʼ����
'    adoPrm8.Type = adVarChar
'    adoPrm8.Size = 20
'    adoPrm8.Direction = adParamInput
'    adoPrm8.Value = Format(Trim(DTP(0).Value), "YYYY-MM-DD")
'    adoCmd.Parameters.Append adoPrm8
'    Set adoPrm9 = New ADODB.Parameter             '��������
'    adoPrm9.Type = adVarChar
'    adoPrm9.Size = 20
'    adoPrm9.Direction = adParamInput
'    adoPrm9.Value = Format(Trim(DTP(1).Value), "YYYY-MM-DD")
'    adoCmd.Parameters.Append adoPrm9
'
'    Set adoprm10 = New ADODB.Parameter             '���������
'    adoprm10.Type = adVarChar
'    adoprm10.Size = 50
'    adoprm10.Direction = adParamInput
'    adoprm10.Value = Trim(txt(1).Text)
'    adoCmd.Parameters.Append adoprm10
'
'    Set adoPrm11 = New ADODB.Parameter             '�ǳɱ���
'    adoPrm11.Type = adInteger
'    adoPrm11.Direction = adParamInput
'    adoPrm11.Value = chk.Value
'    adoCmd.Parameters.Append adoPrm11
'
'
'    Set rs = adoCmd.Execute
''    rs.Sort = "item ASC"
''    If Cob(1).Text <> "SG005" Then
''        rs.Sort = "MPN_DESC,������ ASC"
''    Else
'     rs.Sort = "MPN_DESC,С��� ASC"
''    End If
''
''    rs.CursorLocation = 3
'    If strdj <> "" Then '��ֵ�˵���
'
'        With Fps(0)
'
'            For i = 1 To .MaxRows
'                .Row = i
'                .Col = fpSDetail.e_DJBH
'
'                If InStr(strdj, Trim$(.Text)) > 0 Then
'
'                    .SetText fpSDetail.e_YDH, i, Trim$("" & rs!�˵���)
'
'                End If
'
'            Next
'
'        End With
'
'        Exit Sub
'
'    End If
'
'    '��ѯ�����ݸ�ֵ��Fps����¼���ĵ�������
'    If adoPrmReturn.Value = 0 Then
'
'        '����������
'        With Fps(0)
'            .MaxRows = 0
'            Set .DataSource = rs
'            .MaxRows = rs.RecordCount
'
'        End With
'
'        Fra(1).Caption = Trim$(Cob(0).Text)
'    Else
'        GoTo ErrHandle
'
'    End If
  End If
    Exit Sub
ErrHandle:
    MsgBox "ִ��ʧ�ܣ�" + Chr(13) + "ԭ��:" + Err.DESCRIPTION, vbInformation, Me.Caption

End Sub

'Invoice
Public Sub InvoiceExportPrintExcel(ByVal rs As ADODB.Recordset, ByVal strdj As String)

    Dim strSql         As String

    Dim lngRows        As Long

    Dim rsQuery        As Excel.QueryTable

    'Dim ClsP                As New ClsProgress
    Dim ExApp          As Excel.Application

    Dim wkbk           As New Workbook

    Dim wkst           As New Worksheet
    
    Dim i              As Long

    Dim j              As Long
    
    Dim m              As Long
    
    Dim N              As Long

    Dim IntCols        As Integer

    Dim strCols        As String

    Dim strFileName    As String

    Dim IntInertRow    As Integer, IntMaxDetailRow As Integer

    Dim DblNum         As Double
    
    Dim DblNum2        As Double
    
    Dim DblPnum        As Double

    Dim DblAmt         As Double '�ܽ��
    
    Dim DblPamt        As Double

    Dim DblWamt        As Double

    Dim RsNew          As New ADODB.Recordset

    Dim RsNew1         As New ADODB.Recordset

    Dim strShipTo()    As String

    Dim strSoldBy()    As String

    Dim Specification1 As String

    Dim waferid1       As String
    
    Dim Fcount         As Long

    Dim Fcount1        As Long
    
    Dim b()            As String
    
    Dim acpn           As String
    
    Dim asum           As Integer
       
    Dim gdh As String
                    
    Dim ngdh As String
    
    Dim ShipOrder      As String
    
    Dim S_I            As Integer
    
    Dim ShipOrderFlag  As Boolean
    
    Dim DieNoFound     As Boolean
    
    Dim TOTALWAFER As Integer
    
    Dim rs075           As New ADODB.Recordset
    
    Dim strSql075 As String
    
    Dim waferqty075 As Integer
    
    Dim strPONUM_075 As String
    Dim strLineitem_075 As String
    
    Dim DblQty_075 As Long
    Dim DblAmount_075 As Double
    Dim DblDieQty_075 As Long
    Dim strlot_075 As String
    Dim strpn_075 As String
    Dim strMPN_DESC_075 As String
    Dim strNC_075 As String
    Dim strprice_075  As String
    Dim strsono_SG005 As String
    
    Dim strPONUM_TY As String
    Dim strMPN_DESC_TY As String
    Dim strpn_TY As String
    Dim strLot_TY As String
    Dim strprice_TY As String
    Dim DblDieQty_TY As Long
    Dim DblAmount_TY As Double
    Dim strsql_Getnewlotid As String
    
    
    waferqty075 = 0
    
    If rs.RecordCount <= 0 Then
        MsgBox "û��Ҫ���������ϣ�", vbInformation, "��ʾ��"
        Exit Sub

    End If

    '    ClsP.Init 100, True
    '    ClsP.ShowProgress 10, "��ʼ������..."
    strSysPath = App.Path

    Select Case UCase(Combo1.text)
        
        Case "68"
        
            strFileName = strSysPath & "\TEMPLET\68_Invoice.xls" 'Ҫ�򿪵��ļ�
            
        Case "76", "US026"
        
            strFileName = strSysPath & "\TEMPLET\76_Invoice.xls" 'Ҫ�򿪵��ļ�
        Case "SG005"
        
            strFileName = strSysPath & "\TEMPLET\SG005_Invoice.xls" 'Ҫ�򿪵��ļ�
        Case "SG005_SO"
        
            strFileName = strSysPath & "\TEMPLET\SG005_SO_Invoice.xls" 'Ҫ�򿪵��ļ�
        Case "TW067", "ͨ��ģ��"
        
            strFileName = strSysPath & "\TEMPLET\TW067_Invoice.xls" 'Ҫ�򿪵��ļ�
            
        Case "BD", "EQ"
        
            strFileName = strSysPath & "\TEMPLET\BD_Invoice.xls" 'Ҫ�򿪵��ļ�
            
        Case "HK005"
        
            strFileName = strSysPath & "\TEMPLET\HK005_Invoice.xls" 'Ҫ�򿪵��ļ�
            
        Case "HK075"
        
            strFileName = strSysPath & "\TEMPLET\HK075_Invoice.xls" 'Ҫ�򿪵��ļ�
            
        Case "HK080"
        
            strFileName = strSysPath & "\TEMPLET\HK080_Invoice.xls" 'Ҫ�򿪵��ļ�
            
        Case "GC"

            strFileName = strSysPath & "\TEMPLET\GC_Invoice.xls" 'Ҫ�򿪵��ļ�

    End Select

    If rs.RecordCount > 0 Then
        '        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
        
        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblNum2 = 0
        DblPnum = 0
        DblAmt = 0
        
        DblPamt = 0
        DblWamt = 0
        '��ֵ��Excel�У���ͷ
        
        Select Case UCase(Combo1.text)
            
            Case "68"
                
                wkst.Cells(3, 8) = strdj   'Trim$("" & Rs!���۵����)
                wkst.Cells(3, 11) = DATE
            
            Case "TW067", "ͨ��ģ��"
                
                wkst.Cells(3, 7) = strdj   'Trim$("" & Rs!���۵����)
                wkst.Cells(3, 10) = DATE
            
            Case "76", "US026", "GC"
                
                wkst.Cells(3, 5) = strdj   'Trim$("" & Rs!���۵����)
                wkst.Cells(3, 8) = DATE
                
            Case "SG005"
            
                wkst.Cells(3, 5) = strdj   'Trim$("" & Rs!���۵����)
                wkst.Cells(3, 7) = DATE
                
            Case "SG005_SO"
            
                wkst.Cells(3, 7) = strdj   'Trim$("" & Rs!���۵����)
                wkst.Cells(3, 9) = DATE
                
            Case "BD", "EQ"
                
                wkst.Cells(3, 6) = strdj   'Trim$("" & Rs!���۵����)
                wkst.Cells(3, 9) = DATE
                
            Case "HK005"
                
                wkst.Cells(3, 8) = strdj   'Trim$("" & Rs!���۵����)
                wkst.Cells(3, 13) = DATE
                                                       
        End Select

        wkst.Cells(7, 1) = "CONTACT:" & Trim(strUserName)
        wkst.Cells(10, 10) = Trim$("" & rs!�˵���)
        '��ѯ���ų�����ַ------------------------------------------
        strSql = "SELECT DISTINCT SHIP_TO_AD,SOLD_BY,SHIP_TO FROM erpdata..Vw_CustomerShipTo WHERE ���۵���� IN('" & Replace(strdj, ",", "','") & "')"

        If RsNew.State = adStateOpen Then RsNew.Close
        RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

        If RsNew.RecordCount > 0 Then

            'ShipTo
            If InStr(Trim$("" & RsNew!SHIP_TO_AD), "@") > 0 Then
                strShipTo = Split(Trim$("" & RsNew!SHIP_TO_AD), "@")
                
                For i = 0 To UBound(strShipTo)

                    If i + 9 > 14 Then Exit For
                    wkst.Cells(i + 9, 1) = strShipTo(i)
                    If UCase(Combo1.text) = "SG005" Or UCase(Combo1.text) = "SG005_SO" Then
                       If i + 9 = 11 Then
                             wkst.Cells(11, 1) = Mid(strShipTo(1), 54, 24)
                       End If
                       If i + 9 = 10 Then
                             wkst.Cells(10, 1) = Mid(strShipTo(i), 1, 53)
                       End If
                    End If
                Next
            Else
                wkst.Cells(9, 1) = Trim$("" & RsNew!SHIP_TO_AD)

            End If

            'SoldBy
            If InStr(Trim$("" & RsNew!SOLD_BY), "@") > 0 Then
                strSoldBy = Split(Trim$("" & RsNew!SOLD_BY), "@")

                For i = 0 To UBound(strSoldBy)

                    If i + 16 >= 19 Then
                        'Exit For
                        wkst.Cells(18, 1) = wkst.Cells(18, 1) & " " & strSoldBy(i)
                    Else
                       wkst.Cells(i + 16, 1) = strSoldBy(i)
                    End If
                Next
            Else
                wkst.Cells(16, 1) = Trim$("" & RsNew!SOLD_BY)

            End If
            
            Select Case UCase(Combo1.text)
        
                Case "68", "US026"
            
                    wkst.Cells(8, 8) = "    TO: " & Trim$("" & RsNew!SHIP_TO)
                    wkst.Cells(3, 9) = Trim$("" & rs!��������)
'                Case "SG005"
'
'                    wkst.Cells(11, 1) = Trim$("" & RsNew!ship_to)
                Case "BD", "EQ", "HK080"
             
                    wkst.Cells(8, 7) = Trim$("" & RsNew!SHIP_TO)
                    wkst.Cells(3, 6) = Trim("" & rs!���۵����)
                    wkst.Cells(3, 8) = Trim$("" & rs!��������)

                 Case "HK075"
             
                    wkst.Cells(8, 8) = Trim$("" & RsNew!SHIP_TO)
                    wkst.Cells(3, 7) = Trim("" & rs!���۵����)
                    wkst.Cells(3, 9) = Trim$("" & rs!��������)
                Case "GC"
                    
                    'wkst.Cells(6, 6) = Trim$("" & RsNew!SHIP_TO)
                    
                Case "HK005"
                    
                    wkst.Cells(8, 8) = "    TO: " & Trim$("" & RsNew!SHIP_TO)
                    wkst.Cells(3, 12) = Trim$("" & rs!��������)
                    
                Case "76"
             
                    wkst.Cells(8, 6) = Trim$("" & RsNew!SHIP_TO)
                    wkst.Cells(3, 6) = Trim("" & rs!���۵����)
                    wkst.Cells(3, 7) = Trim$("" & rs!��������)
                    
                    
            End Select
        
        Else
            wkst.Cells(9, 1) = ""
            wkst.Cells(10, 1) = ""
            wkst.Cells(11, 1) = ""
            wkst.Cells(12, 1) = ""
            wkst.Cells(13, 1) = ""
            wkst.Cells(14, 1) = ""
            wkst.Cells(16, 1) = ""
            wkst.Cells(17, 1) = ""
            wkst.Cells(18, 1) = ""
            wkst.Cells(19, 1) = ""

        End If

        RsNew.Close

        '----------------------------------------------------------
        If UCase(Combo1.text) <> "GC" Then
            lngRows = 21
        Else
            lngRows = 24
        End If
        
        If UCase(Combo1.text) <> "GC" And UCase(Combo1.text) <> "HK075" And UCase(Combo1.text) <> "ͨ��ģ��" Then
            IntInertRow = rs.RecordCount

            For i = 1 To IntInertRow - 1
                wkst.Rows(lngRows & ":" & lngRows).Select
                ExApp.Selection.Copy
                ExApp.Selection.Insert Shift:=xlDown
            Next i

        End If

        IntMaxDetailRow = rs.RecordCount
        If UCase(Combo1.text) = "GC" Then
            Specification1 = "оƬ"
        Else
            Specification1 = "Integrated Circuit chip"
        End If
        waferid1 = ""
        '        ClsP.ShowProgress 50, "���ڵ���..."
        Dim T As Integer
        T = 0
        Select Case UCase(Combo1.text)
            
            Case "68"

                Do While Not rs.EOF
                    
                    T = T + 1
                    wkst.Cells(lngRows, 1) = Trim$("" & T)

                    wkst.Cells(lngRows, 2) = Trim$("" & rs!PO_NUM)
                    wkst.Cells(lngRows, 3) = Trim$("" & rs!MPN_DESC)
                    If (InStr(rs!�Ϻ�, "B") < 8) Then
                        wkst.Cells(lngRows, 4) = Trim$("" & Mid(rs!�Ϻ�, 3, 7))
                    Else
                    wkst.Cells(lngRows, 4) = Trim$("" & Mid(rs!�Ϻ�, 3, InStr(rs!�Ϻ�, "B") - 2))
                    End If
                   
                    wkst.Cells(lngRows, 5) = Trim$("" & rs!�Ϻ�)
                    wkst.Cells(lngRows, 6) = Trim$("" & rs!������)
                   
                    strSql = "select RIGHT(Replace(rtrim(���̿����),'+',''),2) as waferid FROM erpdata..tblstockmovesub x left join erptemp..mps_mark ad on ad.wafer_id = x.���̿���� where  x.�Ϻ� = '" & Trim(rs!�Ϻ�) & "' and  ad.REMARK2 = '" & rs!������ & "' and x.���ݱ�� = '" & Trim(rs!��������) & "' order by RIGHT(Replace(rtrim(x.���̿����),'+',''),2) "
                   
                    If RsNew.State = adStateOpen Then RsNew.Close
                    RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
                    
                    If RsNew.RecordCount = 0 Then
                     '   strSql = "select RIGHT(Replace(rtrim(���̿����),'+',''),2) as waferid FROM erpdata..tblstockmovesub x  where  �Ϻ� = '" & Trim(rs!�Ϻ�) & "' and  ������ = '" & rs!������ & "'and ���ݱ�� = '" & Trim(rs!��������) & "'and ��� = '" & Trim(rs!С���) & "' order by RIGHT(Replace(rtrim(���̿����),'+',''),2) "
                        strSql = "select RIGHT(Replace(rtrim(���̿����),'+',''),2) as waferid FROM erpdata..tblstockmovesub x  where  �Ϻ� = '" & Trim(rs!�Ϻ�) & "' and  ������ = '" & rs!������ & "'and ���ݱ�� = '" & Trim(rs!��������) & " ' order by RIGHT(Replace(rtrim(���̿����),'+',''),2) "

                        If RsNew.State = adStateOpen Then RsNew.Close
                        RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
                    End If
                    RsNew.MoveFirst

                    For i = 1 To RsNew.RecordCount

                        If i = 1 Then
                
                            waferid1 = "#" & RsNew("waferid")
                
                        Else
                    
                            waferid1 = waferid1 & " " & RsNew("waferid")
                    
                        End If
            
                        RsNew.MoveNext
            
                    Next
                    
                    If RsNew.RecordCount = 25 Then
        
                        waferid1 = "#" & "01-25"
        
                    End If
                    gdh = Trim$("" & rs!������)
                    If gdh <> ngdh Then
                        ngdh = Trim$("" & rs!������)
                         Fcount = 0
        
                        Fcount = RsNew.RecordCount

                        RsNew.Clone

                        Set RsNew = Nothing
                        
                        wkst.Cells(lngRows, 7) = waferid1

                        wkst.Cells(lngRows, 8) = Specification1
                        wkst.Cells(lngRows, 9) = Trim$("" & rs!����)
                        wkst.Cells(lngRows, 10) = Fcount
                        wkst.Cells(lngRows, 11) = Trim$("" & rs!����)
                        wkst.Cells(lngRows, 12) = Trim$("" & rs!���)
                
                        Fcount1 = Fcount1 + Fcount
                        DblNum = DblNum + Val(Trim$("" & rs!����))
                        DblAmt = DblAmt + Val(Trim$("" & rs!���))
                        
                        DblPamt = DblPamt + Val(Trim$("" & rs!�ӹ��ѽ��))
                        DblWamt = DblWamt + Val(Trim$("" & rs!�͹��Ͻ��))
                        
                        lngRows = lngRows + 1
                       
                        rs.MoveNext
                    Else
                        wkst.Range(Chr(7 + 64) & lngRows - 1 & ":" & Chr(7 + 64) & lngRows).Merge
                        wkst.Range(Chr(10 + 64) & lngRows - 1 & ":" & Chr(10 + 64) & lngRows).Merge
                         Fcount = 0
        
                    Fcount = RsNew.RecordCount

                    RsNew.Clone

                    Set RsNew = Nothing
                    
                    wkst.Cells(lngRows, 7) = waferid1

                    wkst.Cells(lngRows, 8) = Specification1
                    wkst.Cells(lngRows, 9) = Trim$("" & rs!����)
                    wkst.Cells(lngRows, 10) = Fcount
                    wkst.Cells(lngRows, 11) = Trim$("" & rs!����)
                    wkst.Cells(lngRows, 12) = Trim$("" & rs!���)
            
                    Fcount1 = Fcount1
                    DblNum = DblNum + Val(Trim$("" & rs!����))
                    DblAmt = DblAmt + Val(Trim$("" & rs!���))
                    
                    DblPamt = DblPamt + Val(Trim$("" & rs!�ӹ��ѽ��))
                    DblWamt = DblWamt + Val(Trim$("" & rs!�͹��Ͻ��))
                    
                    lngRows = lngRows + 1
                   
                    rs.MoveNext
                    End If
                    
                   
           
                Loop
                
                '�������
                wkst.Cells(lngRows, 9) = DblNum & "PCS"
        
                wkst.Cells(lngRows, 10) = Fcount1 & "Ƭ"
        
                wkst.Cells(lngRows, 12) = "US$" & DblAmt
                'wkst.Cells(17, 7) = "Process Amount :US$ " & Format(DblPamt, "0.00")
                'wkst.Cells(18, 7) = "Wafer Amount :US$ " & Format(DblWamt, "0.00")
                wkst.Cells(16, 9) = "$" & Format(DblPamt, "0.00")
                wkst.Cells(17, 9) = "$" & Format(DblWamt, "0.00")
            Case "HK005"
            
                Do While Not rs.EOF
        
                     T = T + 1
                    wkst.Cells(lngRows, 1) = Trim$("" & T)
                    wkst.Cells(lngRows, 2) = Trim$("" & rs!PO_NUM)
                    wkst.Cells(lngRows, 3) = Trim$("" & rs!MPN_DESC)
                     If (InStr(rs!�Ϻ�, "B") < 11) Then
                        wkst.Cells(lngRows, 4) = Trim$("" & Mid(rs!�Ϻ�, 3, 9))
                    Else
                    wkst.Cells(lngRows, 4) = Trim$("" & Mid(rs!�Ϻ�, 3, InStr(rs!�Ϻ�, "B") - 2))
                    End If
                    wkst.Cells(lngRows, 5) = Trim$("" & rs!�Ϻ�)
                    wkst.Cells(lngRows, 6) = Trim$("" & rs!������)
     
                    strSql = "select RIGHT(Replace(rtrim(���̿����),'+',''),2) as waferid FROM erpdata..tblstockmovesub where  �Ϻ� = '" & Trim(rs!�Ϻ�) & "' and  ������ = '" & Trim(rs!������) & "'and ���ݱ�� = '" & Trim(rs!��������) & "' order by RIGHT(Replace(rtrim(���̿����),'+',''),2)"
  
                    If RsNew.State = adStateOpen Then RsNew.Close

                    RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                    RsNew.MoveFirst

                    For i = 1 To RsNew.RecordCount

                        If i = 1 Then
                
                            waferid1 = RsNew("waferid")
                
                        Else
                    
                            waferid1 = waferid1 & "," & RsNew("waferid") & ","
                    
                        End If
            
                        RsNew.MoveNext
            
                    Next
        
                    If RsNew.RecordCount = 25 Then
        
                        waferid1 = "#" & "1-25"
        
                    End If
            
                    Fcount = 0
        
                    Fcount = RsNew.RecordCount

                    RsNew.Clone
 
                    Set RsNew = Nothing

                    wkst.Cells(lngRows, 7) = waferid1

                    wkst.Cells(lngRows, 8) = Specification1
                    wkst.Cells(lngRows, 9) = Trim$("" & rs!����)
                    
                    strSql = " SELECT isnull(sum(D.Dies),0) as ��Ʒ�� from (SELECT t.lot_id AS '������', ISNULL(ISNULL(t.BIN1, t.A), K.NDPW) as 'Dies'  " & _
                        " FROM ( SELECT 'HTKS' AS sub_name, d.SHIP_SITE, RTRIM(d.FAB_CONV_ID) as FAB_CONV_ID,  a.cust_device, a.gcversion,d.PO_NUM, " & _
                        " a.create_date, rtrim(a.lot_id) as lot_id, SUBSTRING(REPLACE(b.���̿����, '+', ''), LEN(a.lot_id) + 1, 2) as waferid, " & _
                        " c.FAILBINCOUNT + c.PASSBINCOUNT AS gross_die, CASE WHEN n.KEY_NAME in ( 'BAD1_DIE','BAD2_DIE') THEN 'E'  ELSE 'A' END Grade, " & _
                        " CONVERT(INT,n.KEY_VALUE ) AS qty,  c.PRODUCTID, rtrim(ay.���) as ���,  b.�󹤵�,  a.qbox, b.���̿����, SUBSTRING(ee.SFC_ID, 12, 8) AS SFC   " & _
                        " FROM erptemp .. tblshipreport_new a  " & _
                        " INNER JOIN erpdata .. tblStockNumTree ax  ON ax.��� = a.qbox " & _
                        " INNER JOIN erpdata .. tblStockNumTree ay ON ay.��� = ax.�ϼ����   " & _
                        " INNER JOIN erpdata .. tblStocksqfhsub b ON b.���ݱ�� = a.ship_order  AND b.��� = a.qbox   AND b.������ = a.lot_id " & _
                        " INNER JOIN ERPBASE .. tblmappingData c  ON c.SUBSTRATEID = b.���̿����  AND c.LOTID = b.������ " & _
                        " INNER JOIN erpbase .. tblCustomerOI d  ON CONVERT(VARCHAR(20), CONVERT(int,d.ID)) = c.FILENAME  AND d.SOURCE_BATCH_ID = c.LOTID   " & _
                        " LEFT JOIN  erpdata..tblErpInStockDetailInfo m ON m.KEY_VALUE = b.���  " & _
                        " LEFT JOIN  erpdata..tblErpInStockDetailInfo n  ON n.BOX_ID = m.BOX_ID  and n.KEY_NAME in ( 'BAD1_DIE','BAD2_DIE','GOOD_DIE') and n.KEY_TYPE = 'WAFER' AND   CHARINDEX(c.SUBSTRATEID , n.KEYID ) <> 0  " & _
                        " INNER JOIN erpdata .. tblErpInStockRelation ee ON    ee.BOX_ID = n.BOX_ID AND  ee.WAFER_ID = n.KEYID  WHERE a.ship_order = '" & Trim$(rs!��������) & "' and a.lot_id = '" & Trim$(rs!������) & "' )  AS p  PIVOT(sum(qty) FOR Grade IN(A,BIN1, E)) AS T " & _
                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV k  ON k.QBOXNUMBER = t.qbox AND k.WAFERSCRIBENUMBER = t.���̿���� AND k.CONTAINERNAME LIKE '%-A' " & _
                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV L ON L.QBOXNUMBER = t.qbox AND L.WAFERSCRIBENUMBER = t.���̿���� AND L.CONTAINERNAME LIKE '%-A-01'  " & _
                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV m  ON m.QBOXNUMBER = t.qbox  AND m.WAFERSCRIBENUMBER = t.���̿���� AND m.CONTAINERNAME LIKE '%-A-02' " & _
                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV n  ON n.QBOXNUMBER = t.qbox AND n.WAFERSCRIBENUMBER = t.���̿���� AND n.CONTAINERNAME LIKE '%-E') D "
                    
                    wkst.Cells(lngRows, 10) = Get_SqlStr(strSql)
                    DblNum2 = DblNum2 + Get_SqlStr(strSql)
                    
                    wkst.Cells(lngRows, 11) = Fcount
                    wkst.Cells(lngRows, 12) = Trim$("" & rs!����)
                    wkst.Cells(lngRows, 13) = Trim$("" & rs!���)
            
                    Fcount1 = Fcount1 + Fcount
                    DblNum = DblNum + Val(Trim$("" & rs!����))
                    DblAmt = DblAmt + Val(Trim$("" & rs!���))
                    
                    DblPamt = DblPamt + Val(Trim$("" & rs!�ӹ��ѽ��))
                    DblWamt = DblWamt + Val(Trim$("" & rs!�͹��Ͻ��))
                    
                    lngRows = lngRows + 1
                   
                    rs.MoveNext
           
                Loop
        
                '�������
                wkst.Cells(lngRows, 9) = DblNum & "PCS"
                
                wkst.Cells(lngRows, 10) = DblNum2 & "PCS"
                
                wkst.Cells(lngRows, 11) = Fcount1 & "Ƭ"
        
                wkst.Cells(lngRows, 13) = "US$" & DblAmt
                'wkst.Cells(17, 7) = "Process Amount :US$ " & Format(DblPamt, "0.00")
                'wkst.Cells(18, 7) = "Wafer Amount :US$ " & Format(DblWamt, "0.00")
                wkst.Cells(16, 9) = "$" & Format(DblPamt, "0.00")
                wkst.Cells(17, 9) = "$" & Format(DblWamt, "0.00")
                 
            Case "HK075"
                wkst.Cells(lngRows - 1, 1) = "Item"
                wkst.Cells(lngRows - 1, 2) = "Customer PO No"
                wkst.Cells(lngRows - 1, 3) = "Line Item"
                wkst.Cells(lngRows - 1, 4) = "12 NC"
                wkst.Cells(lngRows - 1, 5) = "Customer P/N"
                wkst.Cells(lngRows - 1, 6) = "HT P/N"
                wkst.Cells(lngRows - 1, 7) = "Wafer Lot No"
                wkst.Cells(lngRows - 1, 8) = "Specification"
                wkst.Cells(lngRows - 1, 9) = "Die Qty(PCS)"
                wkst.Cells(lngRows - 1, 10) = "Qty(Ƭ)"
                wkst.Cells(lngRows - 1, 11) = "Unit Price"
                wkst.Cells(lngRows - 1, 12) = "Amount"
                
                strPONUM_075 = ""
                strLineitem_075 = ""
                DblAmount_075 = 0
                DblQty_075 = 0
                DblDieQty_075 = 0
                Do While Not rs.EOF
                    strSql075 = "select count(RIGHT(Replace(rtrim(���̿����),'+',''),2)) as WQty " & _
                            "FROM erpdata..tblstockmovesub x  where  �Ϻ� = '" & Trim(rs!�Ϻ�) & "'and  ������ = '" & Trim(rs!������) & "'and " & _
                            "���ݱ�� = '" & Trim(rs!��������) & "'"
                    
                    If INIadoCon.State <> adStateOpen Then
                        INIConnectSTART2
                    End If
                    rs075.Open strSql075, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
                    waferqty075 = waferqty075 + rs075!WQTY

                    If Trim(rs!PO_NUM) = strPONUM_075 And Trim(rs!Line_item) = strLineitem_075 And Trim(rs!NC) = strNC_075 And Trim(rs!MPN_DESC) = strMPN_DESC_075 And Trim(rs!�Ϻ�) = strpn_075 And Trim(rs!������) = strlot_075 And Trim(rs!����) = strprice_075 Then
                        DblDieQty_075 = DblDieQty_075 + Val(Trim$("" & rs!����))
                        DblQty_075 = DblQty_075 + Val(Trim$("" & rs075!WQTY))
                        DblAmount_075 = DblAmount_075 + Val(Trim$("" & rs!���))
                        wkst.Cells(lngRows, 9) = DblDieQty_075
                        wkst.Cells(lngRows, 10) = DblQty_075
                        wkst.Cells(lngRows, 12) = DblAmount_075
                        
                    Else
                        strPONUM_075 = Trim(rs!PO_NUM)
                        strLineitem_075 = Trim(rs!Line_item)
                        strlot_075 = Trim(rs!������)
                        strpn_075 = Trim(rs!�Ϻ�)
                        strMPN_DESC_075 = Trim(rs!MPN_DESC)
                        strNC_075 = Trim(rs!NC)
                        strprice_075 = Trim(rs!����)

                        DblDieQty_075 = Val(Trim$("" & rs!����))
                        DblQty_075 = Val(Trim$("" & rs075!WQTY))
                        DblAmount_075 = Val(Trim$("" & rs!���))
                        
                        If T > 0 Then
                            wkst.Rows(lngRows & ":" & lngRows).Select
                            ExApp.Selection.Copy
                            ExApp.Selection.Insert Shift:=xlDown
                            lngRows = lngRows + 1
                        End If
                        T = T + 1
                        wkst.Cells(lngRows, 1) = Trim$("" & T)
                        wkst.Cells(lngRows, 2) = "'" & Trim$("" & rs!PO_NUM)
                        wkst.Cells(lngRows, 3) = "'" & Trim$("" & rs!Line_item)
                        wkst.Cells(lngRows, 4) = "'" & Trim$("" & rs!NC)
                        wkst.Cells(lngRows, 5) = "'" & Trim$("" & rs!MPN_DESC)
                        wkst.Cells(lngRows, 6) = "'" & Trim$("" & rs!�Ϻ�)
                        wkst.Cells(lngRows, 7) = "'" & Trim$("" & rs!������)
                        wkst.Cells(lngRows, 8) = Specification1
                        wkst.Cells(lngRows, 9) = Trim$("" & rs!����)
                        wkst.Cells(lngRows, 10) = Trim$("" & rs075!WQTY)
                        wkst.Cells(lngRows, 11) = Trim$("" & rs!����)
                        wkst.Cells(lngRows, 12) = Val(Trim$("" & rs!���))
                        
                    End If
                    DblNum = DblNum + Val(Trim$("" & rs!����))
                    DblAmt = DblAmt + Val(Trim$("" & rs!���))
                    DblPamt = DblPamt + Val(Trim$("" & rs!�ӹ��ѽ��))
                    DblWamt = DblWamt + Val(Trim$("" & rs!�͹��Ͻ��))
                
                    
                    rs075.Close
                    rs.MoveNext
                    
                Loop
        
                '�������
                lngRows = lngRows + 1
                wkst.Cells(lngRows, 9) = DblNum & "PCS"
                wkst.Cells(lngRows, 10) = waferqty075 & "Ƭ"
                wkst.Cells(lngRows, 12) = "$" & Format(DblAmt, "0.00")
           
                wkst.Cells(17, 7) = "Process Amount :"
                wkst.Cells(18, 7) = "Wafer Amount :"
                wkst.Cells(17, 8) = "$" & Format(DblPamt, "0.00")
                wkst.Cells(18, 8) = "$" & Format(DblWamt, "0.00")
                ExApp.Visible = True
            Case "HK080"
                
                Do While Not rs.EOF
        
                     T = T + 1
                    wkst.Cells(lngRows, 1) = Trim$("" & T)
                    wkst.Cells(lngRows, 2) = Trim$("" & rs!PO_NUM)
                    wkst.Cells(lngRows, 3) = Trim$("" & rs!MPN_DESC)
                    wkst.Cells(lngRows, 4) = Trim$("" & rs!�Ϻ�)
                    wkst.Cells(lngRows, 5) = Trim$("" & rs!������)
                    
                    wkst.Cells(lngRows, 6) = Specification1
                    wkst.Cells(lngRows, 7) = Trim$("" & rs!����)
                    wkst.Cells(lngRows, 8) = Trim$("" & rs!����)
                    wkst.Cells(lngRows, 9) = Trim$("" & rs!���)

                    DblNum = DblNum + Val(Trim$("" & rs!����))
                    DblAmt = DblAmt + Val(Trim$("" & rs!���))
                    
                    DblPamt = DblPamt + Val(Trim$("" & rs!�ӹ��ѽ��))
                    DblWamt = DblWamt + Val(Trim$("" & rs!�͹��Ͻ��))
                    
                    lngRows = lngRows + 1
                    rs.MoveNext

                Loop
        
                '�������
                wkst.Cells(lngRows, 7) = DblNum & "PCS"
        
                wkst.Cells(lngRows, 9) = DblAmt
                
                wkst.Cells(17, 6) = "Process Amount :"
                wkst.Cells(18, 6) = "Wafer Amount :"
                wkst.Cells(17, 7) = "$" & Format(DblPamt, "0.00")
                wkst.Cells(18, 7) = "$" & Format(DblWamt, "0.00")
            
            Case "TW067"
                
                Do While Not rs.EOF
                    
        
                     T = T + 1
                    wkst.Cells(lngRows, 1) = Trim$("" & T)
                    wkst.Cells(lngRows, 2) = Trim$("" & rs!PO_NUM)
                    wkst.Cells(lngRows, 3) = Trim$("" & rs!MPN_DESC)
                    wkst.Cells(lngRows, 4) = Trim$("" & Mid(rs!�Ϻ�, 3, 9))
                    wkst.Cells(lngRows, 5) = Trim$("" & rs!�Ϻ�)
                    wkst.Cells(lngRows, 6) = Trim$("" & rs!������)
                    
                    wkst.Cells(lngRows, 7) = Specification1
                    wkst.Cells(lngRows, 8) = Trim$("" & rs!����)
                    wkst.Cells(lngRows, 9) = Trim$("" & rs!����)
                    wkst.Cells(lngRows, 10) = Trim$("" & rs!���)

                    DblNum = DblNum + Val(Trim$("" & rs!����))
                    DblAmt = DblAmt + Val(Trim$("" & rs!���))
                    
                    DblPamt = DblPamt + Val(Trim$("" & rs!�ӹ��ѽ��))
                    DblWamt = DblWamt + Val(Trim$("" & rs!�͹��Ͻ��))
                    
                    lngRows = lngRows + 1
                    rs.MoveNext
                Loop
        
                '�������
                wkst.Cells(lngRows, 8) = DblNum & "PCS"
        
                wkst.Cells(lngRows, 10) = DblAmt
                
                wkst.Cells(17, 7) = "Process Amount :"
                wkst.Cells(18, 7) = "Wafer Amount :"
                wkst.Cells(17, 8) = "$" & Format(DblPamt, "0.00")
                wkst.Cells(18, 8) = "$" & Format(DblWamt, "0.00")
                
            Case "ͨ��ģ��"
                
                Do While Not rs.EOF
            
                    If Trim(rs!PO_NUM) = strPONUM_TY And Trim(rs!MPN_DESC) = strMPN_DESC_TY And Trim(rs!�Ϻ�) = strpn_TY And Trim(rs!������) = strLot_TY And Trim(rs!����) = strprice_TY Then
                        DblDieQty_TY = DblDieQty_TY + Val(Trim$("" & rs!����))
                        DblAmount_TY = DblAmount_TY + Val(Trim$("" & rs!���))
     
                        wkst.Cells(lngRows, 8) = DblDieQty_TY
                        wkst.Cells(lngRows, 10) = DblAmount_TY
                        
                    Else
                        strPONUM_TY = Trim(rs!PO_NUM)
                        strMPN_DESC_TY = Trim(rs!MPN_DESC)
                        strpn_TY = Trim(rs!�Ϻ�)
                        strLot_TY = Trim(rs!������)
                        strprice_TY = Trim(rs!����)
                        
                        DblDieQty_TY = Val(Trim$("" & rs!����))
                        DblAmount_TY = Val(Trim$("" & rs!���))
                        
                        If T > 0 Then
                            wkst.Rows(lngRows & ":" & lngRows).Select
                            ExApp.Selection.Copy
                            ExApp.Selection.Insert Shift:=xlDown
                            lngRows = lngRows + 1
                        End If
                        T = T + 1
                        wkst.Cells(lngRows, 1) = Trim$("" & T)
                        wkst.Cells(lngRows, 2) = Trim$("" & rs!PO_NUM)
                        wkst.Cells(lngRows, 3) = Trim$("" & rs!MPN_DESC)
                        wkst.Cells(lngRows, 4) = Trim$("" & Mid(rs!�Ϻ�, 3, 9))
                        wkst.Cells(lngRows, 5) = Trim$("" & rs!�Ϻ�)
                        wkst.Cells(lngRows, 6) = Trim$("" & rs!������)
                        wkst.Cells(lngRows, 7) = Specification1
                        wkst.Cells(lngRows, 8) = Trim$("" & rs!����)
                        wkst.Cells(lngRows, 9) = Trim$("" & rs!����)
                        wkst.Cells(lngRows, 10) = Trim$("" & rs!���)
                        
                    End If
            

                    DblNum = DblNum + Val(Trim$("" & rs!����))
                    DblAmt = DblAmt + Val(Trim$("" & rs!���))
                    
                    DblPamt = DblPamt + Val(Trim$("" & rs!�ӹ��ѽ��))
                    DblWamt = DblWamt + Val(Trim$("" & rs!�͹��Ͻ��))
                    
           
                    rs.MoveNext
                Loop
        
                '�������
                lngRows = lngRows + 1
                wkst.Cells(lngRows, 8) = DblNum & "PCS"
        
                wkst.Cells(lngRows, 10) = DblAmt
                
                wkst.Cells(17, 7) = "Process Amount :"
                wkst.Cells(18, 7) = "Wafer Amount :"
                wkst.Cells(17, 8) = "$" & Format(DblPamt, "0.00")
                wkst.Cells(18, 8) = "$" & Format(DblWamt, "0.00")
                                
                
            
            Case "76", "US026", "SG005"

                Do While Not rs.EOF
        
                     T = T + 1
                    wkst.Cells(lngRows, 1) = Trim$("" & T)
                    wkst.Cells(lngRows, 2) = Trim$("" & rs!MPN_DESC)
                    wkst.Cells(lngRows, 3) = Trim$("" & rs!�Ϻ�)
                    If UCase(Combo1.text) = "76" Then
                        strsql_Getnewlotid = "SELECT SUBSTRING(c.Content,CHARINDEX('CUSTOMER_LOT_COMPLEX"",""',c.Content) + 23,10 )   FROM  erpdata..tblME_PrintInfo c INNER JOIN  (SELECT max(b.ID) AS id FROM  erpdata..tblErpInStockDetailInfo  a, erpdata..tblME_PrintInfo  b  WHERE   a.KEYID=b.EVENT_ID AND a.KEY_VALUE ='" & Trim$("" & rs!С���) & "' AND b.LABEL_ID = 'AAMPN4' ) t1 ON c.id=t1.id"
                         wkst.Cells(lngRows, 4) = GetSqlServerStr(strsql_Getnewlotid)
                    Else
                        wkst.Cells(lngRows, 4) = Trim$("" & rs!������)
                    End If
                    
                    
                    wkst.Cells(lngRows, 5) = Specification1
                    wkst.Cells(lngRows, 6) = Trim$("" & rs!����)
                    wkst.Cells(lngRows, 7) = Trim$("" & rs!����)
                    wkst.Cells(lngRows, 8) = Trim$("" & rs!���)
                    
                    DblNum = DblNum + Val(Trim$("" & rs!����))
                    DblAmt = DblAmt + Val(Trim$("" & rs!���))
                    
                    DblPamt = DblPamt + Val(Trim$("" & rs!�ӹ��ѽ��))
                    DblWamt = DblWamt + Val(Trim$("" & rs!�͹��Ͻ��))
                    
                    lngRows = lngRows + 1

                    rs.MoveNext
 
                Loop
        
                '�������
                wkst.Cells(lngRows, 6) = DblNum & "PCS"
        
                wkst.Cells(lngRows, 8) = "US$" & DblAmt
                
                'wkst.Cells(17, 5) = "Process Amount :US$ " & Format(DblPamt, "0.00")
                'wkst.Cells(18, 5) = "Wafer Amount :US$ " & Format(DblWamt, "0.00")
                wkst.Cells(17, 6) = "$" & Format(DblPamt, "0.00")
                wkst.Cells(18, 6) = "$" & Format(DblWamt, "0.00")
             Case "SG005_SO"

                Do While Not rs.EOF
        
                     T = T + 1
                    wkst.Cells(lngRows, 1) = Trim$("" & T)
                    strsono_SG005 = GetSqlServerStr("SELECT distinct isnull(SO_NO,'') +',' + isnull(SO_LINE,'') FROM  erpdata..tblShipOrder_Dn WHERE shiporder='" & rs!�������� & "'")
                    If strsono_SG005 <> "" Then
                        wkst.Cells(lngRows, 2) = "'" & Trim$("" & Split(strsono_SG005, ",")(0))
                        wkst.Cells(lngRows, 3) = "'" & Trim$("" & Split(strsono_SG005, ",")(1))
                    End If
                    wkst.Cells(lngRows, 4) = Trim$("" & rs!MPN_DESC)
                    wkst.Cells(lngRows, 5) = Trim$("" & rs!�Ϻ�)
                    wkst.Cells(lngRows, 6) = Trim$("" & rs!������)
                    wkst.Cells(lngRows, 7) = Specification1
                    wkst.Cells(lngRows, 8) = Trim$("" & rs!����)
                    wkst.Cells(lngRows, 9) = Trim$("" & rs!����)
                    wkst.Cells(lngRows, 10) = Trim$("" & rs!���)
                    
                    DblNum = DblNum + Val(Trim$("" & rs!����))
                    DblAmt = DblAmt + Val(Trim$("" & rs!���))
                    
                    DblPamt = DblPamt + Val(Trim$("" & rs!�ӹ��ѽ��))
                    DblWamt = DblWamt + Val(Trim$("" & rs!�͹��Ͻ��))
                    
                    lngRows = lngRows + 1

                    rs.MoveNext
 
                Loop
        
                '�������
                wkst.Cells(lngRows, 8) = DblNum & "PCS"
        
                wkst.Cells(lngRows, 10) = "US$" & DblAmt
                
                'wkst.Cells(17, 5) = "Process Amount :US$ " & Format(DblPamt, "0.00")
                'wkst.Cells(18, 5) = "Wafer Amount :US$ " & Format(DblWamt, "0.00")
                wkst.Cells(17, 8) = "$" & Format(DblPamt, "0.00")
                wkst.Cells(18, 8) = "$" & Format(DblWamt, "0.00")
            Case "BD", "EQ"
                
                Do While Not rs.EOF
        
                     T = T + 1
                    wkst.Cells(lngRows, 1) = Trim$("" & T)
                    wkst.Cells(lngRows, 2) = Trim$("" & rs!PO_NUM)
                    wkst.Cells(lngRows, 3) = Trim$("" & rs!MPN_DESC)
                    wkst.Cells(lngRows, 4) = Trim$("" & rs!�Ϻ�)
                    wkst.Cells(lngRows, 5) = Trim$("" & rs!������)
                   
                    strSql = "select DISTINCT QTECHPTNO,CUSTOMERPTNO1 from erpdata..tbltsvnpiproduct WHERE QTECHPTNO2 = '" & Trim(rs!�Ϻ�) & "' "

                    If RsNew.State = adStateOpen Then RsNew.Close

                    RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
                    
                    If RsNew.RecordCount > 0 Then
                    
                        Select Case UCase(Combo1.text)
                    
                            Case "BD"
                        
                                wkst.Cells(20, 6) = "HUATIAN NAME"
                                wkst.Cells(lngRows, 6) = Trim$("" & RsNew!qtechPTNo)
                    
                            Case "EQ"
                        
                                wkst.Cells(20, 6) = "DEVICE"
                                wkst.Cells(lngRows, 6) = Trim$("" & RsNew!CustomerPTNo1)
                    
                        End Select

                    Else

                        Select Case UCase(Combo1.text)
                    
                            Case "BD"
                        
                                wkst.Cells(20, 6) = "HUATIAN NAME"
                                wkst.Cells(lngRows, 6) = " "
                    
                            Case "EQ"
                        
                                wkst.Cells(20, 6) = "DEVICE"
                                wkst.Cells(lngRows, 6) = " "
                    
                        End Select

                    End If
                    
                    RsNew.Clone

                    Set RsNew = Nothing
                    
                    
                    DblPamt = DblPamt + Val(Trim$("" & rs!�ӹ��ѽ��))
                    DblWamt = DblWamt + Val(Trim$("" & rs!�͹��Ͻ��))
                    wkst.Cells(lngRows, 7) = Specification1
                    wkst.Cells(lngRows, 8) = Trim$("" & rs!����)
                    wkst.Cells(lngRows, 9) = Trim$("" & rs!����)
                    wkst.Cells(lngRows, 10) = Trim$("" & rs!���)

                    DblNum = DblNum + Val(Trim$("" & rs!����))
                    DblAmt = DblAmt + Val(Trim$("" & rs!���))
                    lngRows = lngRows + 1
                    rs.MoveNext
                Loop
        
                '�������
                wkst.Cells(lngRows, 8) = DblNum & "PCS"
        
                wkst.Cells(lngRows, 10) = DblAmt
                
                wkst.Cells(17, 6) = "Process Amount :"
                wkst.Cells(18, 6) = "Wafer Amount : "
                wkst.Cells(17, 7) = "$" & Format(DblPamt, "0.00")
                wkst.Cells(18, 7) = "$" & Format(DblWamt, "0.00")
            Case "GC"
                Dim s As Integer
                Dim a() As Integer
                Dim lngRows1 As Integer
                lngRows1 = 0
                j = 12
                
                m = 12
                
                DieNoFound = False
                
                AddSql2 ("delete from erptemp.dbo.ksinvoice where 1 = 1 ")
                
                ShipOrder = ""
                
                            Do While Not rs.EOF
                
                    b = Split(Trim$("" & rs!MPN_DESC), "-")

                    acpn = b(0)
                                    
                    AddSql2 (" insert into erptemp.dbo.ksinvoice values('" & Trim$(rs!PO_NUM) & "','" & Trim$(rs!Specification) & "','" & Trim$(acpn) & "','" & Trim$(rs!�Ϻ�) & "','" & Trim$(rs!����) & "','" & Trim$(rs!����) & "','" & Trim$(rs!���) & "','" & Trim$(rs!�ӹ��ѽ��) & "','" & Trim$(rs!�͹��Ͻ��) & "','','0','0','')")
                    ShipOrderFlag = True
                    For S_I = 0 To UBound(Split(ShipOrder, ","))
                        If Trim$(rs!��������) = Split(ShipOrder, ",")(S_I) Then
                            ShipOrderFlag = False '�˷��������Ѳ�ѯ��NG DIE���������ٲ�
                        End If
                    Next

                    
                    If ShipOrderFlag = True Then     '�ж�ShipOrderFlag ��ȥ�ظ�ShipOrder
                            ShipOrder = ShipOrder & Trim$(rs!��������) & "','"
                    End If
                    rs.MoveNext
                Loop
                 ShipOrder = Mid(ShipOrder, 1, Len(ShipOrder) - 3)
                
'                        strSql = " select  ISNULL(ISNULL(T.E, n.NDPW), 0) as ���� " & _
'                        " FROM ( SELECT 'HTKS' AS sub_name, d.SHIP_SITE,a.ship_order, " & _
'                        " RTRIM(d.FAB_CONV_ID) as FAB_CONV_ID,  a.cust_device, a.gcversion,d.PO_NUM, a.create_date, rtrim(a.lot_id) as lot_id, SUBSTRING(REPLACE(b.���̿����, '+', ''), LEN(a.lot_id) + 1, 2) as waferid,  " & _
'                        " c.FAILBINCOUNT + c.PASSBINCOUNT AS gross_die, CASE WHEN n.KEY_NAME in ( 'BAD1_DIE','BAD2_DIE') THEN 'E'  ELSE 'A' END Grade, CONVERT(INT,n.KEY_VALUE ) AS qty,  c.PRODUCTID, rtrim(ay.���) as ���,  " & _
'                        " b.�󹤵�,  a.qbox, b.���̿����, SUBSTRING(ee.SFC_ID, 12, 8) AS SFC " & _
'                        " FROM erptemp .. tblshipreport_new a  " & _
'                        " INNER JOIN erpdata .. tblStockNumTree ax  ON ax.��� = a.qbox  " & _
'                        " INNER JOIN erpdata .. tblStockNumTree ay ON ay.��� = ax.�ϼ����  " & _
'                        " INNER JOIN erpdata .. tblStocksqfhsub b ON b.���ݱ�� = a.ship_order  AND b.��� = a.qbox   AND b.������ = a.lot_id " & _
'                        " INNER JOIN ERPBASE .. tblmappingData c  ON c.SUBSTRATEID = b.���̿���� AND c.LOTID = b.������ " & _
'                        " INNER JOIN erpbase .. tblCustomerOI d  ON CONVERT(VARCHAR(20), CONVERT(int,d.ID)) = c.FILENAME  AND d.SOURCE_BATCH_ID = c.LOTID  " & _
'                        " LEFT JOIN  erpdata..tblErpInStockDetailInfo m ON m.KEY_VALUE = b.��� " & _
'                        " LEFT JOIN  erpdata..tblErpInStockDetailInfo n  ON n.BOX_ID = m.BOX_ID  and n.KEY_NAME in ( 'BAD1_DIE','BAD2_DIE','GOOD_DIE') and n.KEY_TYPE = 'WAFER' AND   CHARINDEX(c.SUBSTRATEID , n.KEYID ) <> 0 " & _
'                        " inner JOIN erpdata .. tblErpInStockRelation ee ON    ee.BOX_ID = n.BOX_ID AND  ee.WAFER_ID = n.KEYID  WHERE a.ship_order = '" & Trim$(rs!��������) & "' )  AS p  PIVOT(sum(qty) FOR Grade IN(A,BIN1, E)) AS T " & _
'                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV k  ON k.QBOXNUMBER = t.qbox AND k.WAFERSCRIBENUMBER = t.���̿���� AND k.CONTAINERNAME LIKE '%-A' " & _
'                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV L ON L.QBOXNUMBER = t.qbox AND L.WAFERSCRIBENUMBER = t.���̿���� AND L.CONTAINERNAME LIKE '%-A-01' " & _
'                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV m  ON m.QBOXNUMBER = t.qbox  AND m.WAFERSCRIBENUMBER = t.���̿���� AND m.CONTAINERNAME LIKE '%-A-02' " & _
'                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV n  ON n.QBOXNUMBER = t.qbox AND n.WAFERSCRIBENUMBER = t.���̿���� AND n.CONTAINERNAME LIKE '%-E' "
'
'
                        strSql = "select  SUM(ISNULL(ISNULL(T.E, n.NDPW), 0)) as ���� ,T.cust_device" & _
                        " From " & _
                        "( SELECT 'HTKS' AS sub_name, d.SHIP_SITE,a.ship_order," & _
                        " RTRIM(d.FAB_CONV_ID) as FAB_CONV_ID,  a.cust_device, a.gcversion,d.PO_NUM, a.create_date, rtrim(a.lot_id) as lot_id, SUBSTRING(REPLACE(b.���̿����, '+', ''), LEN(a.lot_id) + 1, 2) as waferid," & _
                        " c.FAILBINCOUNT + c.PASSBINCOUNT AS gross_die, CASE WHEN n.KEY_NAME in ( 'BAD1_DIE','BAD2_DIE') THEN 'E'  ELSE 'A' END Grade, CONVERT(INT,n.KEY_VALUE ) AS qty,  c.PRODUCTID, rtrim(ay.���) as ���," & _
                        " b.�󹤵�,  a.qbox, b.���̿����, SUBSTRING(ee.SFC_ID, 12, 8) AS SFC" & _
                        " FROM erptemp .. tblshipreport_new a" & _
                        " INNER JOIN erpdata .. tblStockNumTree ax  ON ax.��� = a.qbox" & _
                        " INNER JOIN erpdata .. tblStockNumTree ay ON ay.��� = ax.�ϼ����" & _
                        " INNER JOIN erpdata .. tblStocksqfhsub b ON b.���ݱ�� = a.ship_order  AND b.��� = a.qbox   AND b.������ = a.lot_id" & _
                        " INNER JOIN ERPBASE .. tblmappingData c  ON c.SUBSTRATEID = b.���̿���� AND c.LOTID = b.������" & _
                        " INNER JOIN erpbase .. tblCustomerOI d  ON CONVERT(VARCHAR(20), CONVERT(int,d.ID)) = c.FILENAME  AND d.SOURCE_BATCH_ID = c.LOTID" & _
                        " LEFT JOIN  erpdata..tblErpInStockDetailInfo m ON m.KEY_VALUE = b.���" & _
                        " LEFT JOIN  erpdata..tblErpInStockDetailInfo n  ON n.BOX_ID = m.BOX_ID  and n.KEY_NAME in ( 'BAD1_DIE','BAD2_DIE','GOOD_DIE') and n.KEY_TYPE = 'WAFER' AND   CHARINDEX(c.SUBSTRATEID , n.KEYID ) <> 0" & _
                        " inner JOIN erpdata .. tblErpInStockRelation ee ON    ee.BOX_ID = n.BOX_ID AND  ee.WAFER_ID = n.KEYID  WHERE a.ship_order  in('" & ShipOrder & "')) AS p  PIVOT(sum(qty) FOR Grade IN(A,BIN1, E)) AS T" & _
                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV k  ON k.QBOXNUMBER = t.qbox AND k.WAFERSCRIBENUMBER = t.���̿���� AND k.CONTAINERNAME LIKE '%-A'" & _
                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV L ON L.QBOXNUMBER = t.qbox AND L.WAFERSCRIBENUMBER = t.���̿���� AND L.CONTAINERNAME LIKE '%-A-01'" & _
                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV m  ON m.QBOXNUMBER = t.qbox  AND m.WAFERSCRIBENUMBER = t.���̿���� AND m.CONTAINERNAME LIKE '%-A-02'" & _
                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV n  ON n.QBOXNUMBER = t.qbox AND n.WAFERSCRIBENUMBER = t.���̿���� AND n.CONTAINERNAME LIKE '%-E'" & _
                        " GROUP BY T.cust_device"
                                 
                                                
                        If RsNew.State = adStateOpen Then RsNew.Close

                        RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
                        
                        If RsNew.RecordCount > 0 Then
                            RsNew.MoveFirst
                            ReDim a(RsNew.RecordCount)
                            For N = 1 To RsNew.RecordCount
                                DblNum2 = DblNum2 + Val(Trim$("" & RsNew!����))
                                a(lngRows1) = Val(Trim$("" & RsNew!����))
                                lngRows1 = lngRows1 + 1
                                RsNew.MoveNext
                            Next
                        End If
                        
                        RsNew.Clone
                    
                        Set RsNew = Nothing
                    
           
                strSql = "select ROW_NUMBER() OVER(order by acpn) item, PO_NUM=stuff((SELECT DISTINCT '/' + PO_NUM FROM erptemp.dbo.ksinvoice WHERE acpn=a.acpn AND �Ϻ�=a.�Ϻ� AND Specification=a.Specification " & _
                             "for xml path('')),1, 1, ''),acpn,Specification,�Ϻ�,sum(����) as ����,sum(���) as ���,sum(�ӹ��ѽ��) as �ӹ��ѽ��,sum(�͹��Ͻ��) as �͹��Ͻ�� " & _
                             "FROM  erptemp.dbo.ksinvoice a group by acpn,Specification,�Ϻ�"

                wkst.Cells(lngRows - 1, 10) = "NG:DIE"
                If RsNew.State = adStateOpen Then RsNew.Close

                RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
                
                IntInertRow = RsNew.RecordCount

                For i = 1 To IntInertRow - 1
                    wkst.Rows(lngRows & ":" & lngRows).Select
                    ExApp.Selection.Copy
                    ExApp.Selection.Insert Shift:=xlDown
                Next i
                RsNew.MoveFirst
                
                For N = 1 To RsNew.RecordCount
                    s = N - 1
                    wkst.Cells(lngRows, 10) = a(s)
                    T = T + 1
                    wkst.Cells(lngRows, 1) = Trim$("" & T)
                   
                        wkst.Cells(lngRows, 2) = Replace(Trim$("" & RsNew!PO_NUM), "/", "/" & Chr(10)) 'һ��PO_NUMһ��
                        wkst.Cells(lngRows, 3) = Specification1  'оƬ

                    wkst.Cells(lngRows, 4) = Trim$("" & RsNew!acpn)
                    wkst.Cells(lngRows, 5) = Trim$("" & RsNew!�Ϻ�)
                    wkst.Cells(lngRows, 6) = Trim$("" & RsNew!����)
                    wkst.Cells(lngRows, 9) = Trim$("" & RsNew!���)
                    
                    DblPamt = Format(DblPamt, "0.00") + Val(Trim$("" & RsNew!�ӹ��ѽ��))
                    DblWamt = Format(DblWamt, "0.00") + Val(Trim$("" & RsNew!�͹��Ͻ��))
                
                    strSql = "select DISTINCT DIE from erptemp.dbo.customerkspn WHERE CUSTOMERPN = '" & RsNew!acpn & "' and �ߴ� = '" & Left(Trim$("" & RsNew!�Ϻ�), 2) & "' "
                
                    If RsNew1.State = adStateOpen Then RsNew1.Close
                
                    RsNew1.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
                    
                    asum = 0
                    TOTALWAFER = 0
                
                    If RsNew1.RecordCount > 0 Then
                
                        asum = Trim$("" & RsNew1!die)
                        
                        TOTALWAFER = Val(RsNew!����) / asum
                        
                        DblPnum = DblPnum + Val(Trim$("" & TOTALWAFER))
                    
                    Else
                
                        MsgBox RsNew!acpn & "�ͻ������޶�Ӧ��DIE��,��ά����", vbInformation, "��ʾ��"
                        
                        TOTALWAFER = 0
                        
                        DblPnum = 0
                        
                        DieNoFound = True
                
                    End If
                
                    RsNew1.Clone
                
                    Set RsNew1 = Nothing
                
                    wkst.Cells(lngRows, 7) = Trim$("" & TOTALWAFER)
                                                 
                    If j = 12 Then
                
                        wkst.Cells(j, 5) = "NOTE:1PC " & RsNew!acpn & " = " & asum & " EA;Total: " & TOTALWAFER & " PCS " & RsNew!acpn & " = " & Val(RsNew!����) & " EA"
                
                    Else

                        If j <= 19 And j > 12 Then
                
                            wkst.Cells(j, 5) = "         1PC " & RsNew!acpn & " = " & asum & " EA;Total: " & TOTALWAFER & " PCS " & RsNew!acpn & " = " & Val(RsNew!����) & " EA"
                
                        End If
                
                    End If
                
                    DblNum = DblNum + Val(Trim$("" & RsNew!����))
                    DblAmt = DblAmt + Val(Trim$("" & RsNew!���))
                
                    
                
                    If j > 19 Then
                
                        wkst.Cells(m, 7) = "1PC " & RsNew!acpn & " = " & asum & " EA;Total: " & TOTALWAFER & " PCS " & RsNew!acpn & " = " & Val(RsNew!����) & " EA"
                
                        m = m + 1
                
                        If m > 19 Then
                
                            MsgBox "��ʽ�Ѿ�������Χ��", vbInformation, "��ʾ��"
                
                            Exit Sub
                
                        End If
                
                    End If
                    j = j + 1
               
                    lngRows = lngRows + 1

                    RsNew.MoveNext
                    
                Next
        
                '�������
              '�������
                If DblNum < 1000 Then
                    wkst.Cells(lngRows, 6) = DblNum & "EA"
                Else
                    wkst.Cells(lngRows, 6) = Format(DblNum, "0,###") & "EA"
                End If
                

                If DieNoFound = True Then
                    wkst.Cells(lngRows, 7) = "Ƭ"
                Else
                    If DblPnum < 1000 Then
                         wkst.Cells(lngRows, 7) = DblPnum & "Ƭ"
                    Else
                         wkst.Cells(lngRows, 7) = Format(DblPnum, "0,###") & "Ƭ"
                    End If
                   
                End If
               ' wkst.Cells(lngRows, 7) = DblPnum & "Ƭ"
                'wkst.Cells(lngRows, 9) = "$" & Format(DblAmt, "0.00")
                DblAmt = Format(DblAmt, "0.00")
                If DblAmt < 1000 Then
                         wkst.Cells(lngRows, 9) = "$" & DblAmt
                Else
                        wkst.Cells(lngRows, 9) = "$" & Format(DblAmt, "0,###")
                End If
                
                
                'wkst.Cells(10, 5) = "Process Amount :US$ " & Format(DblPamt, "0.00")
                'wkst.Cells(11, 5) = "Wafer Amount :US$ " & Format(DblWamt, "0.00")

                    wkst.Cells(20, 6) = "$" & Format(DblPamt, "0.00")
                    wkst.Cells(21, 6) = "$" & Format(DblWamt, "0.00")
                If DblNum2 < 1000 Then
                          wkst.Cells(lngRows + 2, 2) = DblNum2 & "EA"
                Else
                         wkst.Cells(lngRows + 2, 2) = Format(DblNum2, "0,###") & "EA"
                End If

                RsNew.Clone

                Set RsNew = Nothing
            
        End Select
        
    Else
        '        ClsP.UnLoad_Form
        MsgBox "���赼�����ݣ�", vbInformation, "��ʾ��"
        Exit Sub

    End If
  
    '    ClsP.ShowProgress 100, "�����ɹ���"
    '    If Not ClsP Is Nothing Then
    '        Set ClsP = Nothing
    '    End If
    ExApp.Visible = True
    
    '    If intFlag = 1 Then
    '        wkst.PrintPreview
    '        wkbk.Close (False)
    '        ExApp.Quit
    '    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing

    End If

    Exit Sub
ErrHandle:

    On Error Resume Next

    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing

    End If

    '    If Not ClsP Is Nothing Then
    '        Set ClsP = Nothing
    '    End If
    MsgBox Err.DESCRIPTION, vbInformation, "��ʾ��"
    Exit Sub

End Sub

'Packing list
Public Sub PackinglistExportPrintExcel(ByVal rs As ADODB.Recordset, ByVal strdj As String)
    Dim strSql         As String
    
    Dim sstrSql        As String

    Dim lngRows        As Long

    Dim rsQuery        As Excel.QueryTable

    'Dim ClsP                As New ClsProgress
    Dim ExApp          As Excel.Application

    Dim wkbk           As New Workbook

    Dim wkst           As New Worksheet

    Dim i              As Long

    Dim j              As Long
    
    Dim m              As Long
    
    Dim N              As Long
    
    Dim p              As Long

    Dim IntCols        As Integer

    Dim strCols        As String

    Dim strFileName    As String
    
    Dim strmeas        As String

    Dim IntInertRow    As Integer, IntMaxDetailRow As Integer
    
    Dim DblNum         As Double
    
    Dim DblNum1        As Double
    
    Dim DblNum2        As Double
    
    Dim DblPnum        As Double

    Dim DblAmt         As Double '�ܽ��
    
    Dim DblPamt        As Double

    Dim DblWamt        As Double

    Dim intBoxNum      As Integer '����

    Dim strPBigBox     As String  'ǰ���

    Dim strNBigBox     As String  '�����
    
    Dim gdh As String
                    
    Dim ngdh As String

    Dim IntBMegerRow   As Integer

    Dim IntEMegerRow   As Integer

    Dim DblJZ          As Double '����

    Dim DblMZ          As Double 'ë��

    Dim intBegin       As Integer

    Dim RsNew          As New ADODB.Recordset
    
    Dim RsNew1         As New ADODB.Recordset
    
    Dim RsNew2         As New ADODB.Recordset

    Dim strShipTo()    As String

    Dim strSoldBy()    As String
    
    Dim Specification1 As String

    Dim waferid1       As String
    
    Dim b()            As String
    
    Dim acpn           As String
    
    Dim asum           As Integer
    
    Dim ShipOrder      As String
    
    Dim S_I            As Integer
    
    Dim ShipOrderFlag  As Boolean
    
    Dim DieNoFound     As Boolean
    
    Dim TOTALWAFER As Integer
    
    Dim CARTONCNT As Integer
    
    Dim ROW_S As Integer
    
    Dim strsono_SG005      As String
    
    Dim shu As String
    
     Dim strsql_Getnewlotid As String
    
    
    ShipOrder = ""
    strPBigBox = ""
    strNBigBox = ""
    intBoxNum = 1
    
    

    If rs.RecordCount <= 0 Then
        MsgBox "û��Ҫ���������ϣ�", vbInformation, "��ʾ��"
        Exit Sub

    End If

    '    ClsP.Init 100, True
    '    ClsP.ShowProgress 10, "��ʼ������..."
    
    strSysPath = App.Path
    
    Select Case UCase(Combo1.text)
        
        Case "68"
        
            strFileName = strSysPath & "\TEMPLET\68_Packing_list.xls" 'Ҫ�򿪵��ļ�
            
        Case "76", "US026"
        
            strFileName = strSysPath & "\TEMPLET\76_Packing_list.xls" 'Ҫ�򿪵��ļ�
        Case "SG005"
        
            strFileName = strSysPath & "\TEMPLET\SG005_Packing_list.xls" 'Ҫ�򿪵��ļ�
            
         Case "SG005_SO"
        
            strFileName = strSysPath & "\TEMPLET\SG005_SO_Packing_list.xls" 'Ҫ�򿪵��ļ�
        Case "TW067", "ͨ��ģ��"
        
            strFileName = strSysPath & "\TEMPLET\TW067_Packing_list.xls" 'Ҫ�򿪵��ļ�
        
        Case "GC"
        
            strFileName = strSysPath & "\TEMPLET\GC_Packing_list.xls" 'Ҫ�򿪵��ļ�
            
        Case "HK005"
            
            strFileName = strSysPath & "\TEMPLET\HK005_Packing_list.xls" 'Ҫ�򿪵��ļ�
            
        Case "HK080"
            
            strFileName = strSysPath & "\TEMPLET\HK080_Packing_list.xls" 'Ҫ�򿪵��ļ�
            
        Case "HK075"
            
            strFileName = strSysPath & "\TEMPLET\HK075_Packing_list.xls" 'Ҫ�򿪵��ļ�

    End Select
    
    If rs.RecordCount > 0 Then
        '        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblNum1 = 0
        DblNum2 = 0
        DblAmt = 0
        DblJZ = 0
        DblMZ = 0
        
   
        If UCase(Combo1.text) = "GC" Then
            Specification1 = "оƬ"
        Else
            Specification1 = "Integrated Circuit chip"
        End If
        waferid1 = ""
        strmeas = ""

        '��ֵ��Excel�У���ͷ
        Select Case UCase(Combo1.text)
        
            Case "68"
            
                wkst.Cells(3, 9) = strdj   'Trim$("" & Rs!���۵����)
               ' wkst.Cells(3, 9) = DATE
                wkst.Cells(3, 11) = Trim$("" & rs!��������)
                
            Case "TW067", "ͨ��ģ��"
            
                wkst.Cells(3, 7) = strdj   'Trim$("" & Rs!���۵����)
                wkst.Cells(3, 9) = DATE
                
            Case "76", "US026", "SG005"
                
                wkst.Cells(3, 5) = strdj   'Trim$("" & Rs!���۵����)
                wkst.Cells(3, 7) = DATE
                
            Case "SG005_SO"
                
                wkst.Cells(3, 7) = strdj   'Trim$("" & Rs!���۵����)
                wkst.Cells(3, 9) = DATE

            Case "GC"
                wkst.Cells(3, 6) = strdj   'Trim$("" & Rs!���۵����)
                wkst.Cells(3, 9) = DATE
            
            Case "HK005"
            
                wkst.Cells(3, 8) = strdj   'Trim$("" & Rs!���۵����)
               ' wkst.Cells(3, 12) = DATE
                wkst.Cells(3, 11) = Trim$("" & rs!��������)
             Case "HK075"
            
                wkst.Cells(3, 8) = strdj   'Trim$("" & Rs!���۵����)
               ' wkst.Cells(3, 12) = DATE
                wkst.Cells(3, 11) = Trim$("" & rs!��������)
        End Select
        wkst.Cells(7, 1) = "CONTACT:" & Trim(strUserName)
        wkst.Cells(10, 10) = Trim$("" & rs!�˵���)
        
        '��ѯ���ų�����ַ------------------------------------------
        strSql = "SELECT DISTINCT SHIP_TO_AD,SOLD_BY,SHIP_TO FROM erpdata..Vw_CustomerShipTo WHERE ���۵���� IN('" & Replace(strdj, ",", "','") & "')"

        If RsNew.State = adStateOpen Then RsNew.Close
        RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

        If RsNew.RecordCount > 0 Then

            'ShipTo
            If InStr(Trim$("" & RsNew!SHIP_TO_AD), "@") > 0 Then
                strShipTo = Split(Trim$("" & RsNew!SHIP_TO_AD), "@")

                For i = 0 To UBound(strShipTo)

                    If i + 9 > 14 Then Exit For
                    wkst.Cells(i + 9, 1) = strShipTo(i)
                    If UCase(Combo1.text) = "SG005" Or UCase(Combo1.text) = "SG005_SO" Then
                       If i + 9 = 11 Then
                             wkst.Cells(11, 1) = Mid(strShipTo(1), 54, 24)
                       End If
                       If i + 9 = 10 Then
                             wkst.Cells(10, 1) = Mid(strShipTo(i), 1, 53)
                       End If
                    End If
                Next
            Else
                wkst.Cells(9, 1) = Trim$("" & RsNew!SHIP_TO_AD)

            End If

            'SoldBy
            If InStr(Trim$("" & RsNew!SOLD_BY), "@") > 0 Then
                strSoldBy = Split(Trim$("" & RsNew!SOLD_BY), "@")

                For i = 0 To UBound(strSoldBy)

                    If i + 16 >= 19 Then
                        wkst.Cells(18, 1) = wkst.Cells(18, 1) & " " & strSoldBy(i)
                        'Exit For
                    Else
                        wkst.Cells(i + 16, 1) = strSoldBy(i)
                    End If
                Next
            Else
                wkst.Cells(16, 1) = Trim$("" & RsNew!SOLD_BY)

            End If
            
            Select Case UCase(Combo1.text)
        
                Case "68", "US026"
            
                    wkst.Cells(8, 9) = "    TO: " & Trim$("" & RsNew!SHIP_TO)
                Case "HK005"
            
                    wkst.Cells(8, 8) = "    TO: " & Trim$("" & RsNew!SHIP_TO)
                Case "76"
                    wkst.Cells(8, 6) = Trim$("" & RsNew!SHIP_TO)
                    wkst.Cells(8, 8) = " "
                'Case "GC"
                '    wkst.Cells(8, 6) = Trim$("" & RsNew!SHIP_TO)
'               Case "SG005"

'                    wkst.Cells(8, 5) = "    TO: " & Trim$("" & RsNew!ship_to)
                Case "BD", "EQ"
                
                    wkst.Cells(8, 7) = Trim$("" & RsNew!SHIP_TO)
                Case "GC"
                   ' wkst.Cells(6, 7) = Trim$("" & RsNew!SHIP_TO)
                Case "HK075"
                    wkst.Cells(8, 9) = Trim$("" & RsNew!SHIP_TO)
            End Select

        Else
            wkst.Cells(9, 1) = ""
            wkst.Cells(10, 1) = ""
            wkst.Cells(11, 1) = ""
            wkst.Cells(12, 1) = ""
            wkst.Cells(13, 1) = ""
            wkst.Cells(14, 1) = ""
            wkst.Cells(16, 1) = ""
            wkst.Cells(17, 1) = ""
            wkst.Cells(18, 1) = ""
            wkst.Cells(19, 1) = ""

        End If

        RsNew.Close
        '----------------------------------------------------------
        lngRows = 21
        
        If UCase(Combo1.text) <> "GC" Then
            IntInertRow = rs.RecordCount
        
            For i = 1 To IntInertRow - 1
                wkst.Rows(lngRows & ":" & lngRows).Select
                ExApp.Selection.Copy
                ExApp.Selection.Insert Shift:=xlDown
            Next i
        
        End If

        IntMaxDetailRow = rs.RecordCount
        
        '        ClsP.ShowProgress 50, "���ڵ���..."
        
        IntBMegerRow = 20
        IntEMegerRow = 22
        intBegin = 1
        

        If UCase(Combo1.text) = "GC" Then
            AddSql2 ("delete from erptemp.dbo.ksinvoice where 1 = 1 ")
            
        End If
        Dim T As Integer
        T = 0
        CARTONCNT = 0
        For i = 0 To rs.RecordCount - 1
        
            Select Case UCase(Combo1.text)
        
                Case "68"
                    T = T + 1
                    wkst.Cells(lngRows, 1) = Trim$("" & T)
                    wkst.Cells(lngRows, 2) = Trim$("" & rs!PO_NUM)
                    wkst.Cells(lngRows, 3) = Trim$("" & rs!MPN_DESC)
                    If (InStr(rs!�Ϻ�, "B") < 8) Then
                        wkst.Cells(lngRows, 4) = Trim$("" & Mid(rs!�Ϻ�, 3, 7))
                    Else
                    wkst.Cells(lngRows, 4) = Trim$("" & Mid(rs!�Ϻ�, 3, InStr(rs!�Ϻ�, "B") - 2))
                    End If
                    wkst.Cells(lngRows, 5) = Trim$("" & rs!�Ϻ�)
                    wkst.Cells(lngRows, 6) = Trim$("" & rs!������)
                    
                    If Not rs.EOF Then
            
                        strSql = "select RIGHT(Replace(rtrim(���̿����),'+',''),2) as waferid FROM erpdata..tblstockmovesub x left join erptemp..mps_mark ad on ad.wafer_id = x.���̿���� where  x.�Ϻ� = '" & Trim(rs!�Ϻ�) & "' and  ad.REMARK2 = '" & rs!������ & "' and x.���ݱ�� = '" & Trim(rs!��������) & "' order by RIGHT(Replace(rtrim(x.���̿����),'+',''),2) "

                        If RsNew.State = adStateOpen Then RsNew.Close
                        RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
                    
                        If RsNew.RecordCount = 0 Then
                        strSql = "select RIGHT(Replace(rtrim(���̿����),'+',''),2) as waferid FROM erpdata..tblstockmovesub x  where  �Ϻ� = '" & Trim(rs!�Ϻ�) & "' and  ������ = '" & rs!������ & "'and ���ݱ�� = '" & Trim(rs!��������) & "' order by RIGHT(Replace(rtrim(���̿����),'+',''),2) "

                        If RsNew.State = adStateOpen Then RsNew.Close
                        RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
                        End If

                        RsNew.MoveFirst

                        For j = 1 To RsNew.RecordCount
                
                            If j = 1 Then
                
                                waferid1 = "#" & RsNew("waferid")
                
                            Else
                    
                                waferid1 = waferid1 & " " & RsNew("waferid")
                    
                            End If

                            RsNew.MoveNext
            
                        Next

                        If RsNew.RecordCount = 25 Then

                            waferid1 = "#" & "01-25"

                        End If
                        gdh = Trim$("" & rs!������)
                        If gdh <> ngdh Then
                            ngdh = Trim$("" & rs!������)
                        Else
                            wkst.Range(Chr(7 + 64) & IntBMegerRow & ":" & Chr(7 + 64) & IntEMegerRow).Merge
                        End If
                        RsNew.Clone
                        
                        Set RsNew = Nothing

                        wkst.Cells(lngRows, 7) = waferid1
                        
                        wkst.Cells(lngRows, 8) = Trim$("" & rs!DC)
                        wkst.Cells(lngRows, 9) = Specification1
                        wkst.Cells(lngRows, 10) = Trim$("" & rs!����)
                        strPBigBox = Trim$("" & rs!�����)
            
                        If strPBigBox <> strNBigBox Then
            
                            strNBigBox = Trim$("" & rs!�����)
                  
                            'ë��
                            wkst.Cells(lngRows, 13) = Val(Trim$("" & rs!����))
                            DblMZ = DblMZ + Val(Trim$("" & rs!����))
                
                            '����
                            wkst.Cells(lngRows, 12) = Val(Trim$("" & FormatNumber(rs!���� * 0.25, 2, vbTrue)))
                            DblJZ = DblJZ + Val(Trim$("" & FormatNumber(rs!���� * 0.25, 2, vbTrue)))
               
                            '���
                            wkst.Cells(lngRows, 14) = Trim$("" & rs!MEAS)
                            '
                            wkst.Cells(lngRows, 11) = "1"
                            intBoxNum = intBoxNum + 1
            
                            '�趨ˮƽ����ֱ����
                           
                            wkst.Range(Chr(11 + 64) & lngRows & ":" & Chr(14 + 64) & lngRows).horizontalAlignment = xlCenter
                            wkst.Range(Chr(11 + 64) & lngRows & ":" & Chr(14 + 64) & lngRows).verticalAlignment = xlCenter
                            '--------------------------
                
                            IntBMegerRow = IntBMegerRow + intBegin
                    
                            intBegin = 1
                        Else
                            '�ϲ�
'                            wkst.Range(Chr(7 + 64) & IntBMegerRow & ":" & Chr(7 + 64) & IntEMegerRow).Merge
                            wkst.Range(Chr(11 + 64) & IntBMegerRow & ":" & Chr(11 + 64) & IntEMegerRow).Merge
                            wkst.Range(Chr(12 + 64) & IntBMegerRow & ":" & Chr(12 + 64) & IntEMegerRow).Merge
                            wkst.Range(Chr(13 + 64) & IntBMegerRow & ":" & Chr(13 + 64) & IntEMegerRow).Merge
                            wkst.Range(Chr(14 + 64) & IntBMegerRow & ":" & Chr(14 + 64) & IntEMegerRow).Merge
                            '�趨ˮƽ����ֱ����
                            wkst.Range(Chr(11 + 64) & IntBMegerRow & ":" & Chr(14 + 64) & IntEMegerRow).horizontalAlignment = xlCenter
                            wkst.Range(Chr(11 + 64) & IntBMegerRow & ":" & Chr(14 + 64) & IntEMegerRow).verticalAlignment = xlCenter
                            '--------------------------
                            intBegin = intBegin + 1

                        End If
                        
                        

                        DblNum = DblNum + Val(Trim$("" & rs!����))
            
                        lngRows = lngRows + 1
                        IntEMegerRow = lngRows
                        rs.MoveNext
            
                    End If

                Case "76", "US026", "SG005"

                    T = T + 1
                    wkst.Cells(lngRows, 1) = Trim$("" & T)
                    wkst.Cells(lngRows, 2) = Trim$("" & rs!MPN_DESC)
                    wkst.Cells(lngRows, 3) = Trim$("" & rs!�Ϻ�)
                    If UCase(Combo1.text) = "76" Then
                        strsql_Getnewlotid = "SELECT SUBSTRING(c.Content,CHARINDEX('CUSTOMER_LOT_COMPLEX"",""',c.Content) + 23,10 )   FROM  erpdata..tblME_PrintInfo c INNER JOIN  (SELECT max(b.ID) AS id FROM  erpdata..tblErpInStockDetailInfo  a, erpdata..tblME_PrintInfo  b  WHERE   a.KEYID=b.EVENT_ID AND a.KEY_VALUE ='" & Trim$("" & rs!С���) & "' AND b.LABEL_ID = 'AAMPN4' ) t1 ON c.id=t1.id"
                        wkst.Cells(lngRows, 4) = GetSqlServerStr(strsql_Getnewlotid)
                    Else
                        wkst.Cells(lngRows, 4) = Trim$("" & rs!������)
                    End If
        
                  '  wkst.Cells(lngRows, 4) = Trim$("" & rs!������)
                    wkst.Cells(lngRows, 5) = Specification1
                    wkst.Cells(lngRows, 6) = Trim$("" & rs!����)
                    strPBigBox = Trim$("" & rs!�����)

                    If strPBigBox <> strNBigBox Then
                        CARTONCNT = CARTONCNT + 1
                        shu = Trim$("" & rs!ITEM)
            
                        strNBigBox = Trim$("" & rs!�����)
                  
                        'ë��
                        wkst.Cells(lngRows, 9) = Val(Trim$("" & rs!����))
                        DblMZ = DblMZ + Val(Trim$("" & rs!����))
                
                        '����
                        wkst.Cells(lngRows, 8) = Val(Trim$("" & FormatNumber(rs!���� * 0.25, 2, vbTrue)))
                        DblJZ = DblJZ + Val(Trim$("" & FormatNumber(rs!���� * 0.25, 2, vbTrue)))
               
                        '���
                        wkst.Cells(lngRows, 10) = Trim$("" & rs!MEAS)
                        '
                        wkst.Cells(lngRows, 7) = "1"

                        intBoxNum = intBoxNum + 1
            
                        '�趨ˮƽ����ֱ����
                        wkst.Range(Chr(10 + 64) & lngRows & ":" & Chr(13 + 64) & lngRows).horizontalAlignment = xlCenter
                        wkst.Range(Chr(10 + 64) & lngRows & ":" & Chr(13 + 64) & lngRows).verticalAlignment = xlCenter
                        '--------------------------
                
                        IntBMegerRow = IntBMegerRow + intBegin
                    
                        intBegin = 1
                    Else
                        '�ϲ�
'                        Application.DisplayAlerts = 0
'                        Selection.Merge
'                        Application.DisplayAlerts = 1
                        
                        
                        wkst.Range(Chr(10 + 64) & IntBMegerRow & ":" & Chr(10 + 64) & IntEMegerRow).Merge
                        wkst.Range(Chr(7 + 64) & IntBMegerRow & ":" & Chr(7 + 64) & IntEMegerRow).Merge
                        wkst.Range(Chr(8 + 64) & IntBMegerRow & ":" & Chr(8 + 64) & IntEMegerRow).Merge
                        wkst.Range(Chr(9 + 64) & IntBMegerRow & ":" & Chr(9 + 64) & IntEMegerRow).Merge
                        '�趨ˮƽ����ֱ����
                        wkst.Range(Chr(10 + 64) & IntBMegerRow & ":" & Chr(13 + 64) & IntEMegerRow).horizontalAlignment = xlCenter
                        wkst.Range(Chr(10 + 64) & IntBMegerRow & ":" & Chr(13 + 64) & IntEMegerRow).verticalAlignment = xlCenter
                        '--------------------------
                        intBegin = intBegin + 1
                        

                    End If

                    DblNum = DblNum + Val(Trim$("" & rs!����))
            
                    lngRows = lngRows + 1
                    IntEMegerRow = lngRows
                    rs.MoveNext
                 Case "SG005_SO"
          
                    T = T + 1
                    wkst.Cells(lngRows, 1) = Trim$("" & T)
                    strsono_SG005 = GetSqlServerStr("SELECT distinct isnull(SO_NO,'') +',' + isnull(SO_LINE,'') FROM  erpdata..tblShipOrder_Dn WHERE shiporder='" & rs!�������� & "'")
                    If strsono_SG005 <> "" Then
                        wkst.Cells(lngRows, 2) = "'" & Trim$("" & Split(strsono_SG005, ",")(0))
                        wkst.Cells(lngRows, 3) = "'" & Trim$("" & Split(strsono_SG005, ",")(1))
                    End If
                    wkst.Cells(lngRows, 4) = Trim$("" & rs!MPN_DESC)
                    wkst.Cells(lngRows, 5) = Trim$("" & rs!�Ϻ�)
                    wkst.Cells(lngRows, 6) = Trim$("" & rs!������)
                    wkst.Cells(lngRows, 7) = Specification1
                    wkst.Cells(lngRows, 8) = Trim$("" & rs!����)
                    strPBigBox = Trim$("" & rs!�����)

                    If strPBigBox <> strNBigBox Then
                        CARTONCNT = CARTONCNT + 1
                        shu = Trim$("" & rs!ITEM)
            
                        strNBigBox = Trim$("" & rs!�����)
                  
                        'ë��
                        wkst.Cells(lngRows, 11) = Val(Trim$("" & rs!����))
                        DblMZ = DblMZ + Val(Trim$("" & rs!����))
                
                        '����
                        wkst.Cells(lngRows, 10) = Val(Trim$("" & FormatNumber(rs!���� * 0.25, 2, vbTrue)))
                        DblJZ = DblJZ + Val(Trim$("" & FormatNumber(rs!���� * 0.25, 2, vbTrue)))
               
                        '���
                        wkst.Cells(lngRows, 12) = Trim$("" & rs!MEAS)
                        '
                        wkst.Cells(lngRows, 9) = "1"

                        intBoxNum = intBoxNum + 1
            
                        '�趨ˮƽ����ֱ����
                        wkst.Range(Chr(12 + 64) & lngRows & ":" & Chr(15 + 64) & lngRows).horizontalAlignment = xlCenter
                        wkst.Range(Chr(12 + 64) & lngRows & ":" & Chr(15 + 64) & lngRows).verticalAlignment = xlCenter
                        '--------------------------
                
                        IntBMegerRow = IntBMegerRow + intBegin
                    
                        intBegin = 1
                    Else
                        '�ϲ�
'                        Application.DisplayAlerts = 0
'                        Selection.Merge
'                        Application.DisplayAlerts = 1
                        
                        
                        wkst.Range(Chr(12 + 64) & IntBMegerRow & ":" & Chr(12 + 64) & IntEMegerRow).Merge
                        wkst.Range(Chr(9 + 64) & IntBMegerRow & ":" & Chr(9 + 64) & IntEMegerRow).Merge
                        wkst.Range(Chr(10 + 64) & IntBMegerRow & ":" & Chr(10 + 64) & IntEMegerRow).Merge
                        wkst.Range(Chr(11 + 64) & IntBMegerRow & ":" & Chr(11 + 64) & IntEMegerRow).Merge
                        '�趨ˮƽ����ֱ����
                        wkst.Range(Chr(12 + 64) & IntBMegerRow & ":" & Chr(15 + 64) & IntEMegerRow).horizontalAlignment = xlCenter
                        wkst.Range(Chr(12 + 64) & IntBMegerRow & ":" & Chr(15 + 64) & IntEMegerRow).verticalAlignment = xlCenter
                        '--------------------------
                        intBegin = intBegin + 1
                        

                    End If

                    DblNum = DblNum + Val(Trim$("" & rs!����))
            
                    lngRows = lngRows + 1
                    IntEMegerRow = lngRows
                    rs.MoveNext
'               Case "HK080"
                Case "HK075"
                    'wkst.Cells(lngRows - 1, 1) = Trim("Item")
                    'wkst.Cells(lngRows - 1, 2) = Trim("Customer PO No")
                    'wkst.Cells(lngRows - 1, 3) = Trim("Line Item")
                    'wkst.Cells(lngRows - 1, 4) = Trim("Customer P/N")
                    'wkst.Cells(lngRows - 1, 5) = Trim("12 NC")
                    'wkst.Cells(lngRows - 1, 6) = Trim("HT P/N")
                    'wkst.Cells(lngRows - 1, 7) = Trim("Wafer Lot No")
                   ' wkst.Cells(lngRows - 1, 8) = Trim("Specification")
                  '  wkst.Cells(lngRows - 1, 8) = Trim("Die Qty(PCS)")
                    'wkst.Cells(lngRows - 1, 8) = Trim("Carton No")
                   ' wkst.Cells(lngRows - 1, 8) = Trim("N.W. ����(Kgs)")
                   ' wkst.Cells(lngRows - 1, 8) = Trim("G.W.ë��(Kgs)")
                   ' wkst.Cells(lngRows - 1, 8) = Trim("MEAS Cm (L*W*H)")
                    
                    T = T + 1
                    wkst.Cells(lngRows, 1) = Trim$("" & T)
                    wkst.Cells(lngRows, 2) = "'" & Trim$("" & rs!PO_NUM)
                    wkst.Cells(lngRows, 3) = "'" & Trim$("" & rs!LineItem)
                    wkst.Cells(lngRows, 4) = "'" & Trim$("" & rs!NC12)
                    wkst.Cells(lngRows, 5) = "'" & Trim$("" & rs!MPN_DESC)
                    wkst.Cells(lngRows, 6) = "'" & Trim$("" & rs!�Ϻ�)
                    wkst.Cells(lngRows, 7) = "'" & Trim$("" & rs!������)
                    wkst.Cells(lngRows, 8) = "'" & Specification1
                    wkst.Cells(lngRows, 9) = Trim$("" & rs!����)
                    DblNum = DblNum + Val(Trim$("" & rs!����))
                    strSql = "select RIGHT(Replace(rtrim(���̿����),'+',''),2) as waferid FROM erpdata..tblstockmovesub where  �Ϻ� = '" & Trim(rs!�Ϻ�) & "' and  ������ = '" & Trim(rs!������) & "'and ���ݱ�� = '" & Trim(rs!��������) & "' order by RIGHT(Replace(rtrim(���̿����),'+',''),2)"
                    wkst.Cells(lngRows, 10) = Get_SqlserverCnt(strSql) 'Ƭ��
                    DblNum2 = wkst.Cells(lngRows, 10) + DblNum2
                    
                    wkst.Cells(lngRows, 12) = Val(Trim$("" & FormatNumber(rs!���� * 0.25, 2, vbTrue))) '����
                    DblJZ = DblJZ + Val(Trim$("" & FormatNumber(rs!���� * 0.25, 2, vbTrue)))
                    strPBigBox = Trim$("" & rs!�����)
                      
                    If strPBigBox <> strNBigBox Then
        
                        strNBigBox = Trim$("" & rs!�����)
              
                        'ë��
                        wkst.Cells(lngRows, 13) = Val(Trim$("" & rs!����))
                        DblMZ = DblMZ + Val(Trim$("" & rs!����))
            
                        '����
                        'wkst.Cells(lngRows, 12) = Val(Trim$("" & FormatNumber(rs!���� * 0.25, 2, vbTrue)))
                   '     DblJZ = DblJZ + Val(Trim$("" & FormatNumber(rs!���� * 0.25, 2, vbTrue)))
           
                        '���
                        wkst.Cells(lngRows, 14) = Trim$("" & rs!MEAS)
                        '
                        wkst.Cells(lngRows, 11) = "1"
                        intBoxNum = intBoxNum + 1
        
                        '�趨ˮƽ����ֱ����
                        wkst.Range(Chr(11 + 64) & lngRows & ":" & Chr(11 + 64) & lngRows).horizontalAlignment = xlCenter
                        wkst.Range(Chr(11 + 64) & lngRows & ":" & Chr(11 + 64) & lngRows).verticalAlignment = xlCenter
                        wkst.Range(Chr(13 + 64) & lngRows & ":" & Chr(13 + 64) & lngRows).horizontalAlignment = xlCenter
                        wkst.Range(Chr(13 + 64) & lngRows & ":" & Chr(13 + 64) & lngRows).verticalAlignment = xlCenter
                        wkst.Range(Chr(14 + 64) & lngRows & ":" & Chr(14 + 64) & lngRows).horizontalAlignment = xlCenter
                        wkst.Range(Chr(14 + 64) & lngRows & ":" & Chr(14 + 64) & lngRows).verticalAlignment = xlCenter
                        '--------------------------
            
                        IntBMegerRow = IntBMegerRow + intBegin
                
                        intBegin = 1
                    Else
                        '�ϲ�
                       ' wkst.Range(Chr(10 + 64) & IntBMegerRow & ":" & Chr(10 + 64) & IntEMegerRow).Merge
                        wkst.Range(Chr(11 + 64) & IntBMegerRow & ":" & Chr(11 + 64) & IntEMegerRow).Merge
                     '   wkst.Range(Chr(12 + 64) & IntBMegerRow & ":" & Chr(12 + 64) & IntEMegerRow).Merge
                        wkst.Range(Chr(13 + 64) & IntBMegerRow & ":" & Chr(13 + 64) & IntEMegerRow).Merge
                        wkst.Range(Chr(14 + 64) & IntBMegerRow & ":" & Chr(14 + 64) & IntEMegerRow).Merge
                        '�趨ˮƽ����ֱ����
                      '  wkst.Range(Chr(10 + 64) & IntBMegerRow & ":" & Chr(13 + 64) & IntEMegerRow).horizontalAlignment = xlCenter
                     '   wkst.Range(Chr(10 + 64) & IntBMegerRow & ":" & Chr(13 + 64) & IntEMegerRow).verticalAlignment = xlCenter
                        '--------------------------
                        intBegin = intBegin + 1

                    End If
                       
                  '  strSql = "  SELECT SUM(CONVERT(INT, a.pass_die)) FROM erptemp..tblshipreport_new a WHERE a.ship_order = '" & rs!�������� & "' AND a.qbox = '" & rs!С��� & "' "
                   ' wkst.Cells(lngRows, 10) = Get_SqlStr(strSql)
                  '  DblNum2 = DblNum2 + Get_SqlStr(strSql)

                    lngRows = lngRows + 1
                    IntEMegerRow = lngRows
                    rs.MoveNext
                
                Case "HK005"
                    T = T + 1
                    wkst.Cells(lngRows, 1) = Trim$("" & T)
                    wkst.Cells(lngRows, 2) = Trim$("" & rs!PO_NUM)
                    wkst.Cells(lngRows, 3) = Trim$("" & rs!MPN_DESC)
                    If (InStr(rs!�Ϻ�, "B") < 7) Then
                        wkst.Cells(lngRows, 4) = Trim$("" & Mid(rs!�Ϻ�, 3, 9))
                    Else
                    wkst.Cells(lngRows, 4) = Trim$("" & Mid(rs!�Ϻ�, 3, InStr(rs!�Ϻ�, "B") - 2))
                    End If
                    wkst.Cells(lngRows, 5) = Trim$("" & rs!�Ϻ�)
                    wkst.Cells(lngRows, 6) = Trim$("" & rs!������)
        
                    If Not rs.EOF Then
            
                        strSql = "select RIGHT(Replace(rtrim(���̿����),'+',''),2) as waferid FROM erpdata..tblstockmovesub where  �Ϻ� = '" & Trim(rs!�Ϻ�) & "' and  ������ = '" & Trim(rs!������) & "'and ���ݱ�� = '" & Trim(rs!��������) & "' order by RIGHT(Replace(rtrim(���̿����),'+',''),2)"
  
                        If RsNew.State = adStateOpen Then RsNew.Close

                        RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

                        RsNew.MoveFirst

                        For j = 1 To RsNew.RecordCount
                
                            If j = 1 Then
                
                                waferid1 = RsNew("waferid")
                
                            Else
                    
                                waferid1 = waferid1 & "," & RsNew("waferid") & ","
                    
                            End If

                            RsNew.MoveNext
            
                        Next
        
                        If RsNew.RecordCount = 25 Then
        
                            waferid1 = "#" & "1-25"
        
                        End If

                        RsNew.Clone
                        
                        Set RsNew = Nothing

                        wkst.Cells(lngRows, 7) = waferid1
                        wkst.Cells(lngRows, 8) = Specification1
                        wkst.Cells(lngRows, 9) = Trim$("" & rs!����)
                        strPBigBox = Trim$("" & rs!�����)
            
                        If strPBigBox <> strNBigBox Then
            
                            strNBigBox = Trim$("" & rs!�����)
                  
                            'ë��
                            wkst.Cells(lngRows, 13) = Val(Trim$("" & rs!����))
                            DblMZ = DblMZ + Val(Trim$("" & rs!����))
                
                            '����
                            wkst.Cells(lngRows, 12) = Val(Trim$("" & FormatNumber(rs!���� * 0.25, 2, vbTrue)))
                            DblJZ = DblJZ + Val(Trim$("" & FormatNumber(rs!���� * 0.25, 2, vbTrue)))
               
                            '���
                            wkst.Cells(lngRows, 14) = Trim$("" & rs!MEAS)
                            '
                            wkst.Cells(lngRows, 11) = "1"
                            intBoxNum = intBoxNum + 1
            
                            '�趨ˮƽ����ֱ����
                            wkst.Range(Chr(10 + 64) & lngRows & ":" & Chr(13 + 64) & lngRows).horizontalAlignment = xlCenter
                            wkst.Range(Chr(10 + 64) & lngRows & ":" & Chr(13 + 64) & lngRows).verticalAlignment = xlCenter
                            '--------------------------
                
                            IntBMegerRow = IntBMegerRow + intBegin
                    
                            intBegin = 1
                        Else
                            '�ϲ�
                            wkst.Range(Chr(10 + 64) & IntBMegerRow & ":" & Chr(10 + 64) & IntEMegerRow).Merge
                            wkst.Range(Chr(11 + 64) & IntBMegerRow & ":" & Chr(11 + 64) & IntEMegerRow).Merge
                            wkst.Range(Chr(12 + 64) & IntBMegerRow & ":" & Chr(12 + 64) & IntEMegerRow).Merge
                            wkst.Range(Chr(13 + 64) & IntBMegerRow & ":" & Chr(13 + 64) & IntEMegerRow).Merge
                            '�趨ˮƽ����ֱ����
                            wkst.Range(Chr(10 + 64) & IntBMegerRow & ":" & Chr(13 + 64) & IntEMegerRow).horizontalAlignment = xlCenter
                            wkst.Range(Chr(10 + 64) & IntBMegerRow & ":" & Chr(13 + 64) & IntEMegerRow).verticalAlignment = xlCenter
                            '--------------------------
                            intBegin = intBegin + 1

                        End If
                           
                            strSql = "  SELECT SUM(CONVERT(INT, a.pass_die)) FROM erptemp..tblshipreport_new a WHERE a.ship_order = '" & rs!�������� & "' AND a.qbox = '" & rs!С��� & "' "
                    
                        wkst.Cells(lngRows, 10) = Get_SqlStr(strSql)
                        DblNum2 = DblNum2 + Get_SqlStr(strSql)

                        DblNum = DblNum + Val(Trim$("" & rs!����))
                        'DblNum2 = wkst.Cells(lngRows, 10) + DblNum2
                        lngRows = lngRows + 1
                        IntEMegerRow = lngRows
                        rs.MoveNext
            
                    End If
                 Case "TW067", "ͨ��ģ��"
                     T = T + 1
                    wkst.Cells(lngRows, 1) = Trim$("" & T)
                    wkst.Cells(lngRows, 2) = Trim$("" & rs!PO_NUM)
                    wkst.Cells(lngRows, 3) = Trim$("" & rs!MPN_DESC)
                    wkst.Cells(lngRows, 4) = Trim$("" & Mid(rs!�Ϻ�, 3, 9))
                    wkst.Cells(lngRows, 5) = Trim$("" & rs!�Ϻ�)
                    wkst.Cells(lngRows, 6) = Trim$("" & rs!������)
                    wkst.Cells(lngRows, 7) = Specification1
                    wkst.Cells(lngRows, 8) = Trim$("" & rs!����)
                        strPBigBox = Trim$("" & rs!�����)
            
                        If strPBigBox <> strNBigBox Then
            
                            strNBigBox = Trim$("" & rs!�����)
                            
                             '
                            wkst.Cells(lngRows, 9) = "1"
                            intBoxNum = intBoxNum + 1
                            
                             '����
                            wkst.Cells(lngRows, 10) = Val(Trim$("" & FormatNumber(rs!���� * 0.25, 2, vbTrue)))
                            DblJZ = DblJZ + Val(Trim$("" & FormatNumber(rs!���� * 0.25, 2, vbTrue)))
               
                            
                            'ë��
                            wkst.Cells(lngRows, 11) = Val(Trim$("" & rs!����))
                            DblMZ = DblMZ + Val(Trim$("" & rs!����))
                
                           
                            '���
                            wkst.Cells(lngRows, 12) = Trim$("" & rs!MEAS)
                           
            
                            '�趨ˮƽ����ֱ����
                            wkst.Range(Chr(10 + 64) & lngRows & ":" & Chr(13 + 64) & lngRows).horizontalAlignment = xlCenter
                            wkst.Range(Chr(10 + 64) & lngRows & ":" & Chr(13 + 64) & lngRows).verticalAlignment = xlCenter
                            '--------------------------
                
                            IntBMegerRow = IntBMegerRow + intBegin
                    
                            intBegin = 1
                        Else
                            '�ϲ�
                            wkst.Range(Chr(9 + 64) & IntBMegerRow & ":" & Chr(9 + 64) & IntEMegerRow).Merge
                            wkst.Range(Chr(10 + 64) & IntBMegerRow & ":" & Chr(10 + 64) & IntEMegerRow).Merge
                            wkst.Range(Chr(11 + 64) & IntBMegerRow & ":" & Chr(11 + 64) & IntEMegerRow).Merge
                            wkst.Range(Chr(12 + 64) & IntBMegerRow & ":" & Chr(12 + 64) & IntEMegerRow).Merge
                            '�趨ˮƽ����ֱ����
                            wkst.Range(Chr(10 + 64) & IntBMegerRow & ":" & Chr(13 + 64) & IntEMegerRow).horizontalAlignment = xlCenter
                            wkst.Range(Chr(10 + 64) & IntBMegerRow & ":" & Chr(13 + 64) & IntEMegerRow).verticalAlignment = xlCenter
                            '--------------------------
                            intBegin = intBegin + 1

                        End If

                        DblNum = DblNum + Val(Trim$("" & rs!����))
                        DblNum2 = wkst.Cells(lngRows, 10) + DblNum2
                        lngRows = lngRows + 1
                        IntEMegerRow = lngRows
                        rs.MoveNext


                Case "GC"
                    Dim s As Integer
                    Dim weight As Double
                    Dim netweight As Double
                    Dim lngRows1 As Integer
                    lngRows1 = 0
                    
                    If IsNumeric(Trim$(rs!����)) Then
                        weight = Val(Trim$(rs!����))
                    Else
                        weight = 0
                    End If
                    If IsNumeric(Trim$(rs!����)) Then
                       'netweight = Val(Trim$(rs!����))
                    netweight = Round(Trim$(rs!����) / 60000, 2)
                      
                    Else
                        netweight = 0
                    End If
                    b = Split(Trim$("" & rs!MPN_DESC), "-")

                    acpn = b(0)
                                    
                    AddSql2 (" insert into erptemp.dbo.ksinvoice values('" & Trim$(rs!PO_NUM) & "','" & Trim$(rs!Specification) & "','" & Trim$(acpn) & "','" & Trim$(rs!�Ϻ�) & "','" & Trim$(rs!����) & "','0','0','0','0','" & Trim$(rs!�����) & "','" & netweight & "','" & weight & "','" & Trim$(rs!MEAS) & "')")
                    
                    ShipOrderFlag = True
                    For S_I = 0 To UBound(Split(ShipOrder, ","))
                        If Trim$(rs!��������) = Split(ShipOrder, ",")(S_I) Then
                            ShipOrderFlag = False '�˷��������Ѳ�ѯ��NG DIE���������ٲ�
                        End If
                    Next
                    
                    If ShipOrderFlag = True Then     '�ж�ShipOrderFlag ��ȥ�ظ�ShipOrder
                            ShipOrder = ShipOrder & Trim$(rs!��������) & "','"
                    End If
                    
                    rs.MoveNext
                    End Select
                Next
                  If UCase(Combo1.text) = "GC" Then
                    ShipOrder = Mid(ShipOrder, 1, Len(ShipOrder) - 3)
                    
                        strSql = "select  SUM(ISNULL(ISNULL(T.E, n.NDPW), 0)) as ���� ,T.cust_device" & _
                        " From " & _
                        "( SELECT 'HTKS' AS sub_name, d.SHIP_SITE,a.ship_order," & _
                        " RTRIM(d.FAB_CONV_ID) as FAB_CONV_ID,  a.cust_device, a.gcversion,d.PO_NUM, a.create_date, rtrim(a.lot_id) as lot_id, SUBSTRING(REPLACE(b.���̿����, '+', ''), LEN(a.lot_id) + 1, 2) as waferid," & _
                        " c.FAILBINCOUNT + c.PASSBINCOUNT AS gross_die, CASE WHEN n.KEY_NAME in ( 'BAD1_DIE','BAD2_DIE') THEN 'E'  ELSE 'A' END Grade, CONVERT(INT,n.KEY_VALUE ) AS qty,  c.PRODUCTID, rtrim(ay.���) as ���," & _
                        " b.�󹤵�,  a.qbox, b.���̿����, SUBSTRING(ee.SFC_ID, 12, 8) AS SFC" & _
                        " FROM erptemp .. tblshipreport_new a" & _
                        " INNER JOIN erpdata .. tblStockNumTree ax  ON ax.��� = a.qbox" & _
                        " INNER JOIN erpdata .. tblStockNumTree ay ON ay.��� = ax.�ϼ����" & _
                        " INNER JOIN erpdata .. tblStocksqfhsub b ON b.���ݱ�� = a.ship_order  AND b.��� = a.qbox   AND b.������ = a.lot_id" & _
                        " INNER JOIN ERPBASE .. tblmappingData c  ON c.SUBSTRATEID = b.���̿���� AND c.LOTID = b.������" & _
                        " INNER JOIN erpbase .. tblCustomerOI d  ON CONVERT(VARCHAR(20), CONVERT(int,d.ID)) = c.FILENAME  AND d.SOURCE_BATCH_ID = c.LOTID" & _
                        " LEFT JOIN  erpdata..tblErpInStockDetailInfo m ON m.KEY_VALUE = b.���" & _
                        " LEFT JOIN  erpdata..tblErpInStockDetailInfo n  ON n.BOX_ID = m.BOX_ID  and n.KEY_NAME in ( 'BAD1_DIE','BAD2_DIE','GOOD_DIE') and n.KEY_TYPE = 'WAFER' AND   CHARINDEX(c.SUBSTRATEID , n.KEYID ) <> 0" & _
                        " inner JOIN erpdata .. tblErpInStockRelation ee ON    ee.BOX_ID = n.BOX_ID AND  ee.WAFER_ID = n.KEYID  WHERE a.ship_order  in('" & ShipOrder & "')) AS p  PIVOT(sum(qty) FOR Grade IN(A,BIN1, E)) AS T" & _
                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV k  ON k.QBOXNUMBER = t.qbox AND k.WAFERSCRIBENUMBER = t.���̿���� AND k.CONTAINERNAME LIKE '%-A'" & _
                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV L ON L.QBOXNUMBER = t.qbox AND L.WAFERSCRIBENUMBER = t.���̿���� AND L.CONTAINERNAME LIKE '%-A-01'" & _
                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV m  ON m.QBOXNUMBER = t.qbox  AND m.WAFERSCRIBENUMBER = t.���̿���� AND m.CONTAINERNAME LIKE '%-A-02'" & _
                        " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV n  ON n.QBOXNUMBER = t.qbox AND n.WAFERSCRIBENUMBER = t.���̿���� AND n.CONTAINERNAME LIKE '%-E'" & _
                        " GROUP BY T.cust_device"
                        
                        
                    
                        If RsNew.State = adStateOpen Then RsNew.Close

                        RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
                   
                        If RsNew.RecordCount > 0 Then
                            RsNew.MoveFirst
                            ReDim a(RsNew.RecordCount)
                            For N = 1 To RsNew.RecordCount
                                DblNum2 = DblNum2 + Val(Trim$("" & RsNew!����))
                                a(lngRows1) = Val(Trim$("" & RsNew!����))
                                lngRows1 = lngRows1 + 1
                                RsNew.MoveNext
                            Next
                        End If
                        
                        RsNew.Clone
                    
                        Set RsNew = Nothing

                       
        
        wkst.Cells(lngRows - 1, 12) = "NG:DIE"
        
            j = 12
                
            m = 12
            
            DieNoFound = False
     
                strSql = "SELECT ROW_NUMBER() OVER(order by a.acpn) item,PO_NUM=stuff((SELECT DISTINCT '/' + PO_NUM FROM erptemp.dbo.ksinvoice WHERE acpn=a.acpn AND �Ϻ�=a.�Ϻ� AND Specification=a.Specification " & _
                "for xml path('')),1, 1, ''),a.*,b.���� FROM (SELECT Specification,acpn,�Ϻ�,SUM(����) AS ����,SUM(����) As ����,SUM(ë��) As ë�� " & _
                "FROM erptemp.dbo.ksinvoice  GROUP BY Specification,acpn,�Ϻ�) a " & _
                "INNER JOIN (SELECT acpn,�Ϻ�,COUNT(*) AS ���� FROM erptemp.dbo.ksinvoice GROUP BY Specification,acpn,�Ϻ�) b  ON  b.acpn = b.acpn AND b.�Ϻ� = a.�Ϻ�"
                                                                    
  
            If RsNew.State = adStateOpen Then RsNew.Close

            RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
                
            IntInertRow = RsNew.RecordCount

            For i = 1 To IntInertRow - 1
            
                wkst.Rows(lngRows & ":" & lngRows).Select
                ExApp.Selection.Copy
                ExApp.Selection.Insert Shift:=xlDown
                
            Next i
                    
            RsNew.MoveFirst
            
            For N = 1 To RsNew.RecordCount
                s = N - 1
                wkst.Cells(lngRows, 12) = a(s)
                T = T + 1
                wkst.Cells(lngRows, 1) = Trim$("" & T)

                    wkst.Cells(lngRows, 2) = Replace(Trim$("" & RsNew!PO_NUM), "/", "/" & Chr(10)) 'һ��PO_NUMһ��
                    wkst.Cells(lngRows, 3) = Specification1 'оƬ

                wkst.Cells(lngRows, 4) = Trim$("" & RsNew!acpn)
                wkst.Cells(lngRows, 5) = Trim$("" & RsNew!�Ϻ�)
                wkst.Cells(lngRows, 6) = Trim$("" & RsNew!����)
                
                'wkst.Cells(lngRows, 9) = Trim$("" & RsNew!����)
                
                wkst.Cells(lngRows, 9) = Round(Trim$(RsNew!����) / 60000, 2)
                
                wkst.Cells(lngRows, 10) = Trim$("" & RsNew!ë��)
                strmeas = ""
         
                
                    sstrSql = "select distinct MEAS from erptemp.dbo.ksinvoice where  acpn = '" & Trim$("" & RsNew!acpn) & "' and �Ϻ� = '" & Trim$("" & RsNew!�Ϻ�) & "'"
                

                If RsNew2.State = adStateOpen Then RsNew2.Close

                RsNew2.Open sstrSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
                
                If RsNew2.RecordCount > 0 Then
                    RsNew2.MoveFirst
                    
                    For p = 1 To RsNew2.RecordCount
                    
                     ' If RsNew2.RecordCount > 1 Then
                        If strmeas = "" Then
                            strmeas = Trim$("" & RsNew2!MEAS)
                        Else
                            strmeas = Trim$("" & RsNew2!MEAS) & "/" & strmeas
                        End If
                        
                    '  Else
                        
                    '    strmeas = Trim$("" & RsNew2!MEAS)
                        
                   '   End If
                      
                      RsNew2.MoveNext
                      
                    Next
                End If
                
                RsNew2.Clone

                Set RsNew2 = Nothing
        
                
                wkst.Cells(lngRows, 11) = Replace(Trim$("" & strmeas), "/", "/" & Chr(10))
                
                wkst.Cells(lngRows, 8) = Trim$("" & RsNew!����)
                
                DblNum1 = DblNum1 + Val(Trim$("" & RsNew!����))
                DblMZ = DblMZ + Val(Trim$("" & RsNew!ë��))
                'DblJZ = DblJZ + Val(Trim$("" & RsNew!����))
                
                 DblJZ = DblJZ + Round(Trim$(RsNew!����) / 60000, 2)
                
                strSql = "select DISTINCT DIE from erptemp.dbo.customerkspn WHERE CUSTOMERPN = '" & RsNew!acpn & "' and �ߴ� = '" & Left(Trim$("" & RsNew!�Ϻ�), 2) & "'"
                
                If RsNew1.State = adStateOpen Then RsNew1.Close
                
                RsNew1.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
                
                asum = 0
                TOTALWAFER = 0
                
                If RsNew1.RecordCount > 0 Then
                
                    asum = Trim$("" & RsNew1!die)
                    
                    TOTALWAFER = Val(RsNew!����) / asum

                    DblPnum = DblPnum + Val(Trim$("" & TOTALWAFER))
                Else
                
                    MsgBox RsNew!acpn & "�ͻ������޶�Ӧ��DIE������ά����", vbInformation, "��ʾ��"
                    
                    DieNoFound = True
                End If
                
                RsNew1.Clone
                
                Set RsNew1 = Nothing
                
                wkst.Cells(lngRows, 7) = Trim$("" & TOTALWAFER)
                
                'DblPnum = DblPnum + Val(Trim$("" & Val(RsNew!����) / asum))
                
                If j = 12 Then
                
                    wkst.Cells(j, 6) = "NOTE:1PC " & RsNew!acpn & " = " & asum & " EA;Total: " & TOTALWAFER & " PCS " & RsNew!acpn & " = " & Val(RsNew!����) & " EA"
                
                Else

                    If j <= 19 And j > 12 Then
                
                        wkst.Cells(j, 6) = "         1PC " & RsNew!acpn & " = " & asum & " EA;Total: " & TOTALWAFER & " PCS " & RsNew!acpn & " = " & Val(RsNew!����) & " EA"
                
                    End If
                
                End If
                
                DblNum = DblNum + Val(Trim$("" & RsNew!����))
                
                
                
                If j > 19 Then
                
                    wkst.Cells(m, 9) = "1PC " & RsNew!acpn & " = " & asum & " EA;Total: " & TOTALWAFER & " PCS " & RsNew!acpn & " = " & Val(RsNew!����) & " EA"
                
                    m = m + 1
                
                    If m > 19 Then
                
                        MsgBox "��ʽ�Ѿ�������Χ��", vbInformation, "��ʾ��"
                
                        Exit Sub
                
                    End If
                
                End If
                j = j + 1
                lngRows = lngRows + 1

                RsNew.MoveNext
                    
            Next

            RsNew.Clone

            Set RsNew = Nothing
        
        End If
        

        
        If UCase(Combo1.text) = "68" Then
            '�������
            wkst.Cells(lngRows, 10) = DblNum & "PCS"
            wkst.Cells(lngRows, 11) = (intBoxNum - 1) & "CARTONS"
            wkst.Cells(lngRows, 12) = FormatNumber(DblJZ, 2, vbTrue) & "KGS"
            wkst.Cells(lngRows, 13) = FormatNumber(DblMZ, 2, vbTrue) & "KGS"
        
            wkst.Cells(lngRows + 3, 3) = FormatNumber(DblJZ, 2, vbTrue) & "KGS"
            wkst.Cells(lngRows + 4, 3) = FormatNumber(DblMZ, 2, vbTrue) & "KGS"
            wkst.Cells(lngRows + 5, 3) = (intBoxNum - 1) & "CARTONS"
        ElseIf UCase(Combo1.text) = "HK005" Then
            '�������
            wkst.Cells(lngRows, 9) = DblNum & "PCS"
            wkst.Cells(lngRows, 10) = DblNum2 & "PCS"
            wkst.Cells(lngRows, 11) = (intBoxNum - 1) & "CARTONS"
            wkst.Cells(lngRows, 12) = DblJZ & "KGS"
            wkst.Cells(lngRows, 13) = DblMZ & "KGS"
            
            wkst.Cells(lngRows + 2, 3) = (DblNum - DblNum2) & "PCS"
            wkst.Cells(lngRows + 3, 3) = DblJZ & "KGS"
            wkst.Cells(lngRows + 4, 3) = DblMZ & "KGS"
            wkst.Cells(lngRows + 5, 3) = (intBoxNum - 1) & "CARTONS"
        ElseIf UCase(Combo1.text) = "TW067" Or UCase(Combo1.text) = "ͨ��ģ��" Then
            '�������
            wkst.Cells(lngRows, 8) = DblNum & "PCS"
            wkst.Cells(lngRows, 9) = (intBoxNum - 1) & "CARTONS"
            wkst.Cells(lngRows, 10) = DblJZ & "KGS"
            wkst.Cells(lngRows, 11) = DblMZ & "KGS"
        
            wkst.Cells(lngRows + 3, 3) = DblJZ & "KGS"
            wkst.Cells(lngRows + 4, 3) = DblMZ & "KGS"
            wkst.Cells(lngRows + 5, 3) = (intBoxNum - 1) & "CARTONS"
        ElseIf UCase(Combo1.text) = "76" Then
            '�������
            wkst.Cells(lngRows, 6) = DblNum & "PCS"
            
            'wkst.Cells(lngRows, 7) = shu & "CARTONS"
            wkst.Cells(lngRows, 7) = CARTONCNT & "CARTONS"
            wkst.Cells(lngRows, 8) = DblJZ & "KGS"
            wkst.Cells(lngRows, 9) = DblMZ & "KGS"
        
            wkst.Cells(lngRows + 3, 3) = DblJZ & "KGS"
            wkst.Cells(lngRows + 4, 3) = DblMZ & "KGS"
            'wkst.Cells(lngRows + 5, 3) = shu & "CARTONS"
            wkst.Cells(lngRows + 5, 3) = CARTONCNT & "CARTONS"
            
        ElseIf UCase(Combo1.text) = "SG005" Then
            wkst.Cells(lngRows, 6) = DblNum & "PCS"
            wkst.Cells(lngRows, 7) = (intBoxNum - 1) & "CARTONS"
            wkst.Cells(lngRows, 8) = Format(DblJZ, "0.00") & "KGS"
            wkst.Cells(lngRows, 9) = Format(DblMZ, "0.00") & "KGS"
        
            wkst.Cells(lngRows + 3, 3) = DblJZ & "KGS"
            wkst.Cells(lngRows + 4, 3) = DblMZ & "KGS"
            wkst.Cells(lngRows + 5, 3) = (intBoxNum - 1) & "CARTONS"
            
        ElseIf UCase(Combo1.text) = "SG005_SO" Then
            wkst.Cells(lngRows, 8) = DblNum & "PCS"
            wkst.Cells(lngRows, 9) = (intBoxNum - 1) & "CARTONS"
            wkst.Cells(lngRows, 10) = Format(DblJZ, "0.00") & "KGS"
            wkst.Cells(lngRows, 11) = Format(DblMZ, "0.00") & "KGS"
        
            wkst.Cells(lngRows + 3, 5) = DblJZ & "KGS"
            wkst.Cells(lngRows + 4, 5) = DblMZ & "KGS"
            wkst.Cells(lngRows + 5, 5) = (intBoxNum - 1) & "CARTONS"
  
        ElseIf UCase(Combo1.text) = "GC" Then
            '�������
         
            wkst.Cells(lngRows, 6) = DblNum & "EA"
            If DieNoFound = True Then
                wkst.Cells(lngRows, 7) = "Ƭ"
            Else
                wkst.Cells(lngRows, 7) = DblPnum & "Ƭ"
            End If
            wkst.Cells(lngRows, 8) = DblNum1 & "CARTONS"
            wkst.Cells(lngRows, 9) = DblJZ & "KGS"
            wkst.Cells(lngRows, 10) = DblMZ & "KGS"
            
            wkst.Cells(lngRows + 3, 3) = DblNum2 & "EA"
            wkst.Cells(lngRows + 4, 3) = DblJZ & "KGS"
            wkst.Cells(lngRows + 5, 3) = DblMZ & "KGS"
            wkst.Cells(lngRows + 6, 3) = DblNum1 & "CARTONS"
        ElseIf UCase(Combo1.text) = "HK075" Then
            '�������
            wkst.Cells(lngRows, 9) = DblNum & "PCS"
            wkst.Cells(lngRows, 10) = DblNum2 & "Ƭ"
            wkst.Cells(lngRows, 11) = (intBoxNum - 1) & "CARTONS"
            wkst.Cells(lngRows, 12) = DblJZ & "KGS"
            wkst.Cells(lngRows, 13) = DblMZ & "KGS"
            
          '  wkst.Cells(lngRows + 2, 3) = (DblNum - DblNum2) & "PCS"
          '  wkst.Cells(lngRows + 3, 3) = DblJZ & "KGS"
            wkst.Cells(lngRows + 4, 3) = DblJZ & "KGS"
            wkst.Cells(lngRows + 5, 3) = DblMZ & "KGS"
            wkst.Cells(lngRows + 6, 3) = (intBoxNum - 1) & "CARTONS"
            
        Else
            '�������
         
            wkst.Cells(lngRows, 6) = DblNum & "PCS"
            wkst.Cells(lngRows, 7) = DblPnum & "Ƭ"
            wkst.Cells(lngRows, 8) = DblNum1 & "CARTONS"
            wkst.Cells(lngRows, 9) = DblJZ & "KGS"
            wkst.Cells(lngRows, 10) = DblMZ & "KGS"
            
            wkst.Cells(lngRows + 3, 3) = DblNum2 & "EA"
            wkst.Cells(lngRows + 4, 3) = DblJZ & "KGS"
            wkst.Cells(lngRows + 5, 3) = DblMZ & "KGS"
            wkst.Cells(lngRows + 6, 3) = DblNum1 & "CARTONS"

        End If
        
    Else
        '        ClsP.UnLoad_Form
        MsgBox "���赼�����ݣ�", vbInformation, "��ʾ��"
        Exit Sub

    End If

    '
    '    ClsP.ShowProgress 100, "�����ɹ���"
    '
    '    If Not ClsP Is Nothing Then
    '        Set ClsP = Nothing
    '
    '    End If

    ExApp.Visible = True
    
    '    If intFlag = 1 Then
    '        wkst.PrintPreview
    '        wkbk.Close (False)
    '        ExApp.Quit
    '    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing

    End If

    Exit Sub
ErrHandle:

    On Error Resume Next

    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing

    End If

    '    If Not ClsP Is Nothing Then
    '        Set ClsP = Nothing
    '
    '    End If

    MsgBox Err.DESCRIPTION, vbInformation, "��ʾ��"
    Exit Sub

End Sub

Private Sub Command3_Click()


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

    Dim i        As Integer

    Dim j        As Integer

    Dim tempVal  As String

    Dim temp1    As String

    Dim temp2    As String

    Dim temp3    As String

    Dim temp4    As String

    Dim temp5    As String

    Dim temp6    As String

    Dim strChar  As String
    
    Dim SumCount As Integer
    
    Dim VBExcel  As Excel.Application

    Dim xlBook   As Excel.Workbook

    Dim xlSheet  As Excel.Worksheet
    
    If Text2.text = "" Then
        MsgBox "��ѡ����ϴ����ļ�"
        Exit Sub

    End If
    
    SumCount = 0

    'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text2.text)    '���ļ�

    Set xlSheet = xlBook.Worksheets("sheet1")        '��sheet�еı�

    '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 3 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        Exit Sub

    End If

    For i = 1 To xlSheet.Range("A1").CurrentRegion.Rows.count
   
        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
        
            strChar = Chr(96 + j)
            tempVal = Trim(xlSheet.Range(strChar & i).Value)   '��ʱ����ֵ
        
            If j = 1 Then
            
                temp1 = tempVal
            
            End If
        
            If j = 2 Then
                temp2 = tempVal
        
            End If
            
            If j = 3 Then
                temp3 = tempVal
        
            End If
        
        Next j
         
        If Get_SqlserverCnt(" SELECT * FROM erptemp.dbo.customerkspn WHERE  CUSTOMERPN='" & temp1 & "' and �ߴ�='" & temp3 & "'") > 0 Then
            MsgBox temp1 & "�ͻ���������ά����DIE��", vbInformation, "��ʾ"
        Else

            If IsNumeric(temp2) = True Then
                AddSql2 ("insert into erptemp.dbo.customerkspn values('" & temp1 & "','" & temp2 & "','" & temp3 & "')")
                
                SumCount = SumCount + 1
            End If
        End If
               
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

End Sub





















