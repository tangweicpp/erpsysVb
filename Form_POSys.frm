VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form Form_POSys 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�г���������Ϣά��ϵͳ"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13920
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
   ScaleWidth      =   13920
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13440
      Top             =   1200
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
            Picture         =   "Form_POSys.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_POSys.frx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_POSys.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_POSys.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_POSys.frx":2848
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_POSys.frx":349A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTTab0 
      Height          =   12495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   22040
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "������������"
      TabPicture(0)   =   "Form_POSys.frx":40EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Toolbar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "com"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ProgressBar1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "com2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Fra_Datail"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Fra1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "��ѯ/�޸�/ɾ����������"
      TabPicture(1)   =   "Form_POSys.frx":4108
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblKeyID"
      Tab(1).Control(1)=   "lblCusCode(0)"
      Tab(1).Control(2)=   "txtCusCode"
      Tab(1).Control(3)=   "lblCusDev"
      Tab(1).Control(4)=   "txtCusDev"
      Tab(1).Control(5)=   "Fps(1)"
      Tab(1).Control(6)=   "Toolbar2"
      Tab(1).Control(7)=   "txtKID"
      Tab(1).Control(8)=   "cmdSwitch"
      Tab(1).Control(9)=   "chk"
      Tab(1).Control(10)=   "Opt(0)"
      Tab(1).Control(11)=   "Opt(1)"
      Tab(1).Control(12)=   "cmdExportSql"
      Tab(1).Control(13)=   "txtMsg2"
      Tab(1).ControlCount=   14
      Begin VB.TextBox txtMsg2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   1890
         Left            =   -68280
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   30
         Top             =   1200
         Width           =   6615
      End
      Begin VB.Frame Fra1 
         Caption         =   "����ѡ��(OPTION)"
         ForeColor       =   &H000000FF&
         Height          =   3015
         Left            =   0
         TabIndex        =   18
         Top             =   960
         Width           =   14895
         Begin VB.TextBox txtPOQTY 
            Height          =   495
            Left            =   4920
            TabIndex        =   33
            Top             =   1320
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtcust_device 
            Height          =   495
            Left            =   4920
            TabIndex        =   32
            Top             =   2280
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtPo_Price 
            Height          =   495
            Left            =   4920
            TabIndex        =   31
            Top             =   1800
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkMsgAppend 
            Caption         =   "�Ƿ���Ҫ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   2160
            TabIndex        =   29
            Top             =   2625
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.ComboBox cbCusCode 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1755
            TabIndex        =   23
            Top             =   900
            Width           =   1695
         End
         Begin VB.ComboBox cbUploadType 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Form_POSys.frx":4124
            Left            =   1755
            List            =   "Form_POSys.frx":413D
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   570
            Width           =   1695
         End
         Begin VB.ComboBox cbTaxType 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Form_POSys.frx":417E
            Left            =   1755
            List            =   "Form_POSys.frx":4188
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1245
            Width           =   1695
         End
         Begin VB.TextBox txtFilePath 
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Left            =   1755
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   225
            Width           =   4575
         End
         Begin VB.TextBox txtMsg 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H000000FF&
            Height          =   2610
            Left            =   6360
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   19
            Top             =   240
            Width           =   8055
         End
         Begin VB.Label lblFileName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ļ���(N)          "
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   240
            TabIndex        =   28
            Top             =   330
            Width           =   1515
         End
         Begin VB.Label lblCustomerCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ͻ�����     "
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   240
            TabIndex        =   27
            Top             =   1005
            Width           =   1500
         End
         Begin VB.Label lblUploadtype 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������           "
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   240
            TabIndex        =   26
            Top             =   660
            Width           =   1500
         End
         Begin VB.Label lblBand 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��˰/�Ǳ�˰        "
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   240
            TabIndex        =   25
            Top             =   1335
            Width           =   1485
         End
         Begin VB.Label lblMsg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ʼ����Ĳ���(M)                                                                                                "
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   240
            TabIndex        =   24
            Top             =   2625
            Width           =   13230
         End
      End
      Begin VB.Frame Fra_Datail 
         Caption         =   "����״̬(STATUS)"
         ForeColor       =   &H000000FF&
         Height          =   7455
         Left            =   0
         TabIndex        =   16
         Top             =   3960
         Width           =   14895
         Begin FPSpreadADO.fpSpread Fps 
            Height          =   8775
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   14415
            _Version        =   524288
            _ExtentX        =   25426
            _ExtentY        =   15478
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
            MaxCols         =   21
            MaxRows         =   0
            SpreadDesigner  =   "Form_POSys.frx":419A
            Appearance      =   1
            TextTip         =   2
            AppearanceStyle =   0
         End
      End
      Begin VB.CommandButton cmdExportSql 
         BackColor       =   &H00FFC0C0&
         Caption         =   "������ѯ����"
         Height          =   360
         Left            =   -71280
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin VB.OptionButton Opt 
         Caption         =   "��׼��ѯ"
         Height          =   195
         Index           =   1
         Left            =   -73440
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton Opt 
         Caption         =   "ģ����ѯ"
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   8
         Top             =   2130
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CheckBox chk 
         Caption         =   "ȫѡ/��ѡ"
         Height          =   195
         Left            =   -74640
         TabIndex        =   7
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdSwitch 
         BackColor       =   &H00FFC0FF&
         Caption         =   "�������л�: LOTID <--> WAFERID <-->  PO <--> �ͻ�����"
         Height          =   360
         Left            =   -73200
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1260
         Width           =   4935
      End
      Begin VB.TextBox txtKID 
         BackColor       =   &H0080FFFF&
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
         Left            =   -74640
         TabIndex        =   4
         Top             =   1680
         Width           =   3255
      End
      Begin MSComDlg.CommonDialog com2 
         Left            =   13920
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   524800
         MaxFileSize     =   8000
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   615
         Left            =   8160
         TabIndex        =   2
         Top             =   360
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComDlg.CommonDialog com 
         Left            =   14400
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         MaxFileSize     =   800
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   600
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1058
         ButtonWidth     =   3678
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��������            "
               Key             =   "UPLOAD"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��������            "
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   " �˳�����           "
               Key             =   "EXIT"
               ImageIndex      =   3
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   600
         Left            =   -74640
         TabIndex        =   5
         Top             =   480
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   1058
         ButtonWidth     =   4577
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ѯWO                   "
               Key             =   "UPLOAD"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�WO                   "
               ImageIndex      =   5
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "  ɾ��WO                   "
               Key             =   "EXIT"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�����                 "
               ImageIndex      =   3
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   9135
         Index           =   1
         Left            =   -74640
         TabIndex        =   10
         Top             =   3240
         Width           =   14175
         _Version        =   524288
         _ExtentX        =   25003
         _ExtentY        =   16113
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
         MaxCols         =   20
         MaxRows         =   0
         SpreadDesigner  =   "Form_POSys.frx":46B8
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin MSForms.TextBox txtCusDev 
         Height          =   315
         Left            =   -70320
         TabIndex        =   14
         Top             =   2490
         Width           =   1575
         VariousPropertyBits=   746604563
         BorderStyle     =   1
         Size            =   "2778;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCusDev 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ�����:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   -71280
         TabIndex        =   13
         Top             =   2520
         Width           =   1200
      End
      Begin MSForms.TextBox txtCusCode 
         Height          =   315
         Left            =   -73560
         TabIndex        =   12
         Top             =   2490
         Width           =   2175
         VariousPropertyBits=   746604563
         BackColor       =   8454143
         BorderStyle     =   1
         Size            =   "3836;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCusCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ�����:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   0
         Left            =   -74640
         TabIndex        =   11
         Top             =   2520
         Width           =   1200
      End
      Begin VB.Label lblKeyID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOTID:"
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
         Left            =   -74640
         TabIndex        =   3
         Top             =   1320
         Width           =   795
      End
   End
End
Attribute VB_Name = "Form_POSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : Form_POSys
'    Project    : ��ʽ����1
'
'    Description: [type_description_here]
'
'    Modified   :
'<Modified by: Project Administrator at 2019/4/4-11:10:34 on machine: DESKTOP-91AFCV3>
'-------------------------------------------------------------------------------- ' Changed by: Project Administrator at: 2019/4/4-11:10:43 on machine: DESKTOP-91AFCV3
'</Modified by: Project Administrator at 2019/4/4-11:10:34 on machine: DESKTOP-91AFCV3>
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Changed by: Project Administrator at: 2019/4/4-11:10:50 on machine: DESKTOP-91AFCV3
Private Enum fpSDetail

    e_Item = 1
    e_PO_NO = 2
    e_Supplier = 3
    e_Ship_To
    e_FAB_Device
    e_Customer_Device
    e_Wafer_Version
    e_MARKING_CODE
    e_date
    E_LOTID
    E_WAFERID
    e_GoodDieQty
    e_TotalDies
    e_HT_DEVICES
    E_REMARK
    E_Type
    e_Reserved1
    e_Reserved2
    e_Reserved3
    e_Reserved4
    e_price_w
    e_price_d
    e_MCol

End Enum

Dim lPartSec    As Long
Dim bBonded     As Boolean
Dim strFileName As String
Dim gUpID       As String
Dim strCusCode  As String
Dim strCusDev   As String
Dim gBackID     As String
Dim strRealName As String

Private Sub cbTaxType_Click()

Select Case cbTaxType.ListIndex

    Case 0
        bBonded = True

    Case Else
        bBonded = False

End Select

End Sub

Private Sub cbUploadType_Click()

Select Case cbUploadType.ListIndex

    Case 0, 1, 6
        lblBand.Visible = True
        cbTaxType.Visible = True

    Case 3, 4
        cbCusCode.text = "37"
        lblBand.Visible = False
        cbTaxType.Visible = False

    Case Else
        lblBand.Visible = False
        cbTaxType.Visible = False

End Select

End Sub

Private Sub chk_Click()
Dim i As Integer

If chk.Value = 1 Then

    For i = 1 To Fps(1).MaxRows

        With Fps(1)
            .Row = i
            .Col = 1
            .text = 1

        End With

    Next i

ElseIf chk.Value = 0 Then

    For i = 1 To Fps(1).MaxRows

        With Fps(1)
            .Row = i
            .Col = 1
            .text = 0

        End With

    Next i

End If

End Sub

Private Sub cmdExportSql_Click()
Dim xlsApp      As Excel.Application
Dim xlsBook     As Excel.Workbook
Dim xlsSheet    As Excel.Worksheet
Dim i           As Long
Dim j           As Long
Dim strFileName As String

cmdExportSql.Enabled = False

On Error GoTo Ert

If Fps(1).MaxRows = 0 Then
    MsgBox "û�����ݿ��Ե���", vbInformation, "��ʾ"
    cmdExportSql.Enabled = True
    Exit Sub

End If

Set xlsApp = CreateObject("Excel.Application")
Set xlsBook = xlsApp.Workbooks.Add
Set xlsSheet = xlsBook.Worksheets(1)

With xlsApp
    .Rows(1).Font.Bold = True

End With

With Fps(1)

    For i = 0 To Fps(1).MaxRows
        For j = 1 To Fps(1).MaxCols
            .Col = j
            .Row = i
            xlsSheet.Cells(i + 1, j) = ("" & .text)
        Next j
    Next i

End With

xlsApp.Visible = True
strFileName = "C:\others\" & Format(Now, "YYYY-MMDD-HH-MM-SS") & ".xlsx"
xlsBook.SaveAs strFileName
Set xlsApp = Nothing
cmdExportSql.Enabled = True
Exit Sub
Ert:
If Not (xlsApp Is Nothing) Then
    Set xlsApp = Nothing

End If

cmdExportSql.Enabled = True

End Sub

Private Sub cmdSwitch_Click()
If lblKeyID.Caption = "LOTID:" Then
    lblKeyID.Caption = "WAFERID:"
ElseIf lblKeyID.Caption = "WAFERID:" Then
    lblKeyID.Caption = "PONO:"
ElseIf lblKeyID.Caption = "PONO:" Then
    lblKeyID.Caption = "�ͻ�����:"
ElseIf lblKeyID.Caption = "�ͻ�����:" Then
    lblKeyID.Caption = "LOTID:"

End If

End Sub

Private Sub Form_Activate()
SSTTab0.Tab = 0

End Sub

Private Sub Form_Load()
InitData
InitCtrls

End Sub

Private Sub InitData()
Dim strSql As String

gUpID = ""
gBackID = ""

Select Case gUserName

    Case "15507"
        strRealName = "����"

    Case "16452"
        strRealName = "����"

    Case "18035"
        strRealName = "����ܿ"

    Case "7433", "07433"
        strRealName = "���"

    Case "14117"
        strRealName = "����"

    Case "8240", "08240"
        strRealName = "����"

    Case "16368"
        strRealName = "����"

    Case "12089"
        strRealName = "��ǿ"

    Case "12725"
        strRealName = "ȫ����"

    Case "15236"
        strRealName = "�ε�Ƽ"

    Case "16642"
        strRealName = "�ⷼ"

    Case "07885"
        strRealName = "����Ա"

    Case "18420"
        strRealName = "�����"

    Case "18697"
        strRealName = "����"

    Case "18252"
        strRealName = "������"

    Case "18881"
        strRealName = "������"

    Case "19400"
        strRealName = "����"

End Select

strSql = "select EmpName from XTW..employee where empno = '" & gUserName & "'"
strRealName = Get_SqlStr2(strSql)

End Sub

Private Sub InitCtrls()
InitFps
InitCuscode
cbUploadType.ListIndex = 0

End Sub

Private Sub InitFps()

With Fps(0)
    .TypeMaxEditLen = 500
    .MaxRows = 0
    .Col = -1
    .Row = -1
    .Lock = True
    .SetText fpSDetail.e_Item, 0, "״̬"
    .SetText fpSDetail.e_PO_NO, 0, "PO_NO"
    .SetText fpSDetail.e_Supplier, 0, "Supplier"
    .SetText fpSDetail.e_Ship_To, 0, "Ship_To"
    .SetText fpSDetail.e_FAB_Device, 0, "FAB_Device"
    .SetText fpSDetail.e_Customer_Device, 0, "Customer_Device"
    .SetText fpSDetail.e_Wafer_Version, 0, "Wafer_Version"
    .SetText fpSDetail.e_MARKING_CODE, 0, "MARKING_CODE"
    .SetText fpSDetail.e_date, 0, "Date"
    .SetText fpSDetail.E_LOTID, 0, "LotID"
    .SetText fpSDetail.E_WAFERID, 0, "WaferID"
    .SetText fpSDetail.e_GoodDieQty, 0, "GoodDieQty"
    .SetText fpSDetail.e_TotalDies, 0, "TotalDies"
    .SetText fpSDetail.e_HT_DEVICES, 0, "HT_DEVICES"
    .SetText fpSDetail.E_REMARK, 0, "Remark"
    .SetText fpSDetail.E_Type, 0, "ó������"
    .SetText fpSDetail.e_Reserved1, 0, "��ǩԤ���ֶ�1"
    .SetText fpSDetail.e_Reserved2, 0, "��ǩԤ���ֶ�2"
    .SetText fpSDetail.e_Reserved3, 0, "��ǩԤ���ֶ�3"
    .SetText fpSDetail.e_Reserved4, 0, "��ǩԤ���ֶ�4"
    .SetText fpSDetail.e_MCol, 0, "��ǩԤ���ֶ�5"
    .ColWidth(fpSDetail.e_Item) = 10
    .ColWidth(fpSDetail.e_Supplier) = 6

End With

With Fps(1)
    .TypeMaxEditLen = 500
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
    .Col = 1
    .Lock = False
    .Col = 4
    .Lock = False
    .BackColor = vbGreen
    .Col = 5
    .Lock = False
    .BackColor = vbGreen
    .Col = 6
    .Lock = False
    .BackColor = vbGreen
    .Col = 7
    .Lock = False
    .BackColor = vbGreen
    .Col = 12
    .Lock = False
    .BackColor = vbGreen
    .Col = 13
    .Lock = False
    .BackColor = vbGreen
    .Col = 14
    .Lock = False
    .BackColor = vbGreen
    .Col = 15
    .Lock = False
    .BackColor = vbGreen
    .ColWidth(1) = 4

End With

End Sub

Private Sub InitCuscode()
Dim rs As New ADODB.Recordset, i As Integer

Set rs.ActiveConnection = SqlConnect
rs.Source = "select distinct �ͻ����� from tblxcustomer"
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
cbCusCode.Clear
If rs.RecordCount > 0 Then
    rs.MoveFirst

    For i = 1 To rs.RecordCount
        cbCusCode.AddItem Trim(rs("�ͻ�����"))
        rs.MoveNext
    Next i

End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index

    Case 1

        Select Case cbUploadType.ListIndex

            Case 0, 1, 6
                Call imptWO

            Case 3, 4, 5
                Call imptOther

            Case Else
                MsgBox "��δ�����ù���", vbInformation, "��ʾ"
                Exit Sub

        End Select

    Case 2

        Select Case cbUploadType.ListIndex

            Case 0, 1
                Call exptWO

            Case 3, 4
                Call exptOther

            Case Else
                MsgBox "��δ�����ù���", vbInformation, "��ʾ"
                Exit Sub

        End Select

    Case 3
        Unload Me

End Select

End Sub

Private Sub imptWO()

On Error GoTo ErrHandle

Dim i         As Integer
Dim dT        As tyWO
Dim strCusDev As String
Dim strHtDev  As String
Dim VBExcel   As Excel.Application
Dim xlBook    As Excel.Workbook
Dim xlSheet   As Excel.Worksheet
Dim lColsCnt  As Long
Dim lRowsCnt  As Long
Dim strSql    As String
Dim cust_name As String

If cbCusCode.text = "" Then
    MsgBox "��ѡ��ͻ�����", vbInformation, "��ʾ"
    Exit Sub

End If

If cbTaxType.text = "" Then
    MsgBox "��ѡ�񶩵�-��˰/�Ǳ�˰", vbExclamation, "��ʾ"
    Exit Sub

End If

If chkMsgAppend.Value = 1 And txtMsg.text = "" Then
    MsgBox "�������ʼ����Ĳ�����Ϣ,������ȡ����ѡ��ѡ��" & vbCrLf & "{����:��ݵ��ŵ���Ϣ�򲹳�����˵��" & vbCrLf & "��ϵ�绰....", vbInformation, "����"
    Exit Sub

End If

txtFilePath.text = ""
strCusDev = ""
strHtDev = ""
com.Filter = "�����ļ�(*.*)|*.*|Excel�ļ�(*.xls;*.xlsx)|*.xls;*.xlsx"
com.ShowOpen
If com.filename = "" Then
    Exit Sub

End If

txtFilePath.text = com.filename
com.filename = ""
If txtFilePath.text = "" Then
    MsgBox "�ļ���ʧ��", vbInformation, "��ʾ"
    Exit Sub

End If

If InStr(txtFilePath.text, "-A") > 0 Then
    If cbTaxType.ListIndex = 1 Then
        MsgBox "��ȷ���Ƿ�ѡ��˰�Ǳ�˰����", vbCritical, "����"
        Exit Sub

    End If

End If

If InStr(txtFilePath.text, "�Ǳ�˰") > 0 Or InStr(txtFilePath.text, "-B") > 0 Then
    If cbTaxType.ListIndex = 0 Then
        MsgBox "��ȷ���Ƿ�ѡ��˰�Ǳ�˰����", vbCritical, "����"
        Exit Sub

    End If

End If

Set VBExcel = CreateObject("excel.application")
VBExcel.Visible = False
Set xlBook = VBExcel.Workbooks.Open(txtFilePath.text)
Set xlSheet = xlBook.Worksheets(1)
lColsCnt = xlSheet.Range("A1").CurrentRegion.Columns.count
lRowsCnt = xlSheet.Range("A1").CurrentRegion.Rows.count
If lColsCnt <> fpSDetail.e_MCol Then
    MsgBox "Excel�е�����:" & lColsCnt & "���趨��ģ������:" & fpSDetail.e_MCol & "��һ��" & vbCrLf & "��ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
    GoTo EXITPRO
    Exit Sub

End If

Fps(0).MaxRows = 0
ProgressBar1 = 0
lPartSec = 100 * (1 / lRowsCnt)
gUpID = Get_OracleStr("select PO_ITEM_SEQ.nextval from dual")

For i = 2 To lRowsCnt
    Call updateProgressBar
    ' Call GetWOData(dT, xlSheet, i)
    If GetWOData(dT, xlSheet, i) = False Then
        GoTo EXITUPLOAD

    End If

    If setWOData(dT) = False Then
        GoTo EXITUPLOAD

    End If

    Call showWOData(dT, i)
    If ChkWOData(dT, i) Then
        Call SaveWOData(dT, i)
    Else
        '        If MsgBox("�Ƿ�����ϴ�������?", vbOKCancel, "��ʾ") = vbCancel Then
        '            GoTo EXITUPLOAD
        '
        '        End If
        GoTo EXITUPLOAD

    End If

Next
ProgressBar1 = 100
MousePointer = 0
'Call saveWOData_SO

If ExportExcel(dT) = True Then
    txtPOQTY.text = ""
    txtPo_Price.text = ""
    txtcust_device.text = ""
    Call SentMesToPMC(dT)

End If

EXITUPLOAD:
xlBook.Close
Set xlSheet = Nothing
Set xlBook = Nothing
Set VBExcel = Nothing
Exit Sub
EXITPRO:

On Error Resume Next

MousePointer = 0
If Not VBExcel Is Nothing Then
    xlBook.Close
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing

End If

Exit Sub
ErrHandle:
GoTo EXITPRO

End Sub

Private Sub imptOther()
Dim i           As Integer
Dim strArr()    As String
Dim strFileName As String
Dim sArr()      As ty37PO

If cbCusCode.text = "" Then
    MsgBox "��ѡ��ͻ�����", vbInformation, "��ʾ"
    Exit Sub

End If

Select Case cbUploadType.ListIndex

    Case 3, 4
        If cbCusCode.text <> "37" Then
            MsgBox "�ÿͻ�����: " & cbCusCode.text & " û�п���PO���빦��", vbInformation, "��ʾ"
            Exit Sub

        End If

End Select

If chkMsgAppend.Value = 1 And txtMsg.text = "" Then
    MsgBox "�������ʼ����Ĳ�����Ϣ,������ȡ����ѡ��ѡ��" & vbCrLf & "{����:��ݵ��ŵ���Ϣ�򲹳�����˵��" & vbCrLf & "��ϵ�绰....", vbInformation, "����"
    Exit Sub

End If

txtFilePath.text = ""
com2.Filter = "�����ļ�(*.*)|*.*|Excel�ļ�(*.xls;*.xlsx)|*.xls;*.xlsx"
com2.ShowOpen
If com2.filename = "" Then
    Exit Sub

End If

txtFilePath.text = Replace(com2.filename, Chr(0), ",")
com2.filename = ""
If txtFilePath.text = "" Then
    MsgBox "�ļ���ʧ��", vbInformation, "��ʾ"
    Exit Sub

End If

Fps(0).MaxRows = 0
gUpID = Get_OracleStr("select PO_ITEM_SEQ.nextval from dual")

Select Case cbUploadType.ListIndex

    Case 3  '37һ��PO
        If InStr(txtFilePath.text, ",") > 0 Then
            strArr = Split(Trim(txtFilePath.text), ",")

            For i = 1 To UBound(strArr)
                strFileName = strArr(0) & "\" & strArr(i)
                Call GetData_37PO_1(strFileName)
            Next i

        Else
            strFileName = Trim$(txtFilePath.text)
            Call GetData_37PO_1(strFileName)

        End If

        If list37PO Then
            If ExportExcel_37PO("AS") Then
                SentMesToPMC_37PO ("AS")

            End If

        End If

    Case 4  '37����PO
        If InStr(txtFilePath.text, ",") > 0 Then
            strArr = Split(Trim(txtFilePath.text), ",")

            For i = 1 To UBound(strArr)
                strFileName = strArr(0) & "\" & strArr(i)
                Call GetData_37PO_2(strFileName)
            Next i

        Else
            strFileName = Trim$(txtFilePath.text)
            Call GetData_37PO_2(strFileName)

        End If

        If list37PO Then
            If ExportExcel_37PO("TS") Then
                SentMesToPMC_37PO ("TS")

            End If

        End If

    Case 5  'SP29V

        With Fps(0)
            .MaxCols = 3
            .SetText 1, 0, "WaferID"
            .SetText 2, 0, "BIN2"
            .SetText 3, 0, "״̬"
            .ColWidth(1) = 15
            .ColWidth(2) = 15
            .ColWidth(3) = 15

        End With

        If InStr(txtFilePath.text, ",") > 0 Then
            strArr = Split(Trim(txtFilePath.text), ",")

            For i = 1 To UBound(strArr)
                strFileName = strArr(0) & "\" & strArr(i)
                Call SaveOther_SP29V(strFileName)
            Next i

        Else
            strFileName = Trim$(txtFilePath.text)
            Call SaveOther_SP29V(strFileName)

        End If

End Select

End Sub

Private Function list37PO() As Boolean
Dim rs     As ADODB.Recordset
Dim strSql As String

list37PO = False
strSql = " select distinct '37�ϴ��ɹ�' as �ͻ�����,b.po_num,b.test_mtrl_desc as JOBID,a.lotid as LOTID,to_char(wm_concat(a.wafer_id)) as WAFERID,count(1) as WAFERƬ��,sum(a.failbincount+ a.passbincount) as GROSSDIES,b.mpn_desc as �ͻ�����,b.mpn as PRODUCTION_ORDER from mappingdatatest a " & " inner join customeroitbl_test b on a.filename = to_char(b.id) and a.lotid = b.source_batch_id " & " where a.micronlotid = '" & gUpID & "' group by b.po_num,b.test_mtrl_desc,a.lotid,b.mpn_desc,b.mpn "
Set rs = Get_OracleRs(strSql)
If rs.RecordCount = 0 Then
    MsgBox "û�гɹ��ϴ�", vbInformation, "��ʾ"
    Exit Function

End If

With Fps(0)
    .MaxRows = 0
    Set .DataSource = rs

End With

list37PO = True

End Function

Private Sub exptWO()

On Error GoTo Ert

Dim xlsApp     As Excel.Application
Dim xlsBook    As Excel.Workbook
Dim xlsSheet   As Excel.Worksheet
Dim i          As Long
Dim j          As Long
Dim strFileSeq As String, strPartName As String
Dim rs         As New ADODB.Recordset
Dim strCusCode As String
Dim strCusDev  As String

If gUpID = "" Then
    MsgBox "δ�ϴ�����,�޷�����", vbInformation, "��ʾ"
    Exit Sub

End If

strCusCode = UCase(Trim(cbCusCode.text))
strCusDev = Get_OracleStr("select distinct mpn_desc from customeroitbl_test where wafer_visual_inspect = '" & gUpID & "'")
Set rs.ActiveConnection = OraConnect
rs.Source = "select  row_number() over(ORDER BY  bb.lotid,bb.substrateid) as ���,case bb.substratetype when 'A' then '��˰' else '�Ǳ�˰' end as �Ƿ�˰, bb.customershortname as �ͻ�����, " & _
   "       aa.mpn_desc as �ͻ�����,cc.residual as NPI������Ա, " & _
   "       aa.mtrl_num as ���ڻ���, " & _
   "       aa.po_num as PO_NUM, " & _
   "       bb.lotid as LOT_ID, " & _
   "       bb.wafer_id as WAFER_NO, " & _
   "       bb.substrateid as WAFERID, " & _
   "       bb.passbincount as GOOD_DIES, " & _
   "       bb.failbincount as NG_DIES, " & _
   "       bb.passbincount + bb.failbincount as GROSS_DIES, " & _
   "       bb.productid as �����, " & _
   "       aa.Imager_Customer_Rev as ��������, bb.qtech_created_by as �ϴ���Ա,bb.qtech_created_date as �ϴ�ʱ��,  bb.qtech_lastupdate_by as ������Ա, bb.qtech_lastupdate_date as ����ʱ�� " & _
   " from customeroitbl_test aa " & _
   " left join tbltsvnpiproduct cc on cc.customerptno1 = aa.mpn_desc  and cc.qtechptno = aa.mtrl_num and cc.customershortname = aa.customershortname " & _
   " inner join mappingdatatest bb " & _
   "    on to_char(aa.id) = bb.filename " & _
   "   and aa.wafer_visual_inspect = '" & gUpID & "' " & _
   "   group by  bb.customershortname,aa.mpn_desc,aa.mtrl_num,cc.residual,aa.po_num,bb.lotid,bb.wafer_id,bb.substrateid,bb.passbincount,bb.failbincount,bb.productid,aa.Imager_Customer_Rev ,bb.substratetype,bb.qtech_created_by,bb.qtech_created_date,bb.qtech_lastupdate_by,bb.qtech_lastupdate_date "
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount = 0 Then
    MsgBox "��ѯ����������Ϣ, ��ȷ��", vbCritical, "����"
    Exit Sub

End If

Set xlsApp = CreateObject("Excel.Application")
Set xlsBook = xlsApp.Workbooks.Add
Set xlsSheet = xlsBook.Worksheets(1)

With xlsApp
    .Rows(1).Font.Bold = True

End With

For j = 1 To rs.Fields.count
    xlsSheet.Cells(1, j) = ("" & rs(j - 1).name)
Next
rs.MoveFirst

For i = 2 To rs.RecordCount + 1
    For j = 1 To rs.Fields.count
        xlsSheet.Cells(i, j) = ("" & rs(j - 1).Value)
    Next j

    rs.MoveNext
Next i

rs.Close
Set rs = Nothing
xlsApp.Visible = True
strFileName = "C:\others\WO�ϴ�" & Format(Now, "YYYY-MMDD-HH-MM-SS") & ".xlsx"
xlsBook.SaveAs strFileName
Set xlsApp = Nothing
Exit Sub
Ert:
If Not (xlsApp Is Nothing) Then
    Set xlsApp = Nothing

End If

End Sub

Private Sub exptOther()
Dim strSql As String

Select Case cbUploadType.ListIndex

    Case 3  'һ��PO
        If cbCusCode.text = "37" Then
            strSql = "select ID,PO_NUM as PurchaseOrderNo, MPN as ProductionOrderNo,  CREATED_DATE as PODate,  SHIPPING_MST_260 as Currency,  " & "  SHIP_SITE as ShippingAddress, COUNTRY_OF_ASSEMBLY as Termsofpayment,  PO_ITEM as Item, MPN_DESC as MaterialDescription, SOURCE_MTRL_SLOC as LotNo, JOBNO, " & "  CURRENT_WAFER_QTY as Quantity, DATE_CODE as  DelDate, REF_PO as UnitPrice,  DIE_QTY as NetAmount, " & "  SOURCE_MTRL_NUM as PartNumber, SOURCE_BATCH_ID as WaferLot,COUNTRY_OF_FAB as WaferFAB, IMAGER_CUSTOMER_REV as WaferREV,mtrl_num as BagNo " & "   from customeroitbl_test a  where  customershortname='37' and a.qtech_created_date>to_date('2018-03-26','YYYY-MM-DD') and a.flag='Y' order by id desc "

        End If

    Case 4  '����PO
        If cbCusCode.text = "37" Then
            strSql = "select ID,PO_NUM as PurchaseOrderNo,MPN as ProductionOrderNo,CREATED_DATE as PODate,SHIPPING_MST_260 as Currency,SHIP_SITE as ShippingAddress,COUNTRY_OF_ASSEMBLY as Termsofpayment,  PO_ITEM as Item, MPN_DESC as MaterialDescription,SOURCE_BATCH_ID as LotNo, SOURCE_MTRL_SLOC as WaferLot, " & "  DIE_QTY as Quantity, mtrl_num as BagNo,DATE_CODE as  DateCode, t_price as UnitPrice,  CURRENT_WAFER_QTY as NetAmount, " & "  SOURCE_MTRL_NUM as PartNumber, COUNTRY_OF_FAB as WaferFAB, IMAGER_CUSTOMER_REV as WaferREV " & "   from customeroitbl_test a  where customershortname='37' and a.qtech_created_date>to_date('2018-03-26','YYYY-MM-DD') and a.flag='Y' order by id desc "

        End If

End Select

Call ExporToExcel(strSql)

End Sub

Private Sub saveWOData_SO()
Dim strSql As String, strSql2 As String, strHtDev As String

strSql = "insert into HT_SO_HEAD(SO_NUM,SO_ITEM,CUSTOMER,CUSTOMER_PO,CUSTOMER_PO_ITEM,DEVICE,LOT,WAFER_QTY,GROSS_DIE,FLAG,CREATE_DATE,CREATE_BY,REMARK1) " & "select t1.wafer_visual_inspect,row_number() over(ORDER BY t2.lotid), t1.customershortname,t1.po_num,row_number() over(ORDER BY t2.lotid),t1.mpn_desc,t2.lotid,count(t2.substrateid), sum(t2.passbincount+t2.failbincount),'0',sysdate,'" & gUserName & "',t2.SUBSTRATETYPE from customeroitbl_test t1 " & "inner join mappingdatatest t2 on to_char(t1.id) = t2.filename  " & "where t1.wafer_visual_inspect = '" & gUpID & "' group by t2.lotid,t1.customershortname,t1.wafer_visual_inspect,t1.po_num,t1.mpn_desc,t2.SUBSTRATETYPE "
AddSql (strSql)
strHtDev = Get_OracleStr("select  to_char(wm_concat(distinct b.qtechptno)) from  HT_SO_HEAD a, tbltsvnpiproduct b where a.device = b.customerptno1 and a.so_num = '" & gUpID & "'")
strSql = "update HT_SO_HEAD set HT_DEVICE = '" & strHtDev & "' where so_num = '" & gUpID & "'"
AddSql (strSql)
strSql2 = "insert into erptemp..HT_SO_HEAD SELECT * FROM OPENQUERY(ORACLEDB, 'select *  from  HT_SO_HEAD') X where X.SO_NUM = '" & gUpID & "'"
AddSql2 (strSql2)
strSql = " insert into HT_SO_DETAILED(SO_NUM,SO_ITEM,CUSTOMER,LOT,WAFER_ID,MARKING_CODE,GOOD_DIE,NG_DIE,FLAG,CREATE_DATE,CREATE_BY,REMARK1) " & " select t1.wafer_visual_inspect,t3.so_item,t1.customershortname,t2.lotid,t2.substrateid,t2.productid,t2.passbincount,t2.failbincount, '0', sysdate, '" & gUserName & "', " & " t2.SUBSTRATETYPE from customeroitbl_test t1 inner join mappingdatatest t2 on to_char(t1.id) = t2.filename inner join HT_SO_HEAD t3 on t3.so_num = t1.wafer_visual_inspect " & " and t3.lot = t1.source_batch_id where t1.wafer_visual_inspect = '" & gUpID & "' group by t2.lotid,t1.customershortname, t1.wafer_visual_inspect,t2.SUBSTRATETYPE, t2.passbincount, " & " t2.failbincount,t2.substrateid,t2.productid,t3.so_item "
AddSql (strSql)
strSql2 = "insert into erptemp..HT_SO_DETAILED SELECT * FROM OPENQUERY(ORACLEDB, 'select *  from  HT_SO_DETAILED') X where X.SO_NUM = '" & gUpID & "'"
AddSql2 (strSql2)

End Sub

Private Sub GetData_37PO_1(strFileName As String)
Dim i         As Integer
Dim j         As Integer
Dim strChar   As String
Dim tempVal   As String
Dim semPotemp As SemtechPOHeader
Dim VBExcel   As Excel.Application
Dim xlBook    As Excel.Workbook
Dim xlSheet   As Excel.Worksheet

If (cbUploadType.text = "һ��PO") Then
    If InStr(strFileName, "_PO_AS_") = 0 Then
        MsgBox "��ѡ���ϴ����ļ�����37��һ��PO, ��ȷ���Ƿ�ѡ���ļ�", vbInformation, "��ʾ"
        Exit Sub

    End If

End If

Set VBExcel = CreateObject("excel.application")
VBExcel.Visible = False
Set xlBook = VBExcel.Workbooks.Open(strFileName)
Set xlSheet = xlBook.Worksheets(1)
If xlSheet.Range("A1").CurrentRegion.Columns.count <> 49 Then
    MsgBox "Excel�е��������趨��ģ��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
    GoTo EXITPRO
    Exit Sub

End If

For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
        If j <= 26 Then
            strChar = UCase(Chr(96 + j))
        Else
            strChar = "A" & UCase(Chr(96 + j - 26))

        End If

        tempVal = Replace(Trim(xlSheet.Range(strChar & i).Value), Chr(13) + Chr(10), "")

        Select Case strChar

            Case "D"
                semPotemp.ShippingAddress = tempVal

            Case "I"
                semPotemp.PurchaseOrderNo = tempVal

            Case "J"
                semPotemp.ITEM = CInt(tempVal)

            Case "K"
                semPotemp.YourMaterialNumber = tempVal

            Case "N"
                semPotemp.Quantity = CLng(tempVal)

            Case "O"
                semPotemp.UM = tempVal

            Case "P"    ' һ��JOBID
                semPotemp.JOBID = tempVal

            Case "Q"
                semPotemp.DelDate = tempVal

            Case "R"
                semPotemp.Price = tempVal

            Case "S"
                semPotemp.UnitPrice = semPotemp.Price / tempVal

            Case "T"
                semPotemp.CURRENCY = tempVal

            Case "U"
                semPotemp.NetAmount = CLng(tempVal)

            Case "Z"    ' һ��WaferNO
                semPotemp.WaferNO = tempVal

            Case "AI"
                semPotemp.PartNumber = tempVal

            Case "AK", "AO" ' HT��LOTID
                semPotemp.LOTID = IIf(Len(tempVal) <> 0, tempVal, semPotemp.LOTID)

            Case "AQ"
                semPotemp.WaferFAB = tempVal

            Case "AR"
                semPotemp.WaferREV = tempVal

            Case "AT"
                semPotemp.ProductionOrderNo = tempVal

            Case "AU"
                semPotemp.FabSite = tempVal

            Case "AV"
                semPotemp.AssemblySite = tempVal

            Case "AW"
                semPotemp.TestSite = tempVal

        End Select

    Next j

    '�������
    If Len(semPotemp.PurchaseOrderNo) = 0 Then
        MsgBox "I��PO����Ϊ��", vbInformation, "��ʾ"
        GoTo EXITPRO

    End If

    If Len(semPotemp.LOTID) = 0 Then
        MsgBox "AK��AO��LOTID����Ϊ��", vbInformation, "��ʾ"
        GoTo EXITPRO

    End If

    If Len(semPotemp.JOBID) = 0 Then
        MsgBox "P��JOBID����Ϊ��", vbInformation, "��ʾ"
        GoTo EXITPRO

    End If

    '��������
    semPotemp.QTECH_CREATED_BY = gUserName
    semPotemp.KeyStr = semPotemp.KeyStr = semPotemp.PurchaseOrderNo & "_" & semPotemp.JOBID & "_" & semPotemp.LOTID
    If savePO_PRICE(semPotemp) = False Then
        GoTo EXITPRO

    End If

    Call saveWOData_37PO_1(semPotemp)
Next i

EXITPRO:
xlBook.Close
Set xlSheet = Nothing
Set xlBook = Nothing
Set VBExcel = Nothing
Exit Sub

End Sub

Private Function savePO_PRICE(TEMP As SemtechPOHeader) As Boolean
Dim diestr    As String
Dim diers     As New ADODB.Recordset
Dim postr     As String
Dim postr1    As String
Dim pocheck   As String
Dim pocheck1  As String
Dim rs        As ADODB.Recordset
Dim cust_id   As String
Dim cust_name As String
Dim cust_code As String
Dim PO_ID     As String
Dim po_die    As Double
Dim po_waf    As Integer
Dim PO_UNIT   As String

savePO_PRICE = False
cust_id = "37"
cust_name = "Semtech corporation"
cust_code = "AH"
pocheck = "select * from TSV_MD_POPrice where customershortname = '37' and PO_NUM= '" & TEMP.PurchaseOrderNo & "'  and PT = '" & TEMP.YourMaterialNumber & "' "
Set rs = Get_OracleRs(pocheck)
If rs.RecordCount > 0 Then
    MsgBox "PO: " & TEMP.PurchaseOrderNo & ",�ͻ�����: " & TEMP.YourMaterialNumber & "�Ѿ�����ά����¼,��������" & vbCrLf & "�����޸Ļ��˳�", vbInformation, "��ʾ"
    Exit Function

End If

pocheck1 = "select * from TSV_MD_POPrice_tmp where customershortname = '37' and PO_NUM= '" & TEMP.PurchaseOrderNo & "' and PT = '" & TEMP.YourMaterialNumber & "' "
Set rs = Get_OracleRs(pocheck1)
If rs.RecordCount > 0 Then
    MsgBox "PO: " & TEMP.PurchaseOrderNo & ",�ͻ�����: " & TEMP.YourMaterialNumber & "�Ѿ�������ͬ�Ĵ���˵�ά����¼���޷��ظ�ά��", vbInformation, "��ʾ"
    Exit Function

End If

diestr = " select max(b.passbincount) from mappingdatatest b,customeroitbl_test c where b.lotid = '" & TEMP.LOTID & "' " & " and to_char(c.id) = b.filename and c.po_num is null"
If diers.State = adStateOpen Then diers.Close
diers.Open diestr, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If Not diers.EOF Then  '��ʾ��������
    po_die = Val(diers.Fields(0).Value)
Else

End If

If TEMP.CURRENCY = "USD" Then
    PO_UNIT = "��Ԫ"
Else
    PO_UNIT = "�����"

End If

PO_ID = GetPOPriceID()
postr = " insert into TSV_MD_POPrice (ID, CUSTOMERSHORTNAME,CUSTOMERNAME,PO_NUM,PO_DATE,PO_TYPE,PT,QTY,PRICE,UNIT, " & "  Flag, QTECH_CREATED_BY,QTECH_CREATED_DATE,PeaceQty,CUSTAA, DIE_PRICE) values('" & PO_ID & "','" & cust_id & "', " & "  '" & cust_name & "','" & TEMP.PurchaseOrderNo & "',sysdate,'��������','" & TEMP.YourMaterialNumber & "', '" & TEMP.Quantity & "','" & TEMP.UnitPrice & "',  " & "  '" & PO_UNIT & "','Y', '" & TEMP.QTECH_CREATED_BY & "', sysdate,'" & TEMP.Quantity & "','" & cust_code & "',0 )   "
AddSql (postr)
postr1 = " insert into erptemp .. tblBB_CSRPO values (  '" & cust_id & "' ,'" & TEMP.PurchaseOrderNo & "',10,'',  '" & TEMP.YourMaterialNumber & "' " & " , '" & TEMP.Quantity & "', '" & po_die & "' ,'" & TEMP.UnitPrice & "',0,'" & PO_UNIT & "' ,'',CONVERT(varchar(100), getdate(), 20) , '') "
AddSql2 (postr1)
savePO_PRICE = True

End Function

Private Function savePO_PRICE1(TEMP As SemtechPOHeader, i As Integer) As Boolean
Dim diestr    As String
Dim diers     As New ADODB.Recordset
Dim postr     As String
Dim postr1    As String
Dim pocheck   As String
Dim pocheck1  As String
Dim rs        As New ADODB.Recordset
Dim cust_id   As String
Dim cust_name As String
Dim cust_code As String
Dim PO_ID     As String
Dim po_die    As Double
Dim po_waf    As Integer
Dim PO_UNIT   As String
Dim strrebate As String
Dim rsrebate  As New ADODB.Recordset
Dim waf_price As Integer
Dim DIE_PRICE As Double


savePO_PRICE1 = False
cust_id = "37"
cust_name = "Semtech corporation"
cust_code = "AH"
po_die = TEMP.Quantity * i
DIE_PRICE = TEMP.UnitPrice / TEMP.POPrice
pocheck = "select * from TSV_MD_POPrice where customershortname = '37' and PO_NUM= '" & TEMP.PurchaseOrderNo & "'  and PT = '" & TEMP.YourMaterialNumber & "' "
Set rs = Get_OracleRs(pocheck)
If rs.RecordCount > 0 Then
    MsgBox "PO: " & TEMP.PurchaseOrderNo & ",�ͻ�����: " & TEMP.YourMaterialNumber & "�Ѿ�����ά����¼,��������" & vbCrLf & "�����޸Ļ��˳�", vbInformation, "��ʾ"
    Exit Function

End If

pocheck1 = "select * from TSV_MD_POPrice_tmp where customershortname = '37' and PO_NUM= '" & TEMP.PurchaseOrderNo & "' and PT = '" & TEMP.YourMaterialNumber & "' "
Set rs = Get_OracleRs(pocheck1)
If rs.RecordCount > 0 Then
    MsgBox "PO: " & TEMP.PurchaseOrderNo & ",�ͻ�����: " & TEMP.YourMaterialNumber & "�Ѿ�������ͬ�Ĵ���˵�ά����¼���޷��ظ�ά��", vbInformation, "��ʾ"
    Exit Function

End If

diestr = " select max(a.unit_price) from customeroitbl_test a  left join tsv_md_poprice aa on aa.po_num = a.po_num  where a.test_mtrl_desc = '" & TEMP.JOBID & "' "
If diers.State = adStateOpen Then diers.Close
diers.Open diestr, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If Not diers.EOF Then  '��ʾ��������
    waf_price = Val(diers.Fields(0).Value)
Else
    MsgBox " һ��PO�޼۸� "
    Exit Function

End If

If TEMP.CURRENCY = "USD" Then
    PO_UNIT = "��Ԫ"
Else
    PO_UNIT = "�����"

End If

PO_ID = GetPOPriceID()
strrebate = " SELECT a.cust ,a.rebate_waf,a.rebate_die FROM erptemp..ht_cust_rebate a WHERE a.cust = '" & cust_id & "'  AND flag = 0"
If rsrebate.State = adStateOpen Then rsrebate.Close
rsrebate.Open strrebate, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If Not rsrebate.EOF Then
    waf_price = waf_price * Val(rsrebate.Fields(1).Value) / 100
    DIE_PRICE = DIE_PRICE * Val(rsrebate.Fields(2).Value) / 100

End If

postr = " insert into TSV_MD_POPrice (ID, CUSTOMERSHORTNAME,CUSTOMERNAME,PO_NUM,PO_DATE,PO_TYPE,PT,QTY,PRICE,UNIT, " & "  Flag, QTECH_CREATED_BY,QTECH_CREATED_DATE,PeaceQty,CUSTAA, DIE_PRICE) values('" & PO_ID & "','" & cust_id & "', " & "  '" & cust_name & "','" & TEMP.PurchaseOrderNo & "',sysdate,'��������','" & TEMP.YourMaterialNumber & "', '" & po_die & "','" & waf_price & "',  " & "  '" & PO_UNIT & "','Y', '" & TEMP.QTECH_CREATED_BY & "', sysdate,'" & TEMP.Quantity & "','" & cust_code & "','" & DIE_PRICE & "')   "
AddSql (postr)
postr1 = " insert into erptemp .. tblBB_CSRPO values (  '" & cust_id & "' ,'" & TEMP.PurchaseOrderNo & "',10,'',  '" & TEMP.YourMaterialNumber & "' " & " , '" & i & "', '" & po_die & "' ,'" & waf_price & "','" & DIE_PRICE & "','" & PO_UNIT & "' ,'',CONVERT(varchar(100), getdate(), 20) , '') "
AddSql2 (postr1)
savePO_PRICE1 = True

End Function

Private Sub saveWOData_37PO_1(TEMP As SemtechPOHeader)

On Error GoTo ERRON

INIadoCon.BeginTrans
Cnn.BeginTrans
Dim strSql     As String
Dim strsql1    As String, strSql2 As String
Dim strArray() As String
Dim i          As Integer
Dim strPPR     As String
Dim strWOPN    As String
Dim strPOPN    As String

If InStr(TEMP.WaferNO, "#") > 0 Then
    If InStr(TEMP.WaferNO, "PPR") > 0 Then
        strPPR = Mid$(TEMP.WaferNO, InStr(TEMP.WaferNO, "PPR"), 10)
    ElseIf InStr(TEMP.WaferNO, "NCMR") > 0 Then
        strPPR = Mid$(TEMP.WaferNO, InStr(TEMP.WaferNO, "NCMR"), 11)
    Else
        strPPR = ""

    End If

    strArray = Split(Trim(Split(TEMP.WaferNO, "#")(1)), ",")

    For i = 0 To UBound(strArray)
        TEMP.waferid = TEMP.LOTID & Right("0" & Trim(strArray(i)), 2)
        If Get_OracleCnt("select * from mappingdatatest where substrateid = '" & TEMP.waferid & "'") = 0 Then
            MsgBox "��ѯ������WaferID:" & TEMP.waferid & vbCrLf & "ȷ���Ƿ���WO�ϴ�", vbCritical, "����"
            INIadoCon.RollbackTrans
            Cnn.RollbackTrans
            Exit Sub

        End If

        ' ����PPR
        If strPPR <> "" Then
            strSql = "select * from ERPBASE..TBLWAREHOUSEDB_INFO a where a.wafer_id = '" & TEMP.waferid & "'"
            If Get_SqlserverCnt(strSql) > 0 Then
                strsql1 = " update pj_ncmr set ncmr =  '" & strPPR & "'  where wafer_id = '" & TEMP.waferid & "' "
                strSql2 = " Update ERPBASE..TBLWAREHOUSEDB_INFO set Comment = '" & strPPR & "' + ';' +  replace(Comment,'" & strPPR & "','')   where wafer_id = '" & TEMP.waferid & "'"
                AddSql (strsql1)
                AddSql2 (strSql2)
                strSql2 = "update ERPBASE..TBLWAREHOUSEDB_INFO set Comment = REPLACE(Comment,';;',';')  where wafer_id = '" & TEMP.waferid & "' "
                AddSql2 (strSql2)
            Else
                strsql1 = "insert into pj_ncmr (lot_id,ncmr,wafer_id,flag ) values ('" & TEMP.LOTID & "' ,'" & strPPR & "' ,'" & TEMP.waferid & "','Y')"
                strSql2 = "insert into ERPBASE..TBLWAREHOUSEDB_INFO ( HTLOTID, Comment,wafer_id ,flag)  values ('" & TEMP.LOTID & "' ,'" & strPPR & "' ,'" & TEMP.waferid & "','Y')"
                AddSql (strsql1)
                AddSql2 (strSql2)

            End If

            strSql = "select mes_dn_pkg.MES_NCMR_37('" & TEMP.waferid & "') from dual"
            AddSql (strSql)

        End If

        '����PO
        strWOPN = Trim(Get_OracleStr("select b.mpn_desc from mappingdatatest a inner join customeroitbl_test  b on a.filename = to_char(b.id) and a.lotid = b.source_batch_id where a.substrateid = '" & TEMP.waferid & "'"))
        strPOPN = Trim(TEMP.YourMaterialNumber)
        If strPOPN <> strWOPN Then
            MsgBox "waferID: " & TEMP.waferid & " һ��PO�Ļ���Ϊ: " & strPOPN & vbCrLf & "WO�Ļ���Ϊ:" & strWOPN & "���߲�һ��,��ȷ���Ƿ�����", vbInformation, "��ʾ!!!"

        End If

        strsql1 = "update mappingdatatest set micronlotid = '" & gUpID & "' where substrateid = '" & TEMP.waferid & "' "
        AddSql (strsql1)
        strsql1 = "update CUSTOMEROITBL_TEST set " & _
           "PO_NUM = '" & TEMP.PurchaseOrderNo & "',PO_ITEM = '" & TEMP.ITEM & "',SOURCE_MTRL_NUM = '" & TEMP.PartNumber & "'," & _
           " MPN = '" & TEMP.ProductionOrderNo & "', MPN_DESC= '" & TEMP.YourMaterialNumber & "',SOURCE_MTRL_SLOC= '" & TEMP.JOBID & "',OFFSHORE_ASM_COMPANY = '" & TEMP.MfgPlant & "'," & _
           "OFFSHORE_TEST_COMPANY = '" & TEMP.ReceivingPlant & "',CURRENT_WAFER_QTY ='" & TEMP.Quantity & "', COUNTRY_OF_FAB =  '" & TEMP.WaferFAB & "'," & _
           "RETICLE_LEVEL_71= '" & TEMP.Version & "',IMAGER_CUSTOMER_REV  = '" & TEMP.WaferREV & "',PACKAGE_TYPE=  '" & TEMP.TypeService & "',BOX_TYPE= '" & TEMP.UM & "'," & _
           "SHIPPING_MST_260= '" & TEMP.CURRENCY & "', SHIPPING_MST_LEVEL = '" & TEMP.FreightCarrier & "',SHIP_COMMENT = '" & TEMP.TermsDelivery & "',unit_price = '" & TEMP.UnitPrice & "'," & _
           "COUNTRY_OF_ASSEMBLY= '" & TEMP.TermsPayment & "',DATE_CODE = '" & TEMP.DelDate & "',SHIP_SITE = '" & TEMP.ShippingAddress & "',CUSTOM_PART_NO =  '" & TEMP.KeyStr & "'," & _
           "FLAG = 'Y',QTECH_CREATED_BY  = '" & TEMP.QTECH_CREATED_BY & "',QTECH_CREATED_DATE = sysdate,CustomershortName= '37',test_mtrl_desc= '" & TEMP.JOBID & "',MICRON_MATERIAL = '" & TEMP.FabSite & "',SPECIAL_PROCESS_LOT='" & TEMP.AssemblySite & "',WAFER_VISUAL_INSPECT='" & TEMP.TestSite & "' " & _
           "where id in (select c.filename from mappingDataTest c where c.substrateid = '" & TEMP.waferid & "') and ( po_num is null or po_num = '' )"
        strSql2 = "update [ERPBASE].[dbo].[tblCustomerOI] set " & _
           "PO_NUM = '" & TEMP.PurchaseOrderNo & "',PO_ITEM = '" & TEMP.ITEM & "',SOURCE_MTRL_NUM = '" & TEMP.PartNumber & "'," & _
           " MPN = '" & TEMP.ProductionOrderNo & "', MPN_DESC= '" & TEMP.YourMaterialNumber & "',SOURCE_MTRL_SLOC= '" & TEMP.JOBID & "',OFFSHORE_ASM_COMPANY = '" & TEMP.MfgPlant & "'," & _
           "OFFSHORE_TEST_COMPANY = '" & TEMP.ReceivingPlant & "',CURRENT_WAFER_QTY ='" & TEMP.Quantity & "',COUNTRY_OF_FAB =  '" & TEMP.WaferFAB & "'," & _
           "RETICLE_LEVEL_71= '" & TEMP.Version & "',IMAGER_CUSTOMER_REV  = '" & TEMP.WaferREV & "',PACKAGE_TYPE=  '" & TEMP.TypeService & "',BOX_TYPE= '" & TEMP.UM & "'," & _
           "SHIPPING_MST_260= '" & TEMP.CURRENCY & "', SHIPPING_MST_LEVEL = '" & TEMP.FreightCarrier & "',SHIP_COMMENT = '" & TEMP.TermsDelivery & "',unit_price = '" & TEMP.UnitPrice & "'," & _
           "COUNTRY_OF_ASSEMBLY= '" & TEMP.TermsPayment & "',DATE_CODE = '" & TEMP.DelDate & "',SHIP_SITE = '" & TEMP.ShippingAddress & "',CUSTOM_PART_NO =  '" & TEMP.KeyStr & "'," & _
           "FLAG = 'Y',QTECH_CREATED_BY = '" & TEMP.QTECH_CREATED_BY & "',QTECH_CREATED_DATE = getdate(),CustomershortName= '37',test_mtrl_desc= '" & TEMP.JOBID & "',MICRON_MATERIAL = '" & TEMP.FabSite & "',SPECIAL_PROCESS_LOT='" & TEMP.AssemblySite & "',WAFER_VISUAL_INSPECT='" & TEMP.TestSite & "' " & _
           "where id in (select c.filename from [ERPBASE].[dbo].[tblmappingData] c where c.substrateid = '" & TEMP.waferid & "') and (PO_NUM is null or PO_NUM = '') "
        If AddSql(strsql1) = 0 Or AddSql2(strSql2) = 0 Then
            MsgBox "WaferID:" & TEMP.waferid & "һ��PO�ϴ�ʧ��" & vbCrLf & "WOδ�ϴ�;���߸�Wafer��һ��PO�Ѿ�����,�����ظ�����", vbCritical, "ʧ��!!!"
            GoTo ERRON

        End If

    Next i

ElseIf TEMP.WaferNO = "" Then
    strsql1 = "update mappingdatatest set micronlotid = '" & gUpID & "'  where lotid = '" & TEMP.LOTID & "' "
    AddSql (strsql1)
    strWOPN = Trim(Get_OracleStr("select distinct mpn_desc from  customeroitbl_test  where source_batch_id = '" & TEMP.LOTID & "' and po_num is null"))
    strPOPN = Trim(TEMP.YourMaterialNumber)
    If strPOPN <> strWOPN Then
        MsgBox "LOTID: " & TEMP.LOTID & " һ��PO�Ļ���Ϊ: " & strPOPN & vbCrLf & "WO�Ļ���Ϊ:" & strWOPN & "���߲�һ��,��ȷ���Ƿ�����", vbInformation, "��ʾ!!!"

    End If

    strsql1 = "update CUSTOMEROITBL_TEST set " & _
       "PO_NUM = '" & TEMP.PurchaseOrderNo & "',PO_ITEM = '" & TEMP.ITEM & "',SOURCE_MTRL_NUM = '" & TEMP.PartNumber & "'," & _
       " MPN = '" & TEMP.ProductionOrderNo & "', MPN_DESC= '" & TEMP.YourMaterialNumber & "',SOURCE_MTRL_SLOC= '" & TEMP.JOBID & "',OFFSHORE_ASM_COMPANY = '" & TEMP.MfgPlant & "'," & _
       "OFFSHORE_TEST_COMPANY = '" & TEMP.ReceivingPlant & "',CURRENT_WAFER_QTY ='" & TEMP.Quantity & "',COUNTRY_OF_FAB =  '" & TEMP.WaferFAB & "'," & _
       "RETICLE_LEVEL_71= '" & TEMP.Version & "',IMAGER_CUSTOMER_REV  = '" & TEMP.WaferREV & "',PACKAGE_TYPE=  '" & TEMP.TypeService & "',BOX_TYPE= '" & TEMP.UM & "'," & _
       "SHIPPING_MST_260= '" & TEMP.CURRENCY & "', SHIPPING_MST_LEVEL = '" & TEMP.FreightCarrier & "',SHIP_COMMENT = '" & TEMP.TermsDelivery & "',unit_price = '" & TEMP.UnitPrice & "'," & _
       "COUNTRY_OF_ASSEMBLY= '" & TEMP.TermsPayment & "',DATE_CODE = '" & TEMP.DelDate & "',SHIP_SITE = '" & TEMP.ShippingAddress & "',CUSTOM_PART_NO =  '" & TEMP.KeyStr & "'," & _
       "FLAG = 'Y',QTECH_CREATED_BY = '" & TEMP.QTECH_CREATED_BY & "',QTECH_CREATED_DATE = sysdate,CustomershortName= '37',test_mtrl_desc= '" & TEMP.JOBID & "'" & _
       "where source_batch_id = '" & TEMP.LOTID & "' and po_num is null"
    strSql2 = "update [ERPBASE].[dbo].[tblCustomerOI] set " & _
       "PO_NUM = '" & TEMP.PurchaseOrderNo & "',PO_ITEM = '" & TEMP.ITEM & "',SOURCE_MTRL_NUM = '" & TEMP.PartNumber & "'," & _
       " MPN = '" & TEMP.ProductionOrderNo & "', MPN_DESC= '" & TEMP.YourMaterialNumber & "',SOURCE_MTRL_SLOC= '" & TEMP.JOBID & "',OFFSHORE_ASM_COMPANY = '" & TEMP.MfgPlant & "'," & _
       "OFFSHORE_TEST_COMPANY = '" & TEMP.ReceivingPlant & "',CURRENT_WAFER_QTY ='" & TEMP.Quantity & "',COUNTRY_OF_FAB =  '" & TEMP.WaferFAB & "'," & _
       "RETICLE_LEVEL_71= '" & TEMP.Version & "',IMAGER_CUSTOMER_REV  = '" & TEMP.WaferREV & "',PACKAGE_TYPE=  '" & TEMP.TypeService & "',BOX_TYPE= '" & TEMP.UM & "'," & _
       "SHIPPING_MST_260= '" & TEMP.CURRENCY & "', SHIPPING_MST_LEVEL = '" & TEMP.FreightCarrier & "',SHIP_COMMENT = '" & TEMP.TermsDelivery & "',unit_price = '" & TEMP.UnitPrice & "'," & _
       "COUNTRY_OF_ASSEMBLY= '" & TEMP.TermsPayment & "',DATE_CODE = '" & TEMP.DelDate & "',SHIP_SITE = '" & TEMP.ShippingAddress & "',CUSTOM_PART_NO =  '" & TEMP.KeyStr & "'," & _
       "FLAG = 'Y',QTECH_CREATED_BY = '" & TEMP.QTECH_CREATED_BY & "',QTECH_CREATED_DATE = getdate(),CustomershortName= '37',test_mtrl_desc= '" & TEMP.JOBID & "'" & _
       "where source_batch_id = '" & TEMP.LOTID & "' and ( po_num is null or PO_NUM = '') "
    If AddSql(strsql1) = 0 Or AddSql2(strSql2) = 0 Then
        MsgBox "LOTID:" & TEMP.LOTID & "һ��PO�ϴ�ʧ��, ���߸�LOT��һ��PO�Ѿ�����", vbCritical, "ʧ��!!!"
        GoTo ERRON

    End If

Else
    MsgBox "Z�и�ʽ����", vbCritical, "����"
    GoTo ERRON

End If

INIadoCon.CommitTrans
Cnn.CommitTrans
Exit Sub
ERRON:
INIadoCon.RollbackTrans
Cnn.RollbackTrans

End Sub

Private Sub GetData_37PO_2(strFileName As String)
Dim i       As Integer
Dim j       As Integer
Dim strChar As String
Dim tempVal As String
Dim VBExcel As Excel.Application
Dim xlBook  As Excel.Workbook
Dim xlSheet As Excel.Worksheet

If (cbUploadType.text = "����PO") Then
    If InStr(strFileName, "_PO_TS_") = 0 Then
        MsgBox "��ѡ���ϴ����ļ�����37�Ķ���PO, ��ȷ���Ƿ�ѡ���ļ�", vbInformation, "��ʾ"
        Exit Sub

    End If

End If

Set VBExcel = CreateObject("excel.application")
VBExcel.Visible = False
Set xlBook = VBExcel.Workbooks.Open(strFileName)
Set xlSheet = xlBook.Worksheets(1)
If xlSheet.Range("A1").CurrentRegion.Columns.count <> 49 Then
    MsgBox "Excel�е��������趨��ģ��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
    GoTo EXITPRO
    Exit Sub

End If

For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
        If j <= 26 Then
            strChar = UCase(Chr(96 + j))
        Else
            strChar = "A" & UCase(Chr(96 + j - 26))

        End If

        Dim semPotemp As SemtechPOHeader

        tempVal = Replace(Trim(xlSheet.Range(strChar & i).Value), Chr(13) + Chr(10), "")

        Select Case strChar

            Case "A"
                semPotemp.DATE = tempVal

            Case "D"
                semPotemp.MfgPlant = tempVal

            Case "E"
                semPotemp.MfgPlant = semPotemp.MfgPlant & "-" & tempVal

            Case "H"
                semPotemp.TypeService = tempVal

            Case "I"
                semPotemp.PurchaseOrderNo = tempVal

            Case "J"
                semPotemp.ITEM = CInt(tempVal)

            Case "K"
                semPotemp.MaterialDes = tempVal

            Case "L"
                semPotemp.YourMaterialNumber = tempVal

            Case "P"    ' ����JOBID
                semPotemp.JobID_2 = tempVal

            Case "Q"
                semPotemp.DelDate = tempVal

            Case "R"
                semPotemp.UnitPrice = tempVal

            Case "S"
                semPotemp.POPrice = tempVal

            Case "T"
                semPotemp.CURRENCY = tempVal
            
            Case "Z"
                semPotemp.PPR = tempVal

            Case "U"
                semPotemp.NetAmount = CLng(tempVal)

            Case "W"
                semPotemp.TermsPayment = tempVal

            Case "AA"
                semPotemp.ItemLineText = tempVal

            Case "AH"
                semPotemp.Plant = tempVal

            Case "AI"
                semPotemp.PartNumber = tempVal

            Case "AJ"
                semPotemp.Quantity = CLng(tempVal)

            Case "AL"   ' һ��JOBID
                semPotemp.JOBID = tempVal

            Case "AN"   ' ����WaferNO
                semPotemp.WaferNO = Replace(tempVal, "+", "")

            Case "AO", "AS" ' HT_LOTID
                semPotemp.LOTID = IIf(Len(tempVal) <> 0, tempVal, semPotemp.LOTID)

            Case "AT"
                semPotemp.ProductionOrderNo = tempVal

            Case "AU"
                semPotemp.FabSite = tempVal

            Case "AV"
                semPotemp.AssemblySite = tempVal

            Case "AW"
                semPotemp.TestSite = tempVal

        End Select

    Next j

    '�������
    If Len(semPotemp.PurchaseOrderNo) = 0 Then
        MsgBox "I��PO����Ϊ��", vbInformation, "��ʾ"
        GoTo EXITPRO

    End If

    If Len(semPotemp.LOTID) = 0 And Len(semPotemp.JOBID) = 0 Then
        MsgBox "AO��AS��LOTID����Ϊ�ջ�AL��һ�ε�JOBID����Ϊ��", vbInformation, "��ʾ"
        GoTo EXITPRO

    End If

    If Len(semPotemp.JobID_2) = 0 Then
        MsgBox "P�ж���JOBID����Ϊ��", vbInformation, "��ʾ"
        GoTo EXITPRO

    End If

    If Len(semPotemp.WaferNO) = 0 Then
        MsgBox "AN�ж���WAFER_ID����Ϊ��", vbInformation, "��ʾ"
        GoTo EXITPRO

    End If

    If Len(semPotemp.JOBID) = 0 Then
        MsgBox "AL��һ��JOBID����Ϊ��", vbInformation, "��ʾ"
        GoTo EXITPRO
    Else
        If Left$(semPotemp.WaferNO, 1) = "0" Then
            semPotemp.WaferNO = Right$(semPotemp.WaferNO, 1)

        End If

    End If

    '��������
    If semPotemp.LOTID = "" Then
        semPotemp.LOTID = GetLot(semPotemp.JOBID)

    End If

    semPotemp.id = GetMaxID()
    semPotemp.waferid = semPotemp.LOTID & Right("0" & semPotemp.WaferNO, 2)
    semPotemp.QTECH_CREATED_BY = gUserName
    semPotemp.fab_conv_id = Getcustpart(semPotemp.waferid)
    semPotemp.BondOrNot = Get37Bonded(semPotemp.waferid)
    If semPotemp.fab_conv_id = "" Then
        MsgBox "һ��POfab_conv_id����Ϊ��", vbCritical, "����"
        GoTo EXITPRO

    End If

    If semPotemp.BondOrNot = "" Then
        MsgBox "һ��PO��˰�Ǳ�˰��ʶ����Ϊ��", vbCritical, "����"
        GoTo EXITPRO

    End If

    semPotemp.KeyStr = semPotemp.PurchaseOrderNo & "_" & semPotemp.JobID_2 & "_" & semPotemp.waferid
    semPotemp.WaferID_2 = Get_OracleStr("select max(substrateid) as substrateid from mappingdatatest where wafer_id in ('" & semPotemp.WaferNO & "', '0' || '" & semPotemp.WaferNO & "') and lotid = '" & semPotemp.LOTID & "'")
    If Get_OracleCnt(" select * from customeroitbl_test where id in (select filename from mappingdatatest where substrateid = '" & semPotemp.WaferID_2 & "') and test_mtrl_desc <> source_mtrl_sloc and test_mtrl_desc is not null") Then
        If Get_OracleCnt("select * from ib_waferlist where waferid = '" & semPotemp.WaferID_2 & "'") = 0 Then
            MsgBox "ϵͳ�д���WaferID:" & semPotemp.WaferID_2 & "  ����POδ������,�����ٴ�ά������PO", vbCritical, "����"
            GoTo EXITPRO

        End If

    End If

    semPotemp.WaferID_2 = Get_OracleStr("select max(substrateid) || '+' as substrateid from mappingdatatest where wafer_id in ('" & semPotemp.WaferNO & "', '0' || '" & semPotemp.WaferNO & "') and lotid = '" & semPotemp.LOTID & "'")
    Call saveWOData_37PO_2(semPotemp)
Next i

If savePO_PRICE1(semPotemp, i) = False Then
    GoTo EXITPRO

End If

EXITPRO:
xlBook.Close
Set xlSheet = Nothing
Set xlBook = Nothing
Set VBExcel = Nothing

End Sub

Private Sub saveWOData_37PO_2(TEMP As SemtechPOHeader)

On Error GoTo ERRON

Dim strsql1 As String
Dim strSql2 As String
Dim strSql3 As String
Dim strSql4 As String
Dim strPPR As String
Dim strSql As String

Cnn.BeginTrans
INIadoCon.BeginTrans

If InStr(TEMP.PPR, "PPR") > 0 Then
    strPPR = Mid$(TEMP.PPR, InStr(TEMP.PPR, "PPR"), 10)
ElseIf InStr(TEMP.PPR, "NCMR") > 0 Then
    strPPR = Mid$(TEMP.PPR, InStr(TEMP.PPR, "NCMR"), 11)
Else
    strPPR = ""

End If

'����PO
strsql1 = "insert into mappingdatatest(substrateid,lotid,flag,qtech_created_date,wafer_id,customershortname,filename,passbincount,failbincount, QTECH_CREATED_BY,micronlotid)" & " values( '" & TEMP.WaferID_2 & "','" & TEMP.LOTID & "','Y',sysdate,'" & TEMP.WaferNO & "','37','" & TEMP.id & "','" & TEMP.Quantity & "','0', '" & TEMP.QTECH_CREATED_BY & "','" & gUpID & "')"
strSql2 = "insert into CUSTOMEROITBL_TEST (ID ,PO_NUM ,PO_ITEM ,SOURCE_BATCH_ID ,SOURCE_MTRL_NUM, mtrl_num," & _
   " MPN ,MPN_DESC ,SOURCE_MTRL_SLOC,OFFSHORE_ASM_COMPANY,OFFSHORE_TEST_COMPANY," & _
   " CURRENT_WAFER_QTY,DIE_QTY ,COUNTRY_OF_FAB,RETICLE_LEVEL_72 ,IMAGER_CUSTOMER_REV ," & _
   " PACKAGE_TYPE , BOX_TYPE,SHIPPING_MST_260 ,SHIPPING_MST_LEVEL ,SHIP_COMMENT," & _
   " CREATED_DATE  ,REF_PO  ,COUNTRY_OF_ASSEMBLY  ,DATE_CODE  ,  SHIP_SITE   ," & _
   " CUSTOM_PART_NO , FLAG, QTECH_CREATED_BY, QTECH_CREATED_DATE, CustomershortName ,BATCH_COMMENT_TEST,t_price,MTRL_DESC,test_mtrl_desc, fab_conv_id,reticle_level_71, MICRON_MATERIAL,SPECIAL_PROCESS_LOT,WAFER_VISUAL_INSPECT,JOBNO) values(" & _
   " '" & TEMP.id & "', '" & TEMP.PurchaseOrderNo & "','" & TEMP.ITEM & "', '" & TEMP.LOTID & "', '" & TEMP.PartNumber & "', '" & TEMP.BagNo & "'," & _
   " '" & TEMP.ProductionOrderNo & "', '" & TEMP.YourMaterialNumber & "','" & TEMP.JOBID & "', '" & TEMP.MfgPlant & "', '" & TEMP.ReceivingPlant & "', " & _
   " '" & TEMP.NetAmount & "', '" & TEMP.Quantity & "', '" & TEMP.WaferFAB & "', '" & TEMP.Version & "', '" & TEMP.WaferREV & "', " & _
   " '" & TEMP.TypeService & "', '" & TEMP.UM & "','" & TEMP.CURRENCY & "', '" & TEMP.FreightCarrier & "', '" & TEMP.TermsDelivery & "', " & _
   " '" & TEMP.DATE & "', '" & TEMP.UnitPrice & "','" & TEMP.TermsPayment & "', '" & TEMP.DATECODE & "', '" & TEMP.ShippingAddress & "', " & _
   " '" & TEMP.KeyStr & "', 'Y','" & TEMP.QTECH_CREATED_BY & "', sysdate, '37','" & TEMP.DelDate & "'," & TEMP.POPrice & ",'" & TEMP.MaterialDes & "','" & TEMP.JobID_2 & "', '" & TEMP.fab_conv_id & "' ,'" & TEMP.ItemLineText & "', '" & TEMP.FabSite & "', '" & TEMP.AssemblySite & "','" & TEMP.TestSite & "','" & TEMP.BondOrNot & "') "
strSql3 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,lotid,flag,qtech_created_date,wafer_id,customershortname ,filename,passbincount,failbincount,QTECH_CREATED_BY)" & " values( '" & TEMP.WaferID_2 & "','" & TEMP.LOTID & "','Y',getdate(),'" & TEMP.WaferNO & "','37','" & TEMP.id & "','" & TEMP.Quantity & "','0', '" & TEMP.QTECH_CREATED_BY & "')"
strSql4 = "insert into [ERPBASE].[dbo].[tblCustomerOI] (ID ,PO_NUM ,PO_ITEM ,SOURCE_BATCH_ID ,SOURCE_MTRL_NUM, mtrl_num," & _
   " MPN ,MPN_DESC ,SOURCE_MTRL_SLOC,OFFSHORE_ASM_COMPANY,OFFSHORE_TEST_COMPANY," & _
   " CURRENT_WAFER_QTY,DIE_QTY ,COUNTRY_OF_FAB,RETICLE_LEVEL_72 ,IMAGER_CUSTOMER_REV ," & _
   " PACKAGE_TYPE , BOX_TYPE,SHIPPING_MST_260 ,SHIPPING_MST_LEVEL ,SHIP_COMMENT," & _
   " CREATED_DATE  ,REF_PO  ,COUNTRY_OF_ASSEMBLY  ,DATE_CODE  ,  SHIP_SITE   ," & _
   " CUSTOM_PART_NO , FLAG, QTECH_CREATED_BY, QTECH_CREATED_DATE, CustomershortName ,BATCH_COMMENT_TEST,t_price,MTRL_DESC,test_mtrl_desc, fab_conv_id,reticle_level_71,MICRON_MATERIAL,SPECIAL_PROCESS_LOT,WAFER_VISUAL_INSPECT,JOBNO) values(" & _
   " '" & TEMP.id & "', '" & TEMP.PurchaseOrderNo & "','" & TEMP.ITEM & "', '" & TEMP.LOTID & "', '" & TEMP.PartNumber & "', '" & TEMP.BagNo & "'," & _
   " '" & TEMP.ProductionOrderNo & "', '" & TEMP.YourMaterialNumber & "','" & TEMP.JOBID & "', '" & TEMP.MfgPlant & "', '" & TEMP.ReceivingPlant & "', " & _
   " '" & TEMP.NetAmount & "', '" & TEMP.Quantity & "', '" & TEMP.WaferFAB & "', '" & TEMP.Version & "', '" & TEMP.WaferREV & "', " & _
   " '" & TEMP.TypeService & "', '" & TEMP.UM & "','" & TEMP.CURRENCY & "', '" & TEMP.FreightCarrier & "', '" & TEMP.TermsDelivery & "', " & _
   " '" & TEMP.DATE & "', '" & TEMP.UnitPrice & "','" & TEMP.TermsPayment & "', '" & TEMP.DATECODE & "', '" & TEMP.ShippingAddress & "', " & _
   " '" & TEMP.KeyStr & "', 'Y','" & TEMP.QTECH_CREATED_BY & "',  getdate(), '37','" & TEMP.DelDate & "'," & TEMP.POPrice & ",'" & TEMP.MaterialDes & "','" & TEMP.JobID_2 & "', '" & TEMP.fab_conv_id & "' ,'" & TEMP.ItemLineText & "','" & TEMP.FabSite & "', '" & TEMP.AssemblySite & "','" & TEMP.TestSite & "','" & TEMP.BondOrNot & "') "

If AddSql(strsql1) = 0 Or AddSql(strSql2) = 0 Or AddSql2(strSql3) = 0 Or AddSql2(strSql4) = 0 Then
    MsgBox "WaferID:" & TEMP.WaferID_2 & "����PO�ϴ�ʧ��", vbCritical, "ʧ��!!!"
    GoTo ERRON

End If

' ����PPR
Dim strWaferID As String

strWaferID = Replace$(TEMP.WaferID_2, "+", "")

If strPPR <> "" Then
    strSql = "select * from ERPBASE..TBLWAREHOUSEDB_INFO a where a.wafer_id = '" & strWaferID & "'"

    If Get_SqlserverCnt(strSql) > 0 Then
        strsql1 = " update pj_ncmr set ncmr =  '" & strPPR & "'  where wafer_id = '" & strWaferID & "' "
        strSql2 = " Update ERPBASE..TBLWAREHOUSEDB_INFO set Comment = '" & strPPR & "' + ';' +  replace(Comment,'" & strPPR & "','')   where wafer_id = '" & strWaferID & "'"
        AddSql (strsql1)
        AddSql2 (strSql2)
        strSql2 = "update ERPBASE..TBLWAREHOUSEDB_INFO set Comment = REPLACE(Comment,';;',';')  where wafer_id = '" & strWaferID & "' "
        AddSql2 (strSql2)
    Else
        strsql1 = "insert into pj_ncmr (lot_id,ncmr,wafer_id,flag ) values ('" & TEMP.LOTID & "' ,'" & strPPR & "' ,'" & strWaferID & "','Y')"
        strSql2 = "insert into ERPBASE..TBLWAREHOUSEDB_INFO ( HTLOTID, Comment,wafer_id ,flag)  values ('" & TEMP.LOTID & "' ,'" & strPPR & "' ,'" & strWaferID & "','Y')"
        AddSql (strsql1)
        AddSql2 (strSql2)

    End If

    strSql = "select mes_dn_pkg.MES_NCMR_37('" & strWaferID & "') from dual"
    AddSql (strSql)
    
End If

Cnn.CommitTrans
INIadoCon.CommitTrans
Exit Sub
ERRON:
Cnn.RollbackTrans
INIadoCon.RollbackTrans

End Sub

Private Sub SaveOther_SP29V(dirtemp As String)
Dim strCode    As String
Dim i          As Integer
Dim cnt        As Integer
Dim strTmp     As String
Dim strWaferID As String

cnt = 0
Open dirtemp For Input As #1

Do While Not EOF(1)
    Line Input #1, strCode
    If InStr(strCode, "=") = 0 And strCode <> "" Then

        For i = 1 To Len(strCode)
            strTmp = Mid(strCode, i, 1)
            If strTmp = "2" Then
                cnt = cnt + 1

            End If

        Next i

    End If

Loop
If cnt > 0 Then
    strWaferID = Replace(Replace(Split(dirtemp, "\")(UBound(Split(dirtemp, "\"))), ".txt", ""), "-", "")
    If AddSql("update mappingdatatest set remark = '" & cnt & "' where substrateid = '" & strWaferID & "'") > 0 Then

        With Fps(0)
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1
            .text = strWaferID
            .Col = 2
            .text = cnt
            .Col = 3
            .text = "�ѳɹ�����"

        End With

    End If

    AddSql2 ("update [ERPBASE].[dbo].[tblmappingData] set remark = '" & cnt & "' where substrateid = '" & strWaferID & "'")

End If

Close #1

End Sub

Private Sub SaveMappingData(strFileName As String, strCusCode As String)

End Sub

Private Function GetWOData(ByRef dT As tyWO, _
                           xlSheet As Excel.Worksheet, _
                           i As Integer) As Boolean
'Private Sub GetWOData(ByRef dT As tyWO, xlSheet As Excel.Worksheet, i As Integer)
Dim strSql      As String
Dim strSqlfab   As String
Dim rs          As New ADODB.Recordset
Dim rsfab       As New ADODB.Recordset
Dim lRevID      As Long
Dim lID         As Long

Dim strdevice_prcie       As String
Dim rsdevice    As New ADODB.Recordset
Dim rsdevice1   As New ADODB.Recordset
Dim rsdevice_prcie    As New ADODB.Recordset
Dim w_price As String
Dim d_price As String


Dim strprdevice As String
Dim strdevice   As String
Dim price_w     As String
Dim price_d     As String
Dim price_unit  As String
Dim pocheck     As String
Dim pocheck1    As String
Dim cust_name   As String
Dim PO_ID       As String
Dim postr       As String
Dim postr1      As String







GetWOData = True
dT.TAX_TYPE = IIf(cbTaxType.ListIndex = 0, "A", "B")
dT.CUSTOMER_CODE = UCase(Trim(cbCusCode.text))
dT.ITEM = Trim("" & Replace(Replace(xlSheet.Range("A" & i), Chr(10), ""), Chr(13), ""))
dT.po_no = Trim("" & Replace(Replace(xlSheet.Range("B" & i), Chr(10), ""), Chr(13), ""))
dT.SUPPLIER = Trim("" & Replace(Replace(xlSheet.Range("C" & i), Chr(10), ""), Chr(13), ""))
dT.SHIP_TO = Trim("" & Replace(Replace(xlSheet.Range("D" & i), Chr(10), ""), Chr(13), ""))
dT.Fab_Device = Trim("" & Replace(Replace(xlSheet.Range("E" & i), Chr(10), ""), Chr(13), ""))
dT.Customer_Device = Trim("" & Replace(Replace(xlSheet.Range("F" & i), Chr(10), ""), Chr(13), ""))
dT.WAFER_VERSION = Trim("" & Replace(Replace(Replace(xlSheet.Range("G" & i), Chr(10), ""), Chr(13), ""), "'", ""))
dT.MARKING_CODE = Trim("" & Replace(Replace(xlSheet.Range("H" & i), Chr(10), ""), Chr(13), ""))
dT.WO_DATE = Trim("" & Replace(Replace(xlSheet.Range("I" & i), Chr(10), ""), Chr(13), ""))
dT.Lot_id = Trim("" & Replace(Replace(Replace(xlSheet.Range("J" & i), Chr(10), ""), Chr(13), ""), "+", ""))
dT.wafer_id = Trim("" & Replace(Replace(xlSheet.Range("K" & i), Chr(10), ""), Chr(13), ""))
dT.GOOD_DIES_PCS = CLng(Replace(Replace(xlSheet.Range("L" & i), Chr(10), ""), Chr(13), ""))
dT.GROSS_DIES_PCS = CLng(Replace(Replace(xlSheet.Range("M" & i), Chr(10), ""), Chr(13), ""))
dT.HT_DEVICE = Trim("" & Replace(Replace(xlSheet.Range("N" & i), Chr(10), ""), Chr(13), ""))
dT.REMARK = Trim("" & Replace(Replace(xlSheet.Range("O" & i), Chr(10), ""), Chr(13), ""))
dT.TRADE_TYPE = Trim("" & Replace(Replace(xlSheet.Range("P" & i), Chr(10), ""), Chr(13), ""))
dT.DATA1 = Trim("" & Replace(Replace(xlSheet.Range("Q" & i), Chr(10), ""), Chr(13), ""))
dT.DATA2 = Trim("" & Replace(Replace(xlSheet.Range("R" & i), Chr(10), ""), Chr(13), ""))
dT.DATA3 = Trim("" & Replace(Replace(xlSheet.Range("S" & i), Chr(10), ""), Chr(13), ""))
dT.DATA4 = Trim("" & Replace(Replace(xlSheet.Range("T" & i), Chr(10), ""), Chr(13), ""))
dT.DATA5 = Trim("" & Replace(Replace(xlSheet.Range("U" & i), Chr(10), ""), Chr(13), ""))
price_w = Trim("" & Replace(Replace(xlSheet.Range("V" & i), Chr(10), ""), Chr(13), ""))
price_d = Trim("" & Replace(Replace(xlSheet.Range("W" & i), Chr(10), ""), Chr(13), ""))
'price_unit = Trim("" & Replace(Replace(xlSheet.Range("X" & I), Chr(10), ""), Chr(13), ""))
If Len(dT.wafer_id) = 1 Then
    dT.lot_wafer_id = dT.Lot_id & "0" & dT.wafer_id
ElseIf Len(dT.wafer_id) = 2 Then
    dT.lot_wafer_id = dT.Lot_id & dT.wafer_id
    If Left$(dT.wafer_id, 1) = "0" Then
        dT.wafer_id = Right$(dT.wafer_id, 1)

    End If

Else
    dT.lot_wafer_id = dT.Lot_id & dT.wafer_id

End If

'WO���ݰ汾���洢
lRevID = Get_OracleNo("select nvl(max(REV_ID)+1,1) from TBL_WO_TEMPLATE_DATA_REP where J_LOT_ID = '" & dT.Lot_id & "' and K_WAFER_ID = '" & dT.wafer_id & "' ")
lID = Get_OracleNo("select nvl(max(id)+1,1) from TBL_WO_TEMPLATE_DATA_REP ")
strSql = "insert into TBL_WO_TEMPLATE_DATA_REP(A_ITEM,B_PO_NO,C_SUPPLIER,D_SHIP_TO,E_FAB_DEVICE,F_CUSTOMER_DEVICE,G_WAFER_VERSION,H_MARKING_LOT_ID,I_DATE,J_LOT_ID,K_WAFER_ID " & " ,L_GOOD_DIES,M_TOTAL_DIES,N_HT_PN,O_REMARK,P_REMARK,Q_REMARK,R_REMARK,S_REMARK,T_REMARK,U_REMARK,REV_ID,CREATE_BY,CREATE_DATE,TAX_TYPE,ID) " & " values('" & dT.ITEM & "','" & dT.po_no & "','" & dT.SUPPLIER & "','" & dT.SHIP_TO & "','" & dT.Fab_Device & "','" & dT.Customer_Device & "','" & dT.WAFER_VERSION & "','" & dT.MARKING_CODE & "','" & dT.WO_DATE & "','" & dT.Lot_id & "', " & " '" & dT.wafer_id & "'," & dT.GOOD_DIES_PCS & "," & dT.GROSS_DIES_PCS & ",'" & dT.HT_DEVICE & "','" & dT.REMARK & "','" & dT.TRADE_TYPE & "','" & dT.DATA1 & "','" & dT.DATA2 & "','" & dT.DATA3 & "','" & dT.DATA4 & "','" & dT.DATA5 & "'," & lRevID & ",'" & gUserName & "' || '" & gUserRealName & "',sysdate,'" & dT.TAX_TYPE & "'," & lID & ")     "
AddSql (strSql)
strSql = "SELECT * FROM erptemp..CONFIG a WHERE a.CUSTOMER = '" & UCase(Trim(cbCusCode.text)) & "'  AND a.REMARK1 = 'Y'"
If rs.State = adStateOpen Then rs.Close
rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF Then  '��ʾ��������
    strSqlfab = " select p.customershortname,p.customerptno1,p.customerptno2,count(p.qtechptno2 )  from tbltsvnpiproduct p where p.customershortname =  '" & UCase(Trim(cbCusCode.text)) & "'      " & " and p.customerptno1 = '" & dT.Customer_Device & "'   and  p.customerptno2 = '" & dT.Fab_Device & "'   group by p.customershortname,p.customerptno1,p.customerptno2 "
    If rsfab.State = adStateOpen Then rsfab.Close
    rsfab.Open strSqlfab, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rsfab.EOF Then
        If rsfab.Fields(3).Value <> "1" Then
            MsgBox "�ͻ�����+FAB_DEVICE ������Ψһ��Ʒ�Ϻ�"
            GetWOData = False
            Exit Function

        End If

    Else
        MsgBox "�ͻ�����+FAB_DEVICE ������Ψһ��Ʒ�Ϻ�"
        GetWOData = False
        Exit Function

    End If

End If


If cbCusCode.text <> "37" Then

If Trim(dT.po_no) = "" Then
    MsgBox "WO����PO_NUM,�������ϴ�WO,��ȷ��WO��Ϣ!"
    GetWOData = False
    Exit Function

End If

 strdevice_prcie = "SELECT a.wafer_price,a.die_price FROM erptemp..HT_PRICE_CONTROL A  WHERE a.cust_device  = '" & dT.Customer_Device & "' AND a.cust_id = '" & UCase(Trim(cbCusCode.text)) & "' AND FLAG = 0 "


If rsdevice_prcie.State = adStateOpen Then rsdevice_prcie.Close
rsdevice_prcie.Open strdevice_prcie, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

If Not rsdevice_prcie.EOF Then
 
 w_price = Trim(rsdevice_prcie.Fields(0).Value)
 d_price = Trim(rsdevice_prcie.Fields(1).Value)
 


cust_name = Get_SqlStr("SELECT a.�ͻ����� FROM erpdata..tblXCustomer a WHERE a.�ͻ����� = '" & cbCusCode.text & "'")


strdevice = "  SELECT a.wafer_price,a.die_price,a.currency  FROM erptemp..ht_price_control a ,erptemp..ht_price_config b   WHERE a.cust_id = '" & UCase(Trim(cbCusCode.text)) & "'  " & _
            "   AND a.cust_device =  '" & dT.Customer_Device & "'  AND a.flag = 0  AND  b.cust_id = a.cust_id   AND b.po_price = 'Y'  AND  b.openpo = 'N'   "



If rsdevice1.State = adStateOpen Then rsdevice1.Close
rsdevice1.Open strdevice, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If Not rsdevice1.EOF Then

If UCase(Trim(cbCusCode.text)) = "68" Or UCase(Trim(cbCusCode.text)) = "HK075" Then


If price_w = Trim(rsdevice1.Fields(0).Value) And price_d = Trim(rsdevice1.Fields(1).Value) Then

   If Trim(dT.po_no) <> Trim(txtPo_Price.text) Or Trim(dT.Customer_Device) <> Trim(txtcust_device.text) Then

 pocheck = "select peaceqty, po_type from TSV_MD_POPrice where customershortname = '" & UCase(Trim(cbCusCode.text)) & "'  and PO_NUM= '" & Trim(dT.po_no) & "'  and PT = '" & dT.Customer_Device & "' "
Set rs = Get_OracleRs(pocheck)
If rs.RecordCount = 0 Then
   
   
 PO_ID = GetPOPriceID()



 postr = " insert into TSV_MD_POPrice (ID, CUSTOMERSHORTNAME,CUSTOMERNAME,PO_NUM,PO_DATE,PO_TYPE,PT,QTY,PRICE,UNIT, " & _
         "  Flag, QTECH_CREATED_BY,QTECH_CREATED_DATE,PeaceQty,CUSTAA, DIE_PRICE) values('" & PO_ID & "','" & UCase(Trim(cbCusCode.text)) & "', " & _
         "  '" & cust_name & "','" & Trim(dT.po_no) & "',sysdate,'��������', '" & dT.Customer_Device & "', 99999,'" & price_w & "',  " & _
         "  '" & price_unit & "','Y', '', sysdate,999999,'NA','" & price_d & "' )   "

 AddSql (postr)
 
 
 postr1 = " insert into erptemp .. tblBB_CSRPO values (  '" & UCase(Trim(cbCusCode.text)) & "' ,'" & Trim(dT.po_no) & "',10,'',  '" & dT.Customer_Device & "'  " & _
          " , 99999, 99999 ,'" & price_w & "','" & price_d & "','" & price_unit & "' ,'',CONVERT(varchar(100), getdate(), 20) , '') "
 
AddSql2 (postr1)
   
ElseIf rs.Fields(1).Value = "NRE����" Then
   
    If Trim(txtPo_Price.text) = "" Then
    
    txtPo_Price.text = dT.po_no
    txtPOQTY.text = 1
    
    
    ElseIf Trim(txtPo_Price.text) <> dT.po_no Then
    
     txtPo_Price.text = dT.po_no
     txtPOQTY.text = 1
    Else
 
     txtPOQTY.text = Val(txtPOQTY.text) + 1
     
    End If
 
   If Val(rs.Fields(0).Value) < Val(txtPOQTY.text) Then
    
     MsgBox "WO�������� NERPO" & Trim(dT.po_no) & "����" & dT.Customer_Device & "����!"
      GetWOData = False
     Exit Function
    
   End If
   
   
   
End If


  txtPo_Price.text = dT.po_no
  txtcust_device.text = dT.Customer_Device
    
End If


Else

 MsgBox "WO�ϵ��ۺͲ�Ʒ�۸�һ��,��ȷ�ϼ۸���Ϣ!"
        GetWOData = False
        Exit Function
    
End If

Else
    
    
pocheck = "select peaceqty, po_type from TSV_MD_POPrice where customershortname = '" & UCase(Trim(cbCusCode.text)) & "'  and PO_NUM= '" & Trim(dT.po_no) & "'  and PT = '" & dT.Customer_Device & "' "
Set rs = Get_OracleRs(pocheck)
If rs.RecordCount = 0 Then

    MsgBox "PO" & Trim(dT.po_no) & "����" & dT.Customer_Device & "δά���۸�,�������ϴ�WO!"
    GetWOData = False
    
   Unload FrmPOPriceSys_NEW
   FrmPOPriceSys_NEW.Show 1
    
    Exit Function
   
   
ElseIf rs.Fields(1).Value = "NRE����" Then
   
    If Trim(txtPo_Price.text) = "" Then
    
    txtPo_Price.text = dT.po_no
    txtPOQTY.text = 1
    
    
    ElseIf Trim(txtPo_Price.text) <> dT.po_no Then
    
     txtPo_Price.text = dT.po_no
     txtPOQTY.text = 1
    Else
 
     txtPOQTY.text = Val(txtPOQTY.text) + 1
     
    End If
 
   If Val(rs.Fields(0).Value) < Val(txtPOQTY.text) Then
    
     MsgBox "WO�������� NERPO" & Trim(dT.po_no) & "����" & dT.Customer_Device & "����!"
      GetWOData = False
     Exit Function
    
   End If
   
   

End If
    
    
    
    
End If


Else


pocheck = "select peaceqty, po_type from TSV_MD_POPrice where customershortname = '" & UCase(Trim(cbCusCode.text)) & "'  and PO_NUM= '" & Trim(dT.po_no) & "'  and PT = '" & dT.Customer_Device & "' "
Set rs = Get_OracleRs(pocheck)
If rs.RecordCount = 0 Then
   
   
   
 PO_ID = GetPOPriceID()



 postr = " insert into TSV_MD_POPrice (ID, CUSTOMERSHORTNAME,CUSTOMERNAME,PO_NUM,PO_DATE,PO_TYPE,PT,QTY,PRICE,UNIT, " & _
         "  Flag, QTECH_CREATED_BY,QTECH_CREATED_DATE,PeaceQty,CUSTAA, DIE_PRICE) values('" & PO_ID & "','" & UCase(Trim(cbCusCode.text)) & "', " & _
         "  '" & cust_name & "','" & Trim(dT.po_no) & "',sysdate,'��������', '" & dT.Customer_Device & "', 99999,'" & w_price & "',  " & _
         "  '" & price_unit & "','Y', '', sysdate,999999,'NA','" & d_price & "' )   "

 AddSql (postr)
 
 
 postr1 = " insert into erptemp .. tblBB_CSRPO values (  '" & UCase(Trim(cbCusCode.text)) & "' ,'" & Trim(dT.po_no) & "',10,'',  '" & dT.Customer_Device & "'  " & _
          " , 99999, 99999 ,'" & w_price & "','" & d_price & "','" & price_unit & "' ,'',CONVERT(varchar(100), getdate(), 20) , '') "
 
AddSql2 (postr1)


ElseIf rs.Fields(1).Value = "NRE����" Then
   
    If Trim(txtPo_Price.text) = "" Then
    
    txtPo_Price.text = dT.po_no
    txtPOQTY.text = 1
    
    
    ElseIf Trim(txtPo_Price.text) <> dT.po_no Then
    
     txtPo_Price.text = dT.po_no
     txtPOQTY.text = 1
    Else
 
     txtPOQTY.text = Val(txtPOQTY.text) + 1
     
    End If
 
   If Val(rs.Fields(0).Value) < Val(txtPOQTY.text) Then
    
     MsgBox "WO�������� NERPO" & Trim(dT.po_no) & "����" & dT.Customer_Device & "����!"
      GetWOData = False
     Exit Function
    
   End If
   
   
   
   
   

End If
End If
End If
End If


End Function

Private Function setWOData(dT As tyWO) As Boolean
setWOData = False
If SetMarkingCode(dT) = False Then
    Exit Function

End If

Call SetWaferVersion(dT)
Call SetWaferDies(dT)
setWOData = True

End Function

Private Function showWOData(dT As tyWO, i As Integer)
Dim j As Integer

With Fps(0)
    .MaxRows = .MaxRows + 1
    .SetText 1, i - 1, dT.ITEM
    .SetText 2, i - 1, dT.po_no
    .SetText 3, i - 1, dT.SUPPLIER
    .SetText 4, i - 1, dT.SHIP_TO
    .SetText 5, i - 1, dT.Fab_Device
    .SetText 6, i - 1, dT.Customer_Device
    .SetText 7, i - 1, dT.WAFER_VERSION
    .SetText 8, i - 1, dT.MARKING_CODE
    .SetText 9, i - 1, dT.WO_DATE
    .SetText 10, i - 1, dT.Lot_id
    .SetText 11, i - 1, dT.wafer_id
    .SetText 12, i - 1, dT.GOOD_DIES_PCS
    .SetText 13, i - 1, dT.GROSS_DIES_PCS
    .SetText 14, i - 1, dT.HT_DEVICE
    .SetText 15, i - 1, dT.REMARK
    .SetText 16, i - 1, dT.TRADE_TYPE
    .SetText 17, i - 1, dT.DATA1
    .SetText 18, i - 1, dT.DATA2
    .SetText 19, i - 1, dT.DATA3
    .SetText 20, i - 1, dT.DATA4
    .SetText 21, i - 1, dT.DATA5

End With

End Function

Private Function SetMarkingCodeByPN(ByRef dT As tyWO) As Boolean
SetMarkingCodeByPN = False
SetMarkingCodeByPN = True

End Function

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       SetMarkingCode
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/14-14:27:34
'
' Parameters :       dT (tyWO)
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       SetMarkingCode
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       0-354AD8C194ED4
' Date-Time  :       2020-1-2-9:55:44
'
' Parameters :       dT (tyWO)
'--------------------------------------------------------------------------------
Private Function SetMarkingCode(ByRef dT As tyWO) As Boolean
SetMarkingCode = False
Dim strMarkingCodeWO As String

strMarkingCodeWO = dT.MARKING_CODE

Select Case dT.CUSTOMER_CODE

    Case "SH50"
        If dT.Customer_Device = "WS14DZ03" Then
            dT.MARKING_CODE = Left(dT.MARKING_CODE, 3) & "\\" & Right$(dT.MARKING_CODE, 3)

        End If

    Case "SX", "HJ", "TJ003", "JS140", "BJ153"

        Select Case dT.Customer_Device

            Case "OV02A", "OV02A-E", "SP5506-M", "SP5506", "SP5506-E", "SP5506-EM", "SP8407-E", "SP8407", "SP5407-E", "SP5407", "SP2735", "OV02B10", "OV02B1B-E"
                dT.MARKING_CODE = GetSX8CodeID(dT.Lot_id, dT.wafer_id)

            Case Else
                dT.MARKING_CODE = GetSXCodeID()

        End Select

        Select Case dT.HT_DEVICE

            Case "YSX005M", "YSX006M", "YSX004M"
                'dT.MARKING_CODE = GetSX8CodeID(dT.Lot_id, dT.wafer_id)

        End Select

    Case "81"

        Select Case dT.Customer_Device

            Case "1103A_A"
                dT.MARKING_CODE = "HS" & Mid(Year(Now), 3, 1) & "A" & Mid(Year(Now), 4, 1) & "S" & Right("0" & DatePart("WW", Now), 2)

            Case "110F_A"
                dT.MARKING_CODE = "EHD" & "\\" & "510"

        End Select

    Case "GT"

        Select Case dT.Customer_Device

            Case "SIV121DU"
                dT.MARKING_CODE = GetGTCodeID()

        End Select

    Case "GD108", "HK080"

        Select Case dT.Customer_Device

            Case "GW1N-LV1CS30C6/I5"
                If dT.DATA1 = "" Or dT.DATA2 = "" Or dT.DATA3 = "" Then
                    MsgBox "Q��,R��,S��Ϊ���������ֶ�,����Ϊ��", vbInformation, "����"
                    Exit Function

                End If

                dT.MARKING_CODE = dT.DATA1 & "\\" & dT.DATA2 & "\\" & dT.DATA3 & "\\" & dT.Lot_id

            Case "GW1N-LV4CS72"
                If dT.DATA1 = "" Or dT.DATA2 = "" Or dT.DATA3 = "" Then
                    MsgBox "Q��,R��,S��Ϊ���������ֶ�,����Ϊ��", vbInformation, "����"
                    Exit Function

                End If

                dT.MARKING_CODE = dT.DATA1 & "\\" & dT.DATA2 & "\\" & Right(Year(Now), 2) & Right("0" & DatePart("WW", Now), 2) & "B" & "\\" & dT.Lot_id
                dT.MARKING_CODE2 = dT.DATA1 & "\\" & dT.DATA3 & "\\" & Right(Year(Now), 2) & Right("0" & DatePart("WW", Now), 2) & "B" & "\\" & dT.Lot_id

        End Select

    Case "69"
        dT.MARKING_CODE = Mid(dT.Lot_id, 2, 6) & Mid("ABCDEFGHIJKLMNOPQRSTUVWXY", dT.wafer_id, 1)

    Case "SG005", "TW079"
        dT.MARKING_CODE = Mid$(dT.Customer_Device, InStr(dT.Customer_Device, "-") + 2, 1) & Right(Year(Now), 1) & Hex(Month(Now)) & Mid$("123456789ABCDEFGHIJKLMNOP", dT.wafer_id, 1)
        If InStr(dT.Lot_id, ".") > 0 Then
            dT.MARKING_CODE = dT.MARKING_CODE & Mid$(dT.Lot_id, InStr(dT.Lot_id, ".") - 4, 4)
        Else
            dT.MARKING_CODE = dT.MARKING_CODE & Right$(dT.Lot_id, 4)

        End If

    Case "US026"
        If dT.Customer_Device = "TM2G1" Then
            dT.MARKING_CODE = " " & Right(Year(Now), 1) & Mid("123456789ABC", Month(Now), 1) & Mid$("123456789ABCDEFGHIJKLMNOP", dT.wafer_id, 1) & Right(Left(dT.Lot_id, InStr(dT.Lot_id, ".") - 1), 4)
        Else
            dT.MARKING_CODE = Mid$(dT.Customer_Device, InStr(dT.Customer_Device, "-") + 2, 1) & Right(Year(Now), 1) & Hex(Month(Now)) & Mid$("123456789ABCDEFGHIJKLMNOP", dT.wafer_id, 1)
            If InStr(dT.Lot_id, ".") > 0 Then
                dT.MARKING_CODE = dT.MARKING_CODE & Mid$(dT.Lot_id, InStr(dT.Lot_id, ".") - 4, 4)
            Else
                dT.MARKING_CODE = dT.MARKING_CODE & Right$(dT.Lot_id, 4)

            End If

        End If

    Case "TW067"
        If Len(dT.DATA1) <> 5 Then
            MsgBox "Q�б�����5λ��Ϣ�������ȡ��", vbInformation, "��ʾ"
            Exit Function

        End If

        dT.MARKING_CODE = dT.DATA1 & Mid$("123456789ABCDEFGHJKLMNPQRSTUVW", dT.wafer_id, 1)

        '        Select Case dT.Customer_Device �������
        '
        '            Case "PS5250LT", "PS5250LT-AA", "PS5260LT", "PS5250LT-AA"
        '                dT.MARKING_CODE = dT.DATA1 & Mid$("123456789ABCDEFGHJKLMNPQRSTUVW", dT.Wafer_id, 1)
        '
        '             Case Else
        '
        '
        '
        '
        '        End Select
    Case "SH192"
        If dT.HT_DEVICE = "XSH192002" Then
            If InStr(dT.Lot_id, ".") > 0 Then
                dT.MARKING_CODE = "HTG6C" + "\\" + Mid(dT.Lot_id, InStr(dT.Lot_id, ".") - 4, 4) + "\\" + Trim(Right(Year(Now), 2)) + Right("0" & DatePart("WW", Now), 2)
            Else
                dT.MARKING_CODE = "HTG6C" + "\\" + Right(dT.Lot_id, 4) + "\\" + Trim(Right(Year(Now), 2)) + Right("0" & DatePart("WW", Now), 2)

            End If

        End If

    Case "SH115"
        dT.MARKING_CODE = Mid(dT.Customer_Device, 3, 4) + "\\" + Trim(Right(Year(Now), 2)) + Right("0" & DatePart("WW", Now), 2)

    Case "KR001"

        Select Case dT.HT_DEVICE

            Case "XKR00103"
                dT.MARKING_CODE = GetKRMark(dT.Lot_id, dT.wafer_id)

                ' Changed by: Project Administrator at: 2019/9/9-13:28:52 on machine: DESKTOP-MSUG5JD ��� Ҫ�����PC7090K,��������ͳһ��ʽ
                '            Case "PS1130K", "PS4210K", "PC7080D", "PK2130K", "PCB030K", "PK3130K", "PV3109K"
                '                dT.MARKING_CODE = GetKRMarkP(dT.Lot_id, dT.Wafer_id)
            Case Else
                dT.MARKING_CODE = GetKRMarkP(dT.Lot_id, dT.wafer_id)

        End Select

    Case "KR002"
        dT.MARKING_CODE = Right$(dT.Lot_id, 2) & Right$("0" & dT.wafer_id, 2)

    Case "KR009"
        If UCase(dT.Customer_Device) = "HI-1A1" Then
            dT.MARKING_CODE = Right("0" & dT.wafer_id, 2) & "2" & Mid(dT.Lot_id, 5, 3)

        End If

    Case "HY"
        If UCase(dT.Customer_Device) = "HI-258" Then
            dT.MARKING_CODE = Right("0" & dT.wafer_id, 2) & "2" & Mid(dT.Lot_id, 5, 3)

        End If

    Case "AT71", "AH033", "SZ280"

        Select Case dT.Customer_Device

            Case "FP5513E4"
                If dT.DATA2 = "" Or dT.DATA3 = "" Then
                    MsgBox "�г���������дR�к�S�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = dT.DATA2 & dT.DATA3

            Case "FP5510EE4"
                If dT.DATA2 = "" Or dT.DATA3 = "" Then
                    MsgBox "�г���������дR�к�S�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "8a" & Right$(Year(Now), 1) & "\\" & Hex(Month(Now)) & Right$(dT.DATA3, 2)

            Case "FP5510FE4"
                If dT.DATA2 = "" Or dT.DATA3 = "" Then
                    MsgBox "�г���������дR�к�S�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "9-" & Right$(Year(Now), 1) & "\\" & Hex(Month(Now)) & Right$(dT.DATA3, 2)

            Case "FP5519E4"
                If dT.DATA2 = "" Or dT.DATA3 = "" Then
                    MsgBox "�г���������дR�к�S�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "7-" & Right$(Year(Now), 1) & "\\" & Hex(Month(Now)) & Right$(dT.DATA3, 2)

            Case "FP5510E2"
                If dT.DATA2 = "" Or dT.DATA3 = "" Then
                    MsgBox "�г���������дR�к�S�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "2=" & Right$(Year(Now), 1) & "\\" & Hex(Month(Now)) & Right$(dT.DATA3, 2)

            Case "FP5510EE4AEE"
                If dT.DATA2 = "" Or dT.DATA3 = "" Then
                    MsgBox "�г���������дR�к�S�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "8a" & Right$(Year(Now), 1) & "\\" & Hex(Month(Now)) & Right$(dT.DATA3, 2)

            Case "FP5516WE4"
                If dT.DATA2 = "" Or dT.DATA3 = "" Then
                    MsgBox "�г���������дR�к�S�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "5+" & Right$(Year(Now), 1) & "\\" & Hex(Month(Now)) & Right$(dT.DATA3, 2)

        End Select

        Select Case dT.HT_DEVICE

            Case "XAT71023B"
                If dT.DATA2 = "" Or dT.DATA3 = "" Then
                    MsgBox "�г���������дR�к�S�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "1+" & Right$(Year(Now), 1) & "\\" & Hex(Month(Now)) & Right$(dT.DATA3, 2)

            Case "XAT71019B"
                If dT.DATA2 = "" Or dT.DATA3 = "" Then
                    MsgBox "�г���������дR�к�S�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "5-" & Right$(Year(Now), 1) & "\\" & Hex(Month(Now)) & Right$(dT.DATA3, 2)

            Case "XAT71024B"
                If dT.DATA2 = "" Or dT.DATA3 = "" Then
                    MsgBox "�г���������дR�к�S�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "3=" & Right$(Year(Now), 1) & "\\" & Hex(Month(Now)) & Right$(dT.DATA3, 2)

        End Select

    Case "RD"
        If dT.Customer_Device = "RDA2216" Then
            dT.MARKING_CODE = "RDA" & "2216" & Mid(dT.Lot_id, 3, 4) & Right$("0" & dT.wafer_id, 2)

        End If

    Case "AB18"
        dT.MARKING_CODE = Replace(dT.MARKING_CODE, "****", Trim(Right(Year(Now), 2)) + Right("0" & DatePart("WW", Now), 2))

    Case "SD"
        If dT.Customer_Device = "SD12" Then
            dT.MARKING_CODE = "SD12" & "\\" & Mid(dT.Lot_id, 2, 6)

        End If

    Case "SH103"

        'dT.MARKING_CODE = Right$(dT.Customer_Device, 4) & "\\" & Left$(dT.WAFER_VERSION, 4) & "\\" & Right$(dT.WAFER_VERSION, 2)
        Select Case dT.HT_DEVICE

            Case "XSH103003"    ' ��ѩ��, ' Changed by: Project Administrator at: 2019/8/14-14:28:23 on machine: DESKTOP-MSUG5JD
                If Len(dT.WAFER_VERSION) <> 4 Then
                    MsgBox "�г���������дG�е�ֵ,��G�б�����4λ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "K" & Left(dT.WAFER_VERSION, 2) & "\\" & Right(dT.WAFER_VERSION, 2)

        End Select

    Case "DA69"

        Select Case dT.HT_DEVICE

            Case "XDA69001B"
                dT.MARKING_CODE = "46A" & "\\" & Right$(Year(Now), 1) & Right("0" & DatePart("WW", Now), 2)

            Case "XDA69002B"
                dT.MARKING_CODE = "772" & "\\" & Right$(Year(Now), 1) & Right("0" & DatePart("WW", Now), 2)

            Case "XDA69003B"
                dT.MARKING_CODE = "96B" & "\\" & Right$(Year(Now), 1) & Right("0" & DatePart("WW", Now), 2)

            Case "XDA69004B"
                dT.MARKING_CODE = "64BA" & "\\" & Right$(Year(Now), 1) & Right("0" & DatePart("WW", Now), 2)

            Case "XDA69006B"
                ' Changed by: Project Administrator at: 2019/8/19-10:23:30 on machine: DESKTOP-MSUG5JD ��ѩ
                dT.MARKING_CODE = "97C" & "\\" & Right$(Year(Now), 1) & Right("0" & DatePart("WW", Now), 2)

            Case "XDA69A03B"
                dT.MARKING_CODE = "96U" & "\\" & Right$(Year(Now), 1) & Right("0" & DatePart("WW", Now), 2)

            Case "XDA69005B"    ' 2019�°���C,2020�ϰ���D,2020�°���E,���ε���
                'dT.MARKING_CODE = "769" & "\\" & "W" & "C" & Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ3BCDEFGHIJKLMNOPQRSTUVWXY456", DatePart("WW", Now), 1)
                dT.MARKING_CODE = "769" & "\\" & "W" & "D" & Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ3BCDEFGHIJKLMNOPQRSTUVWXY456", DatePart("WW", Now), 1)

            Case "XDA69007B"
                dT.MARKING_CODE = "768" & "\\" & Right$(Year(Now), 1) & Right("0" & DatePart("WW", Now), 2)

            Case "XDA69B03B"
                dT.MARKING_CODE = "96W" & "\\" & Right$(Year(Now), 1) & Right("0" & DatePart("WW", Now), 2)

            Case "XDA69A06B"
                dT.MARKING_CODE = "97U" & "\\" & Right$(Year(Now), 1) & Right("0" & DatePart("WW", Now), 2)

        End Select

    Case "AC64"

        Select Case dT.HT_DEVICE

            Case "XAC64005B", "XAC64002B", "XAC64009B", "XAC64008B", "XAC64014B", "XAC64A08B", "XAC64B08B", "XAC64C08B"
                If dT.WAFER_VERSION = "" Then
                    MsgBox "�г���������дG�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "LUB" & "\\" & dT.WAFER_VERSION

            Case "XAC64011B"
                If dT.WAFER_VERSION = "" Then
                    MsgBox "�г���������дG�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "FLB" & "\\" & dT.WAFER_VERSION

            Case "XAC64013B"
                If dT.WAFER_VERSION = "" Then
                    MsgBox "�г���������дG�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "LVB" & "\\" & dT.WAFER_VERSION

            Case "XAC64006B"
                If dT.WAFER_VERSION = "" Then
                    MsgBox "�г���������дG�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "92011" & "\\" & dT.WAFER_VERSION

            Case "XAC64007B"
                If dT.WAFER_VERSION = "" Then
                    MsgBox "�г���������дG�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "92012" & "\\" & dT.WAFER_VERSION

            Case "XAC64012B"
                If dT.WAFER_VERSION = "" Then
                    MsgBox "�г���������дG�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "LYB" & "\\" & dT.WAFER_VERSION

            Case "XAC64A01B"
                If dT.WAFER_VERSION = "" Then
                    MsgBox "�г���������дG�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "KLB" & "\\" & dT.WAFER_VERSION

            Case "XAC64A12B"
                If dT.WAFER_VERSION = "" Then
                    MsgBox "�г���������дG�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                '  dT.MARKING_CODE = "LYB" & "\\" & dT.WAFER_VERSION
                ' dT.MARKING_CODE = "KLB" & "\\" & dT.WAFER_VERSION
                dT.MARKING_CODE = "LYB" & "\\" & dT.WAFER_VERSION

            Case "XAC64B01B"
                If dT.WAFER_VERSION = "" Then
                    MsgBox "�г���������дG�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "KLB" & "\\" & dT.WAFER_VERSION

            Case "XAC64C01B"
                If dT.WAFER_VERSION = "" Then
                    MsgBox "�г���������дG�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "QLC" & "\\" & dT.WAFER_VERSION

            Case "XAC64A13B"
                If dT.WAFER_VERSION = "" Then
                    MsgBox "�г���������дG�е�ֵ,�Թ������ƴ��", vbExclamation, "��ʾ"
                    Exit Function

                End If

                dT.MARKING_CODE = "LVB" & "\\" & dT.WAFER_VERSION

        End Select

    Case "QR"

        Select Case dT.Customer_Device

            Case "MT01", "AX01"
                dT.MARKING_CODE = Right(dT.Lot_id, 4) & Mid$("ABCDEFGHIJKLMNOPQRSTUVWXY", dT.wafer_id, 1)

        End Select

    Case "SH267"    ' �뾲20190506
        '        Select Case dT.Customer_Device
        '
        '            Case "SC2238", "VENUS", "EIAR"
        '                dT.MARKING_CODE = Right(Year(Now), 1) & UCase(Hex(Month(Now))) & Mid$(dT.lot_wafer_id, 3, 4) & Mid$("123456789ABCDEFGHJKLMNPQR", dT.Wafer_id, 1)
        '
        '        End Select
        ' �뾲20190530 ���пͻ�������һ��
        dT.MARKING_CODE = Right(Year(Now), 1) & UCase(Hex(Month(Now))) & Mid$(dT.lot_wafer_id, 3, 4) & Mid$("123456789ABCDEFGHJKLMNPQR", dT.wafer_id, 1)

    Case "HD"

        Select Case dT.Customer_Device

            Case "GH610", "GH611", "GH612"   ' �ƺ��� 20190523
                dT.MARKING_CODE = dT.Customer_Device & "\\" & Split(dT.DATA5, "-")(0) & "\\" & Split(dT.DATA5, "-")(1)

        End Select

        Select Case dT.HT_DEVICE

            Case "XHD004B"
                dT.MARKING_CODE = dT.Customer_Device & "\\" & Split(dT.DATA5, "-")(0) & "-" & Split(dT.DATA5, "-")(1) & "\\" & Right$("00" & dT.wafer_id, 2)

        End Select

        '    Case "AH017"
        '        If Len(dT.Customer_Device) = 11 Then
        '            dT.MARKING_CODE = Mid(dT.Customer_Device, 3, 5) & "\\" & Right(Year(Now), 2) & Right("0" & DatePart("WW", Now), 2) & "\\" & Mid$("ABCDEFGHIJKLMNOPQRSTUVWXY", dT.Wafer_id, 1) & Mid$(dT.Lot_id, InStr(dT.Lot_id, ".") - 3, 3) & "\\" & Mid(dT.Customer_Device, 7, 3)
        '        ElseIf Len(dT.Customer_Device) = 10 Then
        '            dT.MARKING_CODE = Mid(dT.Customer_Device, 3, 4) & "\\" & Right(Year(Now), 2) & Right("0" & DatePart("WW", Now), 2) & "\\" & Mid$("ABCDEFGHIJKLMNOPQRSTUVWXY", dT.Wafer_id, 1) & Mid$(dT.Lot_id, InStr(dT.Lot_id, ".") - 3, 3) & "\\" & Mid(dT.Customer_Device, 7, 3)
        '        Else
        '            MsgBox "�ͻ�����λ������ȷ", vbCritical, "����"
        '            Exit Function
        '
        '        End If
    Case "SZ217"    ' ��ѩ 20190611
        dT.MARKING_CODE = "ST2018"

    Case "AC70"

        Select Case dT.HT_DEVICE

            Case "XAC7013B"
                If Len(Trim(dT.REMARK)) <> 4 Then
                    MsgBox "O���б������4λ�������Ϣ", vbInformation, "����"
                    Exit Function

                End If

                dT.MARKING_CODE = "7F7L" & "\\" & dT.REMARK

            Case "XAC7018B"
                If Len(Trim(dT.REMARK)) <> 4 Then
                    MsgBox "O���б������4λ�������Ϣ", vbInformation, "����"
                    Exit Function

                End If

                dT.MARKING_CODE = "K318" & "\\" & dT.REMARK

            Case "XAC7016B"
                If Len(Trim(dT.REMARK)) <> 4 Then
                    MsgBox "O���б������4λ�������Ϣ", vbInformation, "����"
                    Exit Function

                End If

                dT.MARKING_CODE = "AWINIC" & "\\" & "87339" & "\\" & dT.REMARK

            Case "XAC7009B"
                If Len(Trim(dT.REMARK)) <> 4 Then
                    MsgBox "O���б������4λ�������Ϣ", vbInformation, "����"
                    Exit Function

                End If

                dT.MARKING_CODE = "K327" & "\\" & dT.REMARK

                '            Case "XAC70A2B"
                '                If Len(Trim(dT.REMARK)) <> 4 Then
                '                    MsgBox "O���б������4λ�������Ϣ", vbInformation, "����"
                '                    Exit Function
                '
                '                End If
                '
                '                dT.MARKING_CODE = "3805" & "\\" & dT.REMARK
            Case "XAC7015B"
                If Len(Trim(dT.REMARK)) <> 4 Then
                    MsgBox "O���б������4λ�������Ϣ", vbInformation, "����"
                    Exit Function

                End If

                dT.MARKING_CODE = "C031" & "\\" & dT.REMARK

            Case "XAC7006B"
                If Len(Trim(dT.REMARK)) <> 4 Then
                    MsgBox "O���б������4λ�������Ϣ", vbInformation, "����"
                    Exit Function

                End If

                dT.MARKING_CODE = "K37S" & "\\" & dT.REMARK

            Case "XAC7017B"
                If Len(Trim(dT.REMARK)) <> 4 Then
                    MsgBox "O���б������4λ�������Ϣ", vbInformation, "����"
                    Exit Function

                End If

                dT.MARKING_CODE = "DGY3" & "\\" & dT.REMARK

            Case "XAC7024B"
                If Len(Trim(dT.REMARK)) <> 4 Then
                    MsgBox "O���б������4λ�������Ϣ", vbInformation, "����"
                    Exit Function

                End If

                dT.MARKING_CODE = "UV25" & "\\" & dT.REMARK

            Case "XAC7022B"
                If Len(Trim(dT.REMARK)) <> 4 Then
                    MsgBox "O���б������4λ�������Ϣ", vbInformation, "����"
                    Exit Function

                End If

                dT.MARKING_CODE = "X4LV" & "\\" & dT.REMARK

            Case "XAC7019B"
                If Len(Trim(dT.REMARK)) <> 4 Then
                    MsgBox "O���б������4λ�������Ϣ", vbInformation, "����"
                    Exit Function

                End If

                dT.MARKING_CODE = "MYCOCY" & "\\" & dT.REMARK

        End Select

    Case "ZJ116"

        Select Case dT.HT_DEVICE

            Case "XZJ11601B"
                If Len(Trim(dT.MARKING_CODE)) <> 10 Then
                    MsgBox "H�в���ȷ", vbInformation, "����"
                    Exit Function

                End If

                If InStr(dT.MARKING_CODE, "\\") = 0 Then
                    MsgBox "H�и�ʽ����ȷ", vbInformation, "����"
                    Exit Function

                End If

        End Select

    Case "HW106", "HK093"

        Select Case dT.HT_DEVICE

            Case "XHW10601M", "XHW10602M", "XHW10603M", "XHW10604M" ' Changed by: Project Administrator at: 2019/8/14-14:28:58 on machine: DESKTOP-MSUG5JD ̷˫ǿ
                dT.MARKING_CODE = "A" & Right(Year(Now), 1) & Hex(Month(Now)) & Mid$("123456789ABCDEFGHIJKLMNOP", dT.wafer_id, 1) & Mid$(dT.Lot_id, 3, 5)

        End Select

    Case "SH105"

        Select Case dT.HT_DEVICE

            Case "XSH10501B"
                If dT.DATA1 = "" Then
                    MsgBox "Q�в���Ϊ��", vbInformation, "����"
                    Exit Function

                End If

                dT.MARKING_CODE = Mid(dT.DATA1, 5, 6) & "\\" & Mid(dT.Lot_id, 2, 6)

        End Select

    Case "AC51"

        Select Case dT.HT_DEVICE

            Case "XAC51008B", "XAC51007B"
                If Len(dT.MARKING_CODE) <> 3 Then
                    MsgBox "H�б����ṩ3λ�����", vbInformation, "����"
                    Exit Function

                End If

                dT.MARKING_CODE = "1646" & "\\" & dT.MARKING_CODE

        End Select

End Select

Select Case dT.HT_DEVICE

    Case "XFJ05701B"
        dT.MARKING_CODE = Right$(dT.Customer_Device, 5) & "\\" & Right(Year(Now), 2) & Right("0" & DatePart("WW", Now), 2)    ' 20190926 ��ѩ�� OA

    Case "XSH103A01", "XSH103001"
        If Len(dT.WAFER_VERSION) <> 6 Then
            MsgBox "�г���������дG�е�ֵ,��G�б�����6λ,�Թ������ƴ��", vbExclamation, "��ʾ"
            Exit Function

        End If

        dT.MARKING_CODE = "7983" & "\\" & Left(dT.WAFER_VERSION, 4) & "\\" & Right(dT.WAFER_VERSION, 2) ' 20190925 ��ѩ�� OA

    Case "XAC7023B"
        If Len(dT.REMARK) <> 2 Then
            MsgBox "�г���������дO�е�ֵ,��O�б�����2λ,�Թ������ƴ��", vbExclamation, "��ʾ"
            Exit Function

        End If

        '  dT.MARKING_CODE = dT.REMARK & "\\" & "Z8" ' 20190930 �޼��� MAIL
        dT.MARKING_CODE = "8Z" & "\\" & dT.REMARK '20191212 �ſ����¹���

    Case "X76006B"
        If CLng(dT.wafer_id) < 13 Or CLng(dT.wafer_id) > 19 Then
            MsgBox "waferID����С��13�����19,����ϵIT", vbCritical, "����"
            Exit Function

        End If

        If CLng(dT.wafer_id) >= 13 And CLng(dT.wafer_id) <= 15 Then
            dT.MARKING_CODE = "DC-1" & "\\" & Mid$("KABCDE", Year(Now) - 2018, 1) & Mid("0123456789ACDEFHKLNPRSTUXYZ", (DatePart("WW", Now) + 1) \ 2, 1) & "P2"
        ElseIf CLng(dT.wafer_id) >= 16 And CLng(dT.wafer_id) <= 17 Then
            dT.MARKING_CODE = "DC-2" & "\\" & Mid$("KABCDE", Year(Now) - 2018, 1) & Mid("0123456789ACDEFHKLNPRSTUXYZ", (DatePart("WW", Now) + 1) \ 2, 1) & "P2"
        Else
            dT.MARKING_CODE = "DC-3" & "\\" & Mid$("KABCDE", Year(Now) - 2018, 1) & Mid("0123456789ACDEFHKLNPRSTUXYZ", (DatePart("WW", Now) + 1) \ 2, 1) & "P2"

        End If

    Case "X76008B"
        dT.MARKING_CODE = "6D" & Mid(dT.Lot_id, 5, 2) & "\\" & Mid$("KABCDE", Year(Now) - 2018, 1) & Mid("0123456789ACDEFHKLNPRSTUXYZ", (DatePart("WW", Now) + 1) \ 2, 1) & "P2"

    Case "X76007B"
        dT.MARKING_CODE = "VJ" & Mid(dT.Lot_id, 9, 2) & "\\" & Mid$("KABCDE", Year(Now) - 2018, 1) & Mid("0123456789ACDEFHKLNPRSTUXYZ", (DatePart("WW", Now) + 1) \ 2, 1) & "P2"

    Case "X76010B"
        dT.MARKING_CODE = "6F" & Mid(dT.Lot_id, 5, 2) & "\\" & Mid$("KABCDE", Year(Now) - 2018, 1) & Mid("0123456789ACDEFHKLNPRSTUXYZ", (DatePart("WW", Now) + 1) \ 2, 1) & "P2"

    Case "XSH21801B"
        dT.MARKING_CODE = "SCE"

    Case "Y68559B"
        dT.MARKING_CODE = "BNA" & "\\" & Mid$("KMNPRSTVWXYZ", Year(Now) - 2018, 1) & Right("00" & DatePart("WW", Now), 2) & "\\" & Right(dT.Fab_Device, 3)

    Case "XSH48002B"
        dT.MARKING_CODE = "7" & Right(Year(Now), 1) & Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", DatePart("WW", Now), 1)

End Select

Dim strMarkingCodeSys As String
Dim strMarkingCodeYF  As String
Dim strMarkingCodeTW  As String

strMarkingCodeYF = GetMarkingCodeYF(dT)
strMarkingCodeTW = GetMarkingCodeTW(dT)
If strMarkingCodeYF <> "" And strMarkingCodeTW <> "" Then
    If strMarkingCodeYF <> strMarkingCodeTW Then
        MsgBox "����벻һ��,����ϵITȷ��", vbCritical, "����"
        Exit Function

    End If

End If

If strMarkingCodeYF <> "" Then
    dT.MARKING_CODE = strMarkingCodeYF

End If

If strMarkingCodeTW <> "" Then
    dT.MARKING_CODE = strMarkingCodeTW

End If

SetMarkingCode = True

End Function

Private Function GetMarkingCodeTW(dT As tyWO) As String
Dim strMarkingTemp As String
Dim strSql         As String
Dim strYear        As String
Dim strMonth       As String
Dim strWeek        As String

strYear = Year(Now)
strMonth = Month(Now)
strWeek = Right("00" & DatePart("WW", Now), 2)
strSql = "select Get_MarkingCode('" & dT.CUSTOMER_CODE & "', '" & dT.Customer_Device & "', '" & dT.HT_DEVICE & "', '" & dT.Fab_Device & "', '" & dT.Lot_id & "', '" & dT.wafer_id & "', '" & dT.MARKING_CODE & "', '" & dT.REMARK & "', '" & dT.TRADE_TYPE & "', '" & dT.DATA1 & "', '" & dT.DATA2 & "', '" & dT.DATA3 & "', '" & dT.DATA4 & "', '" & dT.DATA5 & "','" & strYear & "','" & strMonth & "','" & strWeek & "','" & dT.WAFER_VERSION & "') from dual"
GetMarkingCodeTW = Get_OracleStr(strSql)

End Function

Private Function GetMarkingCodeYF(dT As tyWO) As String
Dim strMarkingTemp As String
Dim strSql         As String
Dim strYear        As String
Dim strMonth       As String
Dim strWeek        As String

strYear = Year(Now)
strMonth = Month(Now)
strWeek = Right("00" & DatePart("WW", Now), 2)
strSql = "select Get_MarkingCode_YF('','" & dT.CUSTOMER_CODE & "', '" & dT.Customer_Device & "', '" & dT.HT_DEVICE & "', '" & dT.Fab_Device & "', '" & dT.Lot_id & "', '" & dT.wafer_id & "', '" & dT.MARKING_CODE & "', '" & dT.REMARK & "', '" & dT.TRADE_TYPE & "', '" & dT.DATA1 & "', '" & dT.DATA2 & "', '" & dT.DATA3 & "', '" & dT.DATA4 & "', '" & dT.DATA5 & "','" & strYear & "','" & strMonth & "','" & strWeek & "','" & dT.WAFER_VERSION & "') from dual"
GetMarkingCodeYF = Get_OracleStr(strSql)

End Function

Private Sub SetWaferVersion(dT As tyWO)

End Sub

Private Sub SetWaferDies(dT As tyWO)

End Sub

Private Function ChkWOData(dT As tyWO, i As Integer) As Boolean
Dim rs     As New ADODB.Recordset
Dim strSql As String

ChkWOData = False
'1. ���ͻ�����Ϳͻ������Ƿ��Ӧ
strSql = "select * from tbltsvnpiproduct where customershortname = '" & dT.CUSTOMER_CODE & "' and customerptno1 = '" & dT.Customer_Device & "' and qtechptno = '" & dT.HT_DEVICE & "' "
If Get_OracleCnt(strSql) = 0 Then

    With Fps(0)
        .Row = i - 1
        .Col = 1
        .ForeColor = vbRed
        .text = "NPIδά���ÿͻ�����,�ͻ�����,���ڻ���"

    End With

    Exit Function

End If

If Check_MarkingcodeByHT(dT) = False Then
    Exit Function

End If

'2. �������
If dT.CUSTOMER_CODE <> "37" Then
    If chkMarkingCodeLen(dT) = False Then

        With Fps(0)
            .Row = i - 1
            .Col = 1
            .ForeColor = vbRed
            .text = "��������,����ϵITȷ��"

        End With

        Exit Function

    End If

End If

'3.�������
Dim strCheckMarkingCodeRes As String

If dT.CUSTOMER_CODE <> "SH48" Then
    strCheckMarkingCodeRes = Get_OracleStr("select CHECK_MARKINGCODE('" & dT.MARKING_CODE & "','" & dT.CUSTOMER_CODE & "','" & dT.Customer_Device & "','" & dT.HT_DEVICE & "','" & dT.Lot_id & "','" & dT.wafer_id & "','" & dT.lot_wafer_id & "') from dual ")
    If strCheckMarkingCodeRes <> "0" Then
        MsgBox strCheckMarkingCodeRes, vbCritical, "��ʾ"
        Exit Function

    End If

End If

'4.���AC70���ֶ��ձ�
Dim strPackage As String

If dT.CUSTOMER_CODE = "AC70" Then
    strPackage = Get_OracleStr("SELECT PACKAGE FROM EU010_REFERENCE where CUST_DEVICE = '" & dT.Customer_Device & "'")
    If strPackage = "" Then
        MsgBox "AC70���ֶ��ձ���û���ҵ��û��ֵ���Ϣ,����ϵIT", vbInformation, "��ʾ"
        Exit Function

    End If

End If

ChkWOData = True

End Function

Private Function Check_MarkingcodeByHT(dT As tyWO) As Boolean
Dim strSql     As String
Dim strKeyWord As String
Dim i          As Integer
Dim keyChar1   As String
Dim keyChar2   As String

Check_MarkingcodeByHT = False
' DEFINED_FLAG = "N"˵���ǿ���������,�˴������
If Get_OracleStr("SELECT DEFINED_FLAG FROM TBL_MARKINGCODE_REP  WHERE HT_PN = '" & dT.HT_DEVICE & "'  and APPLY_FLAG = 'Y' ") = "N" Then
    Check_MarkingcodeByHT = True
    Exit Function

End If

strKeyWord = Get_OracleStr("SELECT REMARK FROM TBL_MARKINGCODE_REP  WHERE HT_PN = '" & dT.HT_DEVICE & "'  and APPLY_FLAG = 'Y' ")
If strKeyWord <> "" Then
    If Len(dT.MARKING_CODE) <> Len(strKeyWord) Then
        MsgBox "����볤�ȴ���,�涨����:" & Len(strKeyWord) & vbCrLf & "��ǰ����:" & Len(dT.MARKING_CODE), vbCritical, "����"
        Exit Function

    End If

    For i = 1 To Len(strKeyWord)
        keyChar1 = Mid$(dT.MARKING_CODE, i, 1)
        keyChar2 = Mid$(strKeyWord, i, 1)
        If keyChar2 <> "*" Then
            If keyChar1 <> keyChar2 Then
                MsgBox dT.HT_DEVICE & "�涨�ĵ�" & i & "λ���ַ�:" & keyChar2 & vbCrLf & "��ǰWafer:" & dT.lot_wafer_id & "�����ĵ�" & i & "λ���ַ�:" & keyChar1 & vbCrLf & "����벻���Ϲ淶", vbCritical, "����"
                Exit Function

            End If

        End If

    Next

End If

Check_MarkingcodeByHT = True

End Function

Private Function SaveWOData(dT As tyWO, i As Integer)

On Error GoTo ErrHandle

Dim rs        As New ADODB.Recordset, lKeyID As String, strSql As String
Dim strsqlin3 As String, strsqlin4 As String

' �ػ�WO�ϴ�: waferid�Զ�׷��+  ����
If cbUploadType.ListIndex = 1 Or cbUploadType.ListIndex = 6 Then
    If Get_OracleCnt("select * from ib_waferlist where waferid = '" & dT.lot_wafer_id & "'") = 0 Then

        With Fps(0)
            .Row = i - 1
            .Col = 1
            .ForeColor = vbRed
            .text = "һ�ζ���WaferIDδ������,����ѡ�ػ��ϴ�"

        End With

        Exit Function

    End If

    If (dT.CUSTOMER_CODE = "GD108" Or dT.CUSTOMER_CODE = "HK080") And dT.Customer_Device = "GW1N-LV4CS72" Then
        dT.MARKING_CODE = dT.DATA1 & "\\" & dT.DATA2 & "\\" & Right(Year(Now), 2) & Right("0" & DatePart("WW", Now), 2) & "B" & "\\" & dT.Lot_id
        dT.MARKING_CODE2 = dT.DATA1 & "\\" & dT.DATA3 & "\\" & Right(Year(Now), 2) & Right("0" & DatePart("WW", Now), 2) & "B" & "\\" & dT.Lot_id
        dT.MARKING_CODE = dT.MARKING_CODE & "@@" & dT.MARKING_CODE2
    Else
        If Get_OracleStr("select productid from mappingdatatest where substrateid = '" & dT.lot_wafer_id & "'") <> "" Then
            dT.MARKING_CODE = Get_OracleStr("select productid from mappingdatatest where substrateid = '" & dT.lot_wafer_id & "'")

        End If

        '3.�������
        Dim strCheckMarkingCodeRes As String

        strCheckMarkingCodeRes = Get_OracleStr("select CHECK_MARKINGCODE('" & dT.MARKING_CODE & "','" & dT.CUSTOMER_CODE & "','" & dT.Customer_Device & "','" & dT.HT_DEVICE & "','" & dT.Lot_id & "','" & dT.wafer_id & "','" & dT.lot_wafer_id & "') from dual ")
        If strCheckMarkingCodeRes <> "0" Then
            MsgBox strCheckMarkingCodeRes, vbCritical, "��ʾ"
            Exit Function

        End If

    End If

    Do
        dT.lot_wafer_id = dT.lot_wafer_id & "+"
    Loop Until (Get_OracleCnt("select * from ib_waferlist where waferid = '" & dT.lot_wafer_id & "'") = 0)

End If

strSql = "select * from mappingdatatest a where a.substrateid = '" & dT.lot_wafer_id & "' and filename is not null "
If rs.State = adStateOpen Then rs.Close
rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF Then
    If Get_OracleCnt("select * from ib_Waferlist where waferid = '" & dT.lot_wafer_id & "'") > 0 Then

        With Fps(0)
            .Row = i - 1
            .Col = 1
            .ForeColor = vbRed
            .text = "�ѿ�����,�����ظ��ϴ�"

        End With

        Exit Function
    Else
        '        Cnn.BeginTrans
        '        INIadoCon.BeginTrans
        '        lKeyID = Trim(rs("filename"))
        '        Call BackupWaferID(lKeyID, dT.lot_wafer_id)
        '        Call DelWaferID(lKeyID, dT.lot_wafer_id)
        '        Call InsertHeaderTbl(dT, lKeyID)
        '        Call InsertDetailTbl(dT, lKeyID)
        '
        '        With Fps(0)
        '            .Row = i - 1
        '            .Col = 1
        '            .ForeColor = vbGreen
        '            .Text = "���³ɹ�"
        '
        '        End With
        MsgBox "�Ѿ��ϴ���,����ɾ��֮ǰ�ϴ��Ķ���,�����޷��ٴ��ϴ�", vbCritical, "����"
        Exit Function

    End If

Else
    Cnn.BeginTrans
    INIadoCon.BeginTrans
    lKeyID = GetMaxID()
    Call InsertHeaderTbl(dT, lKeyID)
    Call InsertDetailTbl(dT, lKeyID)

    With Fps(0)
        .Row = i - 1
        .Col = 1
        .ForeColor = vbBlue
        .text = "�����ɹ�"

    End With

End If

rs.Close
Cnn.CommitTrans
INIadoCon.CommitTrans
Exit Function
ErrHandle:
Cnn.RollbackTrans
INIadoCon.RollbackTrans
MsgBox Err.DESCRIPTION, vbCritical + vbInformation, "����"

End Function

Private Sub InsertHeaderTbl(dT As tyWO, lKeyID As String)
Dim strora         As String, strSql As String
Dim strLastWaferID As String
Dim strPackage     As String

Select Case dT.CUSTOMER_CODE

    Case "68", "70", "HK006"
        strora = "insert into CustomerOItbl_test(id,po_num,wafer_visual_inspect,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & " CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,probe_ship_part_type,RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,ASSEMBLY_FACILITY,BATCH_COMMENT_ASSY,jobno,date_code,shipping_mst_level,shipping_mst_260,TARGET_WAF_THICKNESS,COMP_CODE,SHIP_COMMENT) " & _
           "values ('" & lKeyID & "','" & dT.po_no & "','" & gUpID & "','" & dT.Lot_id & "','" & dT.SUPPLIER & "','" & dT.SHIP_TO & "','" & dT.Fab_Device & "'," & "  '" & dT.Customer_Device & "','" & dT.WAFER_VERSION & "','" & dT.WO_DATE & "','" & dT.HT_DEVICE & "','" & dT.CUSTOMER_CODE & "','Y','" & gUserName & "',sysdate,'" & dT.TRADE_TYPE & "'," & "  '" & dT.DATA1 & "','" & dT.DATA2 & "','" & dT.DATA3 & "','" & dT.DATA4 & "','" & dT.DATA5 & "','" & dT.TAX_TYPE & "','" & dT.DATA3 & "', '" & dT.TRADE_TYPE & "', '" & dT.DATA1 & "','" & dT.Fab_Device & "','" & dT.SHIP_TO & "','" & dT.REMARK & "')"
        strSql = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,po_num,wafer_visual_inspect,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & " CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,probe_ship_part_type,RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,ASSEMBLY_FACILITY,BATCH_COMMENT_ASSY,jobno,date_code,TARGET_WAF_THICKNESS,COMP_CODE,SHIP_COMMENT) " & " values ('" & lKeyID & "','" & dT.po_no & "','" & gUpID & "','" & dT.Lot_id & "','" & dT.SUPPLIER & "','" & dT.SHIP_TO & "','" & dT.Fab_Device & "', " & " '" & dT.Customer_Device & "','" & dT.WAFER_VERSION & "','" & dT.WO_DATE & "','" & dT.HT_DEVICE & "','" & dT.CUSTOMER_CODE & "','Y','" & gUserName & "',GETDATE(),'" & dT.TRADE_TYPE & "' ," & " '" & dT.DATA1 & "','" & dT.DATA2 & "','" & dT.DATA3 & "','" & dT.DATA4 & "','" & dT.DATA5 & "','" & dT.TAX_TYPE & "','" & dT.DATA3 & "','" & dT.Fab_Device & "','" & dT.SHIP_TO & "','" & dT.REMARK & "')"

    Case "AC70"
        strPackage = Get_OracleStr("SELECT PACKAGE FROM EU010_REFERENCE where CUST_DEVICE = '" & dT.Customer_Device & "'")
        If strPackage = "" Then
            '    MsgBox "AC70���ֶ��ձ���û���ҵ��û��ֵ���Ϣ,����ϵIT", vbInformation, "��ʾ"
        Else
            dT.DATA2 = strPackage

        End If

        strora = "insert into CustomerOItbl_test(id,po_num,wafer_visual_inspect,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & " CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,probe_ship_part_type,RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,ASSEMBLY_FACILITY,BATCH_COMMENT_ASSY,jobno,date_code,shipping_mst_level,shipping_mst_260,TARGET_WAF_THICKNESS,COMP_CODE,SHIP_COMMENT) " & _
           " values ('" & lKeyID & "','" & dT.po_no & "','" & gUpID & "','" & dT.Lot_id & "','" & dT.SUPPLIER & "','" & dT.SHIP_TO & "','" & dT.Fab_Device & "'," & "  '" & dT.Customer_Device & "','" & dT.WAFER_VERSION & "','" & dT.WO_DATE & "','" & dT.HT_DEVICE & "','" & dT.CUSTOMER_CODE & "','Y','" & gUserName & "',sysdate,'" & dT.TRADE_TYPE & "'," & "  '" & dT.DATA1 & "','" & dT.DATA2 & "','" & dT.DATA3 & "','" & dT.DATA4 & "','" & dT.DATA5 & "','" & dT.TAX_TYPE & "','" & dT.DATA3 & "', '" & dT.TRADE_TYPE & "', '" & dT.DATA1 & "','" & dT.Lot_id & "','" & dT.SHIP_TO & "','" & dT.REMARK & "')"
        strSql = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,po_num,wafer_visual_inspect,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & " CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,probe_ship_part_type,RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,ASSEMBLY_FACILITY,BATCH_COMMENT_ASSY,jobno,date_code,TARGET_WAF_THICKNESS,COMP_CODE) " & " values ('" & lKeyID & "','" & dT.po_no & "','" & gUpID & "','" & dT.Lot_id & "','" & dT.SUPPLIER & "','" & dT.SHIP_TO & "','" & dT.Fab_Device & "', " & " '" & dT.Customer_Device & "','" & dT.WAFER_VERSION & "','" & dT.WO_DATE & "','" & dT.HT_DEVICE & "','" & dT.CUSTOMER_CODE & "','Y','" & gUserName & "',GETDATE(),'" & dT.TRADE_TYPE & "' ," & " '" & dT.DATA1 & "','" & dT.DATA2 & "','" & dT.DATA3 & "','" & dT.DATA4 & "','" & dT.DATA5 & "','" & dT.TAX_TYPE & "','" & dT.DATA3 & "','" & dT.Lot_id & "','" & dT.SHIP_TO & "' )"

    Case Else
        strora = "insert into CustomerOItbl_test(id,po_num,wafer_visual_inspect,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & " CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,probe_ship_part_type,RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,ASSEMBLY_FACILITY,BATCH_COMMENT_ASSY,jobno,date_code,shipping_mst_level,shipping_mst_260,TARGET_WAF_THICKNESS,COMP_CODE,SHIP_COMMENT) " & _
           " values ('" & lKeyID & "','" & dT.po_no & "','" & gUpID & "','" & dT.Lot_id & "','" & dT.SUPPLIER & "','" & dT.SHIP_TO & "','" & dT.Fab_Device & "'," & "  '" & dT.Customer_Device & "','" & dT.WAFER_VERSION & "','" & dT.WO_DATE & "','" & dT.HT_DEVICE & "','" & dT.CUSTOMER_CODE & "','Y','" & gUserName & "',sysdate,'" & dT.TRADE_TYPE & "'," & "  '" & dT.DATA1 & "','" & dT.DATA2 & "','" & dT.DATA3 & "','" & dT.DATA4 & "','" & dT.DATA5 & "','" & dT.TAX_TYPE & "','" & dT.DATA3 & "', '" & dT.TRADE_TYPE & "', '" & dT.DATA1 & "','" & dT.Lot_id & "','" & dT.SHIP_TO & "','" & dT.REMARK & "')"
        strSql = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,po_num,wafer_visual_inspect,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & " CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,probe_ship_part_type,RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,ASSEMBLY_FACILITY,BATCH_COMMENT_ASSY,jobno,date_code,TARGET_WAF_THICKNESS,COMP_CODE,SHIP_COMMENT) " & " values ('" & lKeyID & "','" & dT.po_no & "','" & gUpID & "','" & dT.Lot_id & "','" & dT.SUPPLIER & "','" & dT.SHIP_TO & "','" & dT.Fab_Device & "', " & " '" & dT.Customer_Device & "','" & dT.WAFER_VERSION & "','" & dT.WO_DATE & "','" & dT.HT_DEVICE & "','" & dT.CUSTOMER_CODE & "','Y','" & gUserName & "',GETDATE(),'" & dT.TRADE_TYPE & "' ," & " '" & dT.DATA1 & "','" & dT.DATA2 & "','" & dT.DATA3 & "','" & dT.DATA4 & "','" & dT.DATA5 & "','" & dT.TAX_TYPE & "','" & dT.DATA3 & "','" & dT.Lot_id & "','" & dT.SHIP_TO & "','" & dT.REMARK & "' )"

End Select

' 37�ع�
If dT.CUSTOMER_CODE = "37" And cbUploadType.ListIndex = 6 Then
    strLastWaferID = Left$(dT.lot_wafer_id, Len(dT.lot_wafer_id) - 1)
    strora = " insert into CustomerOItbl_test(id, po_num,po_item,source_batch_id,source_mtrl_num,mtrl_num,mtrl_desc,test_mtrl_num,test_mtrl_desc,mpn,mpn_desc,source_mtrl_sloc, " & _
       " mtrl_num_mtrlgrp,probe_ship_part_type,offshore_asm_company,offshore_test_company,current_wafer_qty,die_qty,design_id,country_of_fab,fab_conv_id,fab_excr_id,reticle_level_71, " & _
       " reticle_level_72,reticle_level_73,wafer_size,imager_customer_rev,chromaticity,micro_lens_shift,temperature_spec,prb_containment_type,fabrication_facility,prb_excr_id,batch_comment_probe, " & _
       " assy_process_id,dark_bond_pad_assy,assy_serial_type,sticky_backs_to_save,optical_quality,encoded_mark_id,planned_laser_scribe,package_lid_type,package_type,pb_free_package,target_waf_thickness, " & _
       " reliability_sampling,lot_priority,wafer_box_type,test_site,assembly_facility,batch_comment_assy,tst_process_id,elec_special_test,box_type,protective_film_apld,shipping_mst_260,shipping_mst_level, " & _
       " t_price,ship_comment,batch_comment_test,created_date,created_time,unit_price,ref_po,ref_po_item,country_of_assembly,micron_material,date_code,ship_site,special_process_lot,lot_status,custom_part_no, " & _
       " flag,qtech_created_by,qtech_created_date,qtech_lastupdate_by,qtech_lastupdate_date,customershortname,downqty,invflag,wafer_visual_inspect,comp_code,eqdatacode,jobno,zx_fromsite,zx_invoice)   " & _
       " select   '" & lKeyID & "',ct.po_num,ct.po_item,ct.source_batch_id,ct.source_mtrl_num,ct.mtrl_num,ct.mtrl_desc,ct.test_mtrl_num,ct.test_mtrl_desc,ct.mpn,ct.mpn_desc,ct.source_mtrl_sloc,ct.mtrl_num_mtrlgrp, " & _
       " ct.probe_ship_part_type,ct.offshore_asm_company,ct.offshore_test_company,ct.current_wafer_qty,ct.die_qty,ct.design_id,ct.country_of_fab,ct.fab_conv_id,ct.fab_excr_id,ct.reticle_level_71,ct.reticle_level_72, " & _
       " ct.reticle_level_73,ct.wafer_size,ct.imager_customer_rev,ct.chromaticity,ct.micro_lens_shift,ct.temperature_spec,ct.prb_containment_type,ct.fabrication_facility,ct.prb_excr_id,ct.batch_comment_probe, " & _
       " ct.assy_process_id,ct.dark_bond_pad_assy,ct.assy_serial_type,ct.sticky_backs_to_save,ct.optical_quality,ct.encoded_mark_id,ct.planned_laser_scribe,ct.package_lid_type,ct.package_type,ct.pb_free_package, " & _
       " ct.target_waf_thickness,ct.reliability_sampling,ct.lot_priority,ct.wafer_box_type,ct.test_site,ct.assembly_facility,ct.batch_comment_assy,ct.tst_process_id,ct.elec_special_test,ct.box_type, " & _
       " ct.protective_film_apld,ct.shipping_mst_260,ct.shipping_mst_level,ct.t_price,ct.ship_comment,ct.batch_comment_test,ct.created_date,ct.created_time,ct.unit_price,ct.ref_po,ct.ref_po_item, " & _
       " ct.country_of_assembly,ct.micron_material,ct.date_code,ct.ship_site,ct.special_process_lot,ct.lot_status, " & _
       " ct.custom_part_no,ct.flag,'" & gUserName & "',sysdate,ct.qtech_lastupdate_by,ct.qtech_lastupdate_date,ct.customershortname,ct.downqty,ct.invflag,'" & gUpID & "', " & _
       " ct.comp_code,ct.eqdatacode,ct.jobno,ct.zx_fromsite,ct.zx_invoice from CustomerOItbl_test ct, MAPPINGDATATEST mt  where mt.substrateid =  '" & strLastWaferID & "' and to_char(ct.id) = mt.filename and rownum = 1 "
    strSql = " insert into [ERPBASE].[dbo].[tblCustomerOI](id, po_num,po_item,source_batch_id,source_mtrl_num,mtrl_num,mtrl_desc,test_mtrl_num,test_mtrl_desc,mpn,mpn_desc,source_mtrl_sloc, " & _
       " mtrl_num_mtrlgrp,probe_ship_part_type,offshore_asm_company,offshore_test_company,current_wafer_qty,die_qty,design_id,country_of_fab,fab_conv_id,fab_excr_id,reticle_level_71, " & _
       " reticle_level_72,reticle_level_73,wafer_size,imager_customer_rev,chromaticity,micro_lens_shift,temperature_spec,prb_containment_type,fabrication_facility,prb_excr_id,batch_comment_probe, " & _
       " assy_process_id,dark_bond_pad_assy,assy_serial_type,sticky_backs_to_save,optical_quality,encoded_mark_id,planned_laser_scribe,package_lid_type,package_type,pb_free_package,target_waf_thickness, " & _
       " reliability_sampling,lot_priority,wafer_box_type,test_site,assembly_facility,batch_comment_assy,tst_process_id,elec_special_test,box_type,protective_film_apld,shipping_mst_260,shipping_mst_level, " & _
       " t_price,ship_comment,batch_comment_test,created_date,created_time,unit_price,ref_po,ref_po_item,country_of_assembly,micron_material,date_code,ship_site,special_process_lot,lot_status,custom_part_no, " & _
       " flag,qtech_created_by,qtech_created_date,qtech_lastupdate_by,qtech_lastupdate_date,customershortname,downqty,wafer_visual_inspect,comp_code,eqdatacode,jobno,zx_fromsite,zx_invoice)   " & _
       " select   '" & lKeyID & "',ct.po_num,ct.po_item,ct.source_batch_id,ct.source_mtrl_num,ct.mtrl_num,ct.mtrl_desc,ct.test_mtrl_num,ct.test_mtrl_desc,ct.mpn,ct.mpn_desc,ct.source_mtrl_sloc,ct.mtrl_num_mtrlgrp, " & _
       " ct.probe_ship_part_type,ct.offshore_asm_company,ct.offshore_test_company,ct.current_wafer_qty,ct.die_qty,ct.design_id,ct.country_of_fab,ct.fab_conv_id,ct.fab_excr_id,ct.reticle_level_71,ct.reticle_level_72, " & _
       " ct.reticle_level_73,ct.wafer_size,ct.imager_customer_rev,ct.chromaticity,ct.micro_lens_shift,ct.temperature_spec,ct.prb_containment_type,ct.fabrication_facility,ct.prb_excr_id,ct.batch_comment_probe, " & _
       " ct.assy_process_id,ct.dark_bond_pad_assy,ct.assy_serial_type,ct.sticky_backs_to_save,ct.optical_quality,ct.encoded_mark_id,ct.planned_laser_scribe,ct.package_lid_type,ct.package_type,ct.pb_free_package, " & _
       " ct.target_waf_thickness,ct.reliability_sampling,ct.lot_priority,ct.wafer_box_type,ct.test_site,ct.assembly_facility,ct.batch_comment_assy,ct.tst_process_id,ct.elec_special_test,ct.box_type, " & _
       " ct.protective_film_apld,ct.shipping_mst_260,ct.shipping_mst_level,ct.t_price,ct.ship_comment,ct.batch_comment_test,ct.created_date,ct.created_time,ct.unit_price,ct.ref_po,ct.ref_po_item, " & _
       " ct.country_of_assembly,ct.micron_material,ct.date_code,ct.ship_site,ct.special_process_lot,ct.lot_status, " & _
       " ct.custom_part_no,ct.flag,'" & gUserName & "',GetDate(),ct.qtech_lastupdate_by,ct.qtech_lastupdate_date,ct.customershortname,ct.downqty,'" & gUpID & "', " & _
       " ct.comp_code,ct.eqdatacode,ct.jobno,ct.zx_fromsite,ct.zx_invoice from [ERPBASE].[dbo].[tblCustomerOI] ct, [ERPBASE].[dbo].[tblmappingData] mt  where mt.substrateid =  '" & strLastWaferID & "' and convert(varchar,ct.id) = mt.filename"

End If

AddSql (strora)
AddSql2 (strSql)

End Sub

Private Sub InsertDetailTbl(dT As tyWO, lKeyID As String)
Dim strora As String, strSql As String

strora = "insert into mappingDataTest(id,substrateid,substratetype,productid,micronlotid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename) values( mappingData_SEQ.Nextval,'" & dT.lot_wafer_id & "','" & dT.TAX_TYPE & "','" & dT.MARKING_CODE & "','" & dT.MARKING_CODE2 & "','" & dT.Lot_id & "','" & dT.wafer_id & "','" & dT.GOOD_DIES_PCS & "','" & (dT.GROSS_DIES_PCS - dT.GOOD_DIES_PCS) & "','" & dT.CUSTOMER_CODE & "','Y','" & gUserName & "',sysdate,'" & lKeyID & "')"
strSql = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,substratetype,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & " values('" & dT.lot_wafer_id & "','" & dT.TAX_TYPE & "','" & dT.MARKING_CODE & "','" & dT.Lot_id & "','" & dT.wafer_id & "','" & dT.GOOD_DIES_PCS & "','" & (dT.GROSS_DIES_PCS - dT.GOOD_DIES_PCS) & "','" & dT.CUSTOMER_CODE & "','Y','" & gUserName & "',GETDATE(),'" & lKeyID & "')"
AddSql (strora)
AddSql2 (strSql)

End Sub

Private Sub UpdateHeaderTbl(dT As tyWO, lKeyID As String)
Dim strora As String, strSql As String, strCusCode As String

strCusCode = UCase(Trim(cbCusCode.text))
strora = "update CustomerOItbl_test set po_num = '" & dT.po_no & "', wafer_visual_inspect = '" & gUpID & "', SHIP_SITE = '" & dT.SUPPLIER & "',Test_site = '" & dT.SHIP_TO & "',FAB_CONV_ID = '" & dT.Fab_Device & "', " & "mpn_desc = '" & dT.Customer_Device & "',Imager_Customer_Rev = '" & dT.WAFER_VERSION & "',Created_Date ='" & dT.WO_DATE & "' , mtrl_num = '" & dT.HT_DEVICE & "',CustomerShortName = '" & dT.CUSTOMER_CODE & "', " & "probe_ship_part_type = '" & dT.TRADE_TYPE & "',RETICLE_LEVEL_71= '" & dT.DATA1 & "' ,RETICLE_LEVEL_72 = '" & dT.DATA2 & "',RETICLE_LEVEL_73 = '" & dT.DATA3 & "',ASSEMBLY_FACILITY = '" & dT.DATA4 & "', " & "BATCH_COMMENT_ASSY = '" & dT.DATA5 & "',jobno = '" & dT.TAX_TYPE & "',date_code = '" & dT.DATA3 & "',shipping_mst_level = '" & dT.TRADE_TYPE & "',shipping_mst_260 = '" & dT.DATA1 & "' " & "where id = '" & lKeyID & "' "
strSql = "update [ERPBASE].[dbo].[tblCustomerOI] set po_num = '" & dT.po_no & "', wafer_visual_inspect = '" & gUpID & "',SHIP_SITE = '" & dT.SUPPLIER & "',Test_site = '" & dT.SHIP_TO & "',FAB_CONV_ID = '" & dT.Fab_Device & "', " & "mpn_desc = '" & dT.Customer_Device & "',Imager_Customer_Rev = '" & dT.WAFER_VERSION & "',Created_Date ='" & dT.WO_DATE & "' , mtrl_num = '" & dT.HT_DEVICE & "',CustomerShortName = '" & dT.CUSTOMER_CODE & "', " & "probe_ship_part_type = '" & dT.TRADE_TYPE & "',RETICLE_LEVEL_71= '" & dT.DATA1 & "' ,RETICLE_LEVEL_72 = '" & dT.DATA2 & "',RETICLE_LEVEL_73 = '" & dT.DATA3 & "',ASSEMBLY_FACILITY = '" & dT.DATA4 & "', " & "BATCH_COMMENT_ASSY = '" & dT.DATA5 & "',jobno = '" & dT.TAX_TYPE & "',date_code = '" & dT.DATA3 & "',shipping_mst_level = '" & dT.TRADE_TYPE & "',shipping_mst_260 = '" & dT.DATA1 & "' " & "where id = '" & lKeyID & "' "
AddSql (strora)
AddSql2 (strSql)

End Sub

Private Sub UpdateDetailTbl(dT As tyWO, lKeyID As String)
Dim strora As String, strSql As String, strCusCode As String

strCusCode = UCase(Trim(cbCusCode.text))
strora = "update mappingDataTest set substratetype = '" & dT.TAX_TYPE & "', productid = '" & dT.MARKING_CODE & "', passbincount ='" & dT.GOOD_DIES_PCS & "',failbincount = '" & (dT.GROSS_DIES_PCS - dT.GOOD_DIES_PCS) & "',CustomerShortName = '" & dT.CUSTOMER_CODE & "', " & "qtech_lastupdate_by = '" & gUserName & "', qtech_lastupdate_date = sysdate where filename = '" & lKeyID & "' "
strSql = "update [ERPBASE].[dbo].[tblmappingData] set substratetype = '" & dT.TAX_TYPE & "',productid = '" & dT.MARKING_CODE & "', passbincount ='" & dT.GOOD_DIES_PCS & "',failbincount = '" & (dT.GROSS_DIES_PCS - dT.GOOD_DIES_PCS) & "',CustomerShortName = '" & dT.CUSTOMER_CODE & "', " & "qtech_lastupdate_by = '" & gUserName & "', qtech_lastupdate_date = GETDATE() where filename = '" & lKeyID & "' "
AddSql (strora)
AddSql2 (strSql)

End Sub

Private Sub updateProgressBar()
If ProgressBar1.Value + lPartSec >= 100 Then
    ProgressBar1.Value = 100
Else
    ProgressBar1.Value = ProgressBar1.Value + lPartSec

End If

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index

    Case 1
        QueData

    Case 2
        ModData

    Case 3
        delData

    Case 4
        Unload Me

End Select

End Sub

Private Sub QueData()
Dim strKid As String
Dim strSql As String
Dim rsWO   As New ADODB.Recordset

Fps(1).MaxRows = 0
If txtKID.text = "" Then
    MsgBox "������ID", vbInformation, "��ʾ"
    Exit Sub

End If

strKid = UCase(Trim$(txtKID.text))
If Len(strKid) < 5 And lblKeyID.Caption <> "�ͻ�����:" Then
    MsgBox "����������5λID", vbInformation, "��ʾ"
    Exit Sub

End If


Select Case lblKeyID.Caption

    Case "LOTID:"
        If Opt(0).Value = True Then
            strSql = "select  '' as ѡ�� ,'' as ״̬,aa.filename as ID,aa.customershortname as �ͻ�����,bb.mpn_desc as �ͻ�����, bb.po_num as PO_NUM, bb.test_mtrl_desc as JOBID, aa.lotid as LOTID, aa.wafer_id as WAFERNO,  " & "aa.substrateid as WAFERID,aa.passbincount+ aa.failbincount as GROSSDIES,aa.passbincount as GOODIES, aa.failbincount as NGDIES, aa.productid  as ����� , bb.imager_customer_rev as ��������, bb.RETICLE_LEVEL_71 as Q��,bb.RETICLE_LEVEL_72 as R��,RETICLE_LEVEL_73 as S��,  " & "aa.qtech_created_by as ������, aa.qtech_created_date as ��������, aa.qtech_lastupdate_by as ������, aa.qtech_lastupdate_date as ��������,bb.flow as GC��ʽ,bb.MTRL_DESC AS GC���ڻ���   " & "from mappingdatatest aa inner join customeroitbl_test bb on aa.filename = to_char(bb.id) and aa.lotid like '%" & strKid & "%' order by LOTID,WAFERID"
        Else
            strSql = "select  '' as ѡ�� ,'' as ״̬,aa.filename as ID,aa.customershortname as �ͻ�����,bb.mpn_desc as �ͻ�����, bb.po_num as PO_NUM, bb.test_mtrl_desc as JOBID, aa.lotid as LOTID, aa.wafer_id as WAFERNO,  " & "aa.substrateid as WAFERID,aa.passbincount+ aa.failbincount as GROSSDIES,aa.passbincount as GOODIES, aa.failbincount as NGDIES, aa.productid  as ����� , bb.imager_customer_rev as ��������, bb.RETICLE_LEVEL_71 as Q��,bb.RETICLE_LEVEL_72 as R��,RETICLE_LEVEL_73 as S��,  " & "aa.qtech_created_by as ������, aa.qtech_created_date as ��������, aa.qtech_lastupdate_by as ������, aa.qtech_lastupdate_date as ��������,bb.flow as GC��ʽ,bb.MTRL_DESC AS GC���ڻ���      " & "from mappingdatatest aa inner join customeroitbl_test bb on aa.filename = to_char(bb.id) and aa.lotid = '" & strKid & "' order by LOTID,WAFERID"

        End If

    Case "WAFERID:"
        If Opt(0).Value = True Then
            strSql = "select  '' as ѡ��,'' as ״̬,aa.filename as ID,aa.customershortname as �ͻ�����,bb.mpn_desc as �ͻ�����, bb.po_num as PO_NUM, bb.test_mtrl_desc as JOBID, aa.lotid as LOTID, aa.wafer_id as WAFERNO,  " & "aa.substrateid as WAFERID,aa.passbincount+ aa.failbincount as GROSSDIES,aa.passbincount as GOODIES, aa.failbincount as NGDIES, aa.productid  as ����� , bb.imager_customer_rev as ��������, bb.RETICLE_LEVEL_71 as Q��,bb.RETICLE_LEVEL_72 as R��,RETICLE_LEVEL_73 as S��,  " & "aa.qtech_created_by as ������, aa.qtech_created_date as ��������, aa.qtech_lastupdate_by as ������, aa.qtech_lastupdate_date as ��������,bb.flow as GC��ʽ,bb.MTRL_DESC AS GC���ڻ���      " & "from mappingdatatest aa inner join customeroitbl_test bb on aa.filename = to_char(bb.id) and aa.substrateid like '%" & strKid & "%' order by LOTID,WAFERID"
        Else
            strSql = "select  '' as ѡ��,'' as ״̬,aa.filename as ID,aa.customershortname as �ͻ�����,bb.mpn_desc as �ͻ�����, bb.po_num as PO_NUM, bb.test_mtrl_desc as JOBID, aa.lotid as LOTID, aa.wafer_id as WAFERNO,  " & "aa.substrateid as WAFERID,aa.passbincount+ aa.failbincount as GROSSDIES,aa.passbincount as GOODIES, aa.failbincount as NGDIES, aa.productid  as ����� , bb.imager_customer_rev as ��������, bb.RETICLE_LEVEL_71 as Q��,bb.RETICLE_LEVEL_72 as R��,RETICLE_LEVEL_73 as S��,  " & "aa.qtech_created_by as ������, aa.qtech_created_date as ��������, aa.qtech_lastupdate_by as ������, aa.qtech_lastupdate_date as ��������,bb.flow as GC��ʽ,bb.MTRL_DESC AS GC���ڻ���      " & "from mappingdatatest aa inner join customeroitbl_test bb on aa.filename = to_char(bb.id) and aa.substrateid = '" & strKid & "' order by LOTID,WAFERID"

        End If

    Case "PONO:"
        If Opt(0).Value = True Then
            strSql = "select  '' as ѡ��,'' as ״̬,aa.filename as ID,aa.customershortname as �ͻ�����,bb.mpn_desc as �ͻ�����, bb.po_num as PO_NUM, bb.test_mtrl_desc as JOBID, aa.lotid as LOTID, aa.wafer_id as WAFERNO,  " & "aa.substrateid as WAFERID,aa.passbincount+ aa.failbincount as GROSSDIES,aa.passbincount as GOODIES, aa.failbincount as NGDIES, aa.productid  as ����� , bb.imager_customer_rev as ��������, bb.RETICLE_LEVEL_71 as Q��,bb.RETICLE_LEVEL_72 as R��,RETICLE_LEVEL_73 as S��,  " & "aa.qtech_created_by as ������, aa.qtech_created_date as ��������, aa.qtech_lastupdate_by as ������, aa.qtech_lastupdate_date as ��������,bb.flow as GC��ʽ,bb.MTRL_DESC AS GC���ڻ���      " & "from mappingdatatest aa inner join customeroitbl_test bb on aa.filename = to_char(bb.id) and bb.po_num = '" & strKid & "' order by LOTID,WAFERID"
        Else
            strSql = "select  '' as ѡ��,'' as ״̬,aa.filename as ID,aa.customershortname as �ͻ�����,bb.mpn_desc as �ͻ�����, bb.po_num as PO_NUM, bb.test_mtrl_desc as JOBID, aa.lotid as LOTID, aa.wafer_id as WAFERNO,  " & "aa.substrateid as WAFERID,aa.passbincount+ aa.failbincount as GROSSDIES,aa.passbincount as GOODIES, aa.failbincount as NGDIES, aa.productid  as ����� , bb.imager_customer_rev as ��������, bb.RETICLE_LEVEL_71 as Q��,bb.RETICLE_LEVEL_72 as R��,RETICLE_LEVEL_73 as S��,  " & "aa.qtech_created_by as ������, aa.qtech_created_date as ��������, aa.qtech_lastupdate_by as ������, aa.qtech_lastupdate_date as ��������,bb.flow as GC��ʽ,bb.MTRL_DESC AS GC���ڻ���      " & "from mappingdatatest aa inner join customeroitbl_test bb on aa.filename = to_char(bb.id) and bb.po_num = '" & strKid & "' order by LOTID,WAFERID"

        End If

    Case "�ͻ�����:"
        If Opt(0).Value = True Then
            strSql = "select  '' as ѡ��,'' as ״̬,aa.filename as ID,aa.customershortname as �ͻ�����,bb.mpn_desc as �ͻ�����, bb.po_num as PO_NUM, bb.test_mtrl_desc as JOBID, aa.lotid as LOTID, aa.wafer_id as WAFERNO,  " & "aa.substrateid as WAFERID,aa.passbincount+ aa.failbincount as GROSSDIES,aa.passbincount as GOODIES, aa.failbincount as NGDIES, aa.productid  as ����� , bb.imager_customer_rev as ��������, bb.RETICLE_LEVEL_71 as Q��,bb.RETICLE_LEVEL_72 as R��,RETICLE_LEVEL_73 as S��,  " & "aa.qtech_created_by as ������, aa.qtech_created_date as ��������, aa.qtech_lastupdate_by as ������, aa.qtech_lastupdate_date as ��������,bb.flow as GC��ʽ,bb.MTRL_DESC AS GC���ڻ���      " & "from mappingdatatest aa inner join customeroitbl_test bb on aa.filename = to_char(bb.id) and bb.customershortname = '" & strKid & "' and  bb.qtech_created_date>sysdate-30 order by �������� desc"
        Else
            strSql = "select  '' as ѡ��,'' as ״̬,aa.filename as ID,aa.customershortname as �ͻ�����,bb.mpn_desc as �ͻ�����, bb.po_num as PO_NUM, bb.test_mtrl_desc as JOBID, aa.lotid as LOTID, aa.wafer_id as WAFERNO,  " & "aa.substrateid as WAFERID,aa.passbincount+ aa.failbincount as GROSSDIES,aa.passbincount as GOODIES, aa.failbincount as NGDIES, aa.productid  as ����� , bb.imager_customer_rev as ��������, bb.RETICLE_LEVEL_71 as Q��,bb.RETICLE_LEVEL_72 as R��,RETICLE_LEVEL_73 as S��,  " & "aa.qtech_created_by as ������, aa.qtech_created_date as ��������, aa.qtech_lastupdate_by as ������, aa.qtech_lastupdate_date as ��������,bb.flow as GC��ʽ,bb.MTRL_DESC AS GC���ڻ���      " & "from mappingdatatest aa inner join customeroitbl_test bb on aa.filename = to_char(bb.id) and bb.customershortname = '" & strKid & "' and bb.qtech_created_date>sysdate-30 order by �������� desc"

        End If

End Select

Set rsWO = Get_OracleRs(strSql)

With Fps(1)
    .MaxRows = 0
    If rsWO.RecordCount > 0 Then
        txtCusCode.text = Trim(rsWO(3).Value)
        txtCusDev.text = Trim(rsWO(4).Value)
        Set .DataSource = rsWO
    Else
        MsgBox "û�в�ѯ����Ч����", vbInformation, "��ʾ"

    End If

End With

End Sub

Private Sub ModData()
Dim strSql As String, bChoose As Boolean, i As Integer
Dim rs     As New ADODB.Recordset, strKeyID As String, strWaferID As String
Dim dT     As WO_MOD

gBackID = Get_OracleStr("select WO_BACK_SEQ.Nextval from dual")
bChoose = False

With Fps(1)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .text = 1 Then
            bChoose = True

        End If

    Next i

End With

If bChoose = False Then
    MsgBox "�빴ѡ��Ҫ�޸ĵ�Wafer��", vbInformation, "��ʾ"
    Exit Sub

End If

If txtMsg2.text = "" Then
    MsgBox "����д�޸�WO��ԭ��", vbInformation, "��ʾ"
    Exit Sub

End If

With Fps(1)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .text = 1 Then
            .Col = 3
            strKeyID = Trim$(.text)
            .Col = 10
            strWaferID = Trim(.text)
            If IsWaferID_OnWorking(strWaferID) = True And gUserName <> "07885" And gUserName <> "16642" And gUserName <> "20418" And gUserName <> "13258" Then
                .Col = 2
                .ForeColor = vbRed
                .text = "�ѿ������������޸�"
                GoTo NextRecord
            Else
                If IsWaferID_OnWorking(strWaferID) = True Then
                    .Col = 2
                    .ForeColor = vbRed
                    .text = "�ѿ�����"

                End If

                Call BackupWaferID(strKeyID, strWaferID)
                .Col = 4
                dT.strCusCode = Trim$(UCase$(.text))
                .Col = 5
                dT.strCUSDEVICE = Trim$(UCase$(.text))
                .Col = 6
                dT.strpo = Trim$(UCase$(.text))
                .Col = 7
                dT.strJobID = Trim$(UCase$(.text))
                .Col = 10
                dT.strWaferID = Trim$(UCase$(.text))
                .Col = 12
                dT.strGoodDies = Trim$(UCase$(.text))
                .Col = 13
                dT.strBadDies = Trim$(UCase$(.text))
                .Col = 14
                dT.strPRODUCTID = Trim$(UCase$(.text))
                .Col = 15
                dT.strVERSION = Trim$(UCase$(.text))
                If dT.strWaferID = "" Then
                    .Col = 2
                    .ForeColor = vbRed
                    .text = "WAFERID����Ϊ��"
                    GoTo NextRecord

                End If

                Call UpdateWaferInfo(strKeyID, dT)
                .Col = 2
                .ForeColor = vbBlue
                .text = "�޸ĳɹ�"

            End If

        End If

NextRecord:
    Next i

End With

Sleep (200)
Call QueData
Call SentMesToPMC_MOD

End Sub

Private Sub delData()

On Error GoTo Ert

Dim strSql   As String, bChoose As Boolean, i As Integer, j As Integer, k As Integer
Dim rs       As New ADODB.Recordset, strKeyID As String, strWaferID As String
Dim Rs2      As New ADODB.Recordset
Dim xlsApp   As Excel.Application
Dim xlsBook  As Excel.Workbook
Dim xlsSheet As Excel.Worksheet

j = 2
bChoose = False
gBackID = Get_OracleStr("select WO_BACK_SEQ.Nextval from dual")

With Fps(1)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .text = 1 Then
            bChoose = True

        End If

    Next i

End With

If bChoose = False Then
    MsgBox "�빴ѡ��Ҫɾ����Wafer��", vbInformation, "��ʾ"
    Exit Sub

End If

If txtMsg2.text = "" Then
    MsgBox "����дɾ��WO��ԭ��", vbInformation, "��ʾ"
    Exit Sub

End If

Set xlsApp = CreateObject("Excel.Application")
Set xlsBook = xlsApp.Workbooks.Add
Set xlsSheet = xlsBook.Worksheets(1)

With xlsApp
    .Rows(1).Font.Bold = True

End With

With xlsSheet
    .Cells(1, 1) = "״̬"
    .Cells(1, 2) = "�ͻ�����"
    .Cells(1, 3) = "�ͻ�����"
    .Cells(1, 4) = "PO_NUM"
    .Cells(1, 5) = "JOBID"
    .Cells(1, 6) = "LOTID"
    .Cells(1, 7) = "WAFERNO"
    .Cells(1, 8) = "WAFERID"
    .Cells(1, 9) = "GROSSDIES"
    .Cells(1, 10) = "GOODIES"
    .Cells(1, 11) = "NGDIES"
    .Cells(1, 12) = "�����"
    .Cells(1, 13) = "��������"
    .Cells(1, 14) = "������"
    .Cells(1, 15) = "��������"
    .Cells(1, 16) = "������"
    .Cells(1, 17) = "��������"

End With

With Fps(1)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .text = 1 Then
            .Col = 3
            strKeyID = Trim$(.text)
            .Col = 10
            strWaferID = Trim(.text)
            If gUserName <> "07885" And IsWaferID_OnWorking(strWaferID) = True Then
                .Col = 2
                .ForeColor = vbRed
                .text = "�ѿ�����������ɾ��"
                GoTo NextRecord
            Else
                If IsWaferID_OnWorking(strWaferID) = True Then
                    .Col = 2
                    .ForeColor = vbBlue
                    .text = "�ѿ�����,��Ȩɾ���ɹ�"
                Else
                    .Col = 2
                    .ForeColor = vbBlue
                    .text = "ɾ���ɹ�"

                End If

                Dim strTemp As String

                strTemp = " select 'ɾ��' as ״̬,aa.customershortname as �ͻ�����,bb.mpn_desc as �ͻ�����,bb.po_num as PO_NUM,bb.test_mtrl_desc as JOBID, " & " aa.lotid as LOTID,aa.wafer_id as WAFERNO,aa.substrateid as WAFERID,aa.passbincount + aa.failbincount as GROSSDIES, " & " aa.passbincount as GOODIES,aa.failbincount as NGDIES,aa.productid as �����,bb.imager_customer_rev as ��������,aa.qtech_created_by as ������, " & " aa.qtech_created_date as ��������,aa.qtech_lastupdate_by as ������,aa.qtech_lastupdate_date as �������� from mappingdatatest aa inner join customeroitbl_test bb on aa.filename = to_char(bb.id) and aa.filename = '" & strKeyID & "' "
                Set Rs2 = Get_OracleRs(strTemp)

                With xlsSheet

                    For k = 0 To Rs2.Fields.count - 1
                        .Cells(j, k + 1) = Rs2(k).Value
                    Next

                End With

                j = j + 1
                Call BackupWaferID(strKeyID, strWaferID)
                Call DelWaferID(strKeyID, strWaferID)

            End If

        End If

NextRecord:
    Next i

End With

If Trim(txtCusCode.text) = "37" And lblKeyID.Caption = "PONO:" And Len(txtKID.text) <> 0 Then
    Dim strpode    As String
    Dim strpode_1  As String
    Dim strpode1   As String
    Dim strpode1_1 As String

    strpode = "delete  from tsv_md_poprice_tmp where po_num = '" & Trim(txtKID.text) & "'"
    strpode_1 = " delete  from tsv_md_poprice where po_num = '" & Trim(txtKID.text) & "'"
    strpode1 = " DELETE FROM erptemp..tblBB_CSRPO WHERE PO_NUM = '" & Trim(txtKID.text) & "' "
    strpode1_1 = " DELETE FROM erptemp..tblBB_CSRPO_TMP WHERE PO_NUM = '" & Trim(txtKID.text) & "' "
    AddSql (strpode)
    AddSql2 (strpode1)
    AddSql (strpode_1)
    AddSql2 (strpode1_1)

End If

Sleep (1000)
Call QueData
If gUserName = "07885" Then
    Exit Sub

End If

Dim strWOPath  As String
Dim strCusCode As String, strCusDev As String

strCusCode = txtCusCode.text
strCusDev = txtCusDev.text
strWOPath = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\��ɾ��\" & strCusCode
If Dir(strWOPath, vbDirectory) = "" Then
    MkDir strWOPath

End If

strWOPath = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\��ɾ��\" & strCusCode & "\" & Replace(strCusDev, "/", "")
If Dir(strWOPath, vbDirectory) = "" Then
    MkDir strWOPath

End If

strFileName = strWOPath & "\" & Format(Now, "YYYY-MMDD-HH-MM-SS") & ".xlsx"
xlsBook.SaveAs strFileName
xlsApp.Visible = True
Set xlsApp = Nothing
Call SentMesToPMC_DEL
Exit Sub
Ert:
If Not (xlsApp Is Nothing) Then
    Set xlsApp = Nothing

End If

End Sub

Private Sub BackupWaferID(strKeyID As String, strWaferID As String)
Dim strSql As String

AddSql ("delete from mappingdatatest_bak where filename = '" & strKeyID & "'")
AddSql ("delete from customeroitbl_test_bak where id = '" & strKeyID & "'")
AddSql2 ("delete from ERPBASE.dbo.tblCustomerOI_TEMP where id = '" & strKeyID & "'")
strSql = "insert into mappingdatatest_bak(id,substrateid,substratetype,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename,remark) select id,substrateid,substratetype,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename,'" & gBackID & "' from mappingdatatest where filename = '" & strKeyID & "' and substrateid = '" & strWaferID & "'"
AddSql (strSql)
strSql = "insert into customeroitbl_test_bak(QTECH_LASTUPDATE_BY,id,po_num,wafer_visual_inspect,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,probe_ship_part_type,  RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,ASSEMBLY_FACILITY,BATCH_COMMENT_ASSY,jobno,date_code,shipping_mst_level,shipping_mst_260) " & "select '" & gBackID & "',id,po_num,wafer_visual_inspect,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,probe_ship_part_type,  RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,ASSEMBLY_FACILITY,BATCH_COMMENT_ASSY,jobno,date_code,shipping_mst_level,shipping_mst_260 from customeroitbl_test where id = '" & strKeyID & "'"
AddSql (strSql)
strSql = " insert into ERPBASE.dbo.tblCustomerOI_TEMP(QTECH_LASTUPDATE_BY,id,po_num,wafer_visual_inspect,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,probe_ship_part_type,  RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,ASSEMBLY_FACILITY,BATCH_COMMENT_ASSY,jobno,date_code,shipping_mst_level,shipping_mst_260) " & " select '" & gBackID & "',id,po_num,wafer_visual_inspect,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,probe_ship_part_type,  RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,ASSEMBLY_FACILITY,BATCH_COMMENT_ASSY,jobno,date_code,shipping_mst_level,shipping_mst_260 from ERPBASE.dbo.tblCustomerOI where id = " & strKeyID & " "
AddSql2 (strSql)

End Sub

Private Sub UpdateWaferInfo(strKeyID As String, dT As WO_MOD)
Dim strSql   As String
Dim strNewID As String

strSql = "select count(1) from mappingdatatest where filename = '" & strKeyID & "'"
If Get_OracleNo(strSql) > 1 Then
    strNewID = SplitWOFileName(strKeyID, dT.strWaferID)
Else
    strNewID = strKeyID

End If

strSql = "update mappingdatatest set remark = '" & gBackID & "',customershortname = '" & dT.strCusCode & "',productid = '" & dT.strPRODUCTID & "', PASSBINCOUNT = '" & dT.strGoodDies & "',FAILBINCOUNT = '" & dT.strBadDies & "', QTECH_LASTUPDATE_BY = '" & gUserName & "',QTECH_LASTUPDATE_DATE = sysdate where filename = '" & strNewID & "' and substrateid = '" & dT.strWaferID & "' "
AddSql (strSql)
strSql = "update [ERPBASE].[dbo].[tblmappingData] set customershortname = '" & dT.strCusCode & "',productid = '" & dT.strPRODUCTID & "', PASSBINCOUNT = '" & dT.strGoodDies & "', FAILBINCOUNT = '" & dT.strBadDies & "', QTECH_LASTUPDATE_BY = '" & gUserName & "',QTECH_LASTUPDATE_DATE = GETDATE() where filename = '" & strNewID & "' and substrateid = '" & dT.strWaferID & "' "
AddSql2 (strSql)
strSql = "update customeroitbl_test set PO_NUM = '" & dT.strpo & "',MPN_DESC = '" & dT.strCUSDEVICE & "', CUSTOMERSHORTNAME = '" & dT.strCusCode & "',test_mtrl_desc = '" & dT.strJobID & "',imager_customer_rev= '" & dT.strVERSION & "',QTECH_LASTUPDATE_BY = '" & gUserName & "', QTECH_LASTUPDATE_DATE = sysdate  where id = '" & strNewID & "'                            "
AddSql (strSql)
strSql = "update [ERPBASE].[dbo].[tblCustomerOI] set PO_NUM = '" & dT.strpo & "',MPN_DESC = '" & dT.strCUSDEVICE & "', CUSTOMERSHORTNAME = '" & dT.strCusCode & "',test_mtrl_desc = '" & dT.strJobID & "',imager_customer_rev= '" & dT.strVERSION & "',QTECH_LASTUPDATE_BY = '" & gUserName & "', QTECH_LASTUPDATE_DATE = GETDATE()  where id = '" & strNewID & "'                            "
AddSql2 (strSql)

End Sub

Private Function SplitWOFileName(strKeyID As String, strWaferID As String) As String
Dim strNewID  As String
Dim strNewSeq As String

strNewID = GetMaxID()
AddSql ("update mappingdatatest set filename = '" & strNewID & "'where filename = '" & strKeyID & "' and substrateid = '" & strWaferID & "'")
AddSql2 ("update [ERPBASE].[dbo].[tblmappingData]  set filename = '" & strNewID & "'where filename = '" & strKeyID & "' and substrateid = '" & strWaferID & "'")
AddSql ("update customeroitbl_test_bak set id = " & strNewID & " where id = " & strKeyID & " ")
AddSql ("insert into customeroitbl_test select distinct * from customeroitbl_test_bak where id = " & strNewID & " ")
AddSql2 ("update ERPBASE.dbo.tblCustomerOI_TEMP set id = " & strNewID & " where id = " & strKeyID & " ")
AddSql2 ("insert into ERPBASE.dbo.tblCustomerOI select distinct * from ERPBASE.dbo.tblCustomerOI_TEMP where id = " & strNewID & " ")
SplitWOFileName = strNewID

End Function

Private Sub DelWaferID(strKeyID As String, strWaferID As String)
Dim strSql As String, rs As New ADODB.Recordset

strSql = "delete from mappingdatatest where filename = '" & strKeyID & "' and substrateid = '" & strWaferID & "' "
AddSql (strSql)
strSql = "delete from erpbase..tblmappingData where filename = '" & strKeyID & "' and substrateid = '" & strWaferID & "' "
AddSql2 (strSql)
If txtCusCode.text = "37" Then
    strSql = "delete from mappingdata37po where substrateid = '" & strWaferID & "' "
    AddSql (strSql)
    strSql = "delete from erpbase..tblmappingData where substrateid = '" & strWaferID & "' "
    AddSql2 (strSql)

End If

'-------------------------------------------------------------------------------
strSql = "select * from mappingdatatest a where a.filename = '" & strKeyID & "' "
If rs.State = adStateOpen Then rs.Close
rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If rs.EOF Then
    strSql = "delete from customeroitbl_test where id = '" & strKeyID & "' "
    AddSql (strSql)
    strSql = "delete from erpbase..tblCustomerOI where id = '" & strKeyID & "' "
    AddSql2 (strSql)

End If

rs.Close

End Sub

Private Function ExportExcel(dT As tyWO) As Boolean

On Error GoTo Ert

Dim xlsApp     As Excel.Application
Dim xlsBook    As Excel.Workbook
Dim xlsSheet   As Excel.Worksheet
Dim i          As Long
Dim j          As Long
Dim iCnt       As Integer
Dim strFileSeq As String, strPartName As String
Dim rs         As New ADODB.Recordset

ExportExcel = False
Set rs.ActiveConnection = OraConnect
rs.Source = "select row_number() over(ORDER BY  bb.lotid,bb.substrateid) as ���,case bb.substratetype when 'A' then '��˰' when 'B' then '�Ǳ�˰' else 'δ֪' end as �Ƿ�˰, bb.customershortname as �ͻ�����, " & _
   "       aa.Fab_conv_id as FAB����,aa.mpn_desc as �ͻ�����,cc.residual as NPI������Ա, " & _
   "       aa.mtrl_num as ���ڻ���, " & _
   "       aa.po_num as PO_NUM, " & _
   "       bb.lotid as LOT_ID, " & _
   "       bb.wafer_id as WAFER_NO, " & _
   "       bb.substrateid as WAFERID, " & _
   "       bb.passbincount as GOOD_DIES, " & _
   "       bb.failbincount as NG_DIES, " & _
   "       bb.passbincount + bb.failbincount as GROSS_DIES, " & _
   "       bb.productid as �����, " & _
   "       aa.Imager_Customer_Rev as ��������, bb.qtech_created_by as �ϴ���Ա,bb.qtech_created_date as �ϴ�ʱ��,  bb.qtech_lastupdate_by as ������Ա, bb.qtech_lastupdate_date as ����ʱ�� " & _
   "  from customeroitbl_test aa " & _
   "  left join tbltsvnpiproduct cc on cc.customerptno1 = aa.mpn_desc  and  cc.qtechptno = aa.mtrl_num  and cc.customershortname = aa.customershortname and cc.residual is not null " & _
   " inner join mappingdatatest bb " & _
   "    on to_char(aa.id) = bb.filename " & _
   "   and aa.wafer_visual_inspect = '" & gUpID & "' and aa.customershortname = '" & dT.CUSTOMER_CODE & "' " & _
   "   group by  bb.customershortname,cc.residual,aa.mtrl_num,aa.Fab_conv_id, aa.mpn_desc,aa.po_num,bb.lotid,bb.wafer_id,bb.substrateid,bb.passbincount,bb.failbincount,bb.productid,aa.Imager_Customer_Rev ,bb.substratetype,bb.qtech_created_by,bb.qtech_created_date,bb.qtech_lastupdate_by,bb.qtech_lastupdate_date "
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount = 0 Then
    MsgBox "��ѯ����������Ϣ, �˴��ϴ�ʧ��, ������ȷ��,�ٴ��ϴ�", vbCritical, "����"
    Exit Function

End If

iCnt = rs.RecordCount
Set xlsApp = CreateObject("Excel.Application")
Set xlsBook = xlsApp.Workbooks.Add
Set xlsSheet = xlsBook.Worksheets(1)
xlsSheet.name = "WO"

With xlsApp
    .Rows(1).Font.Bold = True

End With

For j = 1 To rs.Fields.count
    xlsSheet.Cells(1, j) = ("" & rs(j - 1).name)
Next
rs.MoveFirst

For i = 2 To rs.RecordCount + 1
    For j = 1 To rs.Fields.count
        If j = 9 Or j = 11 Then
            If Left(rs(j - 1).Value, 1) = "0" Then
                xlsSheet.Cells(i, j) = ("'" & rs(j - 1).Value)
            Else
                xlsSheet.Cells(i, j) = ("" & rs(j - 1).Value)

            End If

        Else
            xlsSheet.Cells(i, j) = ("" & rs(j - 1).Value)

        End If

    Next j

    rs.MoveNext
Next i

rs.Close
'--------------------
If dT.CUSTOMER_CODE = "68" Or dT.CUSTOMER_CODE = "HK006" Then
    rs.Source = "select row_number() over(ORDER BY  bb.lotid) as ���, aa.mpn_desc as �ͻ�Ʒ��,aa.Fab_conv_id as �ͻ�LOTNO,bb.customershortname as ��������, aa.mtrl_num as ����Ʒ��,  bb.lotid AS  ����LOTNO,count(bb.wafer_id) AS ���� " & " from customeroitbl_test aa inner join mappingdatatest bb  on to_char(aa.id) = bb.filename " & " and aa.wafer_visual_inspect = '" & gUpID & "' and aa.customershortname = '" & dT.CUSTOMER_CODE & "' " & " group by  bb.customershortname,aa.Fab_conv_id, aa.mpn_desc,bb.lotid,aa.mtrl_num "
Else
    rs.Source = "select row_number() over(ORDER BY  bb.lotid) as ���, '' AS �Ϻ�,  bb.customershortname as �ͻ� ,aa.mtrl_num as �ͺ�,  bb.lotid AS  LOT,count(distinct bb.wafer_id) AS ����    " & " from customeroitbl_test aa inner join mappingdatatest bb on to_char(aa.id) = bb.filename  " & " and aa.wafer_visual_inspect = '" & gUpID & "' and aa.customershortname = '" & dT.CUSTOMER_CODE & "' " & " group by  bb.customershortname,bb.lotid,aa.mtrl_num "

End If

rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
If xlsBook.Worksheets.count = 1 Then
    xlsBook.Worksheets.Add after:=xlsBook.Worksheets(1)

End If

Set xlsSheet = xlsBook.Worksheets(2)
xlsSheet.name = "��ǩ"

With xlsApp
    .Rows(1).Font.Bold = True

End With

For j = 1 To rs.Fields.count
    xlsSheet.Cells(1, j) = ("" & rs(j - 1).name)
Next
rs.MoveFirst
If dT.CUSTOMER_CODE = "68" Or dT.CUSTOMER_CODE = "HK006" Then

    For i = 2 To rs.RecordCount + 1
        For j = 1 To rs.Fields.count
            xlsSheet.Cells(i, j) = ("" & rs(j - 1).Value)
        Next j

        rs.MoveNext
    Next i

Else

    For i = 2 To rs.RecordCount + 1
        For j = 1 To rs.Fields.count
            If j = 2 Then '�Ϻ�
                'Npi���ձ�һ�����ڻ����ж�Ӧ�����Բ�Ϻŵ��������Ӧ���ʱ��������ֵ
                If Get_OracleCnt("select distinct MARKETLASTUPDATE_BY from tbltsvnpiproduct where CUSTOMERSHORTNAME='" & rs(2).Value & "' and  qtechptno='" & rs(3).Value & "'") = 1 Then
                    xlsSheet.Cells(i, j) = ("" & Get_OracleStr("select distinct MARKETLASTUPDATE_BY from tbltsvnpiproduct where CUSTOMERSHORTNAME='" & rs(2).Value & "' and  qtechptno='" & rs(3).Value & "'"))
                Else
                    xlsSheet.Cells(i, j) = ""

                End If

            Else
                xlsSheet.Cells(i, j) = ("" & rs(j - 1).Value)

            End If

        Next j

        rs.MoveNext
    Next i

End If

rs.Close
Set rs = Nothing
'-------------------
xlsBook.Worksheets(1).Activate
xlsApp.Visible = True
Dim strWOPath As String

strWOPath = "\\10.160.1.84\open\FileServer\WO\���ϴ�\" & dT.HT_DEVICE
If Dir(strWOPath, vbDirectory) = "" Then
    MkDir strWOPath

End If

Call CopyFileToFtp(txtFilePath.text, strWOPath & "\")
strFileName = strWOPath & "\" & dT.CUSTOMER_CODE & "_" & dT.HT_DEVICE & "_" & iCnt & "Ƭ" & "-" & Format(Now, "YYYYMMDD-HHMMSS") & ".xlsx"
xlsBook.SaveAs strFileName
Set xlsApp = Nothing
ExportExcel = True
Exit Function
Ert:
If Not (xlsApp Is Nothing) Then
    Set xlsApp = Nothing

End If

End Function

Private Function ExportExcel_37PO(dT As String) As Boolean

On Error GoTo Ert

Dim xlsApp      As Excel.Application
Dim xlsBook     As Excel.Workbook
Dim xlsSheet    As Excel.Worksheet
Dim i           As Long
Dim j           As Long
Dim iCnt        As Integer
Dim strFileSeq  As String
Dim strPartName As String

ExportExcel_37PO = False
Set xlsApp = CreateObject("Excel.Application")
Set xlsBook = xlsApp.Workbooks.Add
Set xlsSheet = xlsBook.Worksheets(1)

With xlsApp
    .Rows(1).Font.Bold = True

End With

xlsSheet.Cells(1, 1) = "�ͻ�����"
xlsSheet.Cells(1, 2) = "PO"
xlsSheet.Cells(1, 3) = "JOBID"
xlsSheet.Cells(1, 4) = "LOTID"
xlsSheet.Cells(1, 5) = "WAFERID"
xlsSheet.Cells(1, 6) = "Ƭ��"
xlsSheet.Cells(1, 7) = "DIE����"
xlsSheet.Cells(1, 8) = "�ͻ�����"
xlsSheet.Cells(1, 9) = "PRODUCTION ORDER"

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i

        For j = 1 To 9
            .Col = j
            xlsSheet.Cells(i + 1, j) = Trim$(.text)
        Next
    Next

End With

xlsApp.Visible = True
Dim strWOPath As String

strWOPath = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\���ϴ�\37"
If Dir(strWOPath, vbDirectory) = "" Then
    MkDir strWOPath

End If

Select Case dT

    Case "AS"
        strWOPath = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\���ϴ�\37\һ��PO"
        strFileName = strWOPath & "\" & "37һ��PO" & "-" & Format(Now, "YYYYMMDD-HHMMSS") & ".xlsx"

    Case "TS"
        strWOPath = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\���ϴ�\37\����PO"
        strFileName = strWOPath & "\" & "37����PO" & "-" & Format(Now, "YYYYMMDD-HHMMSS") & ".xlsx"

End Select

If Dir(strWOPath, vbDirectory) = "" Then
    MkDir strWOPath

End If

xlsBook.SaveAs strFileName
Set xlsApp = Nothing
ExportExcel_37PO = True
Exit Function
Ert:
If Not (xlsApp Is Nothing) Then
    Set xlsApp = Nothing

End If

End Function

Private Sub SentMesToPMC(dT As tyWO)
'�����ʼ����ƻ���
Dim dirtemp             As String
Dim strTemp             As String
Dim i                   As Integer
Dim strBand             As String
Dim rs                  As New ADODB.Recordset
Dim strKHJZ             As String
Dim strCNJZ             As String
Dim strPecs             As String
Dim strDies             As String
Dim strMailRecipients   As String
Dim strMailRecipientsCC As String
Dim strMailSubject      As String
Dim strMailBody         As String
Dim strMailAttachment   As String
Dim strSql              As String

If bBonded = True Then
    strBand = "��˰"
Else
    strBand = "�Ǳ�˰"

End If

i = 0
dirtemp = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\�ʼ�����\SentTo_Upload.cfg"
strSql = " select  t2.mpn_desc, t2.mtrl_num,count(1) Pecs ,sum(t1.passbincount + t1.failbincount) Dies from mappingdatatest t1 " & "  inner join customeroitbl_test t2 on to_char(t2.id) = t1.filename " & "  where t2.wafer_visual_inspect = '" & gUpID & "'  group by t2.mpn_desc, t2.mtrl_num"
strMailSubject = "���ϴ�" & dT.CUSTOMER_CODE & "��" & strBand & "����"
strMailBody = "���ڲ�" & strRealName & ",����:" & gUserName & "���ϴ�" & dT.CUSTOMER_CODE & "��" & strBand & "����" & vbCrLf
Set rs = Get_OracleRs(strSql)

Do While Not rs.EOF
    strKHJZ = Trim("" & rs!MPN_DESC)
    strCNJZ = Trim("" & rs!mtrl_Num)
    strPecs = Trim$("" & rs!Pecs)
    strDies = Trim$("" & rs!Dies)
    strMailSubject = strMailSubject & "," & strKHJZ & "-" & strCNJZ
    strMailBody = strMailBody & "�ͻ�����:" & strKHJZ & "- ���ڻ���:" & strCNJZ & "-" & strPecs & "Ƭ" & "-" & strDies & "��" & vbCrLf
    rs.MoveNext
Loop
strMailSubject = strMailSubject & "��ע�����"
strMailBody = strMailBody & "��ϸ������" & vbCrLf
strMailBody = strMailBody & txtMsg.text
If gUserName = "07885" Then
    strMailRecipients = Get_OracleStr("select a.RECV_USER_TO from ERP_EMAIL_RECV a WHERE a.EMAIL_TYPE = 'WO_UPLOAD_RECV_TEST' ")
    strMailRecipientsCC = Get_OracleStr("select a.RECV_USER_CC from ERP_EMAIL_RECV a WHERE a.EMAIL_TYPE = 'WO_UPLOAD_RECV_TEST' ")
Else
    strMailRecipients = Get_OracleStr("select a.RECV_USER_TO from ERP_EMAIL_RECV a WHERE a.EMAIL_TYPE = 'WO_UPLOAD_RECV' ")
    strMailRecipientsCC = Get_OracleStr("select a.RECV_USER_CC from ERP_EMAIL_RECV a WHERE a.EMAIL_TYPE = 'WO_UPLOAD_RECV' ")

End If

If Dir(strFileName) <> "" Then
    strMailAttachment = Replace(Replace$(strFileName, "\\10.160.1.84\open\FileServer\WO\���ϴ�", "\svn\OpenFolder\FileServer\WO\���ϴ�"), "\", "/")
    Call SentEml(strMailRecipients, strMailRecipientsCC, strMailSubject, txtMsg.text, strMailAttachment)
Else
    Call SentEml(strMailRecipients, strMailRecipientsCC, strMailSubject, txtMsg.text, "")

End If

MsgBox "�ʼ��ѷ���", vbInformation, "��ʾ"

End Sub

'���÷����ʼ�API
Public Sub SentEml(mailTo As String, _
                   mailCc As String, _
                   mailTitle As String, _
                   mailBody As String, _
                   filename As String)
Dim strSql As String

strSql = " insert into tbl_Sent_Mail(MAIL_ID,SENT_FROM,SENT_TIME,SENT_TO,SENT_TO_CC,MAIL_TITLE,MAIL_BODY,MAIL_ATTACHMENT,FLAG,MAIL_KEY) " & " values(MAILID_SEQ.Nextval,'" & gUserRealName & "',sysdate,'" & mailTo & "','" & mailCc & "','" & mailTitle & "','" & mailBody & "','" & filename & "','0','" & gUpID & "') "
AddSql (strSql)

End Sub

Private Sub SentMesToPMC_37PO(strtype As String)
'�����ʼ����ƻ���
Dim strSentTo(100) As String
Dim strSentCC(20)  As String
Dim strSentTitle   As String
Dim strSentText    As String
Dim dirtemp        As String
Dim strTemp        As String
Dim i              As Integer
Dim rs             As New ADODB.Recordset
Dim strSql         As String

If strtype = "AS" Then
    strtype = "һ��"
Else
    strtype = "����"

End If

If gUserName = "07885" Then

    '   Exit Sub
End If

i = 0
dirtemp = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\�ʼ�����\SentTo_Upload_37PO.cfg"
strSentTitle = "37" & strtype & "PO�ϴ��ɹ�"
strSql = "select distinct '���ڻ���:' || d.mtrl_num || ' ' || case c.substratetype  when 'A' then '��˰'  when 'B' then '�Ǳ�˰' else  'δ֪' end || count(1) || 'Ƭ'  from mappingdatatest a  inner join customeroitbl_test b " & "    on to_char(b.id) = a.filename " & " inner join mappingdatatest c " & "    on c.substrateid = replace(a.substrateid, '+', '') " & " inner join customeroitbl_test d " & "    on to_char(d.id) = c.filename " & " where a.micronlotid = '" & gUpID & "' " & " group by d.mtrl_num, c.substratetype, d.jobno "
Set rs = Get_OracleRs(strSql)
If Not rs.EOF Then

    Do While Not rs.EOF
        strSentText = strSentText & rs(0).Value & vbCrLf
        rs.MoveNext
    Loop

End If

strSentText = strSentText & vbCrLf
strSentText = strSentText & txtMsg.text & vbCrLf & "�������: "
Open dirtemp For Input As #1

Do While Not EOF(1)
    Line Input #1, strTemp
    strSentTo(i) = Trim$(strTemp)
    i = i + 1
Loop
Close #1
If SentMes(strSentTitle, strSentText, strSentTo, strFileName, strSentCC) = True Then
    MsgBox "�ʼ��ѷ���", vbInformation, Me.Caption
Else
    MsgBox "�ʼ�����ʧ��", vbCritical, Me.Caption

End If

End Sub

Private Sub SentMesToPMC_DEL()
'�����ʼ����ƻ���
Dim strSentTo(20) As String
Dim strSentCC(10) As String
Dim strSentTitle  As String
Dim strSentText   As String
Dim strCusCode    As String, strCusDev As String
Dim dirtemp       As String
Dim strTemp       As String
Dim i             As Integer
Dim strBand       As String

If gUserName = "07885" Then
    Exit Sub

End If

If bBonded = True Then
    strBand = "��˰"
Else
    strBand = "�Ǳ�˰"

End If

i = 0
dirtemp = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\�ʼ�����\SentTo_Mod_Del.cfg"
Open dirtemp For Input As #1

Do While Not EOF(1)
    Line Input #1, strTemp
    strSentTo(i) = Trim$(strTemp)
    i = i + 1
Loop
Close #1
strCusCode = txtCusCode.text
strCusDev = txtCusDev.text
strSentTitle = "��ɾ������," & strBand & ",�ͻ�����:" & strCusCode & ",�ͻ�����:" & strCusDev & ", ��ע�����"
strSentText = "���ڲ�" & strRealName & ",����:" & gUserName & "��ɾ������," & strBand & ",�ͻ�����:" & strCusCode & "�ͻ�����:" & strCusDev & ",��ϸ������" & vbCrLf
strSentText = strSentText & txtMsg2.text
If SentMes(strSentTitle, strSentText, strSentTo, strFileName, strSentCC) = True Then
    MsgBox "�ʼ��ѷ���", vbInformation, Me.Caption
Else
    MsgBox "�ʼ�����ʧ��", vbCritical, Me.Caption

End If

End Sub

Private Sub SentMesToPMC_MOD()
'�����ʼ����ƻ���
Dim strSentTo(20) As String
Dim strSentCC(10) As String
Dim strSentTitle  As String
Dim strSentText   As String, strPartName As String
Dim xlsApp        As Excel.Application
Dim xlsBook       As Excel.Workbook
Dim xlsSheet      As Excel.Worksheet
Dim rs            As New ADODB.Recordset
Dim i             As Integer, j As Integer, k As Integer
Dim dirtemp       As String
Dim strTemp       As String
Dim strBand       As String

If gUserName = "07885" Then
    Exit Sub

End If

If bBonded = True Then
    strBand = "��˰"
Else
    strBand = "�Ǳ�˰"

End If

i = 0
dirtemp = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\�ʼ�����\SentTo_Mod_Del.cfg"
Open dirtemp For Input As #1

Do While Not EOF(1)
    Line Input #1, strTemp
    strSentTo(i) = Trim$(strTemp)
    i = i + 1
Loop
Close #1

On Error GoTo Ert

Set rs.ActiveConnection = OraConnect
rs.Source = "select '�޸�ǰ' as ״̬,b.mpn_desc as �ͻ�����,a.substratetype as ��˰,a.customershortname as �ͻ�����,b.po_num as PO_NUM,a.lotid as LOTID,a.substrateid as WAFERID, " & "a.passbincount as GOODDIES,a.failbincount as BADDIES,a.productid as �����,b.Imager_Customer_Rev as �������� from mappingdatatest_bak a, customeroitbl_test_bak b where a.filename = to_char(b.id) and a.remark = '" & gBackID & "' " & "union select '�޸ĺ�' as ״̬,b.mpn_desc as �ͻ�����,a.substratetype as ��˰,a.customershortname as �ͻ�����,b.po_num as PO_NUM,a.lotid as LOTID,a.substrateid as WAFERID, " & "a.passbincount as GOODDIES,a.failbincount as BADDIES,a.productid as �����,b.Imager_Customer_Rev as �������� from mappingdatatest a, customeroitbl_test b where a.filename = to_char(b.id) and a.remark = '" & gBackID & "' "
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount = 0 Then
    MsgBox "��ѯ����������Ϣ, ��ȷ��", vbCritical, "����"
    Exit Sub

End If

Set xlsApp = CreateObject("Excel.Application")
Set xlsBook = xlsApp.Workbooks.Add
Set xlsSheet = xlsBook.Worksheets(1)

With xlsApp
    .Rows(1).Font.Bold = True

End With

For j = 1 To rs.Fields.count
    xlsSheet.Cells(1, j) = ("" & rs(j - 1).name)
Next
rs.MoveFirst

For i = 2 To rs.RecordCount + 1
    For j = 1 To rs.Fields.count
        xlsSheet.Cells(i, j) = ("" & rs(j - 1).Value)
    Next j

    rs.MoveNext
Next i

rs.Close
Set rs = Nothing
xlsApp.Visible = True
Dim strWOPath  As String
Dim strCusCode As String, strCusDev As String

strCusCode = txtCusCode.text
strCusDev = txtCusDev.text
strWOPath = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\���޸�\" & strCusCode
If Dir(strWOPath, vbDirectory) = "" Then
    MkDir strWOPath

End If

strWOPath = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\���޸�\" & strCusCode & "\" & Replace(strCusDev, "/", "")
If Dir(strWOPath, vbDirectory) = "" Then
    MkDir strWOPath

End If

strFileName = strWOPath & "\" & Format(Now, "YYYY-MMDD-HH-MM-SS") & ".xlsx"
xlsBook.SaveAs strFileName
Set xlsApp = Nothing
strSentTitle = "���޸Ķ���," & strBand & ",�ͻ�����:" & strCusCode & ",�ͻ�����:" & strCusDev & ", ��ע�����"
strSentText = "���ڲ�" & strRealName & ",����:" & gUserName & "���޸Ķ���," & strBand & ",�ͻ�����:" & strCusCode & ",�ͻ�����:" & strCusDev & ",��ϸ������" & vbCrLf
strSentText = strSentText & txtMsg2.text
If SentMes(strSentTitle, strSentText, strSentTo, strFileName, strSentCC) = True Then
    MsgBox "�ʼ��ѷ���", vbInformation, Me.Caption
Else
    MsgBox "�ʼ�����ʧ��", vbCritical, Me.Caption

End If

Exit Sub
Ert:
If Not (xlsApp Is Nothing) Then
    Set xlsApp = Nothing

End If

End Sub




