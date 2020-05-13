VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_zsgc 
   Caption         =   "机种信息维护"
   ClientHeight    =   10875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10875
   ScaleWidth      =   16080
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTextUpdate 
      Height          =   495
      Left            =   4680
      TabIndex        =   108
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H000080FF&
      Caption         =   "..."
      Height          =   615
      Left            =   4080
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   120
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   15360
      Top             =   600
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
            Picture         =   "frm_zsgc.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_zsgc.frx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_zsgc.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_zsgc.frx":24F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_zsgc.frx":3148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_zsgc.frx":3D9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtQuery2 
      BackColor       =   &H00FFC0FF&
      Height          =   270
      Left            =   12120
      TabIndex        =   106
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtQuery1 
      BackColor       =   &H00FFC0FF&
      Height          =   270
      Left            =   9120
      TabIndex        =   104
      Top             =   960
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "EU010机种信息"
      Height          =   4095
      Left            =   6600
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox TextPROVENANCE 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   4680
         TabIndex        =   110
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtPackage 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   1320
         TabIndex        =   102
         Top             =   2543
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "返回"
         Height          =   495
         Left            =   2880
         TabIndex        =   37
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "保存"
         Height          =   495
         Left            =   1320
         TabIndex        =   36
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   4680
         TabIndex        =   35
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   1320
         TabIndex        =   33
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   4680
         TabIndex        =   31
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   1320
         TabIndex        =   29
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   4680
         TabIndex        =   27
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   1320
         TabIndex        =   25
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblPROVENANCE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROVENANCE"
         Height          =   180
         Left            =   3480
         TabIndex        =   109
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PACKAGE"
         Height          =   180
         Left            =   600
         TabIndex        =   101
         Top             =   2640
         Width           =   630
      End
      Begin VB.Label Label15 
         Caption         =   "ORIG"
         Height          =   255
         Left            =   4080
         TabIndex        =   34
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "PMC"
         Height          =   255
         Left            =   840
         TabIndex        =   32
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "PRODUCT_12NC"
         Height          =   255
         Left            =   3360
         TabIndex        =   30
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "DEVICE_NAME"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "MARKING_CODE"
         Height          =   375
         Left            =   3360
         TabIndex        =   26
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "CUST_DEVICE"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "37阴极线"
      Height          =   5175
      Left            =   6480
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CommandButton Command2 
         Caption         =   "返回"
         Height          =   495
         Left            =   3840
         TabIndex        =   22
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "保存"
         Height          =   495
         Left            =   2160
         TabIndex        =   21
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   1080
         TabIndex        =   20
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1080
         TabIndex        =   16
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   4440
         TabIndex        =   14
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   4440
         TabIndex        =   10
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "日期格式为xxxx/xx/xx或xxxx-xx-xx"
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label Label8 
         Caption         =   "SEQ"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "STATUS"
         Height          =   375
         Left            =   3360
         TabIndex        =   17
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "TIMESTAMP"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "CUSTOMER"
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "CODE"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "BLINE"
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "DEVICE"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "AAMPN2"
      Height          =   6495
      Left            =   4440
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   12015
      Begin VB.CommandButton Command8 
         Caption         =   "返回"
         Height          =   495
         Left            =   7560
         TabIndex        =   98
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "保存"
         Height          =   495
         Left            =   5640
         TabIndex        =   97
         Top             =   5520
         Width           =   975
      End
      Begin VB.TextBox Text41 
         Height          =   375
         Left            =   1800
         TabIndex        =   96
         Top             =   5400
         Width           =   1695
      End
      Begin VB.TextBox Text40 
         Height          =   375
         Left            =   9720
         TabIndex        =   94
         Top             =   4560
         Width           =   1815
      End
      Begin VB.TextBox Text39 
         Height          =   375
         Left            =   5640
         TabIndex        =   92
         Top             =   4560
         Width           =   1815
      End
      Begin VB.TextBox Text38 
         Height          =   375
         Left            =   1800
         TabIndex        =   90
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox Text37 
         Height          =   375
         Left            =   9720
         TabIndex        =   88
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox Text36 
         Height          =   375
         Left            =   5640
         TabIndex        =   86
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox Text35 
         Height          =   375
         Left            =   1800
         TabIndex        =   84
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox Text34 
         Height          =   375
         Left            =   9720
         TabIndex        =   82
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox Text33 
         Height          =   375
         Left            =   5640
         TabIndex        =   80
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox Text32 
         Height          =   375
         Left            =   1800
         TabIndex        =   78
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Text31 
         Height          =   375
         Left            =   9720
         TabIndex        =   76
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text30 
         Height          =   375
         Left            =   5640
         TabIndex        =   74
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text29 
         Height          =   375
         Left            =   1800
         TabIndex        =   72
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text28 
         Height          =   375
         Left            =   9720
         TabIndex        =   70
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text27 
         Height          =   375
         Left            =   5640
         TabIndex        =   68
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text26 
         Height          =   375
         Left            =   1800
         TabIndex        =   66
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text25 
         Height          =   375
         Left            =   9720
         TabIndex        =   64
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text24 
         Height          =   375
         Left            =   5640
         TabIndex        =   62
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text23 
         Height          =   375
         Left            =   1800
         TabIndex        =   60
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label45 
         Caption         =   "日期格式为xxxx/xx/xx或xxxx-xx-xx"
         Height          =   255
         Left            =   1800
         TabIndex        =   99
         Top             =   5040
         Width           =   3135
      End
      Begin VB.Label Label44 
         Caption         =   "MARKINGCODEFIRST"
         Height          =   255
         Left            =   240
         TabIndex        =   95
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label43 
         Caption         =   "QTECH_LASTUPDATE_DATE"
         Height          =   255
         Left            =   7800
         TabIndex        =   93
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label Label42 
         Caption         =   "QTECH_LASTUPDATE_BY"
         Height          =   255
         Left            =   3720
         TabIndex        =   91
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label41 
         Caption         =   "QTECH_CREATED_DATE"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label40 
         Caption         =   "QTECH_CREATED_BY"
         Height          =   255
         Left            =   8160
         TabIndex        =   87
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label39 
         Caption         =   "FLAG"
         Height          =   255
         Left            =   4680
         TabIndex        =   85
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label38 
         Caption         =   "UL_LISTED_FLAG"
         Height          =   255
         Left            =   360
         TabIndex        =   83
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label37 
         Caption         =   "PKG_GRP_CD"
         Height          =   255
         Left            =   8640
         TabIndex        =   81
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "PACKAGING_TYPE"
         Height          =   255
         Left            =   4560
         TabIndex        =   79
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label35 
         Caption         =   "MPQ_QTY"
         Height          =   255
         Left            =   840
         TabIndex        =   77
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label34 
         Caption         =   "PBF_DIE_ATTACH"
         Height          =   255
         Left            =   8280
         TabIndex        =   75
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label33 
         Caption         =   "HALIDE_FREE"
         Height          =   255
         Left            =   4320
         TabIndex        =   73
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label32 
         Caption         =   "TEMP"
         Height          =   255
         Left            =   960
         TabIndex        =   71
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label31 
         Caption         =   "MSL"
         Height          =   255
         Left            =   8880
         TabIndex        =   69
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label30 
         Caption         =   "ECAT"
         Height          =   255
         Left            =   4680
         TabIndex        =   67
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "LEAD_FREE"
         Height          =   255
         Left            =   720
         TabIndex        =   65
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "PART"
         Height          =   255
         Left            =   8880
         TabIndex        =   63
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "LOC"
         Height          =   255
         Left            =   4800
         TabIndex        =   61
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label26 
         Caption         =   "ID"
         Height          =   375
         Left            =   1080
         TabIndex        =   59
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "AAMPN1"
      Height          =   4935
      Left            =   6240
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CommandButton Command6 
         Caption         =   "返回"
         Height          =   495
         Left            =   3840
         TabIndex        =   57
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "保存"
         Height          =   495
         Left            =   2160
         TabIndex        =   56
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox Text22 
         Height          =   375
         Left            =   2040
         TabIndex        =   55
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox Text21 
         Height          =   375
         Left            =   5760
         TabIndex        =   53
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text20 
         Height          =   375
         Left            =   2040
         TabIndex        =   51
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox Text19 
         Height          =   375
         Left            =   5760
         TabIndex        =   49
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text18 
         Height          =   375
         Left            =   2040
         TabIndex        =   47
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text17 
         Height          =   375
         Left            =   5760
         TabIndex        =   45
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   2040
         TabIndex        =   43
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text15 
         Height          =   375
         Left            =   5760
         TabIndex        =   41
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   2040
         TabIndex        =   39
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label25 
         Caption         =   "日期格式为xxxx/xx/xx或xxxx-xx-xx"
         Height          =   255
         Left            =   960
         TabIndex        =   58
         Top             =   2880
         Width           =   3135
      End
      Begin VB.Label Label24 
         Caption         =   "QTECH_LASTUPDATE_DATE"
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label23 
         Caption         =   "QTECH_LASTUPDATE_BY"
         Height          =   255
         Left            =   3960
         TabIndex        =   52
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "QTECH_CREATED_DATE"
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "QTECH_CREATED_BY"
         Height          =   375
         Left            =   4080
         TabIndex        =   48
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "FLAG"
         Height          =   255
         Left            =   1080
         TabIndex        =   46
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "TYPENAME"
         Height          =   255
         Left            =   4440
         TabIndex        =   44
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "NUMBEROFHOURS"
         Height          =   255
         Left            =   720
         TabIndex        =   42
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "MS_LEVEL"
         Height          =   255
         Left            =   4440
         TabIndex        =   40
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "ID"
         Height          =   255
         Left            =   1200
         TabIndex        =   38
         Top             =   480
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   900
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   1588
      ButtonWidth     =   1032
      ButtonHeight    =   1482
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查询"
            Key             =   "s"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "i"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "u"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "删除"
            Key             =   "d"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "上传"
            Key             =   "c"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "t"
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8760
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_zsgc.frx":49EC
               Key             =   "s"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_zsgc.frx":563E
               Key             =   "i"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_zsgc.frx":6290
               Key             =   "u"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_zsgc.frx":6EE2
               Key             =   "d"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_zsgc.frx":7B34
               Key             =   "t"
            EndProperty
         EndProperty
      End
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFC0FF&
      Height          =   300
      ItemData        =   "frm_zsgc.frx":8786
      Left            =   5400
      List            =   "frm_zsgc.frx":879C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   22095
      _Version        =   524288
      _ExtentX        =   38973
      _ExtentY        =   13573
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
      SpreadDesigner  =   "frm_zsgc.frx":87E3
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   15480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblQuery2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查询条件2"
      Height          =   180
      Left            =   11160
      TabIndex        =   105
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lblQuery1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查询条件1"
      Height          =   180
      Left            =   8160
      TabIndex        =   103
      Top             =   960
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "维护类型"
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
      Left            =   4440
      TabIndex        =   100
      Top             =   960
      Width           =   840
   End
End
Attribute VB_Name = "frm_zsgc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUpdate_Click()
On Error Resume Next
Dim FName
    '帅选文件
    CommonDialog2.Filter = "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx;*.xlsm;*.XLSM"
    CommonDialog2.ShowOpen
    '得到文件名
    FName = CommonDialog2.filename
    If FName <> "" Then
       txtTextUpdate.text = FName
    End If
If Combo1.text = "" Then
    MsgBox "请选择上传类型"
End If

End Sub

Private Sub Combo1_Click()

Select Case Combo1.ListIndex

    Case 0

    Case 1

    Case 2

    Case 3
        lblQuery1.Caption = "CUST_DEVICE"
        lblQuery2.Caption = "DEVICE_NAME"

    Case 4

End Select

End Sub

Private Sub Command6_Click()
Frame1.Visible = False

End Sub

Private Sub Command8_Click()
Frame2.Visible = False

End Sub

Private Sub Command2_Click()
Frame3.Visible = False

End Sub

Private Sub Command4_Click()
Frame4.Visible = False

End Sub

Private Sub ClearTxt()
Dim ctrl As Control

For Each ctrl In Me.Controls

    If TypeOf ctrl Is TextBox Then
        ctrl.text = ""

    End If

Next

End Sub

Private Function IsID(id As Long)
Dim rs     As New ADODB.Recordset
Dim strSql As String

Select Case Combo1.text

    Case "AAMPN1"
        strSql = "SELECT * FROM CUSTOMERMSLevelTBL WHERE  ID='" & id & "'  "

    Case "AAMPN2"
        strSql = "select * from CUSTOMERMPNAttributes WHERE ID='" & id & "'"

End Select

If rs.State = adStateOpen Then rs.Close
rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF Then
    IsID = rs.RecordCount

End If

End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "s"
        Query

    Case "i"
        Call Insert

    Case "u"
        Call Update

    Case "d"
        Delete
        
    Case "c"
        Call upload
        
    Case "t"
        Unload Me

    
End Select

End Sub

Private Sub Query()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Dim rs      As New ADODB.Recordset
Dim strsql1 As String
Dim strSql2 As String
Dim strSql3 As String
Dim strSql4 As String
Dim strSql5 As String

If Combo1.text = "" Then
    MsgBox "请选择你要查询的选项名称", vbInformation, "提示"
    Exit Sub

End If

fpSpread1.MaxRows = 0

Select Case Combo1.text

    Case "AAMPN1"
        strsql1 = "select ID,MS_LEVEL,NUMBEROFHOURS,TYPENAME,FLAG,QTECH_CREATED_BY,QTECH_CREATED_DATE,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE,'' √ from CUSTOMERMSLevelTBL "

    Case "AAMPN2"
        strsql1 = " select ID,LOC,PART,LEAD_FREE,ECAT,MSL,TEMP,HALIDE_FREE,PBF_DIE_ATTACH,MPQ_QTY,PACKAGING_TYPE,PKG_GRP_CD,UL_LISTED_FLAG,FLAG,QTECH_CREATED_BY,QTECH_CREATED_DATE,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE,MARKINGCODEFIRST,'' √ from CUSTOMERMPNAttributes ORDER BY ID  "

    Case "37阴极线"
        strsql1 = "select DEVICE,BLINE,CODE,CUSTOMER,TIMESTAMP,STATUS,SEQ,'' as 选择 from code37"

    Case "EU010机种信息"
        strsql1 = "select CUST_DEVICE,MARKING_CODE,DEVICE_NAME,PRODUCT_12NC,PMC,ORIG,PACKAGE,PROVENANCE,'' as 选择 from EU010_REFERENCE where 1 = 1 "
        If Trim(txtQuery1.text) <> "" Then
            strSql2 = " and CUST_DEVICE = '" & Trim(txtQuery1.text) & "'   "

        End If

        If Trim(txtQuery2.text) <> "" Then
            strSql3 = " and DEVICE_NAME = '" & Trim(txtQuery2.text) & "'   "
        End If

        strSql4 = " order by CUST_DEVICE,DEVICE_NAME"
        strSql5 = "and CUST_DEVICE not in('A1810E_AW87359_BUMP','A1810F_AW87359_BUMP','A1725C_BUMP','A1725CS_BUMP','A1802B_BUMP','A1908C_BUMP','A1810G_AW87369_BUMP','A1810G_AW87379_BUMP','A1725C6_BUMP')"

        strsql1 = strsql1 & strSql2 & strSql3 & strSql5 & strSql4
    Case "AC70机种信息"
        strsql1 = "select CUST_DEVICE,DEVICE_NAME,PACKAGE,'' as 选择 from EU010_REFERENCE where 1 = 1 "
        If Trim(txtQuery1.text) <> "" Then
            strSql2 = " and CUST_DEVICE = '" & Trim(txtQuery1.text) & "'   "
        End If
        
        If Trim(txtQuery2.text) <> "" Then
            strSql3 = " and DEVICE_NAME = '" & Trim(txtQuery2.text) & "'   "
        End If

        strSql5 = "and CUST_DEVICE in('A1810E_AW87359_BUMP','A1810F_AW87359_BUMP','A1725C_BUMP','A1725CS_BUMP','A1802B_BUMP','A1908C_BUMP','A1810G_AW87369_BUMP','A1810G_AW87379_BUMP','A1725C6_BUMP')"

        strSql4 = " order by CUST_DEVICE"
 
        strsql1 = strsql1 & strSql2 & strSql3 & strSql5 & strSql4
        
    Case "HD机种信息"
        strsql1 = "select CUST_DEVICE,DEVICE_NAME,ORIG,'' as 选择 from EU010_REFERENCE where 1 = 1 "
        If Trim(txtQuery1.text) <> "" Then
            strSql2 = " and CUST_DEVICE = '" & Trim(txtQuery1.text) & "'   "
        End If
        
        If Trim(txtQuery2.text) <> "" Then
            strSql3 = " and DEVICE_NAME = '" & Trim(txtQuery2.text) & "'   "
        End If

        strSql5 = "and CUST_DEVICE in('GH3100','GH3103')"

        strSql4 = " order by CUST_DEVICE"
 
        strsql1 = strsql1 & strSql2 & strSql3 & strSql5 & strSql4
        
        
    
End Select

If rs.State = adStateOpen Then rs.Close
rs.Open strsql1, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF Then
    Call ShowDesginc(rs)

End If

'Call ClearTxt
End Sub
Private Sub upload()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False

Dim strsql1 As String
Dim strSql2 As String
Dim strSql3 As String
Dim strSql4 As String

    If txtTextUpdate.text = "" Then
        MsgBox "请选择你要上传的文件", vbInformation, "提示"
        Exit Sub
    End If
    If Combo1.text = "" Then
        MsgBox "请选择你要上传类型", vbInformation, "提示"
        Exit Sub
    End If
Select Case Combo1.text

    Case "AAMPN1"
        'strsql1 = "select ID,MS_LEVEL,NUMBEROFHOURS,TYPENAME,FLAG,QTECH_CREATED_BY,QTECH_CREATED_DATE,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE,'' √ from CUSTOMERMSLevelTBL "

    Case "AAMPN2"
        'strsql1 = " select ID,LOC,PART,LEAD_FREE,ECAT,MSL,TEMP,HALIDE_FREE,PBF_DIE_ATTACH,MPQ_QTY,PACKAGING_TYPE,PKG_GRP_CD,UL_LISTED_FLAG,FLAG,QTECH_CREATED_BY,QTECH_CREATED_DATE,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE,MARKINGCODEFIRST,'' √ from CUSTOMERMPNAttributes ORDER BY ID  "

    Case "37阴极线"
        'strsql1 = "select DEVICE,BLINE,CODE,CUSTOMER,TIMESTAMP,STATUS,SEQ,'' √ from code37"

    Case "EU010机种信息"
        Call uoloadEU010
        
End Select

End Sub

Private Function uoloadEU010()

Dim i         As Integer
Dim ii        As String
Dim ierror  As String    '记录出错位置

Dim VBExcel   As Excel.Application
Dim xlBook    As Excel.Workbook
Dim xlSheet   As Excel.Worksheet
Dim lColsCnt  As Long
Dim lRowsCnt  As Long
Dim strSql    As String
Dim rs      As New ADODB.Recordset
Dim PACKAGE As String
Dim ORIG  As String
Dim PMC  As String
Dim cust_device As String
Dim CUST_DEVICES As String
Dim K12_NC  As String
Dim MARKING_CODE As String
Dim device  As String
Dim DEVICES As String
ierror = ""
CUST_DEVICES = ""    '查询导入结果所设置的变量
DEVICES = ""

Set VBExcel = CreateObject("excel.application")
VBExcel.Visible = False
Set xlBook = VBExcel.Workbooks.Open(txtTextUpdate.text)
Set xlSheet = xlBook.Worksheets(1)
lColsCnt = xlSheet.Range("C6").CurrentRegion.Columns.count '列
lRowsCnt = xlSheet.Range("C6").CurrentRegion.Rows.count    '行
lRowsCnt = lRowsCnt - 1
'If lColsCnt <> 8 Then
'    MsgBox "Excel中的列数:" & lColsCnt & "和设定的模版列数:" & fpSDetail.e_MCol & "不一致" & vbCrLf & "请确认Excel是否正确！", vbInformation, "提示"
'    GoTo EXITPRO
'    Exit Sub
'End If
'Mid("0123456789ACDEFHKLNPRSTUXYZ", (DatePart("WW", Now) + 1) \ 2, 1)

For i = 3 To lRowsCnt
        'ii = Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ", i, 1)
        ii = ColumnToNum(i)  '数字转列数方法
        PACKAGE = Trim(xlSheet.Range(ii & 6))
        ORIG = Trim(xlSheet.Range(ii & 8))
        PMC = Trim(xlSheet.Range(ii & 9))
        cust_device = Trim(xlSheet.Range(ii & 10))
        CUST_DEVICES = CUST_DEVICES & cust_device & "','"
        K12_NC = Trim(xlSheet.Range(ii & 11))
        MARKING_CODE = Trim(xlSheet.Range(ii & 12))
        device = Trim(xlSheet.Range(ii & 13))
        DEVICES = DEVICES & device & "','"
        
        strSql = "select * from EU010_REFERENCE where CUST_DEVICE = '" & cust_device & "' and  DEVICE_NAME = '" & device & "'"
        If rs.State = adStateOpen Then rs.Close
            rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
        If cust_device = "" Or device = "" Then
            MsgBox "CUST_DEVICE 和 DEVICE_NAME 不能为空"
            Exit Function
        End If
        
        If rs.RecordCount > 0 Then
            MsgBox "该数据已存在:CUST_DEVICE: " & cust_device & ", DEVICE_NAME :" & device
        Else
            strSql = "INSERT INTO EU010_REFERENCE(CUST_DEVICE,MARKING_CODE,DEVICE_NAME,PRODUCT_12NC,PMC,ORIG,PACKAGE)values('" & cust_device & "','" & MARKING_CODE & "','" & device & "','" & K12_NC & "','" & PMC & "','" & ORIG & "','" & PACKAGE & "')"
            If AddSql(strSql) > 0 Then
                MsgBox "新增成功", vbInformation, "提示"
            Else
                ierror = ierror + i + ","
                MsgBox "新增失败 在第" & i & "列", vbInformation, "提示 CUST_DEVICE: " & cust_device & ", DEVICE_NAME :" & device
            End If
        End If
        rs.Close
Next
        CUST_DEVICES = Mid(CUST_DEVICES, 1, Len(CUST_DEVICES) - 3)
        DEVICES = Mid(DEVICES, 1, Len(DEVICES) - 3)
If i = (lRowsCnt + 1) Then
    MsgBox "导入完毕"
Else
    MsgBox "从第" & ierror & "列开始未导入，导入失败"
End If

'最后查询导入结果
strSql = "select CUST_DEVICE,MARKING_CODE,DEVICE_NAME,PRODUCT_12NC,PMC,ORIG,PACKAGE,'' as 选择 from EU010_REFERENCE where CUST_DEVICE in ('" & CUST_DEVICES & "') and  DEVICE_NAME in ('" & DEVICES & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF Then
    Call ShowDesginc(rs)

End If


If Not VBExcel Is Nothing Then
    xlBook.Close
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing

End If
End Function

Private Sub ShowDesginc(rs As ADODB.Recordset)
Dim i As Long

With fpSpread1
    .MaxRows = 0
    Set .DataSource = rs

End With

With fpSpread1
    If .MaxRows < 0 Then
        MsgBox "没有查询到数据", vbInformation
        Exit Sub

    End If

End With

Select Case Combo1.text

    Case "AAMPN1"

        With fpSpread1

            For i = 1 To .MaxRows
                .Row = i
                .Col = 10
                .ColWidth(10) = 2
                .CellType = CellTypeCheckBox
            Next

        End With

    Case "AAMPN2"

        With fpSpread1

            For i = 1 To .MaxRows
                .Row = i
                .Col = 20
                .ColWidth(20) = 2
                .CellType = CellTypeCheckBox
            Next

        End With

    Case "37阴极线"

        With fpSpread1

            For i = 1 To .MaxRows
                .Row = i
                .Col = 8
                .ColWidth(8) = 2
                .CellType = CellTypeCheckBox
            Next

        End With

    Case "EU010机种信息"

        With fpSpread1

            For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                .Lock = True
                .Col = 3
                .Lock = True
                .Col = 9
                .ColWidth(8) = 7
                .CellType = CellTypeCheckBox
            Next

        End With
        
    Case "AC70机种信息", "HD机种信息"

        With fpSpread1

            For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                .Lock = True
                .Col = 2
                .Lock = True
                .Col = 4
                .ColWidth(8) = 7
                .CellType = CellTypeCheckBox
            Next

        End With

End Select

End Sub

Private Sub Insert()
If Combo1.text = "" Then
    MsgBox "请选择你要新增的选项名称", vbInformation, "提示"
    Exit Sub

End If

Select Case Combo1.text

    Case "AAMPN1"
        Frame1.Visible = True

    Case "AAMPN2"
        Frame2.Visible = True

    Case "37阴极线"
        Frame3.Visible = True

    Case "EU010机种信息"
        Frame4.Visible = True
        
    Case "AC70机种信息"
        Frame4.Visible = True
        Label14.Visible = False
        Text12.Visible = False
        Label11.Visible = False
        Text9.Visible = False
        Label13.Visible = False
        Text11.Visible = False
        Label15.Visible = False
        Text13.Visible = False
        lblPROVENANCE.Visible = False
        TextPROVENANCE.Visible = False
    Case "HD机种信息"
        Frame4.Visible = True
        Label14.Visible = False
        Text12.Visible = False
        Label11.Visible = False
        Text9.Visible = False
        Label13.Visible = False
        Text11.Visible = False
        'Label15.Visible = False
        'Text13.Visible = False
        lblPROVENANCE.Visible = False
        TextPROVENANCE.Visible = False
        txtPackage.Visible = False
        Label46.Visible = False
End Select

End Sub

Private Sub Command1_Click()
Dim strSql    As String
Dim device    As String
Dim BLINE     As Long
Dim CODE      As String
Dim CUSTOMER  As String
Dim TIMESTAMP As Date
Dim STATUS    As Long
Dim SEQ       As Long

If IsNumeric(Trim(Text2.text)) = False Or IsNumeric(Trim(Text6.text)) = False Or IsNumeric(Trim(Text7.text)) = False Then
    MsgBox "BLINE,STATUS,SEQ必须为数值类型", vbInformation, "提示"
    Exit Sub

End If

If Not IsDate(Trim(Text5.text)) Then
    MsgBox "日期格式错误", vbInformation, "提示"
    Exit Sub

End If

device = Trim(Text1.text)
BLINE = Val(Trim(Text2.text))
CODE = Trim(Text3.text)
CUSTOMER = Trim(Text4.text)
TIMESTAMP = CDate(Format(Trim(Text5.text), "00000000"))
STATUS = Val(Trim(Text6.text))
SEQ = Val(Trim(Text7.text))
strSql = "INSERT INTO code37(DEVICE,BLINE,CODE,CUSTOMER,TIMESTAMP,STATUS,SEQ)VALUES('" & device & "', " & BLINE & ",'" & CODE & "','" & CUSTOMER & "','" & TIMESTAMP & "'," & STATUS & "," & SEQ & ")"
AddSql (strSql)
Query

End Sub

Private Sub Command3_Click()
Dim strSql       As String
Dim cust_device  As String
Dim MARKING_CODE As String
Dim DEVICE_NAME  As String
Dim PRODUCT_12NC As String
Dim PMC          As String
Dim ORIG         As String
Dim PACKAGE      As String
Dim PROVENANCE   As String

cust_device = Trim(Text8.text)
MARKING_CODE = Trim(Text9.text)
DEVICE_NAME = Trim(Text10.text)
PRODUCT_12NC = Trim(Text11.text)
PACKAGE = Trim(txtPackage.text)
PMC = Trim(Text12.text)
ORIG = Trim(Text13.text)
PROVENANCE = Trim(TextPROVENANCE.text)

If cust_device = "" Then
    MsgBox "CUST_DEVICE不可为空", vbInformation, "提示"
    Exit Sub
End If

If DEVICE_NAME = "" Then
    MsgBox "DEVICE_NAME不可为空", vbInformation, "提示"
    Exit Sub
End If
If Combo1.text = "HD机种信息" Then
    If ORIG = "" Then
        MsgBox "ORIG", vbInformation, "提示"
        Exit Sub
    End If
Else

    If PACKAGE = "" Then
        MsgBox "PACKAGE不可为空", vbInformation, "提示"
        Exit Sub
    End If
End If

If Combo1.text <> "AC70机种信息" Then
    strSql = "INSERT INTO EU010_REFERENCE(CUST_DEVICE,MARKING_CODE,DEVICE_NAME,PRODUCT_12NC,PMC,ORIG,PACKAGE,PROVENANCE)values('" & cust_device & "','" & MARKING_CODE & "','" & DEVICE_NAME & "','" & PRODUCT_12NC & "','" & PMC & "','" & ORIG & "','" & PACKAGE & "','" & PROVENANCE & "')"
    If AddSql(strSql) > 0 Then
        MsgBox "新增成功", vbInformation, "提示"
    End If
Else
    strSql = "INSERT INTO EU010_REFERENCE(CUST_DEVICE,DEVICE_NAME,PACKAGE)values('" & cust_device & "','" & DEVICE_NAME & "','" & PACKAGE & "')"
    If AddSql(strSql) > 0 Then
        MsgBox "新增成功", vbInformation, "提示"
    End If
End If

Query

End Sub

Private Sub Command5_Click()
Dim strSql                As String
Dim id                    As Long
Dim MS_LEVEL              As String
Dim NUMBEROFHOURS         As String
Dim TYPENAME              As String
Dim FLAG                  As String
Dim QTECH_CREATED_BY      As String
Dim QTECH_CREATED_DATE    As Date
Dim QTECH_LASTUPDATE_BY   As String
Dim QTECH_LASTUPDATE_DATE As Date

If IsNumeric(Trim(Text14.text)) = False Then
    MsgBox "ID必须为数值类型", vbInformation, "提示"
    Exit Sub

End If

If Not IsDate(Trim(Text20.text)) Or Not IsDate(Trim(Text22.text)) Then
    MsgBox "日期格式错误", vbInformation, "提示"
    Exit Sub

End If

If IsID(Val(Trim(Text14.text))) > 0 Then
    MsgBox "ID已存在，请重新输入", vbInformation, "提示"
    Exit Sub

End If

id = Val(Trim(Text14.text))
MS_LEVEL = Trim(Text15.text)
NUMBEROFHOURS = Trim(Text16.text)
TYPENAME = Trim(Text17.text)
FLAG = Trim(Text18.text)
QTECH_CREATED_BY = Trim(Text19.text)
QTECH_CREATED_DATE = CDate(Format(Trim(Text20.text), "00000000"))
QTECH_LASTUPDATE_BY = Trim(Text21.text)
QTECH_LASTUPDATE_DATE = CDate(Format(Trim(Text22.text), "00000000"))
strSql = "INSERT INTO CUSTOMERMSLevelTBL(ID,MS_LEVEL,NUMBEROFHOURS,TYPENAME,FLAG,QTECH_CREATED_BY,QTECH_CREATED_DATE,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE)VALUES('" & id & "','" & MS_LEVEL & "','" & NUMBEROFHOURS & "','" & TYPENAME & "','" & FLAG & "','" & QTECH_CREATED_BY & "','" & QTECH_CREATED_DATE & "','" & QTECH_LASTUPDATE_BY & "','" & QTECH_LASTUPDATE_DATE & "')"
AddSql (strSql)
Query

End Sub

Private Sub Command7_Click()
Dim strSql                As String
Dim id                    As Long
Dim LOC                   As String
Dim PART                  As String
Dim LEAD_FREE             As String
Dim ECAT                  As String
Dim MSL                   As String
Dim TEMP                  As String
Dim HALIDE_FREE           As String
Dim PBF_DIE_ATTACH        As String
Dim MPQ_QTY               As Long
Dim PACKAGING_TYPE        As String
Dim PKG_GRP_CD            As String
Dim UL_LISTED_FLAG        As String
Dim FLAG                  As String
Dim QTECH_CREATED_BY      As String
Dim QTECH_CREATED_DATE    As Date
Dim QTECH_LASTUPDATE_BY   As String
Dim QTECH_LASTUPDATE_DATE As Date
Dim cust_device           As String
Dim MARKING_CODE          As String
Dim DEVICE_NAME           As String
Dim PRODUCT_12NC          As String
Dim PMC                   As String
Dim ORIG                  As String
Dim MARKINGCODEFIRST      As String

If IsNumeric(Trim(Text23.text)) = False Or IsNumeric(Trim(Text32.text)) = False Then
    MsgBox "ID或MPQ_QTY必须为数值类型", vbInformation, "提示"
    Exit Sub

End If

If Not IsDate(Trim(Text38.text)) Or Not IsDate(Trim(Text40.text)) Then
    MsgBox "日期格式错误", vbInformation, "提示"
    Exit Sub

End If

If IsID(Val(Trim(Text23.text))) > 0 Then
    MsgBox "ID已存在，请重新输入", vbInformation, "提示"
    Exit Sub

End If

id = Val(Trim(Text23.text))
LOC = Trim(Text24.text)
PART = Trim(Text25.text)
LEAD_FREE = Trim(Text26.text)
ECAT = Trim(Text27.text)
MSL = Trim(Text28.text)
TEMP = Trim(Text29.text)
HALIDE_FREE = Trim(Text30.text)
PBF_DIE_ATTACH = Trim(Text31.text)
MPQ_QTY = Val(Trim(Text32.text))
PACKAGING_TYPE = Trim(Text33.text)
PKG_GRP_CD = Trim(Text34.text)
UL_LISTED_FLAG = Trim(Text35.text)
FLAG = Trim(Text36.text)
QTECH_CREATED_BY = Trim(Text37.text)
QTECH_CREATED_DATE = CDate(Format(Trim(Text38.text), "00000000"))
QTECH_LASTUPDATE_BY = Trim(Text39.text)
QTECH_LASTUPDATE_DATE = CDate(Format(Trim(Text40.text), "00000000"))
MARKINGCODEFIRST = Trim(Text41.text)
strSql = "INSERT INTO CUSTOMERMPNAttributes(ID,LOC,PART,LEAD_FREE,ECAT,MSL,TEMP,HALIDE_FREE,PBF_DIE_ATTACH,MPQ_QTY,PACKAGING_TYPE,PKG_GRP_CD,UL_LISTED_FLAG,FLAG,QTECH_CREATED_BY,QTECH_CREATED_DATE,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE,MARKINGCODEFIRST)VALUES('" & id & "','" & LOC & "','" & PART & "','" & LEAD_FREE & "','" & ECAT & "','" & MSL & "','" & TEMP & "','" & HALIDE_FREE & "','" & PBF_DIE_ATTACH & "','" & MPQ_QTY & "','" & PACKAGING_TYPE & "','" & PKG_GRP_CD & "','" & UL_LISTED_FLAG & "','" & FLAG & "','" & QTECH_CREATED_BY & "','" & QTECH_CREATED_DATE & "','" & QTECH_LASTUPDATE_BY & "','" & QTECH_LASTUPDATE_DATE & "','" & MARKINGCODEFIRST & "')"
AddSql (strSql)
Query

End Sub

Private Sub Update()
Dim id                    As Long
Dim MS_LEVEL              As String
Dim NUMBEROFHOURS         As String
Dim TYPENAME              As String
Dim FLAG                  As String
Dim QTECH_CREATED_BY      As String
Dim QTECH_CREATED_DATE    As Date
Dim QTECH_LASTUPDATE_BY   As String
Dim QTECH_LASTUPDATE_DATE As Date
Dim LOC                   As String
Dim PART                  As String
Dim LEAD_FREE             As String
Dim ECAT                  As String
Dim MSL                   As String
Dim TEMP                  As String
Dim HALIDE_FREE           As String
Dim PBF_DIE_ATTACH        As String
Dim MPQ_QTY               As Long
Dim PACKAGING_TYPE        As String
Dim PKG_GRP_CD            As String
Dim UL_LISTED_FLAG        As String
Dim MARKINGCODEFIRST      As String
Dim device                As String
Dim BLINE                 As Long
Dim CODE                  As String
Dim CUSTOMER              As String
Dim TIMESTAMP             As Date
Dim STATUS                As Long
Dim SEQ                   As Long
Dim cust_device           As String
Dim MARKING_CODE          As String
Dim DEVICE_NAME           As String
Dim PRODUCT_12NC          As String
Dim PMC                   As String
Dim ORIG                  As String
Dim PACKAGE               As String
Dim strsql1               As String
Dim falg                  As Boolean
Dim i                     As Long
Dim iUpdateRecordCnt      As Long
Dim PROVENANCE            As String


iUpdateRecordCnt = 0
falg = False
If Combo1.text = "" Then
    MsgBox "请选择你要修改的选项名称", vbInformation, "提示"
    Exit Sub

End If

With fpSpread1
    If .MaxRows < 0 Then
        MsgBox "无数据", vbInformation
        Exit Sub

    End If

    For i = 1 To .MaxRows
        .Row = i

        Select Case Combo1.text

            Case "AAMPN1"
                .Col = 10
                If .Value = "1" Then
                    falg = True

                End If

            Case "AAMPN2"
                .Col = 20
                If .Value = "1" Then
                    falg = True

                End If

            Case "37阴极线"
                .Col = 8
                If .Value = "1" Then
                    falg = True

                End If

            Case "EU010机种信息"
                .Col = 9
                If .Value = "1" Then
                    falg = True

                End If
                
            Case "AC70机种信息", "HD机种信息"
                .Col = 4
                If .Value = "1" Then
                    falg = True

                End If
                
        End Select

    Next

End With

If falg = False Then
    MsgBox "请选择修改的行", vbInformation, "提示"
    Exit Sub

End If

With fpSpread1

    Select Case Combo1.text

        Case "AAMPN1"

            For i = 1 To .MaxRows
                .Row = i
                .Col = 10
                If .text = "1" Then
                    .Col = 1
                    id = Trim(.text)
                    .Col = 2
                    MS_LEVEL = Trim(.text)
                    .Col = 3
                    NUMBEROFHOURS = Trim(.text)
                    .Col = 4
                    TYPENAME = Trim(.text)
                    .Col = 5
                    FLAG = Trim(.text)
                    .Col = 6
                    QTECH_CREATED_BY = Trim(.text)
                    .Col = 7
                    QTECH_CREATED_DATE = Trim(.text)
                    .Col = 8
                    QTECH_LASTUPDATE_BY = Trim(.text)
                    .Col = 9
                    QTECH_LASTUPDATE_DATE = Trim(.text)

                End If

                If IsNumeric(Trim(Text14.text)) = False Then
                    MsgBox "ID必须为数值类型", vbInformation, "提示"
                    Exit Sub

                End If

                If Not IsDate(Trim(Text20.text)) Or Not IsDate(Trim(Text22.text)) Then
                    MsgBox "日期格式错误", vbInformation, "提示"
                    Exit Sub

                End If

                If IsID(Val(Trim(Text14.text))) > 0 Then
                    MsgBox "ID已存在，请重新输入", vbInformation, "提示"
                    Exit Sub

                End If

                strsql1 = "UPDATE INSITEQT2.CUSTOMERMSLEVELTBL SET MS_LEVEL='" & MS_LEVEL & "', NUMBEROFHOURS='" & NUMBEROFHOURS & "', TYPENAME='" & TYPENAME & "', FLAG='" & FLAG & "', QTECH_CREATED_BY='" & QTECH_CREATED_BY & "', QTECH_CREATED_DATE='" & QTECH_CREATED_DATE & "', QTECH_LASTUPDATE_BY='" & QTECH_LASTUPDATE_BY & "', QTECH_LASTUPDATE_DATE='" & QTECH_LASTUPDATE_DATE & "' where ID='" & id & "'"
                If AddSql(strsql1) > 0 Then
                    iUpdateRecordCnt = iUpdateRecordCnt + 1

                End If

            Next i

        Case "AAMPN2"

            For i = 1 To .MaxRows
                .Row = i
                .Col = 20
                If .text = "1" Then
                    .Col = 1
                    id = Trim(.text)
                    .Col = 2
                    LOC = Trim(.text)
                    .Col = 3
                    PART = Trim(.text)
                    .Col = 4
                    LEAD_FREE = Trim(.text)
                    .Col = 5
                    ECAT = Trim(.text)
                    .Col = 6
                    MSL = Trim(.text)
                    .Col = 7
                    TEMP = Trim(.text)
                    .Col = 8
                    HALIDE_FREE = Trim(.text)
                    .Col = 9
                    PBF_DIE_ATTACH = Trim(.text)
                    .Col = 10
                    MPQ_QTY = Trim(.text)
                    .Col = 11
                    PACKAGING_TYPE = Trim(.text)
                    .Col = 12
                    PKG_GRP_CD = Trim(.text)
                    .Col = 13
                    UL_LISTED_FLAG = Trim(.text)
                    .Col = 14
                    FLAG = Trim(.text)
                    .Col = 15
                    QTECH_CREATED_BY = Trim(.text)
                    .Col = 16
                    QTECH_CREATED_DATE = Trim(.text)
                    .Col = 17
                    QTECH_LASTUPDATE_BY = Trim(.text)
                    .Col = 18
                    QTECH_LASTUPDATE_DATE = Trim(.text)
                    .Col = 19
                    MARKINGCODEFIRST = Trim(.text)

                End If

                If IsNumeric(Trim(Text23.text)) = False Or IsNumeric(Trim(Text32.text)) = False Then
                    MsgBox "ID或MPQ_QTY必须为数值类型", vbInformation, "提示"
                    Exit Sub

                End If

                If Not IsDate(Trim(Text38.text)) Or Not IsDate(Trim(Text40.text)) Then
                    MsgBox "日期格式错误", vbInformation, "提示"
                    Exit Sub

                End If

                If IsID(Val(Trim(Text23.text))) > 0 Then
                    MsgBox "ID已存在，请重新输入", vbInformation, "提示"
                    Exit Sub

                End If

                strsql1 = "UPDATE INSITEQT2.CUSTOMERMPNATTRIBUTES SET  LOC='" & LOC & "', PART='" & PART & "', LEAD_FREE='" & LEAD_FREE & "', ECAT='" & ECAT & "', MSL='" & MSL & "', TEMP='" & TEMP & "', HALIDE_FREE='" & HALIDE_FREE & "', PBF_DIE_ATTACH='" & PBF_DIE_ATTACH & "', MPQ_QTY='" & MPQ_QTY & "', PACKAGING_TYPE='" & PACKAGING_TYPE & "', PKG_GRP_CD='" & PKG_GRP_CD & "', UL_LISTED_FLAG='" & UL_LISTED_FLAG & "', FLAG='" & FLAG & "', QTECH_CREATED_BY='" & QTECH_CREATED_BY & "', QTECH_CREATED_DATE='" & QTECH_CREATED_DATE & "', QTECH_LASTUPDATE_BY='" & QTECH_LASTUPDATE_BY & "', QTECH_LASTUPDATE_DATE='" & QTECH_LASTUPDATE_DATE & "', MARKINGCODEFIRST='" & MARKINGCODEFIRST & "' WHERE ID='" & id & "'"
                If AddSql(strsql1) > 0 Then
                    iUpdateRecordCnt = iUpdateRecordCnt + 1

                End If

            Next i

        Case "37阴极线"

            For i = 1 To .MaxRows
                .Row = i
                .Col = 8
                If .text = "1" Then
                    .Col = 1
                    device = Trim(.text)
                    .Col = 2
                    BLINE = Trim(.text)
                    .Col = 3
                    CODE = Trim(.text)
                    .Col = 4
                    CUSTOMER = Trim(.text)
                    .Col = 5
                    TIMESTAMP = Trim(.text)
                    .Col = 6
                    STATUS = Trim(.text)
                    .Col = 7
                    SEQ = Trim(.text)

                End If

                If IsNumeric(Trim(Text2.text)) = False Or IsNumeric(Trim(Text6.text)) = False Or IsNumeric(Trim(Text7.text)) = False Then
                    MsgBox "BLINE,STATUS,SEQ必须为数值类型", vbInformation, "提示"
                    Exit Sub

                End If

                If Not IsDate(Trim(Text5.text)) Then
                    MsgBox "日期格式错误", vbInformation, "提示"
                    Exit Sub

                End If

                strsql1 = "UPDATE INSITEQT2.CODE37 SET  BLINE='" & BLINE & "', CODE='" & CODE & "',TIMESTAMP='" & TIMESTAMP & "', STATUS='" & STATUS & "', SEQ='" & SEQ & "' WHERE CUSTOMER='" & CUSTOMER & "' and DEVICE='" & device & "'"
                If AddSql(strsql1) > 0 Then
                    iUpdateRecordCnt = iUpdateRecordCnt + 1

                End If

            Next i

        Case "EU010机种信息"

            For i = 1 To .MaxRows
                .Row = i
                .Col = 9
                If .text = "1" Then
                    .Col = 1
                    cust_device = Trim(.text)
                    .Col = 2
                    MARKING_CODE = Trim(.text)
                    .Col = 3
                    DEVICE_NAME = Trim(.text)
                    .Col = 4
                    PRODUCT_12NC = Trim(.text)
                    .Col = 5
                    PMC = Trim(.text)
                    .Col = 6
                    ORIG = Trim(.text)
                    .Col = 7
                    PACKAGE = Trim(.text)
                    .Col = 8
                    PROVENANCE = Trim(.text)
                    strsql1 = "UPDATE INSITEQT2.EU010_REFERENCE SET  MARKING_CODE='" & MARKING_CODE & "', PRODUCT_12NC='" & PRODUCT_12NC & "', PMC='" & PMC & "', ORIG='" & ORIG & "', PACKAGE = '" & PACKAGE & "', PROVENANCE='" & PROVENANCE & "'  WHERE CUST_DEVICE='" & cust_device & "' and  DEVICE_NAME= '" & DEVICE_NAME & "'"
                    If AddSql(strsql1) > 0 Then
                        iUpdateRecordCnt = iUpdateRecordCnt + 1

                    End If

                End If

            Next i
            
        Case "AC70机种信息"

            For i = 1 To .MaxRows
                .Row = i
                .Col = 4
                If .text = "1" Then
                    .Col = 1
                    cust_device = Trim(.text)
                    .Col = 2
                    DEVICE_NAME = Trim(.text)
                    .Col = 3
                    PACKAGE = Trim(.text)
                    strsql1 = "UPDATE INSITEQT2.EU010_REFERENCE SET   PACKAGE = '" & PACKAGE & "'  WHERE CUST_DEVICE='" & cust_device & "' and  DEVICE_NAME= '" & DEVICE_NAME & "'"
                    If AddSql(strsql1) > 0 Then
                        iUpdateRecordCnt = iUpdateRecordCnt + 1

                    End If

                End If

            Next i
            
        Case "HD机种信息"

            For i = 1 To .MaxRows
                .Row = i
                .Col = 4
                If .text = "1" Then
                    .Col = 1
                    cust_device = Trim(.text)
                    .Col = 2
                    DEVICE_NAME = Trim(.text)
                    .Col = 3
                    ORIG = Trim(.text)
                    strsql1 = "UPDATE INSITEQT2.EU010_REFERENCE SET   ORIG = '" & ORIG & "'  WHERE CUST_DEVICE='" & cust_device & "' and  DEVICE_NAME= '" & DEVICE_NAME & "'"
                    If AddSql(strsql1) > 0 Then
                        iUpdateRecordCnt = iUpdateRecordCnt + 1

                    End If

                End If

            Next i
        End Select

End With

MsgBox "已成功修改:" & iUpdateRecordCnt & "笔数据", vbInformation, "修改提示"
Query

End Sub

Private Sub Delete()
Dim rs            As New ADODB.Recordset
Dim strsql1       As String
Dim id            As String
Dim CUSTOMER      As String
Dim device        As String
Dim cust_device   As String
Dim DEVICE_NAME   As String
Dim i             As Long
Dim falg          As Boolean
Dim iDelRecordCnt As Long

iDelRecordCnt = 0
falg = False

With fpSpread1
    If .MaxRows < 0 Then
        MsgBox "无数据", vbInformation
        Exit Sub

    End If

    For i = 1 To .MaxRows
        .Row = i

        Select Case Combo1.text

            Case "AAMPN1"
                .Col = 10
                If .Value = "1" Then
                    falg = True

                End If

            Case "AAMPN2"
                .Col = 20
                If .Value = "1" Then
                    falg = True

                End If

            Case "37阴极线"
                .Col = 8
                If .Value = "1" Then
                    falg = True

                End If

            Case "EU010机种信息"
                .Col = 8
                If .Value = "1" Then
                    falg = True

                End If
            Case "AC70机种信息", "HD机种信息"
                .Col = 4
                If .Value = "1" Then
                    falg = True

                End If

        End Select

    Next

End With

If falg = False Then
    MsgBox "请选择删除的行", vbInformation, "提示"
    Exit Sub

End If

With fpSpread1

    For i = 1 To .MaxRows
        .Row = i

        Select Case Combo1.text

            Case "AAMPN1"
                .Col = 10
                If .text = "1" Then
                    .Col = 1
                    id = Trim(.text)

                End If

                strsql1 = "DELETE FROM CUSTOMERMSLevelTBL WHERE ID='" & id & "'"
                If AddSql(strsql1) > 0 Then
                    iDelRecordCnt = iDelRecordCnt + 1

                End If

            Case "AAMPN2"
                .Col = 20
                If .text = "1" Then
                    .Col = 1
                    id = Trim(.text)

                End If

                strsql1 = "DELETE FROM CUSTOMERMPNAttributes WHERE ID='" & id & "'"
                If AddSql(strsql1) > 0 Then
                    iDelRecordCnt = iDelRecordCnt + 1

                End If

            Case "37阴极线"
                .Col = 8
                If .text = "1" Then
                    .Col = 1
                    device = Trim(.text)
                    .Col = 4
                    CUSTOMER = Trim(.text)

                End If

                strsql1 = "DELETE FROM code37 WHERE CUSTOMER='" & CUSTOMER & "' and DEVICE='" & device & "'"
                If AddSql(strsql1) > 0 Then
                    iDelRecordCnt = iDelRecordCnt + 1

                End If

            Case "EU010机种信息"
                .Col = 8
                If .text = "1" Then
                    .Col = 1
                    cust_device = Trim(.text)
                    .Col = 3
                    DEVICE_NAME = Trim(.text)

                End If

                strsql1 = "DELETE FROM EU010_REFERENCE WHERE CUST_DEVICE='" & cust_device & "' and  DEVICE_NAME= '" & DEVICE_NAME & "'"
                If AddSql(strsql1) > 0 Then
                    iDelRecordCnt = iDelRecordCnt + 1

                End If

            Case "AC70机种信息", "HD机种信息"
                .Col = 4
                If .text = "1" Then
                    .Col = 1
                    cust_device = Trim(.text)
                    .Col = 2
                    DEVICE_NAME = Trim(.text)

                End If

                strsql1 = "DELETE FROM EU010_REFERENCE WHERE CUST_DEVICE='" & cust_device & "' and  DEVICE_NAME= '" & DEVICE_NAME & "'"
                If AddSql(strsql1) > 0 Then
                    iDelRecordCnt = iDelRecordCnt + 1

                End If
                
        End Select

    Next

End With

MsgBox "已成功删除:" & iDelRecordCnt & "笔数据", vbInformation, "删除提示"
Query

End Sub


Private Function ColumnToNum(char As Variant) ', Optional n As Integer)  excel数字变字母列号
Dim char_times As Integer, char_mod As Integer
Dim a As Integer
'char 可以是字母（大小写），也可以是数字
'功能1：将列数转化为大写字母
'功能2：将大，小写字母转化为列数
    char_mod = 0
    If IsNumeric(char) Then
        char_times = Int(char / 26)
        char_mod = char Mod 26
        Select Case char_times
        Case 0
            ColumnToNum = Chr(char + 64)
        Case Is < 27
            ColumnToNum = Chr(char_times + 64) & Chr(char_mod + 64)
        Case Else
            ColumnToNum = "超出范围"
        End Select
    Else
        If Asc(Mid(char, 1, 1)) < 91 And Asc(Mid(char, 1, 1)) > 64 Then '大写
            For a = Len(char) To 1 Step -1
                char_mod = Application.WorksheetFunction.Power(26, Len(char) - a) * (Asc(Mid(char, a, 1)) - 64) + char_mod
            Next a
        ElseIf Asc(Mid(char, 1, 1)) < 123 And Asc(Mid(char, 1, 1)) > 96 Then '小写
            For a = Len(char) To 1 Step -1
                char_mod = Application.WorksheetFunction.Power(26, Len(char) - a) * (Asc(Mid(char, a, 1)) - 96) + char_mod
            Next a
        End If
          ColumnToNum = char_mod
    End If
    
End Function



