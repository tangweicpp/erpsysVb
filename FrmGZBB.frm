VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmGZBB 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   13890
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   25170
   FillColor       =   &H000000FF&
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
   ScaleHeight     =   13890
   ScaleWidth      =   25170
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combojiaji 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      ItemData        =   "FrmGZBB.frx":0000
      Left            =   14520
      List            =   "FrmGZBB.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   63
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Frame Fra 
      Caption         =   "光照版本查询界面"
      Height          =   3375
      Left            =   1560
      TabIndex        =   30
      Top             =   3600
      Visible         =   0   'False
      Width           =   15855
      Begin VB.ComboBox BMcode1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         ItemData        =   "FrmGZBB.frx":0016
         Left            =   7200
         List            =   "FrmGZBB.frx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         ItemData        =   "FrmGZBB.frx":0098
         Left            =   7200
         List            =   "FrmGZBB.frx":00A2
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Left            =   6120
         TabIndex        =   60
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Left            =   6120
         TabIndex        =   59
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Height          =   135
         Left            =   9480
         TabIndex        =   58
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Left            =   9480
         TabIndex        =   57
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox Check5 
         Height          =   255
         Left            =   6120
         TabIndex        =   56
         Top             =   600
         Width           =   255
      End
      Begin VB.ComboBox CScode1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         ItemData        =   "FrmGZBB.frx":00AE
         Left            =   7200
         List            =   "FrmGZBB.frx":00CD
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   4080
         TabIndex        =   43
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   4080
         TabIndex        =   42
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1200
         TabIndex        =   41
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1200
         TabIndex        =   40
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1200
         TabIndex        =   39
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "执行查询"
         Height          =   495
         Left            =   6600
         TabIndex        =   32
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "关闭"
         Height          =   495
         Left            =   8880
         TabIndex        =   31
         Top             =   2520
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTP2 
         Height          =   375
         Index           =   1
         Left            =   11280
         TabIndex        =   50
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   305266689
         CurrentDate     =   43738
      End
      Begin MSComCtl2.DTPicker DTP4 
         Height          =   375
         Index           =   0
         Left            =   11280
         TabIndex        =   51
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   305266689
         CurrentDate     =   43738
      End
      Begin MSComCtl2.DTPicker DTP3 
         Height          =   375
         Index           =   2
         Left            =   13680
         TabIndex        =   52
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   305266689
         CurrentDate     =   43738
      End
      Begin MSComCtl2.DTPicker DTP5 
         Height          =   375
         Index           =   2
         Left            =   13680
         TabIndex        =   53
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   305266689
         CurrentDate     =   43738
      End
      Begin VB.Label lblxx2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         Height          =   195
         Left            =   13440
         TabIndex        =   55
         Top             =   1440
         Width           =   180
      End
      Begin VB.Label lblxx1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         Height          =   195
         Left            =   13440
         TabIndex        =   54
         Top             =   720
         Width           =   180
      End
      Begin VB.Label lbl10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "是否加急"
         Height          =   195
         Left            =   6360
         TabIndex        =   48
         Top             =   2040
         Width           =   840
      End
      Begin VB.Label lbl9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请部门"
         Height          =   195
         Left            =   6360
         TabIndex        =   47
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label lbl8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预计到厂时间范围："
         Height          =   195
         Left            =   9720
         TabIndex        =   46
         Top             =   1320
         Width           =   1620
      End
      Begin VB.Label lbl7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6960
         TabIndex        =   45
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label lbl6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "订购日期范围："
         Height          =   195
         Left            =   10080
         TabIndex        =   44
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label lblCUST_DEVICE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "料号"
         Height          =   195
         Left            =   720
         TabIndex        =   38
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lblMARKING_CODE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
         Height          =   195
         Left            =   600
         TabIndex        =   37
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label lblDEVICE_NAME 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "需求人"
         Height          =   195
         Left            =   3480
         TabIndex        =   36
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblPRODUCT_12NC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注版本"
         Height          =   195
         Left            =   480
         TabIndex        =   35
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label lblPMC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原因"
         Height          =   195
         Left            =   3480
         TabIndex        =   34
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label lblORIG 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "厂商"
         Height          =   195
         Left            =   6360
         TabIndex        =   33
         Top             =   600
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmd_report 
      BackColor       =   &H0000C000&
      Caption         =   "条件导出"
      Height          =   600
      Left            =   17160
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmd_reportALL 
      BackColor       =   &H0000FF00&
      Caption         =   "导出所有"
      Height          =   600
      Left            =   14880
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2520
      Width           =   2055
   End
   Begin VB.ComboBox BMcode 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      ItemData        =   "FrmGZBB.frx":010B
      Left            =   14520
      List            =   "FrmGZBB.frx":0127
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   240
      Width           =   1455
   End
   Begin VB.ComboBox CScode 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      ItemData        =   "FrmGZBB.frx":018D
      Left            =   10440
      List            =   "FrmGZBB.frx":01AC
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtText5 
      Height          =   435
      Left            =   4920
      TabIndex        =   17
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000080FF&
      Caption         =   "退出"
      Height          =   480
      Left            =   22080
      MaskColor       =   &H008080FF&
      MousePointer    =   1  'Arrow
      TabIndex        =   16
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "清空输入框"
      Height          =   600
      Left            =   2520
      TabIndex        =   15
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdFPQX 
      Caption         =   "分配权限"
      Height          =   720
      Left            =   17040
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtText4 
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton cmd_Modify 
      Caption         =   "修改"
      Height          =   600
      Left            =   10320
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmd_del 
      BackColor       =   &H000000FF&
      Caption         =   "删除"
      Height          =   360
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd_add 
      Caption         =   "新增"
      Height          =   600
      Left            =   8040
      TabIndex        =   8
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton cmd_query 
      Caption         =   "查询"
      Height          =   600
      Left            =   5640
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtText3 
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txtText2 
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtText1 
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTP 
      Height          =   375
      Index           =   0
      Left            =   10440
      TabIndex        =   20
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   156106753
      CurrentDate     =   43738
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   375
      Index           =   1
      Left            =   10440
      TabIndex        =   22
      Top             =   1440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   156106753
      CurrentDate     =   43738
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   10215
      Index           =   0
      Left            =   0
      TabIndex        =   27
      Top             =   3360
      Width           =   23535
      _Version        =   524288
      _ExtentX        =   41513
      _ExtentY        =   18018
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
      SpreadDesigner  =   "FrmGZBB.frx":01EA
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "厂商"
      Height          =   315
      Left            =   9960
      TabIndex        =   24
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "申请部门"
      Height          =   195
      Left            =   13680
      TabIndex        =   23
      Top             =   240
      Width           =   720
   End
   Begin VB.Label label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "预计到厂时间"
      Height          =   555
      Left            =   9240
      TabIndex        =   21
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "是否加急"
      Height          =   255
      Left            =   13680
      TabIndex        =   19
      Top             =   1320
      Width           =   840
   End
   Begin VB.Label label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "原因"
      Height          =   195
      Left            =   4200
      TabIndex        =   18
      Top             =   1560
      Width           =   360
   End
   Begin VB.Line Line2 
      X1              =   -720
      X2              =   21240
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   22200
      X2              =   22200
      Y1              =   5520
      Y2              =   14640
   End
   Begin VB.Label label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "需求人"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   360
      Width           =   585
   End
   Begin VB.Label label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "订购日期"
      Height          =   315
      Left            =   9600
      TabIndex        =   11
      Top             =   720
      Width           =   720
   End
   Begin VB.Label label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注版本"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   585
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "料号"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   360
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   45
   End
End
Attribute VB_Name = "FrmGZBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' Project    :       生产管理系统
' Procedure  :       光罩版本维护
' Description:       增删查改及数据输入纠错，导出excel，自动生成barcode，生成数据附
'                    带操作人员信息
' Created by :       祝祎凡
' Machine    :       DESKTOP-F6L8S2V
' Date-Time  :       2019/10/16-9:37:47
'
' Parameters :       SQLserver表erptemp.dbo.MASK_CODE，数据来源技术部
'--------------------------------------------------------------------------------
Dim strsqlEX As String '全局变量存放预导出excel的SQL

Private Sub cmd_Click()
Fra.Visible = False
End Sub

Private Sub cmd_report_Click()  '条件导出excel

SqlServerExporToExcel (strsqlEX)
'    Dim strSql       As String
'    Dim rs           As New ADODB.Recordset
'
'    Dim lot         As String
'    Dim Barcode      As String
'    Dim BZBB          As String
'    Dim CS         As String
'    Dim SQBM       As String
'    Dim DGRQ       As String
'    Dim isJiaJi    As String
'    Dim XQBM   As String
'    Dim XQR         As String
'    Dim YJDCSJ  As String
'    Dim reason     As String
'
'
'
'    strSql = "select '' AS '选择',ID,lot as '料号', Barcode as 'Barcode', BZBB as '备注版本',CS as '厂商'," & _
'    "DGRQ as '订购日期',isJiaJi as '是否加急', YJDCSJ as '预计到厂时间',SQBM as '申请部门',XQR as '需求人'," & _
'    "reason as '原因',Create_time as '创建时间' from erptemp.dbo.MASK_CODE where isdel = '0'  "
'
'    If Trim(txtText1.Text) <> "" Then
'        strSql = strSql + " AND LOT  = '" & Trim(txtText1.Text) & "'  "
'    End If
'
'    If Trim(txtText2.Text) <> "" Then
'        strSql = strSql + " AND Barcode = '" & Trim(txtText2.Text) & "'  "
'
'    End If
'
'    If Trim(txtText3.Text) <> "" Then
'        strSql = strSql + " AND BZBB  = '" & Trim(txtText3.Text) & "'  "
'    End If
'
'    If chkCheck5 = 1 Then
'        If Trim(CScode.Text) <> "" Then
'            strSql = strSql + " AND CS  = '" & Trim(CScode.Text) & "'  "
'        End If
'    End If
'
'    If chkCheck2 = 1 Then
'    strSql = strSql + " AND DGRQ  = '" & DGRQ & "'  "
'    End If
'
'    If chkCheck6 = 1 Then
'        If Trim(BMcode.Text) <> "" Then
'            strSql = strSql + " AND SQBM  = '" & Trim(BMcode.Text) & "'  "
'
'        End If
'    End If
'
'    If Trim(txtText4.Text) <> "" Then
'        strSql = strSql + " AND XQR  = '" & Trim(txtText4.Text) & "'  "
'
'    End If
'
'    If Trim(txtText5.Text) <> "" Then
'        strSql = strSql + " AND reason  = '" & Trim(txtText4.Text) & "'  "
'
'    End If
'
'    If chkCheck3 = 1 Then
'        strSql = strSql + " AND isJiaJi  = '" & Trim(Combojiaji1.Text) & "'  "
'    End If
'
'    If chkCheck4 = 1 Then
'    strSql = strSql + " AND YJDCSJ = '" & YJDCSJ & "'  "
'    End If
'
'    strSql = strSql + "  order by Create_time desc,DGRQ desc,Barcode desc"
End Sub

Private Sub cmd_reportALL_Click()  '全部导出excel
 Dim TEMP As String
  
  TEMP = "select '' AS '选择',ID,lot as '料号', Barcode as 'Barcode', BZBB as '备注版本',CS as '厂商', DGRQ as '订购日期',isJiaJi as '是否加急', YJDCSJ as '预计到厂时间'," & _
    "SQBM as '申请部门',XQR as '需求人',reason as '原因',Create_time as '创建时间' from erptemp.dbo.MASK_CODE where isdel = '0' order by Create_time desc,DGRQ desc,Barcode desc"
  
 SqlServerExporToExcel (TEMP)
End Sub

Private Sub cmdFPQX_Click()  '开通管理者权限管理窗口
  FrmGZBB_GLY.Show
End Sub

Private Sub CmdClear_Click()   '初始化输入数据
  Initial
End Sub

Private Sub CmdOK_Click()
    Query
End Sub

Private Sub CmdQuit_Click()   '退出
  Unload Me
End Sub

Private Sub cmd_query_Click()  '条件查询
  Fra.Visible = True
End Sub
Private Sub cmd_add_Click()     '新增
    Dim strSql       As String
    Dim rs           As New ADODB.Recordset
    
   '  Dim ID As Integer
    Dim lot         As String
    Dim Barcode      As String
    Dim BZBB          As String
    Dim CS         As String
    Dim SQBM       As String
    Dim DGRQ       As String
    Dim isJiaJi    As String
    Dim XQBM   As String
    Dim XQR         As String
    Dim YJDCSJ  As String
    Dim SRRYGH As String
    Dim reason     As String
    
    Dim count As Integer   'barcode 重复次数
    Dim icount As Integer  '完全重复的数据次数
    count = 0
    icount = 0
    
    If Check5_input(txtText1.Text) = 1 Then
        MsgBox "lot参数输入不合法"
        Exit Sub
    ElseIf Check5_input(txtText2.Text) = 1 Then
        MsgBox "Barcode参数输入不合法"
        Exit Sub
    ElseIf Check5_input(txtText3.Text) = 1 Then
        MsgBox "备注版本参数不合法"
        Exit Sub
    ElseIf txtText4.Text <> "" And Check5_input(txtText4.Text) = 1 Then
        MsgBox "需求人参数输入不合法"
        Exit Sub
    ElseIf txtText5.Text <> "" And Check5_input(txtText5.Text) = 1 Then
        MsgBox "原因参数输入不合法"
        Exit Sub
    End If
    
    'Barcode = getBarcode() + Trim(txtText2.Text)   'Barcode 自动生成
    Barcode = Trim(txtText2.Text)
    lot = Trim(txtText1.Text)
    BZBB = Trim(txtText3.Text)
    CS = Trim(CScode.Text)
    DGRQ = DTP(0).Value
    YJDCSJ = DTP1(1).Value
    SQBM = Trim(BMcode.Text)
    XQR = Trim(txtText4.Text)
    reason = Trim(txtText5.Text)
    isJiaJi = Trim(Combojiaji.Text)
    SRRYGH = gUserName
    
    Create_time = DATE
      
    If DateDiff("d", CDate(DGRQ), CDate(YJDCSJ)) < 0 Then
        MsgBox "到厂时间比订购日期小，错误！"
        Exit Sub
    End If
    
'    strSql = "select max(ID) as ""ID""  from  CUSTOMERMPNAttributes"  '获取ID最大值 数据库中ID自动生成，自动加一
'    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
'    ID = rs.Fields("ID")
'    ID = ID + 1
'    rs.Close


'    strSql = "select * from erptemp.dbo.MASK_CODE where 1=1 and isdel = '0' and Barcode = '" & Barcode & "'"
'    If INIadoCon.State <> adStateOpen Then
'        INIConnectSTART2
'    End If
'    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
'
'
'    If lot = "" Or Barcode = "" Or BZBB = "" Or DGRQ = "" Or XQR = "" Or CS = "" Or reason = "" Then
'        MsgBox "信息都为必填！"
'        Exit Sub
'    ElseIf rs.RecordCount > 0 Then
'         count = rs.RecordCount
'        If MsgBox("提示：Barcode重复是否继续？重复次数为" & count, vbOKCancel, "提示") = vbOK Then
'             strSql = "INSERT INTO erptemp.dbo.MASK_CODE (LOT,  Barcode, BZBB,CS,DGRQ,YJDCSJ,SQBM,XQR,reason,SRRYGH,Create_time,isJiaJi,isdel)" & _
'            "values('" & lot & "','" & Barcode & "','" & BZBB & "','" & CS & "','" & DGRQ & "','" & YJDCSJ & "','" & SQBM & "','" & XQR & _
'            "','" & reason & "','" & SRRYGH & "','" & Create_time & "','" & isJiaJi & "','0" & "')"
'            Exec_Sql (strSql)
'            rs.Close
'            MsgBox "添加成功！"
'
'        Else
'            Exit Sub
'        End If
'
'    Else
'        '信息插入表中
'        rs.Close
'         strSql = "select * from erptemp.dbo.MASK_CODE where 1=1 and isdel = '0' and LOT = '" & lot & "' and CS = '" & CS & _
'        "' and Barcode = '" & Barcode & "' and BZBB = '" & BZBB & "' and DGRQ = '" & DGRQ & "' and XQR = '" & XQR & "'"
'
'        If INIadoCon.State <> adStateOpen Then
'            INIConnectSTART2
'        End If
'        rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
'
'        If rs.RecordCount > 0 Then
'             icount = rs.RecordCount
'             If MsgBox("提示：数据完全重复是否继续？重复次数为" & icount, vbOKCancel, "提示") = vbOK Then
'                strSql = "INSERT INTO erptemp.dbo.MASK_CODE (LOT,  Barcode, BZBB,CS,DGRQ,YJDCSJ,SQBM,XQR,reason,SRRYGH,Create_time,isJiaJi,isdel)" & _
'            "values('" & lot & "','" & Barcode & "','" & BZBB & "','" & CS & "','" & DGRQ & "','" & YJDCSJ & "','" & SQBM & "','" & XQR & _
'            "','" & reason & "','" & gUserName & "','" & Create_time & "','" & isJiaJi & "','0" & "')"
'                Exec_Sql (strSql)
'                rs.Close
'                MsgBox "添加成功！"
'
'            Else
'                Exit Sub
'            End If
'        Else
'           strSql = "INSERT INTO erptemp.dbo.MASK_CODE (LOT,  Barcode, BZBB,CS,DGRQ,YJDCSJ,SQBM,XQR,reason,SRRYGH,Create_time,isJiaJi,isdel)" & _
'            "values('" & lot & "','" & Barcode & "','" & BZBB & "','" & CS & "','" & DGRQ & "','" & YJDCSJ & "','" & SQBM & "','" & XQR & _
'            "','" & reason & "','" & gUserName & "','" & Create_time & "','" & isJiaJi & "','0" & "')"
'            Exec_Sql (strSql)
'            MsgBox "添加成功！"
'            rs.Close
'        End If
'    End If
'
'
    strSql = "select * from erptemp.dbo.MASK_CODE where 1=1 and isdel = '0' and LOT = '" & lot & "' and CS = '" & CS & _
        "' and Barcode = '" & Barcode & "' and BZBB = '" & BZBB & "' and DGRQ = '" & DGRQ & "' and XQR = '" & XQR & "'"

    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2
    End If
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If lot = "" Or Barcode = "" Or BZBB = "" Or DGRQ = "" Or XQR = "" Or CS = "" Or reason = "" Then
        MsgBox "信息都为必填！"
        Exit Sub
    ElseIf rs.RecordCount > 0 Then
        icount = rs.RecordCount
         If MsgBox("提示：数据完全重复!重复次数为" & icount & "是否继续?", vbOKCancel, "提示") = vbOK Then
            strSql = "INSERT INTO erptemp.dbo.MASK_CODE (LOT,  Barcode, BZBB,CS,DGRQ,YJDCSJ,SQBM,XQR,reason,SRRYGH,Create_time,isJiaJi,isdel)" & _
            "values('" & lot & "','" & Barcode & "','" & BZBB & "','" & CS & "','" & DGRQ & "','" & YJDCSJ & "','" & SQBM & "','" & XQR & _
            "','" & reason & "','" & gUserName & "','" & Create_time & "','" & isJiaJi & "','0" & "')"
            Exec_Sql (strSql)
            rs.Close
            MsgBox "添加成功！"
         Else
             Exit Sub
         End If
    Else
        rs.Close
          strSql = "select * from erptemp.dbo.MASK_CODE where 1=1 and isdel = '0' and Barcode = '" & Barcode & "'"
        
        If INIadoCon.State <> adStateOpen Then
            INIConnectSTART2
        End If
        rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
              
        If rs.RecordCount > 0 Then
            count = rs.RecordCount
            If MsgBox("提示：Barcode重复!重复次数为" & count & "是否继续？", vbOKCancel, "提示") = vbOK Then
                strSql = "INSERT INTO erptemp.dbo.MASK_CODE (LOT,  Barcode, BZBB,CS,DGRQ,YJDCSJ,SQBM,XQR,reason,SRRYGH,Create_time,isJiaJi,isdel)" & _
                "values('" & lot & "','" & Barcode & "','" & BZBB & "','" & CS & "','" & DGRQ & "','" & YJDCSJ & "','" & SQBM & "','" & XQR & _
                "','" & reason & "','" & SRRYGH & "','" & Create_time & "','" & isJiaJi & "','0" & "')"
                Exec_Sql (strSql)
                rs.Close
                MsgBox "添加成功！"
            Else
                Exit Sub
            End If
        Else
             strSql = "INSERT INTO erptemp.dbo.MASK_CODE (LOT,  Barcode, BZBB,CS,DGRQ,YJDCSJ,SQBM,XQR,reason,SRRYGH,Create_time,isJiaJi,isdel)" & _
                "values('" & lot & "','" & Barcode & "','" & BZBB & "','" & CS & "','" & DGRQ & "','" & YJDCSJ & "','" & SQBM & "','" & XQR & _
                "','" & reason & "','" & gUserName & "','" & Create_time & "','" & isJiaJi & "','0" & "')"
                Exec_Sql (strSql)
                MsgBox "添加成功！"
                rs.Close
        End If
    End If
    
   query2

End Sub
Private Sub cmd_del_Click()

    '删除
    Dim i      As Integer
    Dim strSql As String
    Dim strSql2 As String
    Dim count  As Integer

    count = 0
    strSql2 = "update erptemp.dbo.MASK_CODE set isdel = '1' where ID  = "

    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 2
                If Trim(.Text) <> "" Then
                    strSql = strSql2 + "'" & Trim(.Text) & "'  "
                End If


                If AddSql2(strSql) > -1 Then
                        count = count + 1
                End If
            End If
            '.Col = 1
            strSql = strSql2
        Next i
    End With

        If count = 0 Then
            MsgBox "删除失败"
        Else
            MsgBox "删除成功" & "删除记录数" & count & "! "
        End If
    query2

End Sub
Private Sub cmd_Modify_Click()

    Dim rs        As New ADODB.Recordset

    '修改
    Dim i         As Integer

    Dim strSql    As String
    
    Dim lot        As String
    Dim Barcode        As String
    Dim BZBB          As String
    Dim DGRQ        As String
    Dim isJiaJi As String
    Dim YJDCSJ   As String
    Dim XQR         As String
    Dim reason     As String
    Dim Create_time As String
    
    Dim count As Integer
    
    count = 0
    
    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 2
                If Trim(.Text) <> "" Then
                    ID = Trim(.Text)
                End If
                
                .Col = 3
                If Trim(.Text) <> "" Then
                    lot = Trim(.Text)
                End If

                .Col = 4
                If Trim(.Text) <> "" Then
                    Barcode = Trim(.Text)
                End If

                .Col = 5
                If Trim(.Text) <> "" Then
                    BZBB = Trim(.Text)
                End If

                .Col = 6
                If Trim(.Text) <> "" Then
                    CS = Trim(.Text)
7                End If

                .Col = 7
                If Trim(.Text) <> "" Then
                    DGRQ = Trim(.Text)
                End If

                .Col = 8
                If Trim(.Text) <> "" Then
                    isJiaJi = Trim(.Text)
                End If
                
                .Col = 9
                If Trim(.Text) <> "" Then
                    YJDCSJ = Trim(.Text)
                End If
                
                .Col = 10
                If Trim(.Text) <> "" Then
                    SQBM = Trim(.Text)
                End If
                
                .Col = 11
                If Trim(.Text) <> "" Then
                    XQR = Trim(.Text)
                End If
                
                .Col = 12
                If Trim(.Text) <> "" Then
                    reason = Trim(.Text)
                End If
                
                .Col = 14
                If Trim(.Text) <> "" Then
                    SRRYGH = Trim(.Text)
                End If
                
                
                If Check5_input(lot) = 1 Then
                    MsgBox "lot参数输入不合法"
                    Exit Sub
                ElseIf Check5_input(Barcode) = 1 Then
                    MsgBox "Barcode参数输入不合法"
                    Exit Sub
                ElseIf Check5_input(BZBB) = 1 Then
                    MsgBox "备注版本参数不合法"
                    Exit Sub
                ElseIf XQR <> "" And Check5_input(XQR) = 1 Then
                    MsgBox "需求人参数输入不合法"
                    Exit Sub
                ElseIf reason <> "" And Check5_input(reason) = 1 Then
                    MsgBox "原因参数输入不合法"
                    Exit Sub
                ElseIf Check5_date(DGRQ) = 1 Then
                    MsgBox "订购日期参数输入不合法 正确格式为：YY-MM-DD"
                    Exit Sub
                ElseIf Check5_date(YJDCSJ) = 1 Then
                    MsgBox "预计到厂时间参数输入不合法 正确格式为：YY-MM-DD"
                    Exit Sub
                End If
  
                 If DateDiff("d", CDate(DGRQ), CDate(YJDCSJ)) < 0 Then
                    MsgBox "到厂时间比订购日期小，错误！该记录的ID为" & ID
                    Exit Sub
                 End If
    
                If lot = "" Or Barcode = "" Or BZBB = "" Or CS = "" Or DGRQ = "" Or YJDCSJ = "" Then
                    MsgBox "主要信息都为必填！"
                Else
                   strSql = "select * from erptemp.dbo.MASK_CODE where ID = '" & ID & "'"

                    If INIadoCon.State <> adStateOpen Then
                        INIConnectSTART2
                    End If

                    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
                    
                    If rs.RecordCount > 0 Then
                        rs.Close
                       strSql = "UPDATE erptemp.dbo.MASK_CODE SET lot='" & lot & "'" & ",  Barcode ='" & Barcode & "'" & ", BZBB ='" & BZBB & _
                       "'" & ",CS ='" & CS & "',DGRQ ='" & DGRQ & "', isJiaJi ='" & isJiaJi & "', YJDCSJ ='" & YJDCSJ & "', SQBM ='" & SQBM & _
                       "',XQR ='" & XQR & "', reason ='" & reason & "',SRRYGH ='" & SRRYGH & "' where ID='" & ID & "'"
                        AddSql2 (strSql)
                        
                        count = count + 1
                    Else
                        '修改了主键光照版本
                        MsgBox "修改失败,不能修改光照版本号 第" & .Row & "行"

                    End If

                End If

                End If

        Next i

    End With

    If count = 0 Then
        MsgBox "修改失败"
    Else
        MsgBox "修改成功" & "修改记录数" & count & "! "
    
    End If

Query

End Sub

Private Sub Form_Load()
    DTP(0).Value = DATE
    DTP1(1).Value = DATE
    DTP2(1).Value = DATE
    DTP3(2).Value = DATE
    DTP4(0).Value = DATE
    DTP5(2).Value = DATE
    Dim strSql       As String

    Dim rs     As New ADODB.Recordset
     strSql = "select * from erptemp.dbo.KJQX where GZBBQX = '1' and GH = '" & gUserName & "'"

    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2

    End If
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        cmd_del.Visible = True
        CMD_Modify.Visible = True
        cmdFPQX.Visible = False
    Else
        cmd_del.Visible = False
        CMD_Modify.Visible = False
        cmdFPQX.Visible = False
    End If
    rs.Close

If gUserName = "07885" Then
   cmd_del.Visible = True
   CMD_Modify.Visible = True
   cmdFPQX.Visible = True
End If

    Fra.Visible = False  '查询界面关闭
End Sub

Private Sub ListDataType(rs As ADODB.Recordset)

    Dim i As Long

    With fpS(0)
        .MaxRows = 0
        Set .DataSource = rs

    End With

    With fpS(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            .BackColor = &HFFFF&
            .ColWidth(1) = 10
            .CellType = CellTypeCheckBox
            .Text = 0
            .Col = 2
            .Lock = True
            .Col = 13
            .Lock = True
            If gUserName <> "07885" Then
                .Col = 14
                .Lock = True
            End If
        Next
        
    End With
    rs.Close
End Sub
Private Function Query()
   '查询
    Dim strSql       As String

    Dim rs           As New ADODB.Recordset

    Dim lot         As String

    Dim Barcode      As String

    Dim BZBB          As String
    
    Dim CS         As String
    
    Dim SQBM       As String
    
    Dim DGRQKS       As String
    Dim DGRQJS     As String
    
    Dim isJiaJi    As String
    
    Dim XQBM   As String
    
    Dim XQR         As String
    
    Dim YJDCSJKS  As String
    Dim YJDCSJJS  As String
    
    Dim reason     As String
    
    

    lot = Trim(Text1.Text)
    Barcode = Trim(Text2.Text)
    BZBB = Trim(Text3.Text)
    CS = Trim(CScode.Text)
    DGRQKS = DTP2(1).Value
    DGRQJS = DTP3(2).Value
    YJDCSJKS = DTP4(0).Value
    YJDCSJJS = DTP5(2).Value
    SQBM = Trim(BMcode1.Text)
    XQR = Trim(Text4.Text)
    reason = Trim(Text5.Text)
    isJiaJi = Trim(Combo1.Text)

    strSql = "select '' AS '选择',ID,lot as '料号', Barcode as 'Barcode', BZBB as '备注版本',CS as '厂商'," & _
    "DGRQ as '订购日期',isJiaJi as '是否加急', YJDCSJ as '预计到厂时间',SQBM as '申请部门',XQR as '需求人'," & _
    "reason as '原因', Create_time as '创建时间',SRRYGH as '输入人员工号' from erptemp.dbo.MASK_CODE where isdel = '0'  "

    If Trim(Text1.Text) <> "" Then
        strSql = strSql + " AND LOT  = '" & Trim(Text1.Text) & "'  "

    End If

    If Trim(Text2.Text) <> "" Then
        strSql = strSql + " AND Barcode = '" & Trim(Text2.Text) & "'  "

    End If

    If Trim(Text3.Text) <> "" Then
        strSql = strSql + " AND BZBB  = '" & Trim(Text3.Text) & "'  "

    End If
    
    If Check5 = 1 Then
        If Trim(CScode1.Text) <> "" Then
            strSql = strSql + " AND CS  = '" & Trim(CScode1.Text) & "'  "
    
        End If
    End If
    
    If Check2 = 1 Then
        strSql = strSql + " AND DGRQ  >= '" & DGRQKS & "' and DGRQ <='" & DGRQJS & "'"
    End If
    
    If Check3 = 1 Then
        If Trim(Combo1.Text) <> "" Then
            strSql = strSql + " AND isJiaJi  = '" & Trim(Combo1.Text) & "'  "
    
        End If
    End If
    
    If Check6 = 1 Then
        If Trim(BMcode1.Text) <> "" Then
            strSql = strSql + " AND SQBM  = '" & Trim(BMcode1.Text) & "'  "
    
        End If
    End If
    
    If Trim(Text4.Text) <> "" Then
        strSql = strSql + " AND XQR  = '" & Trim(Text4.Text) & "'  "

    End If

    If Trim(Text5.Text) <> "" Then
        strSql = strSql + " AND reason  = '" & Trim(Text4.Text) & "'  "

    End If

    If Check4 = 1 Then
    strSql = strSql + " AND YJDCSJ >= '" & YJDCSJKS & "'AND YJDCSJ <= '" & YJDCSJJS & "'"
    End If
    
    strSql = strSql + "  order by Create_time desc,DGRQ desc,Barcode desc"
    
    strsqlEX = strSql '存给全局变量
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2

    End If

    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType(rs)
        Exit Function
    End If

End Function

'初始化
Private Function Initial()

    txtText1.Text = ""
    txtText2.Text = ""
    txtText3.Text = ""
    txtText4.Text = ""
    txtText5.Text = ""

End Function
Private Function query2()
    Dim strSql       As String

    Dim rs           As New ADODB.Recordset

    strSql = "select '' AS '选择',ID,lot as '料号', Barcode as 'Barcode', BZBB as '备注版本',CS as '厂商', DGRQ as '订购日期',isJiaJi as '是否加急', YJDCSJ as '预计到厂时间'," & _
    "SQBM as '申请部门',XQR as '需求人',reason as '原因',Create_time as '创建时间',SRRYGH as '输入人员工号' from erptemp.dbo.MASK_CODE where isdel = '0'  order by Create_time desc,DGRQ desc,Barcode desc "

    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2

    End If
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        rs.Close
        Exit Function

    End If
End Function

Private Function getBarcode() As String 'barcode 自动生成算法

Dim Barcode As String
Barcode = "HT" + Format(DATE, "YYYYMMDD")
'MsgBox "" & Barcode
getBarcode = Barcode
End Function

Private Function Check5_input(input_String As String) As Integer
    If InStr(input_String, "'") > 0 Or InStr(input_String, "’") > 0 Then
       ' MsgBox "输入字符串不合法"
        Check5_input = 1
    Else
        Check5_input = 0
    End If
End Function
Private Function Check5_date(input_String As String) As Integer
   If IsDate(input_String) = False Then
       ' MsgBox "输入字符串不合法"
        Check5_date = 1
    Else
        Check5_date = 0
    End If

End Function


Private Function CheckData()

MsgBox "功能尚未开放"

'    Dim i As Long
'
'    With Fps(0)
'        .MaxRows = 0
'        Set .DataSource = rs
'
'    End With
'
'    With Fps(0)
'
'        For i = 1 To .MaxRows
'            .Row = i
'            .Col = 1
'            .BackColor = &HFFFF&
'            .ColWidth(1) = 10
'            .CellType = CellTypeCheckBox
'            .Text = 0
'            .Col = 10
'            .Lock = False
'
'        Next
'
'    End With


End Function

