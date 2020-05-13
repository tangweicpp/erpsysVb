VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Frm_WORK_ORDER 
   BackColor       =   &H00E0E0E0&
   Caption         =   "新版开工单"
   ClientHeight    =   12660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16080
   DrawStyle       =   6  'Inside Solid
   FillColor       =   &H000000FF&
   FillStyle       =   6  'Cross
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
   ScaleHeight     =   12660
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame FraOrder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "工单明细"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   7695
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   5400
      Width           =   18375
      Begin VB.CommandButton cmdCreatRepWafer 
         BackColor       =   &H00FF80FF&
         Caption         =   "生成重工WaferID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdSaveOrder 
         BackColor       =   &H0080FF80&
         Caption         =   "保存工单"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   16560
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkAllWafers 
         BackColor       =   &H00E0E0E0&
         Caption         =   "全选"
         Height          =   195
         Left            =   13080
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdLeftToRight 
         BackColor       =   &H00FF80FF&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1560
         Width           =   990
      End
      Begin VB.CommandButton cmdLotFinder 
         BackColor       =   &H00FF80FF&
         Caption         =   "LOT搜索"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Width           =   990
      End
      Begin VB.TextBox txtLotIndex 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.ListBox lsLot 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   6000
         ItemData        =   "Frm_WORK_ORDER.frx":0000
         Left            =   360
         List            =   "Frm_WORK_ORDER.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
      Begin FPSpreadADO.fpSpread fpsWafers 
         Height          =   6135
         Index           =   0
         Left            =   3600
         TabIndex        =   12
         Top             =   840
         Width           =   12855
         _Version        =   524288
         _ExtentX        =   22675
         _ExtentY        =   10821
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "Frm_WORK_ORDER.frx":0004
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame FraOrder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "工单选项"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5175
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   19695
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   5640
         TabIndex        =   45
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   257490945
         CurrentDate     =   43290
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1800
         TabIndex        =   44
         Top             =   2550
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   257490945
         CurrentDate     =   43290
      End
      Begin VB.CommandButton cmdEXIT 
         BackColor       =   &H00FF80FF&
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox txtJobNo 
         Height          =   285
         Left            =   13080
         TabIndex        =   41
         Top             =   1925
         Width           =   1575
      End
      Begin VB.TextBox txtPO 
         Height          =   285
         Left            =   9240
         TabIndex        =   40
         Top             =   1925
         Width           =   1935
      End
      Begin VB.TextBox txtWO 
         Height          =   315
         Left            =   1800
         TabIndex        =   39
         Top             =   1910
         Width           =   1695
      End
      Begin VB.TextBox txtWorkOrder 
         Height          =   285
         Left            =   5640
         TabIndex        =   37
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtReproduction 
         Height          =   285
         Left            =   13080
         TabIndex        =   35
         Top             =   1275
         Width           =   1815
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00C0C0C0&
         Caption         =   "重置"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7520
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CommandButton cmdMakeLot 
         BackColor       =   &H008080FF&
         Caption         =   "生成LOT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "DUMMY工单,硅基工单,玻璃工单"
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox txtPieces 
         Height          =   285
         Left            =   16200
         TabIndex        =   31
         Top             =   1260
         Width           =   975
      End
      Begin VB.TextBox txtPlantDevice 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5640
         TabIndex        =   26
         Top             =   1925
         Width           =   1815
      End
      Begin VB.TextBox txtOrderDept 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9240
         TabIndex        =   24
         Top             =   1255
         Width           =   2535
      End
      Begin VB.ComboBox cbOrderType 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1240
         Width           =   1695
      End
      Begin VB.ComboBox cbLotType 
         Height          =   315
         Left            =   17760
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   570
         Width           =   1815
      End
      Begin VB.ComboBox cbPri 
         Height          =   315
         Left            =   12360
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   570
         Width           =   1455
      End
      Begin VB.CommandButton CmdQueryLot 
         BackColor       =   &H00FF8080&
         Caption         =   "查询LOT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "普通工单"
         Top             =   3840
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo dcProductNo 
         Height          =   315
         Left            =   9240
         TabIndex        =   6
         Top             =   570
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcCusDevice 
         Height          =   315
         Left            =   5640
         TabIndex        =   4
         Top             =   570
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcCusCode 
         Height          =   315
         Left            =   1800
         TabIndex        =   38
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblJob 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JOB_ID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   12360
         TabIndex        =   42
         Top             =   1965
         Width           =   660
      End
      Begin VB.Label lblReproduction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "重工LOT_ID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   12000
         TabIndex        =   34
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Label lblCnt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WAFER片数:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   15120
         TabIndex        =   30
         Top             =   1305
         Width           =   1020
      End
      Begin VB.Label lplPo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO_NUM:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8355
         TabIndex        =   29
         Top             =   1970
         Width           =   750
      End
      Begin VB.Label lblProduceStop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预计完工日期:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4320
         TabIndex        =   28
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblProduceStart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "计划开工日期:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   27
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblPlantDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "厂内机种:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4710
         TabIndex        =   25
         Top             =   1970
         Width           =   825
      End
      Begin VB.Label lblOrderDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单部门:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8280
         TabIndex        =   23
         Top             =   1300
         Width           =   825
      End
      Begin VB.Label lblOrderType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单类型:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   870
         TabIndex        =   20
         Top             =   1300
         Width           =   825
      End
      Begin VB.Label lblLotType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "量产批(M)/工程批(E)/客户实验(Q)/DC片(D)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   13920
         TabIndex        =   18
         Top             =   630
         Width           =   3720
      End
      Begin VB.Label LblPri 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRI:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   12000
         TabIndex        =   16
         Top             =   630
         Width           =   345
      End
      Begin VB.Label LblWo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "订单单号:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   870
         TabIndex        =   15
         Top             =   1970
         Width           =   825
      End
      Begin VB.Label LblWorkOrder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单号:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4905
         TabIndex        =   14
         Top             =   1300
         Width           =   630
      End
      Begin VB.Label lblProductNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "产品料号:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   8280
         TabIndex        =   5
         Top             =   630
         Width           =   825
      End
      Begin VB.Label lblCustomerDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户机种:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   4560
         TabIndex        =   3
         Top             =   630
         Width           =   825
      End
      Begin VB.Label lblCustomerCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   870
         TabIndex        =   2
         Top             =   630
         Width           =   825
      End
   End
End
Attribute VB_Name = "Frm_WORK_ORDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' 全局变量声明
Dim rsCusCode   As New ADODB.Recordset

Dim rsCusDevice As New ADODB.Recordset

Dim rsProductNo As New ADODB.Recordset

Dim rsCusLot    As New ADODB.Recordset

Dim rsCusWafer  As New ADODB.Recordset

Dim rsOrderChk  As New ADODB.Recordset

Dim aWOID()     As String

Private Enum E_FPS0          'WaferDetails

    E_ID = 1                 'id
    E_WAFERID                'Waferid
    E_NewWaferId             'NewWaferid
    E_TotalDie               '总数量
    E_GoodDie                'good数量
    E_WaferLot               'wafer
    E_MARKINGCODE            'markingcode
    E_OK                     '选择
    
    E_END

End Enum

Private Sub cbOrderType_Click()

    LblWorkOrder.Visible = True
    txtWorkOrder.Visible = True
    lblOrderDept.Visible = True
    txtOrderDept.Visible = True
    lblCustomerCode.Visible = True
    dcCusCode.Visible = True
    lblCustomerDevice.Visible = True
    dcCusDevice.Visible = True
    lblProductNumber.Visible = True
    dcProductNo.Visible = True
    LblPri.Visible = True
    cbPri.Visible = True
    lblLotType.Visible = True
    cbLotType.Visible = True
    LblWo.Visible = True
    txtWO.Visible = True
    lblPlantDevice.Visible = True
    txtPlantDevice.Visible = True
    lplPo.Visible = True
    txtPO.Visible = True
    lblProduceStart.Visible = True
    lblProduceStop.Visible = True

    Select Case cbOrderType.text

        Case "Dummy工单"
            cmdMakeLot.Visible = True
            CmdQueryLot.Visible = False
            lblCnt.Visible = True
            txtPieces.Visible = True
            lblReproduction.Visible = False
            txtReproduction.Visible = False
            dcCusDevice.Enabled = False
            dcCusDevice.text = ""
            lblJob.Visible = False
            txtJobNo.Visible = False
            cmdSaveOrder.Enabled = True
            cmdCreatRepWafer.Visible = False

        Case "玻璃工单"
            cmdMakeLot.Visible = True
            CmdQueryLot.Visible = False
            lblCnt.Visible = True
            txtPieces.Visible = True
            lblReproduction.Visible = False
            txtReproduction.Visible = False
            dcCusDevice.Enabled = False
            dcCusDevice.text = ""
            lblJob.Visible = False
            txtJobNo.Visible = False
            cmdSaveOrder.Enabled = True
            cmdCreatRepWafer.Visible = False

        Case "硅基工单"
            cmdMakeLot.Visible = True
            CmdQueryLot.Visible = False
            lblCnt.Visible = True
            txtPieces.Visible = True
            lblReproduction.Visible = False
            txtReproduction.Visible = False
            dcCusDevice.Enabled = False
            dcCusDevice.text = ""
            lblJob.Visible = False
            txtJobNo.Visible = False
            cmdSaveOrder.Enabled = True
            cmdCreatRepWafer.Visible = False
            
        Case "FO_CSP工单"
            cmdMakeLot.Visible = True
            CmdQueryLot.Visible = False
            lblCnt.Visible = True
            txtPieces.Visible = True
            lblReproduction.Visible = False
            txtReproduction.Visible = False
            dcCusDevice.Enabled = False
            dcCusDevice.text = ""
            lblJob.Visible = False
            txtJobNo.Visible = False
            cmdSaveOrder.Enabled = True
            cmdCreatRepWafer.Visible = False

        Case "重工工单"
            dcCusDevice.Enabled = True
            CmdQueryLot.Visible = True
            cmdMakeLot.Visible = False
            lblCnt.Visible = True
            txtPieces.Visible = True
            lblReproduction.Visible = True
            txtReproduction.Visible = True
            lblJob.Visible = True
            txtJobNo.Visible = True
            cmdSaveOrder.Enabled = False
            cmdCreatRepWafer.Visible = True

        Case Else
            CmdQueryLot.Visible = True
            cmdMakeLot.Visible = False
            lblCnt.Visible = False
            txtPieces.Visible = False
            lblReproduction.Visible = False
            txtReproduction.Visible = False
            lblJob.Visible = False
            txtJobNo.Visible = False
            cmdSaveOrder.Enabled = True
            cmdCreatRepWafer.Visible = False

    End Select

End Sub

Private Sub chkAllWafers_Click()

    Dim i As Integer

    If chkAllWafers.Value = 1 Then

        For i = 1 To fpsWafers(0).MaxRows

            With fpsWafers(0)
                .Row = i
                .Col = E_FPS0.E_OK
                .text = 1

            End With

        Next i
        
    ElseIf chkAllWafers.Value = 0 Then

        For i = 1 To fpsWafers(0).MaxRows

            With fpsWafers(0)
                .Row = i
                .Col = E_FPS0.E_OK
                .text = 0

            End With

        Next i
        
    End If

End Sub

Private Sub cmdArrayMake_Click()

    ' 0. 清空临时表
    Call ClearOrderTmp
    ' 1. 把勾选的wafer明细存入临时表ORDER_TEMP
    Call TransDataToTmp
    ' 2. 按工单显示明细
    Call ShowDetailByOrderID

    Call BackUpOrderData

    Call ShowThisData

End Sub

Private Sub ShowDetailByOrderID()

    Dim sOra As String

    Dim rs   As ADODB.Recordset

    sOra = "select distinct WORK_ORDER_ID from  WORK_ORDER_TMP order by WORK_ORDER_ID"
    Set rs = Get_OracleRs(sOra)

    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            txtWorkOrder.text = rs.Fields(0).Value
            Call ShowRealOrderDetails(txtWorkOrder.text)
        
            ' 执行开工单
            Call ORDER_Handle
      
            Sleep (5000)
            rs.MoveNext
        Loop

    End If

End Sub

Private Sub BackUpOrderData()

    Dim sOra As String

    sOra = "insert into WORK_ORDER_HISTORY select * from WORK_ORDER_TMP"
    Exec_Ora (sOra)

End Sub

Private Sub ShowThisData()

    Dim sOra As String

    sOra = "select work_order_id as 工单号, lot_id, wafer_id, marking_code, totaldie, gooddie from WORK_ORDER_TMP"

    ExporToExcel (sOra)

End Sub

Private Sub ShowRealOrderDetails(sOrderId As String)

    Dim sOra As String

    Dim rs   As ADODB.Recordset

    sOra = "select id,wafer_id,wafer_id,totaldie,gooddie,lot_id,marking_code,work_order_id,1 from WORK_ORDER_TMP  where work_order_ID = '" & sOrderId & "' "
    Set rs = Get_OracleRs(sOra)

    fpsWafers(0).MaxRows = rs.RecordCount

    For i = 0 To rs.RecordCount - 1

        With fpsWafers(0)
            Set .DataSource = Nothing
            .Row = i + 1
            
            .Col = E_FPS0.E_ID
            .text = CStr(rs.Fields(0).Value)

            .Col = E_FPS0.E_WAFERID
            .text = CStr(rs.Fields(1).Value)

            .Col = E_FPS0.E_NewWaferId
            .text = CStr(rs.Fields(2).Value)

            .Col = E_FPS0.E_TotalDie
            .text = CStr(rs.Fields(3).Value)

            .Col = E_FPS0.E_GoodDie
            .text = CStr(rs.Fields(4).Value)

            .Col = E_FPS0.E_WaferLot
            .text = CStr(rs.Fields(5).Value)
            
            .Col = E_FPS0.E_MARKINGCODE
            .text = CStr(rs.Fields(6).Value)
        
            .Col = E_FPS0.E_OK
            .text = CStr(rs.Fields(8).Value)

        End With
    
        rs.MoveNext
    Next

End Sub

Private Sub ClearOrderTmp()
    Exec_Ora ("DELETE FROM WORK_ORDER_TMP")

End Sub

Private Sub TransDataToTmp()

    Dim wd As WORKORDER_DATA

    With fpsWafers(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_FPS0.E_OK

            If .text = 1 Then
                .Col = 1
                wd.ID = .text
                .Col = 3
                wd.Wafer_id = .text
                .Col = 4
                wd.TOTALDIE = .text
                .Col = 5
                wd.gooddie = .text
                .Col = 6
                wd.Lot_id = .text
                .Col = 7
                wd.MARKING_CODE = .text
                .Col = 8
                wd.WORK_ORDER_ID = .text
        
                Call InsertOrderToTmpTbl(wd)

            End If
        
        Next

    End With

End Sub

Private Sub InsertOrderToTmpTbl(wd As WORKORDER_DATA)

    Dim sOra As String

    sOra = "Insert into WORK_ORDER_TMP(ID, WORK_ORDER_ID, LOT_ID, WAFER_ID, TOTALDIE, gooddie, MARKING_CODE) values('" & wd.ID & "', '" & wd.WORK_ORDER_ID & "', '" & wd.Lot_id & "', '" & wd.Wafer_id & "', '" & wd.TOTALDIE & "', '" & wd.gooddie & "', '" & wd.MARKING_CODE & "')"
    Exec_Ora (sOra)

End Sub

Private Function CheckWorkOrderData(wd As WORKORDER_DATA) As Boolean
    CheckWorkOrderData = False

    CheckWorkOrderData = True

End Function

Private Sub cmdCreatRepWafer_Click()

    Dim bFlag As Boolean

    Dim sVal  As String

    FLAG = False

    With fpsWafers(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_FPS0.E_OK

            If .text = 1 Then
                bFlag = True

            End If

        Next i

    End With

    If bFlag = False Then
        MsgBox "请勾选wafer", vbInformation, "友情提示"
        Exit Sub

    End If

    InitRepSelWaferData

End Sub

Private Sub InitRepSelWaferData()

    Dim iWaferQty    As Integer

    Dim iGrossDies   As Long

    Dim sWaferNo     As String

    Dim sOra         As String

    Dim sLotId       As String

    Dim sWafer       As String

    Dim sOldWaferId  As String

    Dim bCheckFlag   As Boolean

    Dim sLastWaferNo As String
    
    Dim sLastLotNo   As String

    Dim sNextWaferNo As String

    Dim sSubstriteID As String
    
    sLastWaferNo = ""
    sLastLotNo = ""
    sNextWaferNo = ""
    bCheckFlag = False
    bIsNull = False
    sWafer = ""
    sWaferList = ""

    With fpsWafers(0)

        For i = 0 To fpsWafers(0).MaxRows - 1
            
            .Row = i + 1
            
            ' lotID
            .Col = 6
            sLotId = Trim$(.text)
            
            ' WaferNO
            .Col = 1
            If .text <> "" Then
                sWaferNo = Trim(.text)

                If IsNumeric(sWaferNo) = False Then
                    MsgBox "WaferNo请输入数字", vbInformation, "友情提示"
                    Exit Sub
                Else

                    If CInt(sWaferNo) < 1 Or CInt(sWaferNo) > 25 Then
                        MsgBox "WaferNo请输入1-25", vbInformation, "友情提示"
                        Exit Sub

                    End If

                End If

                If Left$(sWaferNo, 1) = "0" Then
                    sWafer = Replace$(sWaferNo, "0", "")
                Else
                    sWafer = sWaferNo

                End If

            Else
                MsgBox "请输入wafer的序号, 范围(1-25)", vbInformation, "友情提示"
                Exit Sub

            End If
            
            ' Old WaferID
            .Col = 2
            .text = Get_OracleStr("select max(substrateid) from mappingdatatest where wafer_id in ('" & sWafer & "', '0'||'" & sWafer & "') and lotid = '" & sLotId & "'")
            sOldWaferId = .text

            If sOldWaferId = "" Then
    
                MsgBox "查询不到该Wafer:" & sLotId & Right("0" & sWafer, 2), vbCritical, "警告"
                Exit Sub

            End If

            ' New WaferID
            .Col = 3
            .text = sOldWaferId & "+"
            
            .Col = 4
            If .text = "" Then
                MsgBox "TatolDies的数量不可以为空", vbInformation, "友情提示"
                Exit Sub
            Else
                .text = Trim(.text)

                If IsNumeric(.text) = False Then
                    MsgBox "TatolDies请输入数字", vbInformation, "友情提示"
                    Exit Sub
                Else
                    iGrossDies = CLng(.text)

                    If iGrossDies < 1 Then
                        MsgBox "Die数量不可以小于1", vbInformation, "友情提示"
                        Exit Sub

                    End If

                End If

                If CheckGrossDie(iGrossDies, sLotId) = False Then
                    MsgBox "TatolDies数量输入有误,不可以大于实际数量", vbInformation, "友情提示"
                    Exit Sub

                End If

            End If

            .Col = 5
            If .text = "" Then
                MsgBox "GoodDies的数量不可以为空", vbInformation, "友情提示"
                Exit Sub
            Else
                .text = Trim(.text)

                If IsNumeric(.text) = False Then
                    MsgBox "GoodDies请输入数字", vbInformation, "友情提示"
                    Exit Sub
                Else
                    iGrossDies = CLng(.text)

                    If iGrossDies < 1 Then
                        MsgBox "Die数量不可以小于1", vbInformation, "友情提示"
                        Exit Sub

                    End If

                End If

                iGoodDies = CLng(.text)

                If iGoodDies <> iGrossDies Then
                    MsgBox "GoodDies数量应该等于GrossDies数量, 请重新输入", vbInformation, "友情提示"
                    Exit Sub

                End If

            End If
            
            ' 打标码
            .Col = 7
            sSubstriteID = sLotId & Right$("0" & sWafer, 2)
            
            .text = Get_OracleStr("select productid from mappingdatatest where substrateid = '" & sSubstriteID & "' ")
            
        Next i

    End With

    cmdSaveOrder.Enabled = True

End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdLotFinder_Click()

    Dim sLot    As String

    Dim i       As Integer

    Dim resFlag As Boolean

    resFlag = False

    sLot = UCase(Trim(txtLotIndex.text))

    With lsLot

        For i = 0 To .ListCount - 1

            If .List(i) = sLot Then
                .Selected(i) = True
                resFlag = True
                Exit For

            End If

        Next

    End With

    If resFlag = False Then
        MsgBox "搜索不到, 请确认LOT号是否有误", vbInformation, "友情提示"
        Exit Sub

    End If

    txtLotIndex.text = ""
    txtLotIndex.SetFocus

End Sub

Private Sub cmdMakeLot_Click()

    Dim sOrderType As String  ' 工单类型

    Dim sCusCode   As String    ' 客户代码

    Dim sCusDevice As String  ' 客户机种

    Dim sProductNo As String  ' 产品料号

    Dim sWaferQty  As String   ' lot数量

    Dim iLotQty    As Long

    Dim iWaferQty  As Long
    
    Dim rs As New ADODB.Recordset
    Dim iWaferQtyP As Long

    Dim sOra       As String

    Dim sSql       As String

    sOrderType = Trim(cbOrderType.text)
    sCusCode = Trim(dcCusCode.text)
    sCusDevice = Trim(dcCusDevice.text)
    sProductNo = Trim(dcProductNo.text)
    sWaferQty = Trim(txtPieces.text)

    If (sCusCode = "") Or (sCusDevice = "") Or (sProductNo = "") Then
        MsgBox "请选择客户代码,客户机种,产品料号", vbInformation, "友情提示"
        Exit Sub

    End If
    
    sOra = "select *  from tbltsvnpiproduct where customershortname = '" & sCusCode & "' and customerptno1 = '" & sCusDevice & "' and qtechptno2 = '" & sProductNo & "'"
    Set rs = Get_OracleRs(sOra)
    
    If rs.RecordCount = 0 Then
        MsgBox "NPI没有维护该机种料号对应关系, 请确认", vbCritical, "警告"
        Exit Sub
    End If
    
    
    ' 0.清空
    lsLot.Clear

    ' 1.判断
    If sWaferQty = "" Then
        MsgBox "请输入wafer片数", vbInformation, "友情提示"
        Exit Sub
    Else

        If IsNumeric(sWaferQty) = False Then
            MsgBox "请输入数字", vbInformation, "友情提示"
            Exit Sub
        Else

            If CLng(sWaferQty) < 1 Then
                MsgBox "请输入至少1片wafer数量", vbInformation, "友情提示"
                Exit Sub

            End If

        End If

    End If

    ' 2.数据整合
    iWaferQty = CLng(sWaferQty)
    iWaferQtyP = iWaferQty

    iLotQty = IIf((iWaferQty Mod 25) = 0, iWaferQty \ 25, iWaferQty \ 25 + 1)

    ' 清空临时表
    sOra = "delete from ORDER_DATA_TEMP_HEADER"
    sSql = "delete from erpdata.dbo.ORDER_DATA_TEMP_HEADER"
    Call Get_OracleRs(sOra)
    Call Get_SqlserveRs(sSql)

    sOra = "delete from ORDER_DATA_TEMP_DETAILS"
    sSql = "delete from erpdata.dbo.ORDER_DATA_TEMP_DETAILS"
    Call Get_OracleRs(sOra)
    Call Get_SqlserveRs(sSql)

    ' 3.打印结果
    If iLotQty > 1 Then

        For i = 1 To (iLotQty - 1)
            lsLot.AddItem (Insert_WoTbl(cbOrderType.text, dcCusCode.text, dcProductNo.text, 25, UCase(Trim(dcCusDevice.text))))
        Next
    
        iWaferQty = IIf((iWaferQty Mod 25) = 0, 25, iWaferQty Mod 25)
        lsLot.AddItem (Insert_WoTbl(cbOrderType.text, dcCusCode.text, dcProductNo.text, iWaferQty, UCase(Trim(dcCusDevice.text))))
    Else
        iWaferQty = IIf((iWaferQty Mod 25) = 0, 25, iWaferQty Mod 25)
        lsLot.AddItem (Insert_WoTbl(cbOrderType.text, dcCusCode.text, dcProductNo.text, iWaferQty, UCase(Trim(dcCusDevice.text))))

    End If

    MsgBox "成功生成" & iLotQty & "个Lot" & ",共有 " & iWaferQtyP & "片wafers", vbInformation, "友情提示"

End Sub

Private Sub CmdQueryLot_Click()
    QueryHandle

End Sub

Private Sub QueryHandle()

    Dim sOrderType As String  ' 工单类型

    Dim sCusCode   As String    ' 客户代码

    Dim sCusDevice As String  ' 客户机种

    Dim sProductNo As String  ' 产品料号

    Dim sRepLotId  As String   ' 重工lotid

    Dim sWaferQty  As String   ' wafer数量

    Dim sOra       As String

    sOrderType = cbOrderType.text
    sCusCode = dcCusCode.text
    sCusDevice = dcCusDevice.text
    sProductNo = dcProductNo.text
    sRepLotId = UCase(Trim(txtReproduction.text))
    sWaferQty = Trim(txtPieces.text)

    If (sCusCode = "") Or (sCusDevice = "") Or (sProductNo = "") Then
        MsgBox "请选择客户代码,客户机种,产品料号", vbInformation, "友情提示"
        Exit Sub

    End If

    ' npi维护校验
    sOra = "select * from tbltsvnpiproduct  where qtechptno2 = '" & sProductNo & "' and customerptno1 = '" & sCusDevice & "'"

    If Get_OracleStr(sOra) = "" Then
        MsgBox "npi没有维护该机种和料号的对应关系, 请联系npi维护", vbInformation, "友情提示"
        Exit Sub

    End If

    ' 1.清空
    lsLot.Clear

    ' 2.判断
    If sOrderType = "重工工单" Then

        ' 重工工单
        If sWaferQty = "" Or sRepLotId = "" Then
            MsgBox "请输入wafer片数,重工LOTID", vbInformation, "友情提示"
            Exit Sub
        Else

            If IsNumeric(sWaferQty) = False Then
                MsgBox "请输入数字", vbInformation, "友情提示"
                Exit Sub
            Else

                If CLng(sWaferQty) < 1 Then
                    MsgBox "请输入至少1片wafer数量", vbInformation, "友情提示"
                    Exit Sub
                Else

                    If CLng(sWaferQty) > 25 Then
                        MsgBox "单LOT不可以开大于25片的wafer", vbInformation, "友情提示"
                        Exit Sub

                    End If

                    If sCusCode = "37" Then
                        If txtJobNo.text = "" Then
                            MsgBox "37请输入JOB号", vbInformation, "友情提示"
                            Exit Sub

                        End If

                    End If

                End If

            End If

        End If

        sOra = "select distinct ct.source_batch_id from customeroitbl_test ct, mappingDataTest mt where  " & " ct.customershortname = '" & sCusCode & "' and ct.source_batch_id = mt.lotid and ct.source_batch_id = '" & sRepLotId & "' and ct.id = mt.filename"
       
        ' 2.打印结果
        Set rsCusLot = Get_OracleRs(sOra)

        If rsCusLot.RecordCount > 0 Then
    
            With lsLot
                .Clear
    
                rsCusLot.MoveFirst
        
                For i = 0 To rsCusLot.RecordCount - 1
                    .AddItem "" & rsCusLot!source_batch_id
                    .Selected(lsLot.ListCount - 1) = True
                    rsCusLot.MoveNext
                Next

            End With

        Else
            MsgBox "查询不到LOT"
    
            With lsLot
                .Clear

            End With

        End If

    Else
        ' 其他工单
        sOra = "select distinct ct.source_batch_id from customeroitbl_test ct, mappingDataTest mt where ct.mpn_desc = '" & sCusDevice & "' " & "and ct.customershortname = '" & sCusCode & "' and ct.source_batch_id = mt.lotid and not exists(select 1 from ib_waferlist al " & "where al.waferid = mt.substrateid)order by ct.source_batch_id "
        
        ' 2.打印结果
        Set rsCusLot = Get_OracleRs(sOra)

        If rsCusLot.RecordCount > 0 Then
    
            With lsLot
                .Clear
                rsCusLot.MoveFirst
        
                For i = 0 To rsCusLot.RecordCount - 1
                    .AddItem "" & rsCusLot!source_batch_id
                    rsCusLot.MoveNext
                Next

            End With

        Else
            MsgBox "查询不到LOT"
    
            With lsLot
                .Clear

            End With

        End If
    
    End If

End Sub

Private Sub cmdLeftToRight_Click()

    Dim sLot              As String

    Dim sLotList          As String

    Dim sCusCode          As String

    Dim sWorkOrder        As String

    Dim sCusDevice        As String

    Dim bSpecialOrderFlag As Boolean

    Dim iCnt              As Integer

    iCnt = 0
    bSpecialOrderFlag = False

    sLotList = ""
    sCusCode = Trim(dcCusCode.text)
    sWorkOrder = Trim(txtWorkOrder.text)
    sCusDevice = Trim$(dcCusDevice.text)

    With lsLot

        For i = 0 To .ListCount - 1

            If .Selected(i) Then
                iCnt = iCnt + 1
                sLot = .List(i) & "','"
                sLotList = sLotList & sLot
            Else
                bSpecialOrderFlag = True

            End If

        Next

    End With

    ReDim aWOID(iCnt)
 
    If sLotList = "" Then
        MsgBox "请先勾选LotId !", vbInformation, "友情提醒"
        Exit Sub
    Else

        If cbOrderType.text = "Dummy工单" Or cbOrderType.text = "玻璃工单" Or cbOrderType.text = "硅基工单" Or cbOrderType.text = "FO_CSP工单" Then
            If bSpecialOrderFlag Then
                MsgBox "该工单类型生成的lot必须全部开掉, 请全部勾选", vbInformation, "友情提示"
                Exit Sub

            End If
        
        End If

    End If
 
    sLotList = Mid(sLotList, 1, Len(sLotList) - 3)
 
    Call SetFpsWafersDetails(sLotList, sCusCode, sWorkOrder, sCusDevice)

End Sub

Private Sub SetFpsWafersDetails(sLotList As String, _
                                sCusCode As String, _
                                sWorkOrder As String, _
                                sCusDevice As String)

    Dim i          As Integer

    Dim sumQty     As Long

    Dim woType     As String

    Dim htTemp     As String

    Dim yearpart   As String

    Dim monthpart  As String

    Dim lotpart    As String

    Dim wfnum      As String

    Dim wfpart     As String

    Dim iWaferQty  As Long

    Dim sOrderType As String

    Dim sJob       As String

    sOrderType = cbOrderType.text

    If txtPieces.text <> "" Then
        iWaferQty = CLng(Trim$(txtPieces.text))

    End If

    sJob = Trim(txtJobNo.text)

    Set rsCusWafer = Get_OrderDetailsFps(sLotList, sCusCode, sOrderType, sJob)

    If rsCusWafer.RecordCount <= 0 Then
        MsgBox "明细表中没有相关数据，请确认"
        Exit Sub

    End If

    Call InitFps(sOrderType, iWaferQty, rsCusWafer)

End Sub

Private Sub InitFps(sOrderType As String, _
                    iWaferQty As Long, _
                    rsCusWafer As ADODB.Recordset)

    Dim iWaferQtyReal As Integer

    Dim sOra          As String

    Dim sLotId        As String

    Dim iLastRows     As Integer

    iLastRows = fpsWafers(0).MaxRows
    
    sLotId = UCase$(Trim$(txtReproduction.text))
    sOra = "select distinct wafer_id from mappingdatatest where lotid = '" & sLotId & "'"

    iWaferQtyReal = Get_OracleCnt(sOra)

    If sOrderType = "重工工单" Then

        ' check 数量
        If iWaferQty > iWaferQtyReal Then
            MsgBox "WO里一次的Wafer片数小于需求的片数, 请重新确认", vbInformation, "友情提示"
            Exit Sub

        End If
    
        fpsWafers(0).MaxRows = iWaferQty + iLastRows

        For i = 0 To iWaferQty - 1

            With fpsWafers(0)
            
                .Row = i + 1 + iLastRows
        
                .Col = E_FPS0.E_ID
                .text = ""
                .Lock = False
                .CellType = CellTypeEdit
                .BackColorStyle = BackColorStyleUnderGrid
                .BackColor = vbCyan
            
                .Col = E_FPS0.E_WAFERID
                .text = ""
            
                .Col = E_FPS0.E_NewWaferId
                .text = ""

                .Col = E_FPS0.E_TotalDie
                .text = ""
                .Lock = False
                .CellType = CellTypeEdit
                .BackColorStyle = BackColorStyleUnderGrid
                .BackColor = vbCyan

                .Col = E_FPS0.E_GoodDie
                .text = ""
                .Lock = False
                .CellType = CellTypeEdit
                .BackColorStyle = BackColorStyleUnderGrid
                .BackColor = vbCyan
            
                .Col = E_FPS0.E_WaferLot
                .text = CStr(rsCusWafer.Fields(5).Value)

                .Col = E_FPS0.E_MARKINGCODE
                .text = ""
        
                .Col = E_FPS0.E_OK
                .text = CStr("1")

            End With
    
            rsCusWafer.MoveNext
        Next
    ElseIf sOrderType = "Dummy工单" Or sOrderType = "玻璃工单" Or sOrderType = "硅基工单" Or sOrderType = "FO_CSP工单" Then

        fpsWafers(0).MaxRows = rsCusWafer.RecordCount

        For i = 0 To rsCusWafer.RecordCount - 1

            With fpsWafers(0)
                .Row = i + 1
        
                .Col = E_FPS0.E_ID
                .text = CStr(rsCusWafer.Fields(0).Value)

                .Col = E_FPS0.E_WAFERID
                .text = CStr(rsCusWafer.Fields(1).Value)

                .Col = E_FPS0.E_NewWaferId
                .text = CStr(rsCusWafer.Fields(1).Value)

                .Col = E_FPS0.E_TotalDie
                .text = CStr(rsCusWafer.Fields(3).Value)
                .Lock = False
                .CellType = CellTypeEdit
                .BackColorStyle = BackColorStyleUnderGrid
                .BackColor = vbCyan

                .Col = E_FPS0.E_GoodDie
                .text = CStr(rsCusWafer.Fields(4).Value)
                .Lock = False
                .CellType = CellTypeEdit
                .BackColorStyle = BackColorStyleUnderGrid
                .BackColor = vbCyan

                .Col = E_FPS0.E_WaferLot
                .text = CStr(rsCusWafer.Fields(5).Value)

                .Col = E_FPS0.E_MARKINGCODE
                .text = "" & rsCusWafer.Fields(6).Value
        
                .Col = E_FPS0.E_OK
                .text = CStr("1")

            End With
    
            rsCusWafer.MoveNext
        Next
    
    Else
        fpsWafers(0).MaxRows = rsCusWafer.RecordCount
    
        Dim sLastLot   As String

        Dim sLastOrder As String
    
        sLastLot = ""
        sLastOrder = ""
    
        For i = 0 To rsCusWafer.RecordCount - 1

            With fpsWafers(0)
                .Row = i + 1
            
                .Col = E_FPS0.E_ID
                .text = CStr(rsCusWafer.Fields(0).Value)

                .Col = E_FPS0.E_WAFERID
                .text = CStr(rsCusWafer.Fields(1).Value)

                .Col = E_FPS0.E_NewWaferId
                .text = CStr(rsCusWafer.Fields(1).Value)

                .Col = E_FPS0.E_TotalDie
                .text = CStr(rsCusWafer.Fields(3).Value)

                .Col = E_FPS0.E_GoodDie
                .text = CStr(rsCusWafer.Fields(4).Value)

                .Col = E_FPS0.E_WaferLot
                .text = CStr(rsCusWafer.Fields(5).Value)
    
                .Col = E_FPS0.E_MARKINGCODE
                .text = "" & rsCusWafer.Fields(6).Value
        
                .Col = E_FPS0.E_OK
                .text = CStr("1")

            End With
    
            rsCusWafer.MoveNext
        Next

    End If

End Sub

Private Sub cmdReset_Click()
    Unload Me
    Frm_WORK_ORDER.Show

End Sub

Private Sub cmdSaveOrder_Click()
    ORDER_Handle

End Sub

Private Sub ORDER_Handle()

    Dim tOrderHeader    As BillHeader

    Dim tOrderDetails() As BillDetail

    Dim tRepData        As ReproductionWaferData

    Dim iOrdertype      As Integer

    Dim i               As Integer

    Dim j               As Integer

    Dim sSubstrateid    As String

    Dim sLotId          As String

    Dim sOra            As String

    Dim iGrossDies      As Long

    Dim iGoodDies       As Long

    Dim bRet1           As Boolean

    Dim bRet2           As Boolean
    
    
    cmdSaveOrder.Enabled = False
    Call addLogTxt(UCase(Trim(txtWorkOrder.text)), " 点击保存按钮 ")
    
    i = 0 ' 总的wafer片数
    j = 0 ' 选择的wafer片数

    If cbLotType.text = "" Then
        MsgBox "请选择LOT类型工单批次", vbInformation, "提示:"
        Exit Sub

    End If

    If cbOrderType.text = "" Then
        MsgBox "工单类型请选择", vbInformation, "友情提示"
        cmdSaveOrder.Enabled = True
        Exit Sub

    End If

    Select Case cbOrderType.text

        Case "一般工单"
            iOrdertype = 1

        Case "再加工工单"
            iOrdertype = 5

        Case "委外工单"
            iOrdertype = 7

        Case "重工委外工单"
            iOrdertype = 8

        Case "拆件式工单"
            iOrdertype = 11

        Case "预测工单"
            iOrdertype = 13

        Case "试产工单"
            iOrdertype = 15

        Case Else
            iOrdertype = 0

    End Select

    With fpsWafers(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_FPS0.E_OK

            If .text = 1 Then
                j = j + 1
                .Row = i
                .Col = 4
                sumQty = sumQty + CLng(.text)

            End If

        Next i

    End With

    If cbOrderType.text = "Dummy工单" Or cbOrderType.text = "玻璃工单" Or cbOrderType.text = "硅基工单" Or cbOrderType.text = "重工工单" Or cbOrderType.text = "FO_CSP工单" Then
    
        With fpsWafers(0)
            i = .MaxRows

        End With

        If i <> j Then
            MsgBox "该工单类型生成的wafer必须全部开掉, 请全部勾选", vbInformation, "友情提示"
            cmdSaveOrder.Enabled = True
            Exit Sub

        End If

    End If

    tOrderHeader.QTY = sumQty
    tOrderHeader.ID = GetSeqID()    ' 序列号:
    tOrderHeader.ORDERNAME = Replace(UCase(Trim(txtWorkOrder.text)), Chr(13) + Chr(10), "")        ' 工单号: 去除空格回车换行字符
    tOrderHeader.PARA8 = Trim(txtOrderDept.text)               ' 工单部门:
    tOrderHeader.ORDERTYPE = iOrdertype                  ' 工单类型
    tOrderHeader.EVENTTYPE = "CREATED"
    tOrderHeader.ERPUSER = "Auto"
    tOrderHeader.product = Trim(dcProductNo.text)              ' 料号
    tOrderHeader.RequestDate = Now
    tOrderHeader.ERPCREATEDATE = DateTime.Now
    tOrderHeader.PLANSTARTDATE = DTPicker1.Value
    tOrderHeader.PLANENDDATE = DTPicker2.Value
    tOrderHeader.CUSTOMER = Replace(dcCusCode.text, "(ICI)", "")
    tOrderHeader.SALESORDER = ""
    tOrderHeader.MODIFYFLAG = 0
    tOrderHeader.CustomerERPN = Trim(dcCusDevice.text)
    tOrderHeader.NGFLAG = 0
    tOrderHeader.PARA2 = 25
    tOrderHeader.sPri = cbPri.text
    tOrderHeader.sLotType = cbLotType.text

    ' OrderDetails数据初始化
    ReDim tOrderDetails(j)
    j = 0

    With fpsWafers(0)

        For i = 1 To .MaxRows
        
            tRepData.ID = GetMaxID()
            tRepData.CUSTOMERSHORTNAME = dcCusCode.text
            tRepData.CUSTOMERDEVICE = Trim(dcCusDevice.text)

            If tRepData.CUSTOMERSHORTNAME = "37" Then
                tRepData.JOBNO = UCase$(Trim$(txtJobNo.text))
            Else
                tRepData.JOBNO = ""

            End If
        
            .Row = i
            .Col = E_FPS0.E_OK
        
            If .text = 1 Then
                tOrderDetails(j).ORDERNAME = UCase(Trim(txtWorkOrder.text))
            
                .Row = i
            
                .Col = 1
                tRepData.Wafer_id = Trim$(.text)
            
                .Col = 3
                tOrderDetails(j).waferid = Trim$(.text)
                tRepData.SUBSTRATEID = Trim$(.text)
            
                .Col = 4
                tOrderDetails(j).DIEQTY = Trim$(.text)
                tRepData.GROSSBINCOUNT = CLng(Trim$(.text))
            
                .Col = 5
                tOrderDetails(j).FGDIEQTY = Trim$(.text)
                tRepData.PASSBINCOUNT = CLng(Trim$(.text))
            
                .Col = 6
                tOrderDetails(j).WAFERLOT = Trim(.text)
                tRepData.LOTID = Trim$(.text)
    
                If InStr(1, UpLotId, tOrderDetails(i).WAFERLOT) = 0 Then
                    UpLotId = UpLotId & "," & tOrderDetails(j).WAFERLOT

                End If

                .Col = 7
                tOrderDetails(j).MARKINGCODE = Trim$(.text)
                tRepData.PRODUCTID = Trim$(.text)
            
                ' 插入重工WO记录
                If cbOrderType.text = "重工工单" Then
                   
                    If Insert_to_repHeader(tRepData) = False Then
                        Exit Sub
                    End If
                    

                    If Insert_to_repDetails(tRepData) = False Then
                        Exit Sub
                    End If

                End If
            
                j = j + 1

            End If

        Next i

    End With

    Call addLogTxt(UCase(Trim(txtWorkOrder.text)), " 工单类 各字段付值成功 ")

    If tOrderHeader.ORDERNAME = "" Then

        MsgBox "工单号不可以为空"
        cmdSaveOrder.Enabled = True
        Exit Sub
    Else

        If Len(tOrderHeader.ORDERNAME) <> 12 Then

            MsgBox "工单号长度不对"
            cmdSaveOrder.Enabled = True
            Exit Sub

        End If

    End If

    Set rsOrderChk = GetWOData(tOrderHeader.ORDERNAME)

    If rsOrderChk.RecordCount > 0 Then
        MsgBox "Mes系统中已存在此工单号，请确认工单号 ！"
        cmdSaveOrder.Enabled = True
        Exit Sub

    End If

    Set rsOrderChk = GetWoData2(tOrderHeader.ORDERNAME)

    If rsOrderChk.RecordCount > 0 Then
        MsgBox "中间表中已存在此工单号，请不要重复提交 ！"
        cmdSaveOrder.Enabled = True
        Exit Sub

    End If

    If Len(tOrderHeader.PARA8) < 5 Then
        MsgBox "工单部门不对"
        cmdSaveOrder.Enabled = True
        Exit Sub
    Else

        If (InStr(tOrderHeader.PARA8, "_") = 0) Or (Not JudgeDept(tOrderHeader.PARA8)) Then
            MsgBox "生产部不存在"
            cmdSaveOrder.Enabled = True
            Exit Sub

        End If

    End If

    If cbOrderType.text <> "Dummy工单" And InStr(UCase(txtOrderDept.text), "BUMP") = 0 Then

        Set rsOrderChk = GetProduct_Check(dcProductNo.text)

        If rsOrderChk.RecordCount <= 0 Then
            MsgBox "料号不存在！请联系相关的人，先维护料号 ！"
            cmdSaveOrder.Enabled = True
            Exit Sub

        End If

        '4.判断料号的bom是否存在
        Set rsOrderChk = GetProductBom(dcProductNo.text)

        If rsOrderChk.RecordCount <= 0 Then
            MsgBox "新系统中这料号的Bom不存在！请联系相关的人，先维护Bom ！"
            cmdSaveOrder.Enabled = True
            Exit Sub

        End If

        '5.判断料号金碟是否有成本对象
        Set rsOrderChk = GetProductJDObject(dcProductNo.text)

        If rsOrderChk.RecordCount <= 0 Then
            MsgBox "此料号在金碟系统中无成本对象，请找相关人员确认 ！"
            cmdSaveOrder.Enabled = True
            Exit Sub

        End If

        '6.判断新ERP料号bom有没有签核过
        Set rsOrderChk = GetProductBomERpSign(dcProductNo.text)

        If rsOrderChk.RecordCount <= 0 Then
            MsgBox "新系统中这料号的Bom没有被审核通过，请联系工程部！"
            cmdSaveOrder.Enabled = True
            Exit Sub

        End If

        Call addLogTxt(UCase(Trim(txtWorkOrder.text)), " 数据检查完成 ")

    End If

    ' Step3: Dummy工单转移至WO
    If cbOrderType.text = "Dummy工单" Or cbOrderType.text = "玻璃工单" Or cbOrderType.text = "硅基工单" Or cbOrderType.text = "FO_CSP工单" Then
        If Trans_LotData(cbOrderType.text) = False Then
            Exit Sub
        End If

    End If

    Sleep (300)

    ' Step4: 上传工单数据
    If Insert_OrderToDb(tOrderHeader, tOrderDetails(), j) = False Then
        Exit Sub
    End If
    
    If Insert_Shop_Order(UCase(Trim(txtWorkOrder.text)), gUserName, Trim(dcCusDevice.text), Trim$(dcProductNo.text), cbOrderType.text) = False Then
        Exit Sub
    End If
    
    cmdSaveOrder.Enabled = True
    Exit Sub
    
End Sub

Private Sub dcCusCode_LostFocus()

    ' Step0: 根据客户代码导出客户机种
    Dim sCusCode As String   ' 客户代码

    sCusCode = Replace(UCase(Trim(dcCusCode.text)), Chr(13) + Chr(10), "")

    If cbOrderType.text = "Dummy工单" Or cbOrderType.text = "玻璃工单" Or cbOrderType.text = "硅基工单" Or cbOrderType.text = "FO_CSP工单" Then
        If sCusCode = "" Then
            MsgBox "请先选择客户代码"
 
            Exit Sub

        End If
    
        Set rsProductNo = Get_ProductNo(sCusCode, "")
        Set dcProductNo.RowSource = rsProductNo
        dcProductNo.ListField = rsProductNo("qtechptno2").name
        Exit Sub

    End If

    ' 客户机种初始化
    Set rsCusDevice = Get_CusDevice(sCusCode)
    Set dcCusDevice.RowSource = rsCusDevice
    dcCusDevice.ListField = rsCusDevice("CUSTOMERPTNO1").name

End Sub

Private Sub dcCusDevice_LostFocus()

    ' 根据客户代码和客户机种导出料号
    Dim sCusCode   As String

    Dim sCusDevice As String

    sCusCode = Replace(UCase(Trim(dcCusCode.text)), Chr(13) + Chr(10), "")
    sCusDevice = Replace(Trim(dcCusDevice.text), Chr(13) + Chr(10), "")

    ' 产品料号初始化
    Set rsProductNo = Get_ProductNo(sCusCode, sCusDevice)
    Set dcProductNo.RowSource = rsProductNo
    dcProductNo.ListField = rsProductNo("qtechptno2").name

End Sub

Private Sub dcProductNo_Change()

    ' 根据料号导出工单部门
    Dim sProductDept As String

    Dim sProductCode As String

    sProductDept = GetWoDept(dcProductNo.text)
    sProductCode = GetGWoDeptID(sProductDept)

    txtOrderDept.text = sProductDept & "_" & sProductCode

    ' 导出厂内机种
    txtPlantDevice.text = Get_PlantDevice(dcProductNo.text)

    txtOrderDept.Enabled = False
    txtPlantDevice.Enabled = False

    ' 如果是Dummy类则反带出客户机种
    If cbOrderType.text = "Dummy工单" Or cbOrderType.text = "玻璃工单" Or cbOrderType.text = "硅基工单" Or cbOrderType.text = "FO_CSP工单" Then
        dcCusDevice.text = Get_CusDeviceP(dcProductNo.text)
       
        dcCusDevice.Enabled = True

    End If

End Sub

Private Sub InitWaferFpsHeader()
    
    With fpsWafers(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_END - 1
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

        .Col = E_FPS0.E_OK
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
    
        .SetText E_FPS0.E_ID, 0, "ID"
        .SetText E_FPS0.E_WAFERID, 0, "LastWaferId"
        .SetText E_FPS0.E_NewWaferId, 0, "NewWaferId"
        .SetText E_FPS0.E_TotalDie, 0, "TotalDie数量"
        .SetText E_FPS0.E_GoodDie, 0, "GoodDie数量"
        .SetText E_FPS0.E_WaferLot, 0, "LotID"
        .SetText E_FPS0.E_MARKINGCODE, 0, "MarkingCode"
        .SetText E_FPS0.E_OK, 0, "选择"

        .ColWidth(E_FPS0.E_ID) = 4
        .ColWidth(E_FPS0.E_WAFERID) = 10
        .ColWidth(E_FPS0.E_NewWaferId) = 14
        .ColWidth(E_FPS0.E_TotalDie) = 10
        .ColWidth(E_FPS0.E_GoodDie) = 10
        .ColWidth(E_FPS0.E_WaferLot) = 10
        .ColWidth(E_FPS0.E_MARKINGCODE) = 15
        .ColWidth(E_FPS0.E_OK) = 6

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
   
        .Col = E_FPS0.E_OK
        .Lock = False

        .ReDraw = True

    End With
    
End Sub

Private Sub InitPriData()

    Dim i     As Integer

    Dim sOra  As String

    Dim rsPri As New ADODB.Recordset

    sOra = "select distinct pri as PRI from PJ_WO_PRI"

    If rsPri.State = adStateOpen Then
        rsPri.Close

    End If

    rsPri.Open sOra, Cnn, adOpenStatic, adLockReadOnly, adCmdText

    cbPri.Clear

    If Not rsPri.EOF Then

        Do While Not rsPri.EOF
            cbPri.AddItem Trim$("" & rsPri!Pri)
            rsPri.MoveNext
        Loop
    
        cbPri.ListIndex = 0

    End If

    rsPri.Close

End Sub

Private Sub InitLotTypeData()

    Dim i         As Integer

    Dim sOra      As String

    Dim rsLotType As New ADODB.Recordset

    sOra = "select distinct a.lot_type as lott from PJ_WO_PRI a where a.lot_type is not null order by  a.lot_type desc"

    If rsLotType.State = adStateOpen Then
        rsLotType.Close

    End If

    rsLotType.Open sOra, Cnn, adOpenStatic, adLockReadOnly, adCmdText

    cbLotType.Clear

    If Not rsLotType.EOF Then

        Do While Not rsLotType.EOF
            cbLotType.AddItem Trim$("" & rsLotType!lott)
            rsLotType.MoveNext
        Loop
    
        cbLotType.ListIndex = 0

    End If

    rsLotType.Close

End Sub

Private Sub Form_Activate()

    cbOrderType.SetFocus

End Sub

Private Sub Form_Load()

    ' Wafer明细表表头初始化
    Call InitWaferFpsHeader

    ' PRI初始化
    ' Call InitPriData
    cbPri.AddItem ("Hot Lot")
    cbPri.AddItem ("Normal Lot")
    cbPri.AddItem ("Super Hot Lot")
    cbPri.text = "Normal Lot"

    ' LOTTYPE初始化
    'Call InitLotTypeData
    cbLotType.AddItem ("E")
    cbLotType.AddItem ("M")
    cbLotType.AddItem ("Q")
    cbLotType.AddItem ("D")

    ' 日期初始化
    DTPicker1.Value = Format(Now(), "yyyy-MM-dd")
    DTPicker2.Value = Format(Year(Now()) & "-" & Month(Now()) & "-" & "28", "yyyy-MM-dd")
 
    ' 工单类型初始化
    cbOrderType.AddItem ("FO_CSP工单")
    cbOrderType.AddItem ("Dummy工单")
    cbOrderType.AddItem ("玻璃工单")
    cbOrderType.AddItem ("硅基工单")
    cbOrderType.AddItem ("一般工单")
    cbOrderType.AddItem ("样品工单")
    cbOrderType.AddItem ("重工工单")
    cbOrderType.AddItem ("再加工工单")
    cbOrderType.AddItem ("委外工单")
    cbOrderType.AddItem ("重工委外工单")
    cbOrderType.AddItem ("拆件式工单")
    cbOrderType.AddItem ("预测工单")
    cbOrderType.AddItem ("试产工单")
    cbOrderType.AddItem ("小批量试产工单")

    LblWorkOrder.Visible = False
    txtWorkOrder.Visible = False
    lblOrderDept.Visible = False
    txtOrderDept.Visible = False
    lblCnt.Visible = False
    txtPieces.Visible = False
    lblReproduction.Visible = False
    txtReproduction.Visible = False
    lblCustomerCode.Visible = False
    dcCusCode.Visible = False
    lblCustomerDevice.Visible = False
    dcCusDevice.Visible = False
    lblProductNumber.Visible = False
    dcProductNo.Visible = False
    LblPri.Visible = False
    cbPri.Visible = False
    lblLotType.Visible = False
    cbLotType.Visible = False
    LblWo.Visible = False
    txtWO.Visible = False
    lblPlantDevice.Visible = False
    txtPlantDevice.Visible = False
    lplPo.Visible = False
    txtPO.Visible = False
    lblProduceStart.Visible = False
    lblProduceStop.Visible = False
    CmdQueryLot.Visible = False
    cmdMakeLot.Visible = False
    lblJob.Visible = False
    txtJobNo.Visible = False
    cmdCreatRepWafer.Visible = False

    ' 客户代码初始化
    Set rsCusCode = Get_CusCode()
    Set dcCusCode.RowSource = rsCusCode
    dcCusCode.ListField = rsCusCode("productname").name
    dcCusCode.BoundColumn = rsCusCode("PID").name

End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)

    tClose.text = MonthView1.Value

    MonthView1.Visible = False

End Sub

Private Sub MonthView1_LostFocus()
    MonthView1.Visible = False

End Sub

Private Sub tClose_GotFocus()
    MonthView1.Visible = True
    MonthView1.SetFocus

End Sub

Private Sub txtWorkOrder_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call RefreshOrderID

    End If

End Sub

Private Function RefreshOrderID() As String

    Dim sOrderPrefix As String

    Dim sOrderDate   As String

    Dim sOrderSeq    As String

    sOrderPrefix = UCase(Trim(txtWorkOrder.text))

    sOrderDate = Right(Year(DateTime.DATE), 2) & Right("0" & Month(DateTime.DATE), 2)
    sOrderSeq = Right("000" & CStr(CInt(GetWoIDTemp(sOrderPrefix))), 4)
    
    RefreshOrderID = sOrderPrefix & "-" & sOrderDate & sOrderSeq

    If Mid$(sOrderPrefix, 2, 1) = "P" Then
        cbOrderType.text = "一般工单"

    End If

    If Mid$(sOrderPrefix, 2, 1) = "T" Then
        cbOrderType.text = "小批量试产工单"

    End If
       
    txtWorkOrder.text = sOrderPrefix & "-" & sOrderDate & sOrderSeq

    If Mid$(Trim(txtWorkOrder.text), 2, 1) = "P" Or Mid$(Trim(txtWorkOrder.text), 2, 1) = "T" Then
        cbLotType.text = "M"

    End If

    If Mid$(Trim(txtWorkOrder.text), 2, 1) = "S" Then
        cbLotType.text = "E"

    End If

    cmdStr = "insert into TSV_WO_SEQ_TAB(wotype,ymonth,sequenceid,flag) values ( '" & sOrderPrefix & "','" & sOrderDate & "'," & CLng(sOrderSeq) & ", 'Y' ) "
    AddSql (cmdStr)

End Function
