VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmUpLoadONForeCast 
   Caption         =   "ON 上传Forecast资料"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   19215
      _ExtentX        =   33893
      _ExtentY        =   13996
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "ASSY的OI指令"
      TabPicture(0)   =   "FormOn_Forecast.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label11"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label12"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label13"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label15"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label16"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label17"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label18"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DTPicker2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TxtStartid"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtDemandType"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtoutId"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtoutQty"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TxtworkWeek"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TxtSiteid"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxtStageid"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Txtctg"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "TxtPti"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "TxtComment"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Frame2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "CmdClear"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "CmdSave"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "TEST的OI指令"
      TabPicture(1)   =   "FormOn_Forecast.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line2"
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(3)=   "Label6"
      Tab(1).Control(4)=   "Label7"
      Tab(1).Control(5)=   "Label8"
      Tab(1).Control(6)=   "Label9"
      Tab(1).Control(7)=   "Label10"
      Tab(1).Control(8)=   "Label19"
      Tab(1).Control(9)=   "Label20"
      Tab(1).Control(10)=   "Label21"
      Tab(1).Control(11)=   "Label22"
      Tab(1).Control(12)=   "Label23"
      Tab(1).Control(13)=   "Label24"
      Tab(1).Control(14)=   "Label25"
      Tab(1).Control(15)=   "Label26"
      Tab(1).Control(16)=   "Label27"
      Tab(1).Control(17)=   "Label28"
      Tab(1).Control(18)=   "DTPicker1"
      Tab(1).Control(19)=   "DTPicker3"
      Tab(1).Control(20)=   "CmdSaveTest"
      Tab(1).Control(21)=   "CmdClearTest"
      Tab(1).Control(22)=   "Frame1"
      Tab(1).Control(23)=   "TxtOracleCD"
      Tab(1).Control(24)=   "TxtAreacd"
      Tab(1).Control(25)=   "TxtNextSite"
      Tab(1).Control(26)=   "TxtconsItem"
      Tab(1).Control(27)=   "TxtProdPartid"
      Tab(1).Control(28)=   "Txtstage"
      Tab(1).Control(29)=   "TxtSite"
      Tab(1).Control(30)=   "Txtpartid"
      Tab(1).Control(31)=   "TxtPTI2"
      Tab(1).Control(32)=   "TxtPkgCd"
      Tab(1).Control(33)=   "TxtPkgGrpCd"
      Tab(1).Control(34)=   "TxtSchComment"
      Tab(1).Control(35)=   "TxtQty"
      Tab(1).Control(36)=   "Txtit"
      Tab(1).Control(37)=   "TxtOnhand"
      Tab(1).ControlCount=   38
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "FormOn_Forecast.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox TxtOnhand 
         Height          =   375
         Left            =   -59760
         TabIndex        =   71
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox Txtit 
         Height          =   375
         Left            =   -61320
         TabIndex        =   69
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox TxtQty 
         Height          =   375
         Left            =   -65280
         TabIndex        =   68
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox TxtSchComment 
         Height          =   375
         Left            =   -69240
         TabIndex        =   64
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox TxtPkgGrpCd 
         Height          =   375
         Left            =   -73200
         TabIndex        =   63
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox TxtPkgCd 
         Height          =   375
         Left            =   -61320
         TabIndex        =   61
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox TxtPTI2 
         Height          =   375
         Left            =   -65280
         TabIndex        =   60
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Txtpartid 
         Height          =   375
         Left            =   -61320
         TabIndex        =   58
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox TxtSite 
         Height          =   375
         Left            =   -69240
         TabIndex        =   45
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox Txtstage 
         Height          =   375
         Left            =   -65280
         TabIndex        =   44
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox TxtProdPartid 
         Height          =   375
         Left            =   -73200
         TabIndex        =   43
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox TxtconsItem 
         Height          =   375
         Left            =   -69240
         TabIndex        =   42
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox TxtNextSite 
         Height          =   375
         Left            =   -61320
         TabIndex        =   41
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox TxtAreacd 
         Height          =   375
         Left            =   -73200
         TabIndex        =   40
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox TxtOracleCD 
         Height          =   375
         Left            =   -69240
         TabIndex        =   39
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         Caption         =   "BEPP_xls"
         Height          =   2895
         Left            =   -74160
         TabIndex        =   33
         Top             =   4320
         Width           =   15015
         Begin VB.CommandButton Command8 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   4080
            TabIndex        =   37
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton Command7 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   1200
            TabIndex        =   36
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   ".."
            Height          =   495
            Left            =   6120
            TabIndex        =   35
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   495
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   34
            Top             =   840
            Width           =   4935
         End
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   3000
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的xls："
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   840
            TabIndex        =   38
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.CommandButton CmdClearTest 
         Caption         =   "清空 "
         Height          =   480
         Left            =   -65040
         TabIndex        =   32
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton CmdSaveTest 
         Caption         =   "保存"
         Height          =   480
         Left            =   -68040
         TabIndex        =   31
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "保存"
         Height          =   480
         Left            =   6960
         TabIndex        =   30
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "清空 "
         Height          =   480
         Left            =   9960
         TabIndex        =   29
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "FEDS_xls"
         Height          =   2895
         Left            =   840
         TabIndex        =   23
         Top             =   4320
         Width           =   15015
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   495
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   840
            Width           =   4935
         End
         Begin VB.CommandButton Command2 
            Caption         =   ".."
            Height          =   495
            Left            =   6120
            TabIndex        =   26
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton Command3 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   1200
            TabIndex        =   25
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton Command5 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   4080
            TabIndex        =   24
            Top             =   1560
            Width           =   1335
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   3000
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的xls："
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   840
            TabIndex        =   28
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.TextBox TxtComment 
         Height          =   375
         Left            =   9720
         TabIndex        =   21
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox TxtPti 
         Height          =   375
         Left            =   5760
         TabIndex        =   19
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Txtctg 
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox TxtStageid 
         Height          =   375
         Left            =   13680
         TabIndex        =   15
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox TxtSiteid 
         Height          =   375
         Left            =   9720
         TabIndex        =   13
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox TxtworkWeek 
         Height          =   375
         Left            =   5760
         TabIndex        =   11
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox TxtoutQty 
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox TxtoutId 
         Height          =   375
         Left            =   9720
         TabIndex        =   5
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox TxtDemandType 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox TxtStartid 
         Height          =   375
         Left            =   5760
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   13680
         TabIndex        =   7
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   219938817
         CurrentDate     =   40882
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   -73200
         TabIndex        =   57
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   219938817
         CurrentDate     =   40882
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -65280
         TabIndex        =   59
         Top             =   1440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   219938817
         CurrentDate     =   40882
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OnHand："
         Height          =   195
         Left            =   -60480
         TabIndex        =   72
         Top             =   2760
         Width           =   765
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I/T："
         Height          =   195
         Left            =   -62040
         TabIndex        =   70
         Top             =   2760
         Width           =   390
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty："
         Height          =   195
         Left            =   -65760
         TabIndex        =   67
         Top             =   2760
         Width           =   450
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SchComments："
         Height          =   195
         Left            =   -70440
         TabIndex        =   66
         Top             =   2760
         Width           =   1185
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PkgGrpCd："
         Height          =   195
         Left            =   -74280
         TabIndex        =   65
         Top             =   2640
         Width           =   885
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PkgCd："
         Height          =   195
         Left            =   -62040
         TabIndex        =   62
         Top             =   2160
         Width           =   630
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Site："
         Height          =   195
         Left            =   -69720
         TabIndex        =   56
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CreateDate："
         Height          =   195
         Left            =   -74400
         TabIndex        =   55
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stage："
         Height          =   195
         Left            =   -66000
         TabIndex        =   54
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"FormOn_Forecast.frx":0054
         Height          =   390
         Left            =   -62040
         TabIndex        =   53
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ProdPartId："
         Height          =   195
         Left            =   -74280
         TabIndex        =   52
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ConsItem："
         Height          =   195
         Left            =   -70200
         TabIndex        =   51
         Top             =   1560
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "StartingWeek："
         Height          =   195
         Left            =   -66480
         TabIndex        =   50
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NextSite："
         Height          =   195
         Left            =   -62160
         TabIndex        =   49
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MfgAreaCd："
         Height          =   195
         Left            =   -74280
         TabIndex        =   48
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OracleLocCd："
         Height          =   195
         Left            =   -70320
         TabIndex        =   47
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PTI："
         Height          =   195
         Left            =   -65760
         TabIndex        =   46
         Top             =   2160
         Width           =   420
      End
      Begin VB.Line Line2 
         X1              =   -74280
         X2              =   -58800
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line1 
         X1              =   720
         X2              =   16200
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comments："
         Height          =   195
         Left            =   8640
         TabIndex        =   22
         Top             =   2160
         Width           =   930
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pti2："
         Height          =   195
         Left            =   5280
         TabIndex        =   20
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctg："
         Height          =   195
         Left            =   1200
         TabIndex        =   18
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stage_id："
         Height          =   195
         Left            =   12840
         TabIndex        =   16
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Site_id："
         Height          =   195
         Left            =   9000
         TabIndex        =   14
         Top             =   1560
         Width           =   660
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Workweek："
         Height          =   195
         Left            =   4800
         TabIndex        =   12
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start_qty："
         Height          =   195
         Left            =   840
         TabIndex        =   10
         Top             =   1440
         Width           =   870
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"FormOn_Forecast.frx":0062
         Height          =   390
         Left            =   12720
         TabIndex        =   8
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start_part_id："
         Height          =   195
         Left            =   8640
         TabIndex        =   6
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Demand_type："
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Out_part_id："
         Height          =   195
         Left            =   4560
         TabIndex        =   3
         Top             =   840
         Width           =   1050
      End
   End
End
Attribute VB_Name = "FrmUpLoadONForeCast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BcRS        As New ADODB.Recordset
Dim forcastTemp As ForeCastRecord


Private Sub CmdDel_Click()
Dim idTemp As String

idTemp = Trim$(TxtBatchId.Text)

If idTemp = "" Then
    MsgBox "BatchId不可以为空"
    Exit Sub
    
End If

'判断输入的Lot号，是否存在于BC表中
If (Not JudgeBCExist(idTemp)) Then
   MsgBox "这笔：" & idTemp & " 不存在，无需删除！"
Exit Sub

End If


Call DelBC(idTemp)

End Sub

Private Sub CmdClearTest_Click()
TxtSite.Text = ""
Txtstage.Text = ""
Txtpartid.Text = ""
TxtProdPartid.Text = ""
TxtconsItem.Text = ""
TxtNextSite.Text = ""
TxtAreacd.Text = ""
TxtOracleCD.Text = ""
TxtPTI2.Text = ""
TxtPkgCd.Text = ""
TxtPkgGrpCd.Text = ""
TxtSchComment.Text = ""
TxtQty.Text = ""
Txtit.Text = ""
TxtOnhand.Text = ""

End Sub

Private Sub Command1_Click()
'修改数量
Dim idTemp As String

idTemp = Trim$(TxtBatchId.Text)

If idTemp = "" Then
    MsgBox "BatchId不可以为空"
    Exit Sub
    
End If

If Trim(TxtQty1.Text) = "" Then
'先根据BatchId带出原来数量
    MsgBox "先输入BatchId，后回车，带出原来数量！"
    Exit Sub
End If

If Trim(TxtQty2.Text) = "" Then
    MsgBox "请输入现在BC中的数量！"
    Exit Sub
End If

'判断数量是否大于原来数量

If CLng(Trim(TxtQty2.Text)) > CLng(TxtQty1.Text) Then
    MsgBox "输入的数量不可以大于原来的数量！"
    Exit Sub
End If



Call ModifyBC(idTemp, CLng(Trim(TxtQty2.Text)))




End Sub

Private Sub CmdClear_Click()
TxtDemandType.Text = ""
TxtStartid.Text = ""
TxtoutId.Text = ""
TxtoutQty.Text = ""
TxtworkWeek.Text = ""
TxtSiteid.Text = ""
TxtStageid.Text = ""
Txtctg.Text = ""
TxtPti.Text = ""
TxtComment.Text = ""


End Sub

Private Sub CmdSave_Click()
Dim cmdStr As String
Dim cmdStr2 As String
SumCount = 0

forcastTemp.id = GetForeCastID()
forcastTemp.QtechCreateBy = gUserName
forcastTemp.TypeName = "FEDS"
forcastTemp.DemandType = Trim(TxtDemandType.Text)
forcastTemp.StartPartId = Trim(TxtStartid.Text)
forcastTemp.OutPartId = Trim(TxtoutId.Text)
forcastTemp.OutDate = DTPicker2.Value
forcastTemp.outQty = CLng(Trim(TxtoutQty.Text))
forcastTemp.WorkWeek = Trim(TxtworkWeek.Text)
forcastTemp.SiteId = Trim(TxtSiteid.Text)
forcastTemp.StageId = Trim(TxtStageid.Text)
forcastTemp.Ctg = Trim(Txtctg.Text)
forcastTemp.Pti2 = Trim(TxtPti.Text)
forcastTemp.Comments = Trim(TxtComment)

'添加导入Sqlserver
On Error GoTo DealError
        
'Cnn.BeginTrans

cmdStr = " insert into CUSTOMERFORECASTTBL( ID , TYPENAME , DEMAND_TYPE  ,OUT_PART_ID,START_PART_ID , " & _
"   OUT_DATE ,OUT_QTY,WORKWEEK ,SITE_ID,STAGE_ID , CTG , PTI2 ,COMMENTS,FLAG ,QTECH_CREATED_BY , QTECH_CREATED_DATE ) " & _
" values(" & forcastTemp.id & " ,'" & forcastTemp.TypeName & "','" & forcastTemp.DemandType & "','" & forcastTemp.StartPartId & "','" & forcastTemp.OutPartId & "'," & _
" '" & forcastTemp.OutDate & "'," & forcastTemp.outQty & ",'" & forcastTemp.WorkWeek & "','" & forcastTemp.SiteId & "'," & _
" '" & forcastTemp.StageId & "','" & forcastTemp.Ctg & "','" & forcastTemp.Pti2 & "','" & forcastTemp.Comments & "','Y','" & forcastTemp.QtechCreateBy & "',sysdate) "

cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerForeCast]( ID , TYPENAME , DEMAND_TYPE  ,OUT_PART_ID,START_PART_ID ," & _
"   OUT_DATE ,OUT_QTY,WORKWEEK ,SITE_ID,STAGE_ID , CTG , PTI2 ,COMMENTS,FLAG ,QTECH_CREATED_BY , QTECH_CREATED_DATE ) " & _
" values(" & forcastTemp.id & " ,'" & forcastTemp.TypeName & "','" & forcastTemp.DemandType & "','" & forcastTemp.StartPartId & "','" & forcastTemp.OutPartId & "'," & _
" '" & forcastTemp.OutDate & "'," & forcastTemp.outQty & ",'" & forcastTemp.WorkWeek & "','" & forcastTemp.SiteId & "'," & _
" '" & forcastTemp.StageId & "','" & forcastTemp.Ctg & "','" & forcastTemp.Pti2 & "','" & forcastTemp.Comments & "','Y','" & forcastTemp.QtechCreateBy & "',getdate()) "



                        
AddSql (cmdStr)
AddSql2 (cmdStr2)
SumCount = SumCount + 1
 
'Cnn.CommitTrans
 MsgBox "保存成功!", vbInformation, "友情提示"
  
Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1





End Sub

Private Sub CmdSaveTest_Click()
Dim cmdStr As String
Dim cmdStr2 As String
SumCount = 0

forcastTemp.id = GetForeCastID()
forcastTemp.QtechCreateBy = gUserName
forcastTemp.TypeName = "BEPP"

forcastTemp.CreateDate = DTPicker3.Value
forcastTemp.Site = Trim(TxtSite.Text)
forcastTemp.Stage = Trim(Txtstage.Text)
forcastTemp.PartId = Trim(Txtpartid.Text)
forcastTemp.ProdPartId = Trim(TxtProdPartid.Text)
forcastTemp.ConsItem = Trim(TxtconsItem.Text)

forcastTemp.StartingWeek = DTPicker1.Value
forcastTemp.NextSite = Trim(TxtNextSite.Text)
forcastTemp.MfgAreaCd = Trim(TxtAreacd.Text)
forcastTemp.OracleLocCd = Trim(TxtOracleCD.Text)
forcastTemp.PTI = Trim(TxtPTI2)

forcastTemp.PkgCd = Trim(TxtPkgCd.Text)
forcastTemp.PkgGrpCd = Trim(TxtPkgGrpCd.Text)
forcastTemp.SchComments = Trim(TxtSchComment.Text)
forcastTemp.qty = CLng(Trim(TxtQty.Text))
forcastTemp.IT = Trim(Txtit)
forcastTemp.OnHand = Trim(TxtOnhand)


'添加导入Sqlserver
On Error GoTo DealError
        
'Cnn.BeginTrans

cmdStr = " insert into CUSTOMERFORECASTTBL( ID , TYPENAME , CREATE_DATE ,SITE ,STAGE ,PART_ID,PROD_PART_ID ," & _
"     CONS_ITEM ,STARTING_WEEK ,NEXT_SITE , MFG_AREA_CD,ORACLE_LOC_CD ," & _
"     PTI , PKG_CD ,PKG_GRP_CD ,SCH_COMMENTS ,QTY ,I_T,ON_HAND , FLAG ,QTECH_CREATED_BY , QTECH_CREATED_DATE ) " & _
" values(" & forcastTemp.id & " ,'" & forcastTemp.TypeName & "','" & forcastTemp.CreateDate & "','" & forcastTemp.Site & "','" & forcastTemp.Stage & "','" & forcastTemp.PartId & "','" & forcastTemp.ProdPartId & "'," & _
" '" & forcastTemp.ConsItem & "','" & forcastTemp.StartingWeek & "','" & forcastTemp.NextSite & "','" & forcastTemp.MfgAreaCd & "','" & forcastTemp.OracleLocCd & "'," & _
" '" & forcastTemp.PTI & "','" & forcastTemp.PkgCd & "','" & forcastTemp.PkgGrpCd & "','" & forcastTemp.SchComments & "'," & forcastTemp.qty & ",'" & forcastTemp.IT & "','" & forcastTemp.OnHand & "','Y','" & forcastTemp.QtechCreateBy & "',sysdate) "

cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerForeCast]( ID , TYPENAME , CREATE_DATE ,SITE ,STAGE ,PART_ID,PROD_PART_ID ," & _
"   CONS_ITEM ,STARTING_WEEK ,NEXT_SITE , MFG_AREA_CD,ORACLE_LOC_CD ," & _
"     PTI , PKG_CD ,PKG_GRP_CD ,SCH_COMMENTS ,QTY ,I_T,ON_HAND , FLAG ,QTECH_CREATED_BY , QTECH_CREATED_DATE ) " & _
" values(" & forcastTemp.id & " ,'" & forcastTemp.TypeName & "','" & forcastTemp.CreateDate & "','" & forcastTemp.Site & "','" & forcastTemp.Stage & "','" & forcastTemp.PartId & "','" & forcastTemp.ProdPartId & "'," & _
" '" & forcastTemp.ConsItem & "','" & forcastTemp.StartingWeek & "','" & forcastTemp.NextSite & "','" & forcastTemp.MfgAreaCd & "','" & forcastTemp.OracleLocCd & "'," & _
" '" & forcastTemp.PTI & "','" & forcastTemp.PkgCd & "','" & forcastTemp.PkgGrpCd & "','" & forcastTemp.SchComments & "'," & forcastTemp.qty & ",'" & forcastTemp.IT & "','" & forcastTemp.OnHand & "','Y','" & forcastTemp.QtechCreateBy & "',getdate()) "

                  
AddSql (cmdStr)
AddSql2 (cmdStr2)
SumCount = SumCount + 1
 
'Cnn.CommitTrans
 MsgBox "保存成功!", vbInformation, "友情提示"
  
Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1






End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim FName
    '帅选文件
    'CommonDialog1.Filter = "EXCEL文件(*.xls)|*.xls"
    CommonDialog1.Filter = "CSV文件(*.csv)|*.csv"
    
    CommonDialog1.ShowOpen
    '得到文件名
    FName = CommonDialog1.FileName
    If FName <> "" Then
       Text2.Text = FName
    End If
End Sub

Private Sub Command3_Click()
'上传资料

Dim source_batch_id_Temp As String
'上传OI的CSV
'处理文件名
If Text2.Text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

'获取文件名
'    If InStrRev(Trim(Text2.Text), "\") > 0 Then
'        strFileName = Mid(Trim(Text2.Text), InStrRev(Trim(Text2.Text), "\") + 1)
'        dirName = Mid$(Trim(Text2.Text), 1, InStrRev(Trim(Text2.Text), "\"))
'    End If
    

'2012-06-27 jiayunzhang 修改读Excel的方式


'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text2.Text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 11 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If


Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String
Dim dieQtyTemp As Long
Dim pcsQtemp As Integer
Dim start_part As String
Dim out_part As String
Dim typenameTemp As String
   


SumCount = 0
BCResultFlag = False

Cnn.BeginTrans

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    pcsQtemp = 0
    
    source_batch_id_Temp = ""
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
           
        If j = 1 Then
            'source_batch_id_Temp = Trim(tempVal)  'LotId
            
            temp = temp & "," & newStr("" & tempVal)
            
        End If
        
        If j = 4 Then
           ' dieQtyTemp = CLng(Trim(tempVal))  'qty
             'temp = temp & "," & newStr("" & tempVal)
             temp = temp & "," & newStr("" & tempVal)
            
            
        End If
        
        If j = 3 Then
   
            
            temp = temp & "," & newStr("" & tempVal)
            out_part = Left(tempVal, InStr(tempVal, "-") + 3)
                       
        End If
         
            
        If j = 2 Then
            temp = temp & "," & newStr("" & tempVal)
            start_part = Left(tempVal, InStr(tempVal, "-") + 3)
          
        End If
        
        If j = 5 Then
        
          dieQtyTemp = CLng(Trim(tempVal))  'qty
             temp = temp & "," & newStr("" & tempVal)
             
          '  temp = temp & "," & newStr("" & tempVal)
          
        End If
        
        If j = 6 Then
            temp = temp & "," & newStr("" & tempVal)
          
        End If
        
        If j = 7 Then
        temp = temp & "," & newStr("" & tempVal)
        
        End If
        
        If j = 8 Then
            temp = temp & "," & newStr("" & tempVal)
          
        End If
        
        If j = 9 Then
        temp = temp & "," & newStr("" & tempVal)
        
        End If
        
         If j = 10 Then
            temp = temp & "," & newStr("" & tempVal)
          
        End If
        
        If j = 11 Then
        temp = temp & "," & newStr("" & tempVal)
        
        End If
        
        
        
        
    Next j

    '取目前DB最大的ID号
    If start_part <> out_part Then
    MsgBox "请检查start_part 和 out_part"
    Exit Sub
    End If
    
    id = GetForeCastID()
    typenameTemp = "FEDS"
    temp = id & ",'" & typenameTemp & "'" & temp
'    temp2 = temp & ",'Y','Auto',GETDATE(),'','','AA',0"

    temp2 = temp & ",'Y','Auto',GETDATE()"
    temp = temp & ",'Y','Auto',sysdate"

'    Debug.Print temp

             '判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
'    If (JudgeFlagStautsBC(source_batch_id_Temp)) Then
'       MsgBox "这笔：" & source_batch_id_Temp & "已存在，无需上传!"
'       GoTo NextRecord2
'
'    End If
    
    
'    If (Not JudgeFlagStautsBCQty(source_batch_id_Temp, dieQtyTemp)) Then
'       MsgBox "这笔：" & source_batch_id_Temp & "与BI中的Die数量不一致!"
'       GoTo NextRecord2
'
'    End If


    Call AddONForcast(temp, temp2)
    SumCount = SumCount + 1
     
    '上传到DB
NextRecord2:

Next i

Cnn.CommitTrans

     
     xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit




If SumCount > 0 Then
    MsgBox "已成功上传" & SumCount & "笔！", , "友情提醒"
    
    Else
        If BCResultFlag = True Then
            MsgBox "上传失败，请确认资料格式！", , "友情提醒"
            Exit Sub
        End If
    
End If


End Sub


Private Function newStr(temp As String)
If temp <> "" Then
newStr = "'" & temp & "'"
Else
newStr = "''"

End If

End Function

Private Sub Command5_Click()

Dim temp As String

temp = "    select  ID ,TYPENAME,DEMAND_TYPE ,OUT_PART_ID,START_PART_ID  ,OUT_DATE as start_date ,OUT_QTY as start_qty ,WORKWEEK ,SITE_ID  ,STAGE_ID ,CTG   ,PTI2  ,COMMENTS , CREATE_DATE  , SITE ,STAGE,PART_ID ,PROD_PART_ID , CONS_ITEM  ,STARTING_WEEK  , " & _
"  NEXT_SITE , MFG_AREA_CD  ,ORACLE_LOC_CD ,PTI   ,PKG_CD ,PKG_GRP_CD ,SCH_COMMENTS , QTY ,I_T  ,ON_HAND , FLAG  , QTECH_CREATED_BY  ,QTECH_CREATED_DATE  from CUSTOMERFORECASTTBL order by id "
      
 ExporToExcel (temp)
       
End Sub


Private Sub TxtBatchId_KeyPress(KeyAscii As Integer)
Dim idTemp As String

If KeyAscii = 13 Then
    idTemp = Trim$(TxtBatchId.Text)

    '判断输入的Lot号，是否存在于BC表中
    If (Not JudgeBCExist(idTemp)) Then
       MsgBox "这笔：" & idTemp & " 不存在，无需删除！"
    Exit Sub
    
    End If
    
    Set BcRS = GetDecBCQty(idTemp)

    TxtQty1.Text = BcRS.fields("dieqty").Value


End If

End Sub


Private Sub TxtQty2_KeyPress(KeyAscii As Integer)
Dim Numbers As String
Numbers = "1234567890" + Chr(8)
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If


End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim FName
    '帅选文件
    CommonDialog2.Filter = "EXCEL文件(*.xls)|*.xls"
    CommonDialog2.ShowOpen
    '得到文件名
    FName = CommonDialog2.FileName
    If FName <> "" Then
       Text1.Text = FName
    End If
End Sub

Private Sub Command7_Click()

'上传资料

Dim source_batch_id_Temp As String
'上传OI的CSV
'处理文件名
If Text1.Text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

'获取文件名
'    If InStrRev(Trim(Text2.Text), "\") > 0 Then
'        strFileName = Mid(Trim(Text2.Text), InStrRev(Trim(Text2.Text), "\") + 1)
'        dirName = Mid$(Trim(Text2.Text), 1, InStrRev(Trim(Text2.Text), "\"))
'    End If
    

'2012-06-27 jiayunzhang 修改读Excel的方式


'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text1.Text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 17 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If


Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String
Dim dieQtyTemp As Long
Dim pcsQtemp As Integer

Dim typenameTemp As String
   


SumCount = 0
BCResultFlag = False

Cnn.BeginTrans

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    pcsQtemp = 0
    
    source_batch_id_Temp = ""
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
           
        If j = 1 Then
            'source_batch_id_Temp = Trim(tempVal)  'LotId
            
            temp = temp & "," & newStr("" & tempVal)
            
        End If
        
        If j = 4 Then
           ' dieQtyTemp = CLng(Trim(tempVal))  'qty
             'temp = temp & "," & newStr("" & tempVal)
             temp = temp & "," & newStr("" & tempVal)
        End If
        
        If j = 3 Then
   
            
            temp = temp & "," & newStr("" & tempVal)
                       
        End If
         
            
        If j = 2 Then
            temp = temp & "," & newStr("" & tempVal)
          
        End If
        
        If j = 5 Then
        
          'dieQtyTemp = CLng(Trim(tempVal))  'qty
             temp = temp & "," & newStr("" & tempVal)
             
          '  temp = temp & "," & newStr("" & tempVal)
          
        End If
        
        If j = 6 Then
            temp = temp & "," & newStr("" & tempVal)
          
        End If
        
        If j = 7 Then
        temp = temp & "," & newStr("" & tempVal)
        
        End If
        
        If j = 8 Then
            temp = temp & "," & newStr("" & tempVal)
          
        End If
        
        If j = 9 Then
        temp = temp & "," & newStr("" & tempVal)
        
        End If
        
         If j = 10 Then
            temp = temp & "," & newStr("" & tempVal)
          
        End If
        
        If j = 11 Then
        temp = temp & "," & newStr("" & tempVal)
        
        End If
        
        If j = 12 Then
        temp = temp & "," & newStr("" & tempVal)
        
        End If
        
         If j = 13 Then
            temp = temp & "," & newStr("" & tempVal)
          
        End If
        
        If j = 14 Then
        temp = temp & "," & newStr("" & tempVal)
        
        End If
        
          If j = 15 Then
        temp = temp & "," & newStr("" & tempVal)
        
        End If
        
         If j = 16 Then
            temp = temp & "," & newStr("" & tempVal)
          
        End If
        
        If j = 17 Then
        temp = temp & "," & newStr("" & tempVal)
        
        End If
        
        
        
        
        
        
    Next j

    '取目前DB最大的ID号
    id = GetForeCastID()
    typenameTemp = "BEPP"
    temp = id & ",'" & typenameTemp & "'" & temp
'    temp2 = temp & ",'Y','Auto',GETDATE(),'','','AA',0"

    temp2 = temp & ",'Y','Auto',GETDATE()"
    temp = temp & ",'Y','Auto',sysdate"

'    Debug.Print temp

             '判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
'    If (JudgeFlagStautsBC(source_batch_id_Temp)) Then
'       MsgBox "这笔：" & source_batch_id_Temp & "已存在，无需上传!"
'       GoTo NextRecord2
'
'    End If
    
    
'    If (Not JudgeFlagStautsBCQty(source_batch_id_Temp, dieQtyTemp)) Then
'       MsgBox "这笔：" & source_batch_id_Temp & "与BI中的Die数量不一致!"
'       GoTo NextRecord2
'
'    End If


    Call AddONForcastBePP(temp, temp2)
    SumCount = SumCount + 1
     
    '上传到DB
NextRecord2:

Next i

Cnn.CommitTrans

     
     xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit




If SumCount > 0 Then
    MsgBox "已成功上传" & SumCount & "笔！", , "友情提醒"
    
    Else
        If BCResultFlag = True Then
            MsgBox "上传失败，请确认资料格式！", , "友情提醒"
            Exit Sub
        End If
    
End If







End Sub

Private Sub Command8_Click()
Dim temp As String

temp = "    select  ID ,TYPENAME,DEMAND_TYPE ,START_PART_ID ,OUT_PART_ID ,OUT_DATE ,OUT_QTY ,WORKWEEK ,SITE_ID  ,STAGE_ID ,CTG   ,PTI2  ,COMMENTS , CREATE_DATE  , SITE ,STAGE,PART_ID ,PROD_PART_ID , CONS_ITEM  ,STARTING_WEEK  , " & _
"  NEXT_SITE , MFG_AREA_CD  ,ORACLE_LOC_CD ,PTI   ,PKG_CD ,PKG_GRP_CD ,SCH_COMMENTS , QTY ,I_T  ,ON_HAND , FLAG  , QTECH_CREATED_BY  ,QTECH_CREATED_DATE  from CUSTOMERFORECASTTBL order by id "
      
 ExporToExcel (temp)
End Sub

Private Sub Form_Activate()
TxtDemandType.SetFocus

DTPicker1.Value = Format(Now, "yyyy-mm-dd")
DTPicker2.Value = Format(Now, "yyyy-mm-dd")
DTPicker3.Value = Format(Now, "yyyy-mm-dd")

'  select * from  CUSTOMERMSLevelTBL for update
End Sub

