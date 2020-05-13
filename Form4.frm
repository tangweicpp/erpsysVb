VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmSetTime 
   Caption         =   "自定义时间与Remark 设定"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   12780
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "时间设定"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Remark设定"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command6"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame5 
         Caption         =   "工单完工时间"
         Height          =   1095
         Left            =   720
         TabIndex        =   28
         Top             =   3600
         Width           =   9615
         Begin VB.CommandButton cmd 
            Caption         =   "修改"
            Height          =   360
            Left            =   6960
            TabIndex        =   34
            Top             =   360
            Width           =   990
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   4800
            TabIndex        =   33
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   364380161
            CurrentDate     =   42612
         End
         Begin VB.TextBox txtText3 
            Height          =   405
            Left            =   1200
            TabIndex        =   31
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "新日期"
            Height          =   195
            Index           =   1
            Left            =   4080
            TabIndex        =   32
            Top             =   480
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "工单号："
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   30
            Top             =   480
            Width           =   720
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "导出报表"
         Height          =   600
         Left            =   -74280
         TabIndex        =   27
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "修改"
         Height          =   1815
         Left            =   -74280
         TabIndex        =   21
         Top             =   2400
         Width           =   9615
         Begin VB.TextBox TxtRemark2 
            Height          =   375
            Left            =   1440
            TabIndex        =   24
            Top             =   1080
            Width           =   5415
         End
         Begin VB.CommandButton Command5 
            Caption         =   "修改"
            Height          =   360
            Left            =   6960
            TabIndex        =   23
            Top             =   480
            Width           =   990
         End
         Begin VB.TextBox TxtWafer2 
            Height          =   375
            Left            =   1440
            TabIndex        =   22
            Top             =   480
            Width           =   5415
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remark："
            Height          =   195
            Left            =   600
            TabIndex        =   26
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WaferID："
            Height          =   195
            Left            =   600
            TabIndex        =   25
            Top             =   600
            Width           =   795
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "新增"
         Height          =   1695
         Left            =   -74280
         TabIndex        =   14
         Top             =   480
         Width           =   9615
         Begin VB.TextBox TxtRemark 
            Height          =   375
            Left            =   1440
            TabIndex        =   17
            Top             =   960
            Width           =   5415
         End
         Begin VB.CommandButton Command4 
            Caption         =   "添加"
            Height          =   360
            Left            =   6960
            TabIndex        =   16
            Top             =   360
            Width           =   990
         End
         Begin VB.TextBox TxtWafer 
            Height          =   375
            Left            =   1440
            TabIndex        =   15
            Top             =   360
            Width           =   5415
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remark："
            Height          =   195
            Left            =   600
            TabIndex        =   20
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   4080
            TabIndex        =   19
            Top             =   480
            Width           =   45
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WaferID："
            Height          =   195
            Left            =   600
            TabIndex        =   18
            Top             =   480
            Width           =   795
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "新增"
         Height          =   975
         Left            =   720
         TabIndex        =   8
         Top             =   600
         Width           =   9615
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1200
            TabIndex        =   10
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton Command1 
            Caption         =   "添加"
            Height          =   360
            Left            =   6960
            TabIndex        =   9
            Top             =   360
            Width           =   990
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   4680
            TabIndex        =   11
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   364445697
            CurrentDate     =   40947
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LOTID："
            Height          =   195
            Left            =   600
            TabIndex        =   13
            Top             =   480
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "日期："
            Height          =   195
            Left            =   4080
            TabIndex        =   12
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "修改"
         Height          =   1335
         Left            =   720
         TabIndex        =   2
         Top             =   1920
         Width           =   9615
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1200
            TabIndex        =   4
            Top             =   480
            Width           =   2295
         End
         Begin VB.CommandButton Command2 
            Caption         =   "修改"
            Height          =   360
            Left            =   6960
            TabIndex        =   3
            Top             =   480
            Width           =   990
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   4680
            TabIndex        =   5
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   364380161
            CurrentDate     =   40947
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LOTID："
            Height          =   195
            Left            =   600
            TabIndex        =   7
            Top             =   600
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "新日期："
            Height          =   195
            Left            =   3960
            TabIndex        =   6
            Top             =   600
            Width           =   720
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "导出报表"
         Height          =   600
         Left            =   600
         TabIndex        =   1
         Top             =   4800
         Width           =   1335
      End
   End
   Begin VB.Label lblLOTID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOTID："
      Height          =   195
      Left            =   1320
      TabIndex        =   29
      Top             =   4440
      Width           =   630
   End
End
Attribute VB_Name = "FrmSetTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oiRS        As New ADODB.Recordset

Private Sub cmd_Click()
Dim strSql As String
Dim workName As String
Dim rs As New Recordset
Dim cmd As New ADODB.Command
Dim dtTemp As Date

dtTemp = DTPicker2.Value
workName = Trim(txtText3.Text)
If workName = "" Then
MsgBox "请输入工单！"
Exit Sub
End If

strSql = "select plan_star_date, plan_end_date from shop_order where shop_order ='" & workName & "'" '先判断工单是否存在
 If Cnn.State = 0 Then
    ConOracle
 End If
 
   rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
   If rs.RecordCount <= 0 Then
   MsgBox "要修改的工单不存在请确认！"
   Exit Sub
   End If
   
 strSql = "update shop_order set plan_end_date=to_date('" & dtTemp & "','yyyy-mm-dd') where shop_order='" & workName & "'" '修改工单完工日期主要为MPSWIP报表服务
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
   MsgBox "修改成功！"

End Sub

Private Sub Command1_Click()
'增加
Dim lotIdTemp As String
Dim dtTemp As Date
Dim sqlTemp As String
Dim remarkTemp As String



If Trim(Text1.Text) <> "" Then

    lotIdTemp = Trim(Text1.Text)
    dtTemp = DTPicker1.Value
    remarkTemp = ""
    
    '判断输入的Lot号，是否正确
    
    If JudgeLot2(lotIdTemp) Then
    
    
        '判断是否存在 存在则提示信息
        If Not (JudgeLot(lotIdTemp)) Then
        sqlTemp = "insert into WipreportDate(lotid,lotdate,remark) values ( '" & lotIdTemp & "',to_date('" & dtTemp & "','yyyy-mm-dd'),'" & remarkTemp & "' ) "
        AddSql (sqlTemp)
        MsgBox "添加成功!"
        
        Else
        
        MsgBox "LotId:" & lotIdTemp & "已存在！"
        End If
        
    Else
         MsgBox "LotId:" & lotIdTemp & "在Mes系统中不存在，请确认Lot号！"
    
    End If
    

Else
MsgBox "请先输入LotId!"
End If


End Sub

Private Sub Command2_Click()
'修改
Dim lotIdTemp As String
Dim dtTemp As Date
Dim sqlTemp As String
Dim remarkTemp As String


If Trim(Text2.Text) <> "" Then

    lotIdTemp = Trim(Text2.Text)
    dtTemp = DTPicker3.Value
    remarkTemp = ""
    
    '判断是否存在 存在则修改，不存在提示
     If JudgeLot(lotIdTemp) Then
     
        sqlTemp = "update WipreportDate set lotdate=to_date('" & dtTemp & "','yyyy-mm-dd'), remark='" & remarkTemp & "'    where lotid='" & lotIdTemp & "' "
        AddSql (sqlTemp)
        MsgBox "修改成功!"
        
    Else
        
          MsgBox "LotId:" & lotIdTemp & "不存在！"
     End If
    

    

Else
MsgBox "请先输入LotId!"
End If


End Sub

Public Function JudgeLot(lotIdTemp As String) As Boolean

Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "  select * from WipreportDate where lotid='" + lotIdTemp + "' "
         
slectResult = QueryStr(cmdStr)
JudgeLot = slectResult
End Function


Public Function JudgeWafer(lotIdTemp As String) As Boolean

Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "  select * from WipreportDateRemark where lotid='" + lotIdTemp + "' "
         
slectResult = QueryStr(cmdStr)
JudgeWafer = slectResult
End Function


Public Function JudgeLot2(lotIdTemp As String) As Boolean

Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "  select * from A_Lotwafers where  wafernumber='" + lotIdTemp + "' "
         
         
slectResult = QueryStr(cmdStr)
JudgeLot2 = slectResult
End Function

Public Function JudgeWafer2(lotIdTemp As String) As Boolean

Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "  select * from A_Lotwafers where  waferscribenumber='" + lotIdTemp + "' "
         
         
slectResult = QueryStr(cmdStr)
JudgeWafer2 = slectResult
End Function



Private Sub Command3_Click()
 ExporToExcel ("select lotid,lotdate,remark,CreateDate from WipreportDate order by CreateDate desc ")
End Sub

Private Sub Command4_Click()


'增加 Remark  2012-06-18
Dim lotIdTemp As String
Dim sqlTemp As String
Dim remarkTemp As String



If Trim(TxtWafer.Text) <> "" Then

    lotIdTemp = Trim(TxtWafer.Text)
    remarkTemp = Trim(TxtRemark.Text)
    
    '判断输入的Lot号，是否正确
    
    If JudgeWafer2(lotIdTemp) Then
    
    
        '判断是否存在 存在则提示信息
        If Not (JudgeWafer(lotIdTemp)) Then
        sqlTemp = "insert into WipreportDateRemark(lotid,remark) values ( '" & lotIdTemp & "','" & remarkTemp & "' ) "
        AddSql (sqlTemp)
        MsgBox "添加成功!"
        
        Else
        
        MsgBox "WaferId:" & lotIdTemp & "已存在！"
        End If
        
    Else
         MsgBox "WaferId:" & lotIdTemp & "在Mes系统中不存在，请确认Wafer号！"
    
    End If
    

Else
MsgBox "请先输入WaferId!"
End If






End Sub

Private Sub Command5_Click()

'修改 Remark 2012-06-18
Dim lotIdTemp As String
Dim sqlTemp As String
Dim remarkTemp As String


If Trim(TxtWafer2.Text) <> "" Then

    lotIdTemp = Trim(TxtWafer2.Text)

    remarkTemp = Trim(TxtRemark2.Text)
    
    '判断是否存在 存在则修改，不存在提示
     If JudgeWafer(lotIdTemp) Then
     
        sqlTemp = "update WipreportDateRemark set  remark='" & remarkTemp & "'    where lotid='" & lotIdTemp & "' "
        AddSql (sqlTemp)
        MsgBox "修改成功!"
        
    Else
        
          MsgBox "WaferId:" & lotIdTemp & "不存在！"
     End If
    

    

Else
MsgBox "请先输入WaferId!"
End If





End Sub

Private Sub Command6_Click()
 ExporToExcel ("select lotid as WaferId,remark,CreateDate from WipreportDateRemark order by CreateDate desc ")
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
DTPicker1.Value = DateTime.Date
DTPicker3.Value = DateTime.Date

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim lotIdTemp As String
lotIdTemp = Trim(Text2.Text)

 If KeyAscii = 13 Then
    
    
    Set oiRS = GetWipSetData(lotIdTemp)
    If (oiRS.RecordCount > 0) Then
    
    DTPicker3.Value = CDate(oiRS.Fields("lotdate").Value)
    Text3.Text = IIf(IsNull(oiRS.Fields("remark").Value), "", oiRS.Fields("remark").Value)

    End If
    
    
    
    
 End If
 

End Sub
