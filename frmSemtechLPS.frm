VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmSemtechLPS 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Semtech标签打印系统"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   10935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   19455
      Begin VB.CheckBox chk 
         BackColor       =   &H00C0C0C0&
         Caption         =   "补打"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   9960
         Width           =   1455
      End
      Begin VB.TextBox txtHTLot 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   6135
         Left            =   7680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton cmdOuterPrinter 
         BackColor       =   &H00FF80FF&
         Caption         =   "外箱  打印"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   8760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdInnerTrayPrinter 
         BackColor       =   &H0080C0FF&
         Caption         =   "内盒/卷盘 打印"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   8760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8880
         Width           =   2175
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H008080FF&
         Caption         =   "初始化"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   8880
         Width           =   2175
      End
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1800
         Width           =   6135
      End
      Begin VB.TextBox txtScan 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtDN 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1155
         Width           =   2655
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   6135
         Left            =   10560
         TabIndex        =   12
         Top             =   1800
         Width           =   6135
         _Version        =   524288
         _ExtentX        =   10821
         _ExtentY        =   10821
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
         SpreadDesigner  =   "frmSemtechLPS.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lblSelMatch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "挑料核对:"
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
         Left            =   10560
         TabIndex        =   11
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态区:"
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
         Top             =   3960
         Width           =   750
      End
      Begin VB.Label lblScan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描框"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   645
         Width           =   540
      End
      Begin VB.Label lblDN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN"
         Height          =   195
         Left            =   480
         TabIndex        =   1
         Top             =   1200
         Width           =   210
      End
   End
End
Attribute VB_Name = "frmSemtechLPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Enum E_LPS
E_JobNO = 1
E_QTY
E_JobNO2
E_Qty2
E_End

End Enum

Private Sub chk_Click()

If chk.Visible = True Then
    cmdOuterPrinter.Visible = True
    cmdInnerTrayPrinter.Visible = True
End If

End Sub

Private Sub CmdExit_Click()

Unload Me
End Sub

Private Sub cmdInnerTrayPrinter_Click()

Dim rs As New ADODB.Recordset
Dim iDeviceQty As Integer
Dim sOra As String
Dim sDevice As String

If txtDN.Text = "" Then
    MsgBox "DN不可以为空", vbInformation
    Exit Sub
End If

sOra = "select * from ST_TR_SEQ order by seqtime, dev"
Set rs = Get_OracleRs(sOra)

If rs.BOF Then
'    PrintOPLable
    Exit Sub
End If

rs.MoveFirst
Do While Not rs.EOF

    Call InsertDB(rs)
    rs.MoveNext
Loop

Sleep (2000)

' 同时打印内盒卷盘标签
PrintInLable

cmdOuterPrinter.Visible = True

End Sub

Private Sub cmdOuterPrinter_Click()

Dim sOra As String
Dim iMax As Integer
Dim i As Integer
Dim rs As New ADODB.Recordset

sOra = "select max(outbox_num) from PACKING_DETAILED where dn_num = '" & txtDN.Text & "'"

iMax = Get_OracleNo(sOra)

Frm_37_QboxLabel.Show
        
Frm_37_QboxLabel.Hide

For i = 1 To iMax

    sOra = "select distinct INBOX_NUM  from PACKING_DETAILED where outbox_num = '" & i & "' and dn_num = '" & txtDN.Text & "'"

    Set rs = Get_OracleRs(sOra)
    
    If Not rs.BOF Then
      
            rs.MoveFirst
            Do While Not rs.EOF
                Frm_37_QboxLabel.ComDN.Text = txtDN.Text
        
                Frm_37_QboxLabel.TxtWaferIDOut.Text = Frm_37_QboxLabel.TxtWaferIDOut.Text & rs.fields(0).Value & vbCrLf
        
                rs.MoveNext
            Loop
                
    End If
        
    Call Frm_37_QboxLabel.CmdOKOut_Click
    Sleep (2000)
    
Next

' 打印总标签
FrmSemtech_LablePrint.Show
FrmSemtech_LablePrint.Hide

FrmSemtech_LablePrint.cmbDN.Text = txtDN.Text

FrmSemtech_LablePrint.Opt(2).Value = True
    ' 查询
    Call FrmSemtech_LablePrint.cmd_Click(0)
    
    With FrmSemtech_LablePrint.fps(0)

        For j = 1 To .MaxRows
            .Row = j
            
            .Col = 1
            .Text = 1
        Next

    End With
        
    ' 打印
    Sleep (2000)
    Call FrmSemtech_LablePrint.cmd_Click(2)
    
End Sub

Private Sub PrintOutLable()








End Sub


Private Sub cmdReset_Click()
Unload Me
frmSemtechLPS.Show
End Sub

Private Sub Form_Activate()
txtScan.SetFocus
End Sub

Private Sub Form_Load()

' 初始化fps
InitFps
InitTxtStatus

End Sub

Private Sub InitFps()

With fps
    .ReDraw = False
    .MaxCols = E_F_PO.E_End - 1
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
        
    ' 定义表头名
    .SetText E_LPS.E_JobNO, 0, "应挑JOB"
    .SetText E_LPS.E_QTY, 0, "应挑数量"
    .SetText E_LPS.E_JobNO2, 0, "已挑JOB"
    .SetText E_LPS.E_Qty2, 0, "已挑数量"
          
    ' 定义宽度
    .ColWidth(1) = 12
    .ColWidth(2) = 10
    .ColWidth(3) = 12
    .ColWidth(4) = 10
    
    ' 定义高度
    .RowHeight(0) = 20
    .RowHeight(-1) = 15

    .ReDraw = True
End With

End Sub

Private Sub InitTxtStatus()

Dim iLotLen As Integer

iLotLen = (Len(txtHTLot.Text) - Len(Replace$(txtHTLot.Text, vbCrLf, ""))) / 2

txtStatus.Text = vbCrLf & iLotLen

End Sub

Private Sub txtScan_KeyPress(KeyAscii As Integer)

' 扫描结束触发
If KeyAscii <> 13 Then
    Exit Sub
End If

txtStatus.ForeColor = vbBlue

' 抓取DN判断赋值
InitDN

' Job数量核对
MatchJobQty

' 清空
ClearTxtScan

End Sub

Private Sub InitDN()

Dim sFsChar As String
Dim sDbChar As String
Dim sSelFsChar As String
Dim sSelDbChar As String

sFsChar = Left$(Trim(txtScan.Text), 1)
sDbChar = Left$(Trim(txtScan.Text), 2)
sSelFsChar = Mid$(Trim(txtScan.Text), 2)
sSelDbChar = Mid$(Trim(txtScan.Text), 3)
sSelFuChar = Trim$(txtScan.Text)

' DN获取
If sFsChar = "I" Then
    txtDN.Text = sSelFsChar
    
    ' 判断是否合法
    If Get_OracleCnt("select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & sSelFsChar & "'") = 0 Then
        
        MsgBox "扫描的DN不正确, 请确认", vbInformation
        Exit Sub
    End If
    
    ' 合法: 初始化fps
    AssignFps (sSelFsChar)
    
    ' 清空: ST_TR_SEQ
    ClearST_TR_SEQ
End If

End Sub

Private Sub MatchJobQty()

Dim sFsChar As String
Dim sDbChar As String
Dim sSelFsChar As String
Dim sSelDbChar As String
Dim sSelFuChar As String
Dim sSql As String
Dim rs As New ADODB.Recordset

sFsChar = Left$(Trim(txtScan.Text), 1)
sDbChar = Left$(Trim(txtScan.Text), 2)
sSelFsChar = Mid$(Trim(txtScan.Text), 2)
sSelDbChar = Mid$(Trim(txtScan.Text), 3)
sSelFuChar = Trim$(txtScan.Text)

' Job获取
If sFsChar = "S" Then
    If InStr(txtHTLot.Text, sSelFsChar) Then
       ' MsgBox "请不要扫描同一个卷盘号", vbInformation
        
        txtStatus.ForeColor = vbRed
        Exit Sub
    End If

    If Get_SqlserverCnt("select * from [erpdata].[dbo].TblTSV_Tray_details where TRAYQBOXNUMBER = '" & sSelFuChar & "'") = 0 Then
        MsgBox "扫描的LotID不正确, 请确认", vbInformation
        Exit Sub
    End If

    ' 判断是否Lot和DN是否挂钩
    sSql = "select * from [erpdata].[dbo].TblTSV_Tray_details where TRAYQBOXNUMBER = '" & sSelFuChar & "'"
    Set rs = Get_SqlserveRs(sSql)
    
    If Get_OracleCnt("select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & txtDN.Text & "' and batchnumber = '" & rs.fields("Customerlotid").Value & "' ") = 0 Then
        MsgBox "扫描的LotID和DN不匹配, 请确认", vbInformation
        Exit Sub
    Else
        txtHTLot = txtHTLot & sSelFsChar & vbCrLf
        
        ' 扫1个卷盘插1笔数据
        sOra = "insert into ST_TR_SEQ values('" & txtDN.Text & "', '" & rs.fields("Customerlotid").Value & "', '" & rs.fields("CUSTOMERPT").Value & "', '" & rs.fields("QTY").Value & "', sysdate,'" & sSelFuChar & "' )"
        Exec_Ora (sOra)
        
    End If
    
    ' 更新Fps
    AssignFps (sSelFuChar)
End If

' 数量状态变更
InitTxtStatus

End Sub

Private Sub AssignFps(sSel As String)

Dim sOra As String
Dim rs As New ADODB.Recordset
Dim iLotLen As Integer
Dim irow As Integer
Dim sJobNo As String
Dim sLotQty As Long
Dim sRightQty As String
Dim sPreQty As String
Dim bPrintCheck As Boolean

iSum = 0
iSumPre = 0
bPrintCheck = True

iLotLen = (Len(txtHTLot.Text) - Len(Replace$(txtHTLot.Text, vbCrLf, ""))) / 2

If iLotLen = 0 Then
    ' 初始化
    sOra = "select batchnumber, sum(quantity), '', '0' from CUSTOMERSHIPPINGUPTBL where delivery = '" & sSel & "' group by batchnumber"

    Set rs = Get_OracleRs(sOra)

    With fps
        .MaxRows = 0
        If rs.RecordCount > 0 Then
            Set .DataSource = rs
        End If
    End With

Else
    ' 更新
    sOra = "select * from [erpdata].[dbo].TblTSV_Tray_details where TRAYQBOXNUMBER = '" & sSel & "'"
    
    Set rs = Get_SqlserveRs(sOra)
    
    With fps
        For irow = 1 To .MaxRows
            .Row = irow
            .Col = 1
            
            If .Text = rs.fields("Customerlotid").Value Then
                .Col = 3
                .Text = rs.fields("Customerlotid").Value
                
                .Col = 4
                .Text = Str(Val(.Text) + rs.fields("Qty").Value)
            End If
        Next
    End With

    ' 判断总数
    With fps
        For irow = 1 To .MaxRows
              .Row = irow
              .Col = 2
              
              sRightQty = .Text
              
              .Col = 4
              sPreQty = .Text
              
              If Val(sPreQty) <> Val(sRightQty) Then
                bPrintCheck = False
              End If
        Next
    End With
    
    If bPrintCheck Then
        cmdInnerTrayPrinter.Visible = True
        'cmdOuterPrinter.Visible = True
        txtScan.Locked = True
    End If
End If

End Sub

' 按机种分类插入数据
Private Sub InsertDB(rs As ADODB.Recordset)

Dim sOra As String
Dim tData As tSTData
Dim lCnt As Long

' 外箱ID
sOra = "select count(outbox_num) from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "' " & _
" and outbox_num in (select nvl(max(outbox_num), '1') from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "')"
lCnt = Get_OracleNo(sOra)

If lCnt <= 107 Then
    sOra = "select * from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "'"
    If Get_OracleCnt(sOra) > 0 Then
        sOra = "select nvl(max(outbox_num), '1') from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "' "
        tData.OUTBOX_NUM = Get_OracleNo(sOra)
    Else
        sOra = " select nvl(max(outbox_num), '0') +1 from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "'"
        tData.OUTBOX_NUM = Get_OracleNo(sOra)
    End If
    
Else
    sOra = "select (nvl(max(outbox_num), '1') + 1) from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "' "
    tData.OUTBOX_NUM = Get_OracleStr(sOra)
End If

' 内盒ID
sOra = "select count(inbox_num) from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "'  and outbox_num = '" & tData.OUTBOX_NUM & "' " & _
"and inbox_num in (select nvl(max(inbox_num), '1') from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "'  and outbox_num = '" & tData.OUTBOX_NUM & "')"
lCnt = Get_OracleNo(sOra)

If lCnt <= 8 Then
    sOra = "select nvl(max(inbox_num), '1') from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "'  and outbox_num = '" & tData.OUTBOX_NUM & "'"
    tData.INBOX_NUM = Get_OracleStr(sOra)
Else
    sOra = "select nvl(max(inbox_num), '1') + 1 from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "'  and outbox_num = '" & tData.OUTBOX_NUM & "'"
    tData.INBOX_NUM = Get_OracleStr(sOra)
End If

' 卷盘ID
tData.TRAYID = rs.fields("LOTID")
tData.CREATE_BY = gUserName
tData.DN_NUM = rs.fields("DN")
tData.JOB_ID = rs.fields("JOB")
tData.CUSTOMER_DEVICE = rs.fields("DEV")
tData.qty = 15000

' 插入数据
Call insertToSql(tData)

' 打印外箱
'PrintOPLable

End Sub


Private Sub insertToSql(tData As tSTData)

Dim sOra As String

sOra = "insert into PACKING_DETAILED values('" & tData.TRAYID & "','" & tData.INBOX_NUM & "','" & tData.OUTBOX_NUM & "','" & tData.DN_NUM & "','" & tData.JOB_ID & "','" & tData.qty & "','" & tData.CUSTOMER_DEVICE & "',sysdate,'" & tData.CREATE_BY & "','0','0','') "

Exec_Ora (sOra)

End Sub

Private Sub ClearST_TR_SEQ()

Dim sOra As String

sOra = "delete from ST_TR_SEQ"

Exec_Ora (sOra)

End Sub

Private Sub ClearTxtScan()

txtScan.Text = ""

End Sub

Private Sub PrintInLable()

' 合内箱
Dim rs As New ADODB.Recordset
Dim sInfo As String
Dim sOra As String
Dim sAppend As String
Dim iOp As Integer
Dim iIp As Integer
Dim iOpMax As Integer
Dim iIpMax As Integer
Dim FName As String
Dim dirtemp As String

sInfo = ""
iOp = 1
iIp = 1

sOra = "select max(outbox_num) from PACKING_DETAILED where dn_num = '" & txtDN.Text & "'  "
iOpMax = Get_OracleNo(sOra)
iIpMax = 9

For iOp = 1 To iOpMax
    For iIp = 1 To iIpMax
        sAppend = "select * from PACKING_DETAILED where dn_num = '" & txtDN.Text & "' and outbox_num = '" & iOp & "' and inbox_num = '" & iIp & "' "
        Set rs = Get_OracleRs(sAppend)
        
        If Not rs.BOF Then
        
            rs.MoveFirst
            Do While Not rs.EOF
                sInfo = sInfo & rs.fields("trayid") & vbCrLf
                rs.MoveNext
            Loop
                
        End If
        
        Call PrintInbox(sInfo, txtDN.Text)
        
        ' 第一个内盒打印完毕
        sInfo = ""
        
        Sleep (8000)
    Next
    
Next

End Sub
