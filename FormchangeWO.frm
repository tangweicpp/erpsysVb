VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FormchangeWO 
   Caption         =   "订单信息修改"
   ClientHeight    =   10905
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14820
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
   ScaleHeight     =   10905
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Fr1 
      Height          =   12375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      Begin VB.TextBox txtDelFrom 
         Height          =   375
         Left            =   8520
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkCheckAll 
         Caption         =   "全选"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10080
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFC0FF&
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H000000FF&
         Caption         =   "删除"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmdquery 
         BackColor       =   &H00C0C000&
         Caption         =   "查询"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H0000C000&
         Caption         =   "提交修改"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtLOT 
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo CmbCustomer 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   6015
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   1560
         Width           =   13695
         _Version        =   524288
         _ExtentX        =   24156
         _ExtentY        =   10610
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
         SpreadDesigner  =   "FormchangeWO.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lbl1223 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请求人员"
         Height          =   195
         Left            =   7800
         TabIndex        =   13
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblLOT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOT号:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4200
         TabIndex        =   3
         Top             =   450
         Width           =   660
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   420
         Width           =   1035
      End
      Begin VB.Label lblC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   45
      End
   End
End
Attribute VB_Name = "FormchangeWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0

    E_CustName = 1
    e_Lot
    e_Wafer
    E_id
    E_PO
    e_DEVICE
    E_GoodDie
    E_NGDie
    E_idd
    E_Check
     
    E_End

End Enum

Dim mainItemRS As New ADODB.Recordset

Dim reportRS   As New ADODB.Recordset

Private Sub chkCheckAll_Click()

    Dim i As Integer

    If chkCheckAll.Value = 1 Then

        For i = 1 To Fps(0).MaxRows

            With Fps(0)
                .Row = i
                .Col = E_FPS0.E_Check
                .Text = 1

            End With

        Next i
        
    ElseIf chkCheckAll.Value = 0 Then

        For i = 1 To Fps(0).MaxRows

            With Fps(0)
                .Row = i
                .Col = E_FPS0.E_Check
                .Text = 0

            End With

        Next i
        
    End If

End Sub

Private Sub cmdDel_Click()

    If txtDelFrom.Text = "" Then
        MsgBox "请输入要求删除的员工姓名", vbInformation, "警告"
        Exit Sub

    End If

    Dim nWoTemp  As WoWafer

    Dim userid   As String

    Dim id       As Long

    Dim Exitflag As Boolean

    Dim nDelSeq  As Long
    
    nDelSeq = Get_OracleNo("select DELWOSEQ.NEXTVAL  from dual")

    userid = UCase(gUserName)
    Exitflag = False

    If Trim(CmbCustomer.Text) = "" Or Trim(txtLOT.Text) = "" Then
        MsgBox "请输入客户及LOT系统！"
        cmdDel.Enabled = True
        Exit Sub

    End If

    With Fps(0)

        For i = 1 To .MaxRows

            .Row = i
            .Col = E_FPS0.E_Check

            If .Text = 1 Then
                Exitflag = True
                '要修改

                .Row = i
                .Col = 1
                nWoTemp.CustName = .Text
    
                .Row = i
                .Col = 2
                nWoTemp.lot = .Text
    
                .Row = i
                .Col = 3
                nWoTemp.wafer = .Text
    
                .Row = i
                .Col = 4
                nWoTemp.id = .Text
    
                .Row = i
                .Col = 5
                nWoTemp.PO = .Text
    
                .Row = i
                .Col = 6
                nWoTemp.device = .Text
    
                .Row = i
                .Col = 7
                nWoTemp.gooddie = .Text
    
                .Row = i
                .Col = 8
                nWoTemp.ngdie = .Text
    
                .Row = i
                .Col = 9
                nWoTemp.idd = .Text
    
                If wowaferdet(nWoTemp) Then
    
                    MsgBox "这笔：" & nWoTemp.wafer & "  已开工单!"
                    Exit Sub
    
                End If
    
                '备份数据
                Call backupWo(userid, nWoTemp.wafer, nDelSeq)
    
                Call delwafer(nWoTemp)
    
                If wolotdet(nWoTemp) = False Then
    
                    Call dellot(nWoTemp)
    
                End If

            End If

        Next i

        If Exitflag = False Then
            MsgBox "请勾选要删除的wafer"
            Exit Sub

        End If

    End With

    '--------------------
    MsgBox "删除成功!", vbInformation, "友情提示"
    
    ' 发送邮件
    
    Dim FSO        As New FileSystemObject

    Dim LogFile    As TextStream
    
    Dim strDatas   As String

    Dim strRowData As String

    Dim strColData As String

    Dim strSql     As String

    Dim j          As Integer
    
    Dim rs         As New ADODB.Recordset
    
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "DEL_WO_" & Format(g_Date, "YYYYMMDD") & nDelSeq & ".csv")
    
    strDatas = "CUSTOMER,WAFER_ID,LOT,GOOD_DIE,BAD_DIE,CUSTOMER_DEVICE, PO_NUM,DATE," & vbCrLf
               
    strSql = " select CUSTOMER,WAFER_ID,LOT,GOOD_DIE,BAD_DIE,CUSTOMER_DEVICE, PO_NUM, sysdate from  WO_BACKUP where DELSEQ = '" & nDelSeq & "' "

    strRowData = ""

    If rs.State = adStateOpen Then rs.Close
    If Cnn.State = 0 Then
        ConOracle

    End If

    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        
        For j = 0 To rs.Fields.Count - 1
             
            strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
                        
        Next
        
        If i = maxRow Then
            strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
            strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    '发邮件
    Dim strRecipient   As String

    Dim strRecipientCC As String

    Dim strTitle       As String
    
    strRecipient = "wei.tang_ks@ht-tech.com"
    strRecipientCC = "xue.liu_ks@ht-tech.com"
    
    strTitle = "<订单删除:" & txtLOT.Text & ">" & "<请求人员:" & txtDelFrom.Text & ">" & "<操作员:" & gUserName & ">"
        
    Call MailDetail_TW(strTitle, strRecipient, g_Path & "\" & "DEL_WO_" & Format(g_Date, "YYYYMMDD") & nDelSeq & ".csv", strRecipientCC)
    
    Call ShowData_WhereCus(UCase(Trim(CmbCustomer.Text)), UCase(Trim(txtLOT.Text)))
 
End Sub

Private Sub cmdExit_Click()
    Unload Me

End Sub

Private Sub Form_Activate()

'    If gUserName <> "07885" Then
'
'        cmdModify.Visible = False
'        cmdDel.Visible = False
'
'    End If

End Sub

Private Sub Form_Load()

    IniCustomerName

    With Fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
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

        .Col = E_FPS0.E_Check
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        .SetText E_FPS0.E_CustName, 0, "客户代码"
        .SetText E_FPS0.e_Lot, 0, "LotID"
        .SetText E_FPS0.e_Wafer, 0, "WaferID"
        .SetText E_FPS0.E_id, 0, "NO"
        .SetText E_FPS0.E_PO, 0, "PO"
        .SetText E_FPS0.e_DEVICE, 0, "客户机种"
        .SetText E_FPS0.E_GoodDie, 0, "GoodDies"
        .SetText E_FPS0.E_NGDie, 0, "NGDies"
        .SetText E_FPS0.E_idd, 0, "ID_FLAG"
        .SetText E_FPS0.E_Check, 0, "选择"
       
        .ColWidth(E_FPS0.E_CustName) = 10
        .ColWidth(E_FPS0.e_Lot) = 10
        .ColWidth(E_FPS0.e_Wafer) = 15
        .ColWidth(E_FPS0.E_id) = 5
        .ColWidth(E_FPS0.E_PO) = 15
        .ColWidth(E_FPS0.e_DEVICE) = 10
        .ColWidth(E_FPS0.E_GoodDie) = 12
        .ColWidth(E_FPS0.E_NGDie) = 12
        .ColWidth(E_FPS0.E_idd) = 10
        .ColWidth(E_FPS0.E_Check) = 5
        
        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .Col = E_FPS0.e_DEVICE
        .Lock = False
        .CellType = CellTypeEdit
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = vbCyan
                
        If gUserName = "16642" Then
            .Col = E_FPS0.E_PO
            .Lock = False
            .CellType = CellTypeEdit
            .BackColorStyle = BackColorStyleUnderGrid
            .BackColor = vbCyan
       
        End If

        .Col = E_FPS0.E_GoodDie
        .Lock = False
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = vbCyan
                
        .Col = E_FPS0.E_NGDie
        .Lock = False
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = vbCyan
                
        .Col = E_FPS0.E_Check
        .Lock = False
        
        .ReDraw = True

    End With
    
End Sub

Private Sub IniCustomerName()
    Set mainItemRS = GetJDCustomerName()
    Set CmbCustomer.RowSource = mainItemRS
    CmbCustomer.ListField = mainItemRS("productname").Name
    CmbCustomer.BoundColumn = mainItemRS("PID").Name

End Sub

Private Sub cmdQuery_Click()

MsgBox "该接口已关闭, 请至新版订单维护系统操作", vbInformation, "提示"
Exit Sub

    If Trim(CmbCustomer.Text) = "" Or Trim(txtLOT.Text) = "" Then
        MsgBox "请输入客户及LOT系统！"
   
        Exit Sub

    End If

    'ShowData_WhereCus(CmbCustomer.Text,txtLOT.Text)
    Call ShowData_WhereCus(UCase(Trim(CmbCustomer.Text)), UCase(Trim(txtLOT.Text)))

End Sub

Private Sub ShowData_WhereCus(customerTemp As String, lottemp As String)

    Set reportRS = Getwo_wafer(customerTemp, lottemp)

    With Fps(0)
        .MaxRows = 0

        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS

       
        End If

    End With

End Sub

Private Sub cmdModify_Click()

    Dim nWoTemp  As WoWafer

    Dim userid   As String

    Dim id       As Long

    Dim Exitflag As Boolean
    
    If txtDelFrom.Text = "" Then
        MsgBox "请输入要求修改订单的人员姓名"
        Exit Sub
    End If
    
    Exitflag = False
    userid = UCase(gUserName)
    id = GetMaxID()
    
    Dim nDelSeq As Long
    
    nDelSeq = Get_OracleNo("select DELWOSEQ.NEXTVAL  from dual")

    If Trim(CmbCustomer.Text) = "" Or Trim(txtLOT.Text) = "" Then
        MsgBox "请输入客户及LOT 点击查询， 编辑！"
        cmdModify.Enabled = True
        Exit Sub

    End If

    If MsgBox("确认数据无误？", vbOKCancel, "提示") = vbCancel Then
        Exit Sub

    End If

    '-----------
    With Fps(0)

        For i = 1 To .MaxRows

            .Row = i
            .Col = E_FPS0.E_Check

            If .Text = 1 Then
                Exitflag = True
        
                '要修改
                .Row = i
                .Col = 1
                nWoTemp.CustName = .Text
    
                .Row = i
                .Col = 2
                nWoTemp.lot = .Text
    
                .Row = i
                .Col = 3
                nWoTemp.wafer = .Text
    
                .Row = i
                .Col = 4
                nWoTemp.id = .Text
    
                .Row = i
                .Col = 5
                nWoTemp.PO = .Text
    
                .Row = i
                .Col = 6
                nWoTemp.device = .Text
    
                .Row = i
                .Col = 7
                nWoTemp.gooddie = .Text
    
                .Row = i
                .Col = 8
                nWoTemp.ngdie = .Text

'                If wowaferdet(nWoTemp) Then
'                    MsgBox "这笔：" & nWoTemp.wafer & "  已开工单, 不可以修改!"
'                    Exit Sub
'
'                End If
            
                '备份数据
                Call backupWo(userid, nWoTemp.wafer, nDelSeq)
    
                Call modifwafer(userid, id, nWoTemp)

            End If

        Next i

        If Exitflag = False Then
            MsgBox "请先选择再提交修改"
            Exit Sub

        End If

        Call Modifywo(userid, id, nWoTemp)

    End With

    '--------------------
    MsgBox "修改成功!", vbInformation, "友情提示"
    
    ' 发送邮件
    Dim FSO     As New FileSystemObject

    Dim LogFile As TextStream
    
    Dim strDatas As String
    

    Dim strRowData  As String
    Dim strColData  As String
    Dim strSql      As String
    Dim j           As Integer
    
    Dim rs As New ADODB.Recordset
    
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "MOD_WO_" & Format(g_Date, "YYYYMMDD") & nDelSeq & ".csv")
    
    strDatas = "CUSTOMER,WAFER_ID,LOT,GOOD_DIE,BAD_DIE,CUSTOMER_DEVICE, PO_NUM,DATE," & vbCrLf
               
    strSql = " select CUSTOMER,WAFER_ID,LOT,GOOD_DIE,BAD_DIE,CUSTOMER_DEVICE, PO_NUM, sysdate from  WO_BACKUP where DELSEQ = '" & nDelSeq & "' "

    strRowData = ""

    If rs.State = adStateOpen Then rs.Close
    If Cnn.State = 0 Then
        ConOracle

    End If

    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        
        For j = 0 To rs.Fields.Count - 1
             
            strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
                        
        Next
        
        If i = maxRow Then
            strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
            strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
     '发邮件
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    Dim strTitle As String
    
    strRecipient = "wei.tang_ks@ht-tech.com"
    strRecipientCC = "xue.liu_ks@ht-tech.com"
    
    strTitle = "<订单修改:" & txtLOT.Text & ">" & "<请求人员:" & txtDelFrom.Text & ">" & "<操作员:" & gUserName & ">"
        
    Call MailDetail_TW(strTitle, strRecipient, g_Path & "\" & "MOD_WO_" & Format(g_Date, "YYYYMMDD") & nDelSeq & ".csv", strRecipientCC)
    
    Call ShowData_WhereCus(UCase(Trim(CmbCustomer.Text)), UCase(Trim(txtLOT.Text)))
 
End Sub

Private Sub backupWo(UID As String, WaferID As String, nDelSeq As Long)

    Dim sOra As String

    Dim woB  As WOBACKUP

    Dim rs   As ADODB.Recordset

    sOra = "select a.CUSTOMERSHORTNAME as CUSTOMER, b.SUBSTRATEID as WAFER_ID, b.LOTID as LOT, b.PASSBINCOUNT as GOOD_DIE, b.FailBinCount as BAD_DIE, a.mpn_desc as CUSTOMER_DEVICE, a.PO_NUM as PO_NUM, a.fab_conv_id as FAB_Device, a.test_site as SHIP_TO,a.imager_customer_rev as SEC_CODE, a.chromaticity as CUST_SEC_CODE, a.QTECH_CREATED_DATE as CREATE_DATE, a.QTECH_CREATED_BY as  CREATE_BY, sysdate as LASTUPDATE, '" & UID & "' as LASTUPDATE_BY, 'D' as EVENT from customeroitbl_test a, mappingdatatest b where a.id = b.filename and b.substrateid IN ('" & WaferID & "')"
    Set rs = Get_OracleRs(sOra)

    woB.Customer = GetRsData(rs, "CUSTOMER")
    woB.WAFER_ID = GetRsData(rs, "WAFER_ID")
    woB.lot = GetRsData(rs, "LOT")
    woB.GOOD_DIE = GetRsData(rs, "GOOD_DIE")
    woB.BAD_DIE = GetRsData(rs, "BAD_DIE")
    woB.CUSTOMER_DEVICE = GetRsData(rs, "CUSTOMER_DEVICE")
    woB.PO_NUM = GetRsData(rs, "PO_NUM")
    woB.FAB_Device = GetRsData(rs, "FAB_Device")
    woB.ship_to = GetRsData(rs, "SHIP_TO")
    woB.SEC_CODE = GetRsData(rs, "SEC_CODE")
    woB.CUST_SEC_CODE = GetRsData(rs, "CUST_SEC_CODE")
    woB.CREATE_DATE = GetRsData(rs, "CREATE_DATE")
    woB.CREATE_BY = GetRsData(rs, "CREATE_BY")
    woB.LASTUPDATE = GetRsData(rs, "LASTUPDATE")
    woB.LASTUPDATE_BY = GetRsData(rs, "LASTUPDATE_BY")
    woB.EVENT = GetRsData(rs, "EVENT")

    sOra = "Insert into wo_backup values('" & woB.Customer & "', '" & woB.WAFER_ID & "', '" & woB.lot & "', '" & woB.GOOD_DIE & "', '" & woB.BAD_DIE & "', '" & woB.CUSTOMER_DEVICE & "', '" & woB.PO_NUM & "','" & woB.FAB_Device & "', " & "'" & woB.ship_to & "', '" & woB.SEC_CODE & "','" & woB.CUST_SEC_CODE & "', '" & woB.CREATE_DATE & "', '" & woB.CREATE_BY & "', '" & woB.LASTUPDATE & "','" & woB.LASTUPDATE_BY & "', '" & woB.EVENT & "','" & nDelSeq & "')"

    Exec_Ora (sOra)

End Sub
