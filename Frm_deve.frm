VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_deve 
   Caption         =   "Form1"
   ClientHeight    =   11145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14745
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
   ScaleHeight     =   11145
   ScaleWidth      =   14745
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox ComboSql 
      Height          =   315
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   7200
      Width           =   2655
   End
   Begin VB.CommandButton cmdDeleteOrder 
      Caption         =   "删除工单"
      Height          =   645
      Left            =   7320
      TabIndex        =   16
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除条目"
      Enabled         =   0   'False
      Height          =   360
      Left            =   11520
      TabIndex        =   13
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame Fra 
      Caption         =   "删除工单"
      Height          =   4455
      Left            =   600
      TabIndex        =   10
      Top             =   6480
      Width           =   16215
      Begin VB.CommandButton cmdFLAG 
         Caption         =   "FLAG"
         Height          =   360
         Left            =   9840
         TabIndex        =   24
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtDelFrom 
         Height          =   375
         Left            =   11040
         TabIndex        =   23
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton cmd 
         Caption         =   "退出"
         Height          =   720
         Left            =   12240
         TabIndex        =   21
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "重抛"
         Height          =   645
         Left            =   9840
         TabIndex        =   20
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "工单还原"
         Height          =   645
         Left            =   3120
         TabIndex        =   19
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdOrderDetails 
         Caption         =   "查询工单明细"
         Height          =   645
         Left            =   360
         TabIndex        =   15
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox ComboOrderTbl 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtOrderID 
         Height          =   285
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblasdsa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "删除需求人员"
         Height          =   195
         Left            =   9720
         TabIndex        =   22
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label lblOrderTblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "表名"
         Height          =   195
         Left            =   3120
         TabIndex        =   17
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单号"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   540
      End
   End
   Begin VB.Frame WO 
      Caption         =   "删除WO"
      Height          =   1215
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   14055
      Begin VB.CommandButton cmdDetails 
         Caption         =   "订单明细"
         Height          =   360
         Left            =   8040
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   3840
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtWaferID 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtLotID 
         Height          =   285
         Left            =   480
         TabIndex        =   2
         Top             =   708
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FileName"
         Height          =   195
         Left            =   4200
         TabIndex        =   5
         Top             =   480
         Width           =   645
      End
      Begin VB.Label lblWaferID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WaferID"
         Height          =   195
         Left            =   2520
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblLotID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LotID"
         Height          =   195
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   390
      End
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   1455
      Index           =   0
      Left            =   600
      TabIndex        =   8
      Top             =   1560
      Width           =   14055
      _Version        =   524288
      _ExtentX        =   24791
      _ExtentY        =   2566
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
      SpreadDesigner  =   "Frm_deve.frx":0000
      Appearance      =   2
      TextTip         =   2
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   3135
      Index           =   1
      Left            =   600
      TabIndex        =   9
      Top             =   3120
      Width           =   14055
      _Version        =   524288
      _ExtentX        =   24791
      _ExtentY        =   5530
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
      SpreadDesigner  =   "Frm_deve.frx":049A
      Appearance      =   2
      TextTip         =   2
   End
End
Attribute VB_Name = "Frm_deve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()

    Dim LOTID    As String

    Dim WaferID  As String

    Dim filename As String

    LOTID = UCase$(Trim$(txtLotID.Text))
    WaferID = Trim$(txtWaferID.Text)
    filename = Trim$(txtFileName.Text)

    If LOTID = "" Or filename = "" Then
        MsgBox ("请输入LOT号, FileName!")
    Else
        Call DelWO(LOTID, WaferID, filename)
    End If

End Sub

Private Sub cmdDeleteOrder_Click()

If txtDelFrom.Text = "" Then
    MsgBox "请输入需求人员姓名(部门)", vbCritical, "提醒"
    Exit Sub
End If

    Dim OrderID As String

    OrderID = UCase(Trim(txtOrderID.Text))

    If OrderID = "" Then
        MsgBox ("请输入工单号")
    Else
        Call DelOrder(OrderID)
    End If

End Sub

Private Sub cmdDetails_Click()

    ' 显示WO明细
    Dim LOTID    As String

    Dim WaferID  As String

    Dim filename As String

    LOTID = UCase$(Trim$(txtLotID.Text))
    WaferID = Trim$(txtWaferID.Text)
    filename = Trim$(txtFileName.Text)

    If LOTID = "" Then
        MsgBox ("请输入LOT号!")
    Else
        Call ShowDetails(LOTID, WaferID, filename)
    End If

End Sub

Private Sub ShowDetails(LOTID As String, WaferID As String, filename As String)

    Dim Str_Sql As String

    ' Header表
    Str_Sql = "select ct.* from customeroitbl_test ct where ct.source_batch_id = '" + LOTID + "'"
    Set mainItemRS = getStr(Str_Sql)

    With fps(0)
        .MaxRows = 0

        If mainItemRS.RecordCount > 0 Then
            Set .DataSource = mainItemRS
        End If

    End With

    Str_Sql = "select a.filename,a.customershortname,a.lotid,a.wafer_id, a.substrateid, a.productid, a.flag, a.qtech_created_date, a.qtech_created_by from mappingdatatest a WHERE a.lotid = '" + LOTID + "'"
    Set mainItemRS = getStr(Str_Sql)

    With fps(1)
        .MaxRows = 0

        If mainItemRS.RecordCount > 0 Then
            Set .DataSource = mainItemRS
        End If

    End With

End Sub

Private Sub DelOrder(OrderID As String)

    Dim Str_Sql       As String

    Dim STr_Sql1      As String

    Dim str_sql2      As String

    Dim Str_sql3      As String

    Dim STr_sql4      As String

    Dim STr_sql5      As String

    Dim str_sql6      As String

    Dim Str_sql7      As String

    Dim str_sql8      As String

    Dim str_sql9      As String

    Dim sty_sql10     As String

    Dim sty_sql11     As String

    Dim sty_sql12     As String

    Dim iRes          As Integer

    Dim Str_sql_Guard As String

    ' 加判断后再删除
    ' 0 是否退料
   Str_sql_Guard = "select SUM(实领数量) from [erpbase].[dbo].[tblllplan] where 工单号 =  '" + OrderID + "'"
    If Get_SqlserverNo(Str_sql_Guard) > 0 Then
    
        iRes = MsgBox("该工单还未全部退料,还要删除吗?", vbYesNoCancel, "提示:")
        If iRes <> vbYes Then

            Exit Sub

        End If
    End If
      
    ' 1 是否抛到金蝶
    Str_sql_Guard = "select * from erpdata..tblTSV_TLInfo a where a.工单号 = '" + OrderID + "'"

    If QuerySqlserver(Str_sql_Guard) Then
        iRes = MsgBox("工单已经抛到金蝶, 要继续删除吗?", vbYesNoCancel, "提示:")

        If iRes <> vbYes Then

            Exit Sub

        End If
    End If

    ' 2 产品是否在机台内
    Str_sql_Guard = "select a.RESOURCENAME from historymainline a,(select max(CONTAINERTXNSEQUENCE) mm, containername from historymainline " & "where containername in ( select conn.containername from container conn where conn.mfgorderid in  (select mfg.mfgorderid from mfgorder mfg where mfg.mfgordername  in ('" + OrderID + "')) ) " & "group by containername) b where a.containername = b.containername and a.CONTAINERTXNSEQUENCE = b.mm and a.RESOURCENAME is not null order by a.RESOURCENAME"

    If QueryStr(Str_sql_Guard) Then
        iRes = MsgBox("在机台内, 要继续删除吗?", vbYesNoCancel, "提示:")

        If iRes <> vbYes Then

            Exit Sub

        End If
    End If

    ' 3 产品是否在生产
    Str_sql_Guard = "select * from mfgorder a, a_lotwafers b, mappingdatatest c,customeroitbl_test d,ib_wohistory e,container f," & "currentstatus g,spec h,operation i,workcenter j, specbase k,container l,product m, productbase n " & "Where b.workordername = a.mfgordername and c.substrateid = b.waferscribenumber and to_char(d.id) = c.filename and e.ordername = b.workordername " & "and f.containerid = b.containerid and g.currentstatusid = f.currentstatusid and h.specid = g.specid and i.operationid = h.operationid " & "and j.workcenterid = i.workcenterid and k.specbaseid = h.specbaseid and a.mfgordername = '" + OrderID + "' and l.containerid = b.containerid " & "and l.status = 1 and m.productid = l.productid and n.productbaseid = m.productbaseid and k.specname <> '3010' "

    If QueryStr(Str_sql_Guard) Then
        iRes = MsgBox("在生产, 要继续删除吗?", vbYesNoCancel, "提示:")

        If iRes <> vbYes Then

            Exit Sub

        End If
    End If

    ' 备份数据
    STr_Sql1 = "insert into container_bak select * from container conn where conn.mfgorderid in  (select mfg.mfgorderid from mfgorder mfg where mfg.mfgordername  in ('" + OrderID + "'))"
    str_sql2 = "insert into mfgorder_bak select * from mfgorder mfg where mfg.mfgordername in ('" + OrderID + "') "
    Str_sql3 = "insert into A_Lotwafers_bak select * from A_Lotwafers al where al.workordername in ('" + OrderID + "')"
    STr_sql4 = "insert into ib_wohistory_bak select * from ib_wohistory where ordername in ('" + OrderID + "') "
    STr_sql5 = "insert into ib_waferlist_bak select * from ib_waferlist where ordername in ('" + OrderID + "') "
    str_sql6 = "insert into [erpdata].[dbo].[tblTSVworkorder_bak] select * from  [erpdata].[dbo].[tblTSVworkorder] where ORDERNAME in ('" + OrderID + "') "
    Str_sql7 = "insert into [erpdata].[dbo].[tblTSVwaferlist_bak] select * from  [erpdata].[dbo].[tblTSVwaferlist] where ORDERNAME in ('" + OrderID + "')"
    str_sql8 = "insert into [erpbase].[dbo].[tblllplan_bak] select * from [erpbase].[dbo].[tblllplan] where 工单号 in ('" + OrderID + "')"
    str_sql9 = "insert into PJ_WO_PRI_bak select * from PJ_WO_PRI where wo in ('" & OrderID & "')"

    AddSql (STr_Sql1)
    AddSql (str_sql2)
    AddSql (Str_sql3)
    AddSql (STr_sql4)
    AddSql (STr_sql5)
    AddSql (str_sql9)

    AddSql2 (str_sql6)
    AddSql2 (Str_sql7)
    AddSql2 (str_sql8)

    MsgBox "备份成功", vbInformation, "提示"

    ' 删除
    STr_Sql1 = "delete from container conn where conn.mfgorderid in  (select mfg.mfgorderid from mfgorder mfg where mfg.mfgordername  in ('" + OrderID + "')) "
    str_sql2 = "delete from mfgorder mfg where mfg.mfgordername in ('" + OrderID + "')"
    Str_sql3 = "delete from A_Lotwafers al where al.workordername in ('" + OrderID + "')"
    STr_sql4 = "delete from ib_wohistory where ordername in ('" + OrderID + "')"
    STr_sql5 = "delete from ib_waferlist where ordername in ('" + OrderID + "')"

    str_sql6 = "delete from  [erpdata].[dbo].[tblTSVworkorder] where ORDERNAME in ('" + OrderID + "') "
    Str_sql7 = "delete from  [erpdata].[dbo].[tblTSVwaferlist] where ORDERNAME in ('" + OrderID + "')"
    str_sql8 = "delete from  [erpbase].[dbo].[tblllplan] where 工单号 in ('" + OrderID + "')"
    str_sql9 = "delete from PJ_WO_PRI where wo in ('" & OrderID & "')"
    AddSql2 ("delete from erpdata..shop_order where shop_order = '" & OrderID & "' ")

    getStr (STr_Sql1)
    getStr (str_sql2)
    getStr (Str_sql3)
    getStr (STr_sql4)
    getStr (STr_sql5)

    getSqlServerStr2 (str_sql6)
    getSqlServerStr2 (Str_sql7)
    getSqlServerStr2 (str_sql8)

    getStr (str_sql9)
    
    ' 发送邮件
    
        '发邮件
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    Dim strTitle As String
    
 '   strRecipient = "wei.tang_ks@ht-tech.com"
    strRecipientCC = "xue.liu_ks@ht-tech.com"
    
    strTitle = "<工单删除:" & OrderID & ">" & "<请求人员:" & txtDelFrom.Text & ">" & "<操作员:" & gUserName & ">"
        
    'Call MailDetail_TW(strTitle, strRecipient, "", strRecipientCC)
    
    MsgBox ("删除成功"), vbInformation, "提示"

End Sub

Private Sub DelWO(LOTID As String, WaferID As String, filename As String)

    Dim Str_Sql As String

    Str_Sql = "delete from mappingdatatest a WHERE a.lotid = '" + LOTID + "'"
    str_sql2 = "delete from [ERPBASE].[dbo].[tblmappingData]   WHERE lotid = '" + LOTID + "'"
    getStr (Str_Sql)
    getSqlServerStr2 (str_sql2)

    MsgBox ("删除成功")

End Sub

Private Sub cmdFLAG_Click()

If txtOrderID.Text = "" Then
    MsgBox "请输入工单号", vbInformation, "提示"
    Exit Sub
End If

Dim strOrderID As String
strOrderID = Trim$(UCase$(txtOrderID.Text))

Dim strSql As String

strSql = "update shop_order set flag = '0' where shop_order = '" & strOrderID & "'"

AddSql (strSql)

MsgBox "FLAG已置0", vbInformation, "提示"

End Sub

Private Sub cmdOrderDetails_Click()

    Dim OrderID  As String

    Dim Str_Sql  As String

    Dim OrderTbl As String

    OrderID = ""
    Str_Sql = ""
    OrderTbl = ""

    OrderID = UCase(Trim(txtOrderID.Text))
    OrderTbl = ComboOrderTbl.Text

    Select Case OrderTbl

        Case "container"
            Str_Sql = " select conn.* ,conn.rowid from container conn where conn.mfgorderid in  (select mfg.mfgorderid from mfgorder mfg where mfg.mfgordername  in ('" + OrderID + "'))"

        Case "mfgorder"
            Str_Sql = " select mfg.* , mfg.rowid from mfgorder mfg where mfg.mfgordername in ('" + OrderID + "') "

        Case "A_Lotwafers"
            Str_Sql = " select * from A_Lotwafers al where al.workordername in ('" + OrderID + "')"

        Case "ib_wohistory"
            Str_Sql = " select * from ib_wohistory where ordername in ('" + OrderID + "') "

        Case "ib_waferlist"
            Str_Sql = "select * from ib_waferlist where ordername in ('" + OrderID + "')"

        Case "shop_order"
            Str_Sql = "select * from shop_order where shop_order = '" + OrderID + "'"

        Case "shop_order_detail"
            Str_Sql = "select * from shop_order_detail where shop_order = '" + OrderID + "'"

        Case "shop_order_property"
            Str_Sql = "select * from shop_order_property where shop_order = '" + OrderID + "'"

        Case Else
            Str_Sql = ""
    End Select

    If Str_Sql <> "" Then
        Set mainItemRS = getStr(Str_Sql)

        With fps(1)
            .MaxRows = 0

            If mainItemRS.RecordCount > 0 Then
                Set .DataSource = mainItemRS
            End If

        End With

    End If

    Select Case ComboSql.Text

        Case "tblTSVworkorder"
            Str_Sql = " select * from  [erpdata].[dbo].[tblTSVworkorder] where ORDERNAME in  ('" + OrderID + "')"

        Case "tblTSVwaferlist"
            Str_Sql = " select * from  [erpdata].[dbo].[tblTSVwaferlist] where ORDERNAME in ('" + OrderID + "') "

        Case "tblllplan"
            Str_Sql = " select *  from  [erpbase].[dbo].[tblllplan] where 工单号 in ('" + OrderID + "')"

        Case "TblERPFLToME"
            Str_Sql = "select * from ERPBASE..TblERPFLToME where shop_order = '" + OrderID + "'"

        Case Else
            Str_Sql = ""
    End Select

    If Str_Sql <> "" Then
        Set mainItemRS = getSqlServerStr2(Str_Sql)

        With fps(1)
            .MaxRows = 0

            If mainItemRS.RecordCount > 0 Then
                Set .DataSource = mainItemRS
            End If

        End With

    End If

End Sub

Private Sub Command1_Click()

    Dim OrderID   As String

    Dim Str_Sql   As String

    Dim STr_Sql1  As String

    Dim str_sql2  As String

    Dim Str_sql3  As String

    Dim STr_sql4  As String

    Dim STr_sql5  As String

    Dim str_sql6  As String

    Dim Str_sql7  As String

    Dim str_sql8  As String

    Dim str_sql9  As String

    Dim sty_sql10 As String

    If txtOrderID.Text = "" Then
        MsgBox "请输入工单号", vbInformation, "警告"

        Exit Sub

    End If

    OrderID = UCase(Trim(txtOrderID.Text))

    STr_Sql1 = "insert into container select * from container_bak conn where conn.mfgorderid in  (select mfg.mfgorderid from mfgorder_bak mfg where mfg.mfgordername  in ('" + OrderID + "'))"
    str_sql2 = "insert into mfgorder select * from mfgorder_bak mfg where mfg.mfgordername in ('" + OrderID + "') "
    Str_sql3 = "insert into A_Lotwafers  select * from A_Lotwafers_bak al where al.workordername in ('" + OrderID + "')"
    STr_sql4 = "insert into ib_wohistory  select * from ib_wohistory_bak where ordername in ('" + OrderID + "') "
    STr_sql5 = "insert into ib_waferlist  select * from ib_waferlist_bak where ordername in ('" + OrderID + "') "
    str_sql6 = "insert into [erpdata].[dbo].[tblTSVworkorder]  select * from [erpdata].[dbo].[tblTSVworkorder_bak] where ORDERNAME in ('" + OrderID + "') "
   
 str_sql6 = "insert into [erpdata].[dbo].[tblTSVworkorder](SEQ_IBWO,ORDERNAME, ORDERTYPE, DESCRIPTION, EVENTTYPE, ERPUSER, PRODUCT, PRODUCTREVISION, QTY, PRODUCTBOM, ERPCREATEDATE, PLANSTARTDATE, PLANENDDATE, " & _
" CUSTOMER, SALESORDER, PRODUCTFAMILY, MODIFYFLAG, CUSTOMERPN, FABFACILITY, IMAGERREV, DESIGNID, MLEVEL235, MLEVEL260, NGFLAG, PARA1, PARA2, PARA3, PARA4, PARA5, " & _
" PARA6, PARA7, PARA8, PARA9, PARA10, PROTECTIVE_FILM_APLD, LOT_STATUS, MPN) " & _
" SELECT   SEQ_IBWO,ORDERNAME, ORDERTYPE, DESCRIPTION, EVENTTYPE, ERPUSER, PRODUCT, PRODUCTREVISION, QTY, PRODUCTBOM, ERPCREATEDATE, PLANSTARTDATE, PLANENDDATE, " & _
" CUSTOMER, SALESORDER, PRODUCTFAMILY, MODIFYFLAG, CUSTOMERPN, FABFACILITY, IMAGERREV, DESIGNID, MLEVEL235, MLEVEL260, NGFLAG, PARA1, PARA2, PARA3, PARA4, PARA5,  " & _
" PARA6, PARA7, PARA8, PARA9, PARA10, PROTECTIVE_FILM_APLD, LOT_STATUS, MPN FROM [erpdata].[dbo].[tblTSVworkorder_bak] WHERE (ORDERNAME IN  ('" + OrderID + "')) "
    
    
    Str_sql7 = "insert into [erpdata].[dbo].[tblTSVwaferlist] select * from [erpdata].[dbo].[tblTSVwaferlist_bak] where ORDERNAME in ('" + OrderID + "')"
    str_sql8 = "insert into [erpbase].[dbo].[tblllplan]  select * from [erpbase].[dbo].[tblllplan_bak] where 工单号 in ('" + OrderID + "')"
    str_sql9 = "insert into PJ_WO_PRI select * from PJ_WO_PRI_bak where wo in ('" & OrderID & "')"

    AddSql (STr_Sql1)
    AddSql (str_sql2)
    AddSql (Str_sql3)
    AddSql (STr_sql4)
    AddSql (STr_sql5)
    AddSql (str_sql9)

    AddSql2 (str_sql6)
    AddSql2 (Str_sql7)
    AddSql2 (str_sql8)

    MsgBox "数据恢复完成", vbInformation, "提示"

End Sub

Private Sub Command2_Click()

    Dim sOra1 As String

    Dim OID   As String

    OID = UCase(Trim(txtOrderID.Text))

    Select Case ComboOrderTbl.Text

        Case "shop_order"
            sOra1 = "insert into shop_order(SHOP_ORDER,PRD_ID, PRD_VER,ERP_ROUTING, ORDER_QTY, CUST_LOT_QTY, PLAN_STAR_DATE, PLAN_END_DATE, MANF_DEPT, MANF_DEPT_DESC, LOT_TYPE, PRIORITY, PKG, CUST_ID,ERP_CREATE_DATE,CREATOR,flag,ht_device,RELEASE_TYPE)" & _
               "select a.ordername as SHOP_ORDER, b.product as PRD_ID, 'A' as PRD_VER, '' as ERP_ROUTING, COUNT(distinct A.WAFERID) as ORDER_QTY, COUNT(distinct A.WAFERLOT) as CUST_LOT_QTY, B.PLANSTARTDATE AS PLAN_STAR_DATE, B.PLANENDDATE AS PLAN_END_DATE, B.PARA8 AS MANF_DEPT, g.manf_dept_desc AS MANF_DEPT_DESC, e.lot_type as LOT_TYPE, decode(e.pri, 'Hot Lot', 1,'Super Hot Lot',1,4) as PRIORITY, f.pkg_type as PKG, shop_order_property_pkg.SHOP_ORDER_CUSTOMER(b.customer) AS CUST_ID, b.erpcreatedate as ERP_CREATE_DATE, '" + gUserName + "' as CREATOR, '0' as flag, f.qtechptno as ht_device, '1' as RELEASE_TYPE " & _
               "from ib_waferlist a, ib_wohistory b, MAPPINGDATATEST C, PJ_WO_PRI e, tbltsvnpiproduct f, MES_DEPT g " & _
               "where b.ordername = a.ordername and b.ordername = '" & OID & "' AND C.SUBSTRATEID = a.waferid and e.wo = b.ordername and f.qtechptno2 = b.product  and g.manf_dept = substr(b.para8,1,instr(b.para8,'_')-1) group by a.ordername,b.product,B.PLANSTARTDATE,B.PLANENDDATE,B.PARA8,e.lot_type,e.pri,f.pkg_type,shop_order_property_pkg.SHOP_ORDER_CUSTOMER(b.customer),b.erpcreatedate,e.creat_by,f.qtechptno,g.manf_dept_desc"
            AddSql (sOra1)
    
        Case "shop_order_detail"
            sOra1 = "insert into shop_order_detail(SHOP_ORDER,CUST_LOT_ID,WAFER_ID,GROSS_DIE_QTY,GOOD_DIE_QTY, MARK_CODE) select a.ordername   as SHOP_ORDER,a.waferlot    as CUST_LOT_ID,a.waferid     as WAFER_ID,a.dieqty      as GROSS_DIE_QTY,a.fgdieqty    as GOOD_DIE_QTY,a.markingcode as MARK_CODE " & " from ib_waferlist a, ib_wohistory b, MAPPINGDATATEST C, PJ_WO_PRI e Where b.ordername = a.ordername and b.ordername = '" & OID & "' AND C.SUBSTRATEID = a.waferid and e.wo = b.ordername"
            AddSql (sOra1)
        
        Case "shop_order_property"
            sOra1 = "select shop_order_property_pkg1.SHOP_ORDER_PROPERTY('" & OID & "') from dual"
            AddSql (sOra1)
          
        Case Else

    End Select
    
    If ComboSql.Text = "TblERPFLToME" Then
        sSql1 = "insert into ERPBASE..TblERPFLToME (STOCK_TYPE,STOCK_ID,PRD_ID,PRD_VER,QTY,PRD_DATE,EFF_DATE,SHOP_ORDER,SupSN,Flag,Memo,CreateDate,FStauts,HeaderID) " & "select 'W',b.ORDERNAME +c.WAFERLOT,e.料号,'A',COUNT(*),GETDATE() - 1,GETDATE() + 300,b.ORDERNAME ,c.WAFERLOT,0,'',GETDATE(),'','' from erpdata .. tblTSVworkorder b, " & "erpdata .. tblTSVwaferlist c,ERPBASE .. tblllplan d,erpdata..tblSmainM2 e where c.ORDERNAME = b.ORDERNAME and b.ORDERNAME in ('" & OID & "') and d.工单号 = c.ORDERNAME " & "and (d.物料编号 like '01.01.01%' or d.物料编号 like '03.06.02%') and e.物料编号 = d.物料编号 group by b.PRODUCT, b.ORDERNAME,e.料号, c.WAFERLOT "

        AddSql2 (sSql1)
    End If
    
    MsgBox "重抛成功", vbInformation, "提示"
    
End Sub

Private Sub Command3_Click()
'
'On Error GoTo ERRON
'
'Cnn.BeginTrans
'Dim start_shop_order  As String
'Dim lv_bonded         As String
'Dim lv_cust_rework    As String
'Dim list_lot          As String
'Dim lv_propertyname   As String
'Dim lv_propertyvalue  As String
'Dim ln_idx            As Integer
'Dim i                 As Integer
'Dim list_wafer        As String
'Dim lv_propertyname1  As String
'Dim lv_propertyvalue1 As String
'Dim ln_idx1           As Integer
'Dim j                 As Integer
'Dim k                 As Integer
'Dim lv_customer       As String
'Dim list_order        As String
'Dim lv_wo_date        As String
'Dim lv_wo_date_string As String
'Dim strSql As String
'
''1.保税非保税
'If Mid(start_shop_order, 1, 1) = "A" Then
'    lv_bonded = "Y"
'Else
'    lv_bonded = "N"
'
'End If
'
'strSql = "insert into shop_order_property(shop_order, propertyname, propertyvalue, levelid) Values('" & start_shop_order & "', 'BONDED', '" & lv_bonded & "', '1') "
'If AddSql(strSql) = 0 Then
'    MsgBox "保税非保税插入失败", vbInformation, "提示"
'    Exit Sub
'End If
'
''2.重工非重工
'If Mid(start_shop_order, 2, 1) = "R" Then
'    lv_cust_rework = "Y"
'Else
'    lv_cust_rework = "N"
'End If
'
'strSql = "insert into shop_order_property(shop_order, propertyname, propertyvalue, levelid) Values('" & start_shop_order & "', 'CUST_REWORK', '" & lv_cust_rework & "', '1') "
'If AddSql(strSql) = 0 Then
'    MsgBox "重工非重工插入失败", vbInformation, "提示"
'    Exit Sub
'End If
'
''3.日期
'lv_customer = Get_OracleStr("select b.customershortname from ib_waferlist a, mappingdatatest b Where a.ordername = '" & start_shop_order & "' and b.substrateid = a.waferid and rownum = 1 ")
'
'lv_wo_date_string = Get_OracleStr("select nvl(to_char(ib.erpcreatedate, 'YYYY-MM-DD'),to_char(sysdate, 'YYYY-MM-DD')) from ib_workorder ib where ib.ordername = '" & start_shop_order & "' ")
'strSql = "insert into shop_order_property(shop_order, propertyname, propertyvalue, levelid) Values('" & start_shop_order & "', 'ERP_CREATE_DATE_STRING', '" & lv_wo_date_string & "', '1')"
'AddSql (strSql)
'
'
'If lv_customer = "US026" Then
'    lv_wo_date = Get_OracleStr("select nvl(to_char(TRUNC(ib.erpcreatedate, 'D') + 5, 'YYYYWW'),to_char(TRUNC(sysdate, 'D') + 5, 'YYYYWW')) from ib_wohistory ib where ib.ordername =  '" & start_shop_order & "'")
'Else
'   lv_wo_date = Get_OracleStr("select nvl(to_char(TRUNC(ib.erpcreatedate, 'D') + 1, 'YYYYWW'),to_char(TRUNC(sysdate, 'D') + 1, 'YYYYWW')) from ib_wohistory ib where ib.ordername =  '" & start_shop_order & "'")
'End If
'
'strSql = "insert into shop_order_property(shop_order, propertyname, propertyvalue, levelid) Values('" & start_shop_order & "', 'ERP_CREATE_DATE', '" & lv_wo_date & "', '1') "
'AddSql (strSql)
'
''4.
' If lv_customer <> "AA(ON)" And lv_customer <> "AA" Then
'    strSql = "select distinct j.pkg_type || '@@' || d.status || '@@' || d.bline || '@@' || d.code || '@@' || g.msl || '@@' || h.numberofhours || '@@' || g.temp || '@@' || g.pbf_die_attach || '@@' || g.ecat || '@@' || " & _
'   " g.lead_free || '@@' || b.customershortname || '@@' || c.mpn_desc || '@@' || c.test_mtrl_desc || '@@' || k.product_12nc || '@@' || k.pmc || '@@' || k.marking_code || '@@' || replace(replace(c.mpn_desc, L.cust_device1 , l.cust_device2), '.P2', '') || '@@' || l.cust_device3 || '@@' || " & _
'   " decode(j.customerptno3,'一步清洗', 'One steps', '二步清洗', 'Two steps') || '@@' || j.customerptno4 || '@@' || j.customerptno5 || '@@' ||  replace(j.customerptno6, '康宁', 'KN') || '@@' || m.bin1_device || '@@' || m.bin2_device || '@@' || " & _
'   " n.return || '@@' || c.zx_invoice || '@@' from ib_waferlist a inner join mappingdatatest b on b.substrateid = a.waferid left join customeroitbl_test c on to_char(c.id) = b.filename left join ib_workorder e on e.ordername = a.ordername left join code37 d on d.device = c.mpn_desc " & _
'   " left join gcrev f on f.prouct = e.product left join CUSTOMERMPNAttributes g on g.part = e.mpn left join CUSTOMERMSLevelTBL h on h.ms_level = g.msl left join gcrev i on i.prouct = e.product and i.version = 'A'  left join tbltsvnpiproduct j on j.qtechptno2 = e.product and j.customerptno1 = c.mpn_desc " & _
'   " and j.customershortname = b.customershortname and j.customershortname = e.customer left join EU010_REFERENCE k on k.cust_device = c.mpn_desc left join device_label l on c.mpn_desc like l.cust_device1 || '%' left join CUST_BIN_DEVICE m on m.cust_part = c.mpn_desc left join pj_wo_pri n on n.wo = a.ordername where a.ordername = '" & start_shop_order & "'"
'
'   AddSql (strSql)
'
'End If
'
'
'
'
'
'
'
'
'
'
'
'
'
'Cnn.CommitTrans
'Exit Sub
'ERRON:
'Cnn.RollbackTrans

End Sub

Private Sub Form_Activate()

    If gUserName <> "07885" Then
        cmdDeleteOrder.Visible = False
        cmdDelete.Visible = False
    End If

End Sub

Private Sub Form_Load()
    ' ORACLE
    ComboOrderTbl.AddItem ("")
    ComboOrderTbl.AddItem ("container")
    ComboOrderTbl.AddItem ("mfgorder")
    ComboOrderTbl.AddItem ("A_Lotwafers")
    ComboOrderTbl.AddItem ("ib_wohistory")
    ComboOrderTbl.AddItem ("ib_waferlist")
    ComboOrderTbl.AddItem ("shop_order")
    ComboOrderTbl.AddItem ("shop_order_detail")
    ComboOrderTbl.AddItem ("shop_order_property")

    ' SQL_SERVER
    ComboSql.AddItem ("")
    ComboSql.AddItem ("tblTSVworkorder")
    ComboSql.AddItem ("tblTSVwaferlist")
    ComboSql.AddItem ("tblllplan")
    ComboSql.AddItem ("TblERPFLToME")

    With fps(1)
        .ReDraw = False
        
        .MaxRows = 0
        
        '砞竚Α
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
'        .Col = -1
'        .Row = -1
'        .Lock = True
'        .OperationMode = OperationModeNormal
'        .TypeVAlign = TypeVAlignCenter
'        .SelForeColor = &HFF8080
'
'        .Col = 1
'        .CellType = CellTypeCheckBox
'        .TypeHAlign = TypeHAlignCenter
'        .TypeVAlign = TypeVAlignCenter
       
    End With

End Sub
