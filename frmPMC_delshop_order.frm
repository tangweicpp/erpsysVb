VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPMC_delshop_order 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   11595
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   20520
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11595
   ScaleWidth      =   20520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmd_quit 
      BackColor       =   &H0000C000&
      Caption         =   "退出"
      Height          =   480
      Left            =   19440
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   0
      Width           =   990
   End
   Begin VB.Frame Fra 
      Caption         =   "删除工单"
      Height          =   2655
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   17655
      Begin VB.CommandButton cmd_exporttoexcel 
         Caption         =   "导出记录"
         Height          =   645
         Left            =   12240
         TabIndex        =   24
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Textshop_order2 
         Height          =   405
         Left            =   5400
         TabIndex        =   22
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Textuser 
         Height          =   375
         Left            =   13800
         TabIndex        =   20
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Textshop_order 
         Height          =   405
         Left            =   10920
         TabIndex        =   18
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmd_delhistory 
         Caption         =   "工单删除记录"
         Height          =   645
         Left            =   10200
         TabIndex        =   15
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdMes 
         Caption         =   "mes关工单"
         Height          =   645
         Left            =   2760
         TabIndex        =   14
         Top             =   1680
         Width           =   1815
      End
      Begin VB.ComboBox Comboshop_order 
         Height          =   315
         ItemData        =   "frmPMC_delshop_order.frx":0000
         Left            =   7680
         List            =   "frmPMC_delshop_order.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtDelFrom 
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton cmdrevert 
         Caption         =   "工单还原"
         Height          =   645
         Left            =   14280
         TabIndex        =   4
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmd 
         Caption         =   "查询工单领退料"
         Height          =   645
         Left            =   6600
         TabIndex        =   3
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtOrderID 
         Height          =   405
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmdERP 
         Caption         =   "ERP删除工单"
         Height          =   645
         Left            =   360
         TabIndex        =   1
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblshop_order2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单号"
         Height          =   195
         Left            =   5400
         TabIndex        =   23
         Top             =   480
         Width           =   540
      End
      Begin VB.Label lbluser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工号"
         Height          =   195
         Left            =   13320
         TabIndex        =   21
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label lblgdh 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单号"
         Height          =   195
         Left            =   10320
         TabIndex        =   19
         Top             =   1080
         Width           =   540
      End
      Begin VB.Line Line2 
         X1              =   9960
         X2              =   9960
         Y1              =   120
         Y2              =   2640
      End
      Begin VB.Label lblquery 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询领退料"
         Height          =   195
         Left            =   5160
         TabIndex        =   13
         Top             =   120
         Width           =   1260
      End
      Begin VB.Line Line1 
         X1              =   5040
         X2              =   5040
         Y1              =   120
         Y2              =   3360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "删除需求人员"
         Height          =   195
         Left            =   2760
         TabIndex        =   11
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label lblshop_order 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单号"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lbltable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "表名"
         Height          =   195
         Left            =   7680
         TabIndex        =   9
         Top             =   480
         Width           =   540
      End
      Begin VB.Label lbluserinfor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前使用者账号信息："
         Height          =   195
         Left            =   11640
         TabIndex        =   8
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label lbluserGH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   1800
      End
      Begin VB.Label lbluserinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "使用者工号："
         Height          =   195
         Left            =   13680
         TabIndex        =   6
         Top             =   240
         Width           =   3000
      End
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   3495
      Index           =   1
      Left            =   480
      TabIndex        =   16
      Top             =   5640
      Width           =   18615
      _Version        =   524288
      _ExtentX        =   32835
      _ExtentY        =   6165
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
      SpreadDesigner  =   "frmPMC_delshop_order.frx":0037
      Appearance      =   2
      TextTip         =   2
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   2535
      Index           =   0
      Left            =   480
      TabIndex        =   17
      Top             =   3000
      Width           =   18615
      _Version        =   524288
      _ExtentX        =   32835
      _ExtentY        =   4471
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
      SpreadDesigner  =   "frmPMC_delshop_order.frx":04A7
      Appearance      =   2
      TextTip         =   2
   End
End
Attribute VB_Name = "frmPMC_delshop_order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQlQJ As String

Private Sub cmd_Click()
'成品的领退料
'原材料的领退料
'金蝶表
    Dim rs    As New ADODB.Recordset
    Dim OrderID  As String
    Dim Str_Sql  As String
    Dim OrderTbl As String

    OrderID = ""
    Str_Sql = ""
    OrderTbl = ""

    OrderID = UCase(Trim(Textshop_order2.text))
    OrderTbl = Comboshop_order.text

    If OrderID = "" Then
        MsgBox "未输入工单号"
        Exit Sub
    End If

    Select Case OrderTbl

        Case "成品的领退料"
            Str_Sql = " SELECT * FROM erpdata..tblStockMove a WHERE a.领料工单号 =  '" + OrderID + "'"
        Case "原材料的领退料"
            Str_Sql = "SELECT * FROM ERPBASE..tblStockMove a WHERE a.工单号 = '" + OrderID + "' "

        Case "金蝶表"
            Str_Sql = " SELECT * FROM AIS20141114094336..cbInQty a WHERE a.FBillNo  = '" + OrderID + "'"
        Case Else
            Str_Sql = ""
    End Select

    rs.Open Str_Sql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType(rs)
        Exit Sub
    End If

End Sub

Private Sub ListDataType(rs As ADODB.Recordset)

    Dim i As Long

    With fps(0)
        .MaxRows = 0
        Set .DataSource = rs

    End With

    With fps(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            .BackColor = &HFFFF&
            .ColWidth(1) = 10
            .CellType = CellTypeCheckBox
            .text = 0
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

Private Sub cmd_delhistory_Click()

    Dim rs    As New ADODB.Recordset
    Dim Str_Sql  As String
    Dim userGH As String
    Dim OrderID  As String
  
    OrderID = UCase(Trim(Textshop_order.text))
    userGH = Trim(Textuser.text)

    Str_Sql = "SELECT id,userGH as '工号',username as '姓名', shop_order as '工单号',create_time as '删除时间', revert_time as '还原时间'," & _
              "ischecked as '是否审核',case erpisdel when '1' then '是' else '否' end as 'ERP是否删除', case mesisdel when '1' then '是' when '2' then '当前工单未下线' else '否' end as 'MES是否删除'" & _
              "from erptemp.dbo.PMC_Del_shop_order_history where 1=1 "

    If userGH <> "" Then
        Str_Sql = Str_Sql + "and UserGH = '" & userGH & "'"
    End If

    If OrderID <> "" Then
        Str_Sql = Str_Sql + "and OrderID = '" & OrderID & "'"
    End If

     Str_Sql = Str_Sql + "ORDER BY create_time desc"
     strSQlQJ = Str_Sql
    rs.Open Str_Sql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType(rs)
        Exit Sub
    End If
End Sub

Private Function Query(SHOP_ORDER As String)

    Dim rs    As New ADODB.Recordset
    Dim Str_Sql  As String

    Str_Sql = "SELECT id,userGH as '工号',username as '姓名', shop_order as '工单号',create_time as '删除时间', revert_time as '还原时间'," & _
              "ischecked as '是否审核',case erpisdel when '1' then '是' else '否' end as 'ERP是否删除', case mesisdel when '1' then '是' when '2' then '当前工单未下线' else '否' end as 'MES是否删除'" & _
              "from erptemp.dbo.PMC_Del_shop_order_history where 1=1 "

    If SHOP_ORDER <> "" Then
        Str_Sql = Str_Sql + "and shop_order = '" & SHOP_ORDER & "'"
    End If

    Str_Sql = Str_Sql + "ORDER BY create_time desc"
    strSQlQJ = Str_Sql
    rs.Open Str_Sql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType(rs)
        Exit Function
    End If
End Function

Private Sub cmd_exporttoexcel_Click()
SqlServerExporToExcel (strSQlQJ)
End Sub

Private Sub cmd_quit_Click()
Unload Me
Frm_ProductionPlan.Show
MDIForm1.Show
End Sub
Private Sub cmdERP_Click()
Dim rs           As New ADODB.Recordset
Dim OrderID As String
Dim time As String
Dim people As String

If txtDelFrom.text = "" Then
    MsgBox "请输入需求人员姓名(部门)", vbCritical, "提醒"
    Exit Sub
End If

    OrderID = UCase(Trim(txtOrderID.text))
    If OrderID = "" Then
        MsgBox ("请输入工单号")
        Exit Sub
    End If
    strSql = "SELECT * from erptemp.dbo.PMC_Del_shop_order_history  where UserGH ='" + gUserName + "' and shop_order ='" + OrderID + "' and erpisdel = '1'"

    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2
    End If
 rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

  If rs.RecordCount > 0 Then
     people = rs.Fields("UserGH")
     time = rs.Fields("create_time")
     iRes = MsgBox("工单在ERP已有人删除 & 删除人员工号：" & people & "删除时间:" & time & " 要继续删除吗?", vbYesNoCancel, "提示:")
     If iRes <> vbYes Then
        Exit Sub
     End If
  End If
  rs.Close
                
  If DelOrder(OrderID) <> "0" Then      '删工单
     Call delshop_orderERPJL(OrderID) '记录删工单
     cmdMES.Visible = True
  End If
End Sub

Private Function DelOrder(OrderID As String) As String
    
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
    
    Dim strRecipient(10)    As String
    Dim strRecipientCC(2)  As String
    Dim strTitle As String
    Dim XSDH As String      'XSDH 想说的话

    DelOrder = "0"
    ' 加判断后再删除
    ' 0 是否退料
   Str_sql_Guard = "select SUM(实领数量) from [erpbase].[dbo].[tblllplan] where 工单号 =  '" + OrderID + "'"
    If Get_SqlserverNo(Str_sql_Guard) > 0 Then
         MsgBox ("该工单还未全部退料!禁止删除，程序退出" & "提示:")
         cmdMES.Visible = False
         DelOrder = "0"
         Exit Function
    End If
    ' 1 是否抛到金蝶
    Str_sql_Guard = "select * from erpdata..tblTSV_TLInfo a where a.工单号 = '" + OrderID + "'"

    If QuerySqlserver(Str_sql_Guard) Then
        iRes = MsgBox("工单已经抛到金蝶, 要继续删除吗?", vbYesNoCancel, "提示:")

        If iRes = vbYes Then
            strRecipient(0) = "mingming.wu_ks@ht-tech.com"
            strRecipient(1) = "hui.song_ks@ht-tech.com"
            strRecipient(2) = "yifan.zhu_ks@ht-tech.com"
            strRecipientCC(0) = "cost.fin_ks@ht-tech.com"
            strRecipientCC(1) = "jian.pan_ks@ht-tech.com"
            strRecipientCC(2) = "allen.xu_ks@ht-tech.com"

            XSDH = "删除工单时发现以下工单已抛转到金蝶，请知悉:工单为:" & SHOP_ORDER & "财务成本请确认"
            strTitle = "<删工单但工单抛转到金蝶提示:" & OrderID & ">" & "<请求人员:" & txtDelFrom.text & ">" & "<操作员:" & gUserName & ">" & "<操作员姓名:" & gUserRealName & ">"
            Call MailDetail_ZYF(strTitle, strRecipient, XSDH, strRecipientCC)
            MsgBox "已邮件告知财务成本和潘健，(如未发送邮件，请及时联系IT)-请知悉"
        Else
            DelOrder = "0"
            Exit Function
        End If
    End If

    ' 2 产品是否在机台内
    Str_sql_Guard = "select a.RESOURCENAME from historymainline a,(select max(CONTAINERTXNSEQUENCE) mm, containername from historymainline " & "where containername in ( select conn.containername from container conn where conn.mfgorderid in  (select mfg.mfgorderid from mfgorder mfg where mfg.mfgordername  in ('" + OrderID + "')) ) " & "group by containername) b where a.containername = b.containername and a.CONTAINERTXNSEQUENCE = b.mm and a.RESOURCENAME is not null order by a.RESOURCENAME"

    If QueryStr(Str_sql_Guard) Then
        MsgBox ("在机台内不能删除吗！" & "提示:")
        DelOrder = "0"
        Exit Function
    End If

    ' 3 产品是否在生产
    Str_sql_Guard = "select * from mfgorder a, a_lotwafers b, mappingdatatest c,customeroitbl_test d,ib_wohistory e,container f," & "currentstatus g,spec h,operation i,workcenter j, specbase k,container l,product m, productbase n " & "Where b.workordername = a.mfgordername and c.substrateid = b.waferscribenumber and to_char(d.id) = c.filename and e.ordername = b.workordername " & "and f.containerid = b.containerid and g.currentstatusid = f.currentstatusid and h.specid = g.specid and i.operationid = h.operationid " & "and j.workcenterid = i.workcenterid and k.specbaseid = h.specbaseid and a.mfgordername = '" + OrderID + "' and l.containerid = b.containerid " & "and l.status = 1 and m.productid = l.productid and n.productbaseid = m.productbaseid and k.specname <> '3010' "

    If QueryStr(Str_sql_Guard) Then
        MsgBox ("在生产，不能删除！" & "提示:")
        DelOrder = "0"
        Exit Function
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
    
    ' 自动发送邮件
    
    strRecipient(0) = "mingming.wu_ks@ht-tech.com"
    strRecipient(1) = "hui.song_ks@ht-tech.com"
    strRecipient(2) = "yifan.zhu_ks@ht-tech.com"
    strRecipient(3) = "shuang.chen_ks@ht-tech.com"
    strRecipient(4) = "yifan.zhu_ks@ht-tech.com"
    strRecipient(5) = "fengying.qin_ks@ht-tech.com"
    strRecipient(6) = "canbin.lou_ks@ht-tech.com"
    strRecipientCC(0) = "allen.xu_ks@ht-tech.com"

    XSDH = "ERP工单删除通知，工单号为:" & OrderID & "请知悉"
    strTitle = "<ERP工单删除:" & OrderID & ">" & "<请求人员:" & txtDelFrom.text & ">" & "<操作员工号:" & gUserName & ">" & "<操作姓名:" & gUserRealName & ">"
    
    Call MailDetail_ZYF(strTitle, strRecipient, XSDH, strRecipientCC)
    MsgBox ("删除成功"), vbInformation, "提示"
    DelOrder = "1"
End Function

Public Function MailDetail_ZYF(Subject As String, _
                               Recipient() As String, _
                               Attachment As String, _
                               RecipientCC() As String) As Boolean

Dim JM As Object
Set JM = CreateObject("JMAIL.Message")

    'Dim JM             As New jmail.Message

    Dim Recipients()   As String

    Dim RecipientCCs() As String

    Dim strBodyinfo    As String

    Dim i              As Integer

    Dim strSql         As String

    Dim j              As Integer

    Dim rs             As New ADODB.Recordset

    Dim RsD            As New ADODB.Recordset

'    On Error GoTo ErrHandler

    MailDetail_ZYF = False

    JM.Charset = "gb2312"
    JM.Silent = False
    JM.Priority = 1
   JM.MailServerUserName = "sqladmin" '帐号
    JM.MailServerPassWord = "ksitadmin" '密码
    JM.From = "sqladmin@ht-tech.com"    '名字
    JM.FromName = "删工单助手提醒:"  '发件人姓名
    
    '收件人
        For i = 0 To UBound(Recipient) - 1
        If Recipient(i) <> "" Then
            JM.AddRecipient Recipient(i)
        End If
        
    Next
 
    '抄送人
    For i = 0 To UBound(RecipientCC) - 1
        If RecipientCC(i) <> "" Then
            JM.AddRecipientCC RecipientCC(i)
        End If
        
    Next
'    JM.AddRecipientCC "hui.song_ks@ht-tech.com;mingming.wu_ks@ht-tech.com;allen.xu_ks@ht-tech.com"
'    JM.AddRecipientCC "mingming.wu_ks@ht-tech.com"
'    JM.AddRecipientCC "allen.xu_ks@ht-tech.com"
'     JM.AddRecipientCC "ruijuan.huang_ks@ht-tech.com"
    '附件
    
    JM.Subject = "HTKS AutoMail  " & Subject
    strBodyinfo = Attachment

    JM.AppendText (strBodyinfo)
    
    MailDetail_ZYF = JM.Send("mail.ht-tech.com")
    
'ErrHandler:
'    Set JM = Nothing
'    Exit Function

End Function
Private Sub cmdMes_Click()
Dim url As String
Dim strsite As String
Dim strsitelist As String
Dim i As Integer
Dim strSql As String
Dim iRes As String
Dim rs           As New ADODB.Recordset
Dim OrderID As String
Dim SCXQRY As String
Dim people As String
Dim time As String
Dim strTitle As String
Dim strRecipient(10) As String
Dim strRecipientCC(2) As String
Dim XSDH As String

    OrderID = UCase(Trim(txtOrderID.text))
    SCXQRY = UCase(Trim(txtDelFrom.text))
    If OrderID = "" Or txtDelFrom = "" Then
        Exit Sub
    End If
    
    strSql = "SELECT * from erptemp.dbo.PMC_Del_shop_order_history  where shop_order ='" + OrderID + "' and mesisdel = '1'"

    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2
    End If
 rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

  If rs.RecordCount > 0 Then
     people = rs.Fields("UserGH")
     time = rs.Fields("create_time")
     iRes = MsgBox("工单在MES已有人删除 & 删除人员工号：" & people & "删除时间:" & time & " 要继续删除吗?", vbYesNoCancel, "提示:")
     If iRes <> vbYes Then
        Exit Sub
     End If
  End If
 rs.Close
 
 url = "http://10.160.2.30:8080/psb.web/api/v1/shopOrders/" & OrderID & "/_close"

    If CheckProc(url, OrderID) = "OK" Then
        MsgBox "工单关闭成功！"
        delshop_orderMESJL (OrderID) '先记录

        strRecipient(0) = "mingming.wu_ks@ht-tech.com"  '发送邮件
        strRecipient(1) = "hui.song_ks@ht-tech.com"
        strRecipient(2) = "yifan.zhu_ks@ht-tech.com"
        strRecipient(3) = "shuang.chen_ks@ht-tech.com"
        strRecipient(4) = "yifan.zhu_ks@ht-tech.com"
        strRecipient(5) = "fengying.qin_ks@ht-tech.com"
        strRecipient(6) = "canbin.lou_ks@ht-tech.com"
        strRecipientCC(0) = "allen.xu_ks@ht-tech.com"
        strRecipientCC(1) = "wei.chen_ks@ht-tech.com"
        strTitle = "<MES工单关闭:" & OrderID & ">" & "<请求人员:" & txtDelFrom.text & ">" & "<操作员工号:" & gUserName & ">" & "<操作姓名:" & gUserRealName & ">"
        XSDH = "MES已关闭工单,工单号为:" & OrderID & ""
        Call MailDetail_ZYF(strTitle, strRecipient(), XSDH, strRecipientCC())
    ElseIf CheckProc(url, OrderID) = "404" Then
        MsgBox "工单未下线,无需删除"
        strSql = "update erptemp.dbo.PMC_Del_shop_order_history set mesisdel = '2' where UserGH ='" + gUserName + "' and shop_order ='" + OrderID + "'"
        Exec_Sql (strSql)
        Exit Sub
    Else
        MsgBox strsite & "工单关闭失败，请联系IT！", vbInformation, "提示"
    End If

End Sub

Private Function CheckProc(url As String, OrderID As String) As String
Dim xmlHttp As Object
Dim XMLDoc As Object
Dim shop_order_return_flag1 As String
Dim shop_order_return_flag2 As String

Dim Result As String
Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
CheckProc = "error"
xmlHttp.Open "GET", url, True
xmlHttp.Send (Null)
While xmlHttp.readyState <> 4
DoEvents
Wend
Result = xmlHttp.responseText

shop_order_return_flag1 = "SAP DBTech JDBC: Object is closed:"
shop_order_return_flag2 = "{""header"":{""code"":1,""message"":""工单" & OrderID & "不存在！！！""},""value"":null}"
If InStr(Result, shop_order_return_flag1) Then
    CheckProc = "OK"
ElseIf InStr(Result, shop_order_return_flag2) Then
    CheckProc = "404"
End If

End Function

Private Sub cmdrevert_Click()
If gUserName <> "07885" Then
    MsgBox "非管理员禁止使用！"
    Exit Sub
End If
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

    If txtOrderID.text = "" Then
        MsgBox "请输入工单号", vbInformation, "警告"

        Exit Sub

    End If

    OrderID = UCase(Trim(txtOrderID.text))

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
Private Function delshop_orderERPJL(SHOP_ORDER As String)
    Dim strSql       As String
    Dim rs           As New ADODB.Recordset
    Dim iRes          As Integer
   '  Dim ID As Integer
    Dim userGH         As String
    Dim Create_time        As String
    Dim revert_time         As String
    Dim time             As String
    userGH = gUserName
    Create_time = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    strSql = "SELECT * from erptemp.dbo.PMC_Del_shop_order_history  where UserGH ='" + userGH + "' and shop_order ='" + SHOP_ORDER + "'"

    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2
    End If
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
            strSql = "SELECT * from erptemp.dbo.PMC_Del_shop_order_history  where UserGH ='" + userGH + "' and shop_order ='" + SHOP_ORDER + "' and mesisdel = '1'"
            If AddSql2(strSql) = 0 Then   '重复删erp
                  MsgBox ("工单在ERP删除记录已新增 & 删除人员工号：" & userGH & "删除时间:" & Create_time & ";-提示:")
                        strSql = "INSERT INTO erptemp.dbo.PMC_Del_shop_order_history (UserGH, shop_order, create_time, revert_time, ischecked, erpisdel, mesisdel,username)" & _
                                "VALUES('" & userGH & "', '" & SHOP_ORDER & "', '" & Create_time & "', null, '0','1','0','" & gUserRealName & "')"
                        Exec_Sql (strSql)
            Else   'mes删但erp没删除
                strSql = "update erptemp.dbo.PMC_Del_shop_order_history set erpisdel = '1' where UserGH ='" + userGH + "' and shop_order ='" + SHOP_ORDER + "'"
                Exec_Sql (strSql)
           End If
    Else      'erp开始删除
        strSql = "INSERT INTO erptemp.dbo.PMC_Del_shop_order_history (UserGH, shop_order, create_time, revert_time, ischecked, erpisdel, mesisdel,username)" & _
                "VALUES('" & userGH & "', '" & SHOP_ORDER & "', '" & Create_time & "', null, '0','1','0','" & gUserRealName & "')"
        Exec_Sql (strSql)
    End If
    rs.Close
    Call Query(SHOP_ORDER)
End Function

Private Function delshop_orderMESJL(SHOP_ORDER As String)
    Dim strSql       As String
    Dim rs           As New ADODB.Recordset
    
   '  Dim ID As Integer
    Dim userGH         As String
    Dim Create_time        As String
    Dim revert_time         As String
    Dim time             As String
    Dim iRes          As Integer
    userGH = gUserName
    Create_time = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    strSql = "SELECT * from erptemp.dbo.PMC_Del_shop_order_history  where UserGH ='" + userGH + "' and shop_order ='" + SHOP_ORDER + "'"

    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2
    End If
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
            strSql = "SELECT * from erptemp.dbo.PMC_Del_shop_order_history  where UserGH ='" + userGH + "'and shop_order ='" + SHOP_ORDER + "' and erpisdel = '1'"
            If AddSql2(strSql) = 0 Then   '重复删mes
                strSql = "INSERT INTO erptemp.dbo.PMC_Del_shop_order_history (UserGH, shop_order, create_time, revert_time, ischecked, erpisdel, mesisdel,username)" & _
                         "VALUES('" & userGH & "', '" & SHOP_ORDER & "', '" & Create_time & "', null, '0','0','1','" & gUserRealName & "')"
                Exec_Sql (strSql)
            Else   'erp删除但mes没删
                strSql = "update erptemp.dbo.PMC_Del_shop_order_history set mesisdel = '1' where UserGH ='" + userGH + "' and shop_order ='" + SHOP_ORDER + "'"
                Exec_Sql (strSql)
            End If
    Else      'mes开始删除
        strSql = "INSERT INTO erptemp.dbo.PMC_Del_shop_order_history (UserGH, shop_order, create_time, revert_time, ischecked, erpisdel, mesisdel,username)" & _
                "VALUES('" & userGH & "', '" & SHOP_ORDER & "', '" & Create_time & "', null, '0','0','1','" & gUserRealName & "')"
        Exec_Sql (strSql)
    End If
    rs.Close
    Call Query(SHOP_ORDER)
End Function

Private Sub Form_Load()
lbluserinfo.Caption = "使用者工号：" + gUserName + "使用者姓名" + gUserRealName
cmdERP.Visible = True
cmdMES.Visible = True
End Sub
