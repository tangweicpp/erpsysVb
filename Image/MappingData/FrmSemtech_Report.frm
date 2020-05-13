VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmSemtech_Report 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Semtech报表查询"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
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
   ScaleHeight     =   7695
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra 
      ForeColor       =   &H000000FF&
      Height          =   7335
      Index           =   1
      Left            =   3480
      TabIndex        =   17
      Top             =   840
      Width           =   9615
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   3255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   6015
         _Version        =   524288
         _ExtentX        =   10610
         _ExtentY        =   5741
         _StockProps     =   64
         EditEnterAction =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   5
         SpreadDesigner  =   "FrmSemtech_Report.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Fra 
      Caption         =   "查询条件"
      ForeColor       =   &H00FF0000&
      Height          =   7335
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   3495
      Begin VB.ComboBox cmbCombo1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         ItemData        =   "FrmSemtech_Report.frx":04E1
         Left            =   1080
         List            =   "FrmSemtech_Report.frx":04E8
         TabIndex        =   19
         Top             =   2400
         Width           =   2355
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   0
         Left            =   1080
         TabIndex        =   15
         Top             =   1080
         Width           =   2355
      End
      Begin VB.ComboBox cmbCombo1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         ItemData        =   "FrmSemtech_Report.frx":04F7
         Left            =   1080
         List            =   "FrmSemtech_Report.frx":04F9
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2355
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   204144643
         CurrentDate     =   41387
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   204144643
         CurrentDate     =   41387
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库房名称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   2460
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发货单号"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据类型"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期起"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期末"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job      No"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   780
         Width           =   720
      End
   End
   Begin VB.Frame Fra 
      Height          =   855
      Index           =   2
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.CommandButton cmdInvRpt 
         Caption         =   "一键导出库存报表"
         Height          =   360
         Left            =   8640
         MaskColor       =   &H8000000F&
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin MSComDlg.CommonDialog Com 
         Left            =   11760
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "上传文件"
         Height          =   480
         Left            =   11160
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "导出报表"
         Height          =   360
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退 出"
         Height          =   360
         Left            =   5760
         TabIndex        =   12
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdExprot 
         Caption         =   "导出当前数据"
         Height          =   360
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "查  询"
         Height          =   360
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   990
      End
   End
End
Attribute VB_Name = "FrmSemtech_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strdjbh         As String
Dim strdjbh1         As String
Dim DirShare        As String
Dim DirFileShare    As String
Dim DirInvRpt       As String
Dim order           As String
Dim RsClone         As New ADODB.Recordset
Const C_Left = 60
Const C_Top = 120

Private Enum FpsDetail
    e_Choose = 1
    e_DJBH = 2
    e_Cust = 3
    e_YDH = 7
End Enum
'更新种子值
Private Function GetExcelName(ByVal strTitle As String) As String
Dim strSql          As String
Dim Rs              As New ADODB.Recordset
Dim strExFileName   As String
Dim strCurDate      As String
    
    If strTitle = "output Invoice" Then
        strCurDate = Format(Date, "YY-MMDD")
    Else
        strCurDate = Format(Date, "YYYYMMDD")
    End If
    strSql = "select nvl(max(para_3),0)+1 para from tblsys_parameter where sysname='TSVSYS' and kind='Semtech报表' and para_1='" & strTitle & "' and para_2='" & strCurDate & "'"
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not Rs.EOF Then
        strExFileName = strCurDate + "_" + Trim$("" & Rs!PARA)
        strSql = "Update tblsys_parameter set para_2='" & strCurDate & "',para_3='" & Trim$("" & Rs!PARA) & "' where sysname='TSVSYS' and kind='Semtech报表' and para_1='" & strTitle & "'"
        Cnn.Execute strSql
    Else
        strExFileName = strCurDate + "_1"
        strSql = "Update tblsys_parameter set para_2='" & strCurDate & "',para_3='" & Trim$("" & Rs!PARA) & "' where sysname='TSVSYS' and kind='Semtech报表' and para_1='" & strTitle & "'"
        Cnn.Execute strSql
    End If
    Rs.Close
    
    If strTitle = "output Invoice" Then
        GetExcelName = "HTKS_SEDC" & strExFileName
    Else
        GetExcelName = strTitle & "_" & strExFileName
    End If
    
End Function

Private Sub cmbCombo1_Click(Index As Integer)
'Dim strSql              As String
'Dim Rs                  As New ADODB.Recordset

    If Index = 0 Then
        If cmbCombo1(0).Text = "库存报表" Then
            '加载仓库
            cmbCombo1(1).Clear
            cmbCombo1(1).AddItem "所有"
            cmbCombo1(1).AddItem "1000"
            cmbCombo1(1).AddItem "6000"
            cmbCombo1(1).AddItem "7000"
            cmbCombo1(1).AddItem "8000"
            cmbCombo1(1).AddItem "9000"
            cmbCombo1(1).AddItem "Scrap"
            cmbCombo1(1).ListIndex = 0
        Else
            cmbCombo1(1).Clear
        End If
    End If
End Sub

Private Sub CmdExit_Click() '退出
    Unload Me
End Sub

Private Sub cmdExprot_Click()
Dim strExportName           As String

    '校验数据
    If Fps(0).MaxRows <= 0 Then
        MsgBox "没有可导出的数据！", vbInformation, "提示"
        Exit Sub
    End If
    '导出报表
    If cmbCombo1(0).ListIndex = 6 Or cmbCombo1(0).ListIndex = 7 Or cmbCombo1(0).ListIndex = 8 Then '库存报表,SMTCList,Shipped
        strExportName = Trim(Fra(1).Caption)
        If cmbCombo1(1).Text <> "" Then
            strExportName = Trim(Fra(1).Caption) + "-" + Trim(cmbCombo1(1).Text)
        End If
    Else
        strExportName = GetExcelName(Trim(Fra(1).Caption))
    End If

    If Not ExportFpspreadToExcel(Fps(0), strExportName, strExportName) Then Exit Sub
    
End Sub

Private Sub cmdInvRpt_Click()
'一键导出库存报表到指定文件夹中
Dim strSql                  As String
Dim strSqlDetail            As String
Dim Rs                      As New ADODB.Recordset
Dim i                       As Integer
Dim strFileName             As String
Dim strMsg                  As String
    
    If MsgBox("确定要导出吗?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
        Exit Sub
    End If
    '导出的是库存
    strSql = "SELECT RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Reel_Size,Remark" & _
             ",Move_in_Date FROM Vw_InvStockRptFor37 Where 2>1 "
    For i = 0 To cmbCombo1(1).ListCount
        If cmbCombo1(1).List(i) <> "所有" And cmbCombo1(1).List(i) <> "" Then
            If cmbCombo1(1).List(i) = "9000" Then
               
'            strSql = "SELECT b.fab_conv_id as [Wafer Type],b.mpn_desc as [Assy Part#],a.工单号 as [Fab Lot], " & _
'            " right(replace(rtrim(a.流程卡编号),'+',''),2) ID,b.date_code as [D/C],b.die as [QTY(称重后的数量)]," & _
'            " b.test_mtrl_desc as [Job#],b.bag as Bag#,a.箱号 as Comment,b.alternatename as [HT Part#]" & _
'            " FROM erpdata..tblStockNumsub a," & _
'            " (SELECT  substring(cast(datepart(year,ora.create_date)as nvarchar(20))+substring(CAST ('100'+datepart(week,ora.create_date) as nvarchar(20)),2,2),3,4) as date_code," & _
'            " ora.waferid,ora.firstname ,ora.test_mtrl_desc,ora.mpn_desc,ora.bag,ora.fab_conv_id,ora.die,ora.alternatename FROM " & _
'            " OPENQUERY(ORACLEDB, 'SELECT d.fab_conv_id,a.waferid,a.die," & _
'            " a.create_date+6 as create_date,b.firstname ,d.test_mtrl_desc,d.mpn_desc,get_37bagid(b.containername) bag,e.alternatename" & _
'            " FROM weight37 a,container b ,mappingdatatest c,customeroitbl_test d,product e" & _
'            " where a.waferid||''-A'' = b.containername  and a.waferid = c.substrateid " & _
'            " and e.productid=b.productid and c.filename = d.id ' ) ora ) b where a.库房编号 IN('44','45')" & _
'            " and a.流程卡编号=b.waferid "
            strSql = " SELECT * FROM DBO.Vw_InvStockRptFor37By9000 "
            strSqlDetail = ""
            Else
              strSql = "SELECT RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Reel_Size,Remark" & _
              ",Move_in_Date FROM Vw_InvStockRptFor37 Where 2>1 "
              strSqlDetail = " And 仓库名称='" & cmbCombo1(1).List(i) & "'"
            End If
            If Rs.State = adStateOpen Then Rs.Close
            Rs.open strSql + strSqlDetail, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
            If Not Rs.EOF Then '表示有数据才导出报表
                strFileName = Format(Now(), "YYMMDD_HHMM") + "_" + Replace(cmbCombo1(1).List(i), "Scrap", "SCR")
                strMsg = strMsg + DirInvRpt + "\" + strFileName + vbCrLf                '提示消息
                RsExporToExcel Rs, cmbCombo1(1).List(i), strFileName                    '导出到Excel
            End If
            Rs.Close
        End If
    Next
    '导出Shipped
    strSql = "SELECT RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Remark,SO as [DN#],SHIPPED_DATE,Cust_Name " & _
             " FROM Vw_InvShippedRptFor37 "
'             " WHERE SHIPPED_DATE>='" & DateAdd("m", -1, Format(Now(), "YYYY-MM-DD")) & "' and SHIPPED_DATE<'" & DateAdd("d", 1, Format(Now(), "YYYY-MM-DD")) & "' "
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If Not Rs.EOF Then '表示有数据才导出报表
        strFileName = Format(Now(), "YYMMDD_HHMM") + "_SHIPPED"
        strMsg = strMsg + DirInvRpt + "\" + strFileName + vbCrLf                '提示消息
        RsExporToExcel Rs, "SHIPPED", strFileName                               '导出到Excel
    End If
    Rs.Close
    
    MsgBox "导出成功，导出文件路径为：" + vbCrLf + strMsg
    
End Sub

Private Sub cmdReport_Click() '导出报表
Dim strExportName           As String

    If Fps(0).MaxRows <= 0 Then
        MsgBox "没有可导出的数据！", vbInformation, "提示"
        Exit Sub
    End If
    '导出报表
    strExportName = GetExcelName(Trim(Fra(1).Caption))
    
    If cmbCombo1(0).ListIndex = 0 Then
        Call SEDCExportPrintExcel(RsClone, strExportName)                     'SEDC报表
    ElseIf cmbCombo1(0).ListIndex = 1 Then
        Call InputPackinglistExportPrintExcel(RsClone, strExportName)         'outputPackinglist报表
    ElseIf cmbCombo1(0).ListIndex = 2 Then
        Call InputInvoiceExportPrintExcel(RsClone, strExportName)             'outputInvoice报表
    ElseIf cmbCombo1(0).ListIndex = 3 Then
        Call Daily_InvExportPrintExcel(RsClone, strExportName)                'Daily_inventory_report
    ElseIf cmbCombo1(0).ListIndex = 4 Then
        If Not CheckData Then Exit Sub
        Call ShippingPackinglistExportPrintExcel(order, strExportName)      'Shipping Packinglist
    ElseIf cmbCombo1(0).ListIndex = 5 Then
        If Not CheckData Then Exit Sub
        Call ShippingInvoiceExportPrintExcel(order, strExportName)          'Shipping invoice
    End If
End Sub

Private Sub cmdSearch_Click() '查询报表
Dim i                   As Long
Dim strSql              As String
Dim Rs                  As New ADODB.Recordset

    '初始化FPS
    InitFps
    '---------------------------------------------
    If cmbCombo1(0).ListIndex = 0 Then  'SEDC
    
              strSql = "select distinct 0 as 选择,a.containername,a.creationtimestamp as receive_date,to_char(ibwo.erpcreatedate,'yyyyww') testdc,'' as invlocation " & _
                ",wo.mpn_desc as device_name,wo.SOURCE_BATCH_ID as job_no,conn.firstname as lot_no,a.moveinqty as qty,wo.date_code,'' as invcomment " & _
                ",'new packing' as remark,'' as so,'' as invoice,Get_37Inv_MergeDetails(a.containername,wo.mtrl_num) as merge " & _
                "from a_wiplothistory a ,a_wiplotdetailshistory b ,container conn,mfgorder mfg,ib_wohistory ibwo,a_lotwafers waf " & _
                ",customeroitbl_test wo " & _
                "where a.specname='5272' and b.wiplothistoryid=a.wiplothistoryid " & _
                "and conn.containername=a.containername " & _
                "and waf.waferscribenumber=substr(conn.containername,1,instr(conn.containername,'-A')-1) " & _
                "and mfg.mfgordername=waf.workordername and ibwo.ordername=mfg.mfgordername " & _
                "and wo.source_batch_id=waf.wafernumber " & _
                " and wo.customershortname='37' and a.containername not like '%-F%' " & _
                "And a.creationtimestamp>='" & DTP(0).Value & "' And a.creationtimestamp<'" & DTP(1).Value + 1 & "'"
                
        If txt(1).Text <> "" Then
            strSql = strSql & " And wo.SOURCE_BATCH_ID='" & Trim(txt(1).Text) & "'"
        End If
        
    ElseIf cmbCombo1(0).ListIndex = 1 Then  'output packing list
     
        strSql = " select 选择,creationtimestamp outdate,po_num,po_item,PKG,DEVICE,lot_no,Job_No,Sublot_No,date_code,sum( qty) as qty,sum(Price) as Price,sum(USD) as USD,Test_Reject,Merge_in_Job,RMA_Number,mtrl_num from (" & _
                " select distinct 0 as 选择,a.containername, wo.po_num, wo.po_item,'' PKG,wo.mpn_desc DEVICE,wo.MTRL_NUM lot_no,waf.wafernumber Job_No," & _
                " conn.firstname||Get_37Inv_MergeStatus(a.containername) Sublot_No, wo.date_code,a.moveinqty qty, 0 as Price,a.moveinqty * 0 as USD," & _
                "  '' as Test_Reject, Get_37Inv_MergeDetails(a.containername,wo.MTRL_NUM) as Merge_in_Job,'' as RMA_Number, wo.mtrl_num,a.creationtimestamp " & _
                "  from a_wiplothistory a, a_wiplotdetailshistory b,container  conn,mfgorder  mfg,ib_wohistory  ibwo,a_lotwafers waf, customeroitbl_test wo " & _
                " where a.specname = '5272' and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername " & _
                "  and waf.waferscribenumber = substr(conn.containername, 1, instr(conn.containername, '-A') - 1) " & _
                "   and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername and wo.source_batch_id = waf.wafernumber and status<>2 " & _
                "   and wo.customershortname = '37' and a.containername not like '%-F%' And a.creationtimestamp >= '" & DTP(0).Value & "'" & _
                "   And a.creationtimestamp < '" & DTP(1).Value & "' and   a.containername<>'10001-A-01' and not  exists (select 1 from cus37_5272qtymergV2 mer where mer.containername=a.containername)  " & _
                " union select distinct 0 as 选择, a.containername, wo.po_num,wo.po_item,'' PKG, wo.mpn_desc DEVICE, wo.MTRL_NUM lot_no, waf.wafernumber as job_no, " & _
                " conn.firstname||Get_37Inv_MergeStatus(a.containername) Sublot_No,wo.date_code, waf.ndpw qty,0 as Price, a.moveinqty * 0 as USD," & _
                " '' as Test_Reject, Get_37Inv_MergeDetails(a.containername,wo.MTRL_NUM) as Merge_in_Job,'' as RMA_Number , wo.mtrl_num,a.creationtimestamp" & _
                "  from a_wiplothistory  a, a_wiplotdetailshistory b, container conn,mfgorder mfg, ib_wohistory ibwo, a_lotwafers  waf, customeroitbl_test  wo " & _
                " where a.specname = '5272' and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername " & _
                "   and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername and wo.source_batch_id = waf.wafernumber " & _
                "   and wo.customershortname = '37' and a.containername not like '%-F%' And a.creationtimestamp >= '" & DTP(0).Value & "' " & _
                "   And a.creationtimestamp < '" & DTP(1).Value & "' and waf.containerid=conn.containerid  and  exists (select 1 from cus37_5272qtymergV2 mer where mer.containername = a.containername)  " & _
                " )X Where 2>1 "
                
        If txt(1).Text <> "" Then
            strSql = strSql & " And Job_No='" & Trim(txt(1).Text) & "'"
        End If
        strSql = strSql & " group by 选择,po_num,po_item,PKG,DEVICE,lot_no,Job_No,Sublot_No,date_code,Test_Reject,Merge_in_Job,RMA_Number,mtrl_num,creationtimestamp order by Sublot_No,Merge_in_Job desc"
    
    ElseIf cmbCombo1(0).ListIndex = 2 Then  'output invoice
      
      strSql = " select 选择,creationtimestamp outdate,po_num,po_item, PKG, DEVICE,lot_no,job_no,Sublot_No,date_code, sum(qty), Price,sum(USD)  from (" & _
                "select distinct 0 as 选择,a.containername,wo.po_num,wo.po_item,'' PKG,wo.mpn_desc DEVICE,wo.MTRL_NUM lot_no,wo.SOURCE_BATCH_ID job_no " & _
                ",conn.firstname||Get_37Inv_MergeStatus(a.containername) Sublot_No,wo.date_code,a.moveinqty qty,wo.t_price as Price,a.moveinqty*wo.t_price as USD,a.creationtimestamp " & _
                "from a_wiplothistory a ,a_wiplotdetailshistory b ,container conn,mfgorder mfg,ib_wohistory ibwo,a_lotwafers waf " & _
                ",customeroitbl_test wo " & _
                "where a.specname='5272' and b.wiplothistoryid=a.wiplothistoryid " & _
                "and conn.containername=a.containername " & _
                "and waf.waferscribenumber=substr(conn.containername,1,instr(conn.containername,'-A')-1) " & _
                "and mfg.mfgordername=waf.workordername and ibwo.ordername=mfg.mfgordername " & _
                "and wo.source_batch_id=waf.wafernumber " & _
                " and wo.customershortname='37' and a.containername not like '%-F%' " & _
                "And a.creationtimestamp>='" & DTP(0).Value & "' And a.creationtimestamp<'" & DTP(1).Value & "'" & _
                "and a.containername<>'10001-A-01'and not  exists (select 1 from cus37_5272qtymergV2 mer where mer.containername=a.containername)   union " & _
                " select  distinct 0 as 选择,a.containername,wo.po_num,wo.po_item,'' PKG,wo.mpn_desc DEVICE,wo.MTRL_NUM lot_no,waf.wafernumber job_no, " & _
                " conn.firstname||Get_37Inv_MergeStatus(a.containername) Sublot_No,wo.date_code,waf.ndpw qty,wo.t_price as Price,waf.ndpw* wo.t_price as USD,a.creationtimestamp " & _
                " from a_wiplothistory a,a_wiplotdetailshistory b,container conn,mfgorder mfg,ib_wohistory ibwo,a_lotwafers waf,customeroitbl_test wo " & _
                " where a.specname = '5272' and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername" & _
                " and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername " & _
                " and wo.source_batch_id = waf.wafernumber and wo.customershortname = '37' and a.containername not like '%-F%' " & _
                " And a.creationtimestamp >= '" & DTP(0).Value & "' And a.creationtimestamp < '" & DTP(1).Value & "' and waf.containerid=conn.containerid   and  exists (select 1 from cus37_5272qtymergV2 mer where mer.containername = a.containername) " & _
                " )X Where 2>1 "
                
        If txt(1).Text <> "" Then
            strSql = strSql & " And job_no='" & Trim(txt(1).Text) & "'"
        End If
        strSql = strSql & " group by 选择,po_num,po_item, PKG, DEVICE, lot_no, job_no ,Sublot_No,date_code, Price,creationtimestamp order by Sublot_No"
    
    ElseIf cmbCombo1(0).ListIndex = 3 Then  ' Daily Inv
             
        strSql = " select 选择,max(RECEIVE_DATE) as RECEIVE_DATE,TESTDC,LOCATION,DEVICE_NAME,job_no,HTLOT_NO,sum(qty) as qty,date_code,CCOMMENT,Reel_Size,Remark,Move_in_Date from (" & _
                "select distinct 0 as 选择,a.containername,a.creationtimestamp as RECEIVE_DATE,to_char(ibwo.erpcreatedate,'yyww') TESTDC,'' as LOCATION " & _
                ",wo.mpn_desc DEVICE_NAME,wo.SOURCE_BATCH_ID||Get_37Inv_MergeStatus(a.containername) job_no " & _
                ",conn.firstname||Get_37Inv_MergeStatus(a.containername) HTLOT_NO,a.moveinqty qty,wo.date_code,'' CCOMMENT,'NEW PACKING' Reel_Size,'' Remark,'' Move_in_Date " & _
                "from a_wiplothistory a ,a_wiplotdetailshistory b ,container conn,mfgorder mfg,ib_wohistory ibwo,a_lotwafers waf " & _
                ",customeroitbl_test wo " & _
                "where a.specname='5272' and b.wiplothistoryid=a.wiplothistoryid " & _
                "and conn.containername=a.containername " & _
                "and waf.waferscribenumber=substr(conn.containername,1,instr(conn.containername,'-A')-1) " & _
                "and mfg.mfgordername=waf.workordername and ibwo.ordername=mfg.mfgordername " & _
                "and wo.source_batch_id=waf.wafernumber " & _
                " and wo.customershortname='37' and a.containername not like '%-F%'  and status<>'2' " & _
                "And a.creationtimestamp>='" & DTP(0).Value & "' And a.creationtimestamp<'" & DTP(1).Value + 1 & "'" & _
                " ) X Where 2>1 "
                
        If txt(1).Text <> "" Then
            strSql = strSql & " And job_no='" & Trim(txt(1).Text) & "'"
        End If
        strSql = strSql & " group by 选择,TESTDC,LOCATION,DEVICE_NAME,job_no,HTLOT_NO,date_code,CCOMMENT,Reel_Size,Remark,Move_in_Date"
        
    ElseIf cmbCombo1(0).ListIndex = 4 Then  'shipping packing list (INNER JOIN tblCustomerOI d ON CASE WHEN charindex('M',RTRIM(c.batchnumber))>0 THEN LEFT(RTRIM(c.batchnumber),LEN(RTRIM(c.batchnumber))-1) ELSE RTRIM(c.batchnumber) END=d.SOURCE_BATCH_ID)
        strSql = "SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & _
                 "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 ",箱号,料号,mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,净重,毛重,尺寸 " & _
                 " FROM Vw_InvShippedPLFor37 a " & _
                 " WHERE 发货日期>='" & Format(DTP(0).Value, "YYYY-MM-DD") & "' and 发货日期<'" & Format(DTP(1).Value + 1, "YYYY-MM-DD") & "' "
        If txt(0).Text <> "" Then
            strSql = strSql & " And 单据编号='" & Trim(txt(0).Text) & "'"
        End If
        If txt(1).Text <> "" Then
            strSql = strSql & " And batchnumber='" & Trim(txt(1).Text) & "'"
        End If
    ElseIf cmbCombo1(0).ListIndex = 5 Then  'shipping invoice (INNER JOIN tblCustomerOI d ON CASE WHEN charindex('M',RTRIM(c.batchnumber))>0 THEN LEFT(RTRIM(c.batchnumber),LEN(RTRIM(c.batchnumber))-1) ELSE RTRIM(c.batchnumber) END=d.SOURCE_BATCH_ID)
        strSql = "SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3" & _
                 ",city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 ",箱号,料号,mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,单价,AMount,customerPartNumber " & _
                 " FROM Vw_InvShippedInvoiceFor37 a " & _
                 " WHERE 发货日期>='" & Format(DTP(0).Value, "YYYY-MM-DD") & "' and 发货日期<'" & Format(DTP(1).Value + 1, "YYYY-MM-DD") & "' "
        If txt(0).Text <> "" Then
            strSql = strSql & " And 单据编号='" & Trim(txt(0).Text) & "'"
        End If
        If txt(1).Text <> "" Then
            strSql = strSql & " And batchnumber='" & Trim(txt(1).Text) & "'"
        End If
    ElseIf cmbCombo1(0).ListIndex = 6 Then  '库存报表
        If cmbCombo1(1).Text = "9000" Then
            strSql = " SELECT * FROM DBO.Vw_InvStockRptFor37By9000 "
        Else
            strSql = "SELECT 0 选择,RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Reel_Size,Remark" & _
                     ",Move_in_Date FROM Vw_InvStockRptFor37 Where 2>1 "
            If txt(1).Text <> "" Then
                strSql = strSql & " And JOB_NO='" & Trim(txt(1).Text) & "'"
            End If
            If cmbCombo1(1).Text <> "所有" Then
                strSql = strSql & " And 仓库名称='" & Trim(cmbCombo1(1).Text) & "'"
            End If
        End If
    ElseIf cmbCombo1(0).ListIndex = 7 Then  'SMTCList
        strSql = "SELECT 0 选择,Invoice_No,Carton_No,PartName,LotID,QTY,Job_No,DATE_CODE " & _
                 " FROM Vw_InvShippedSMTCListFor37 " & _
                 " WHERE 发货日期>='" & DTP(0).Value & "' and 发货日期<'" & DTP(1).Value + 1 & "' "
        If txt(0).Text <> "" Then
            strSql = strSql & " And 单据编号='" & Trim(txt(0).Text) & "'"
        End If
        If txt(1).Text <> "" Then
            strSql = strSql & " And Job_No='" & Trim(txt(1).Text) & "'"
        End If
    ElseIf cmbCombo1(0).ListIndex = 8 Then  'Shipped
        strSql = "SELECT 0 选择,RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Remark,SO as [DN#],SHIPPED_DATE,Cust_Name " & _
                 " FROM Vw_InvShippedRptFor37 " & _
                 " WHERE SHIPPED_DATE>='" & DTP(0).Value & "' and SHIPPED_DATE<'" & DTP(1).Value + 1 & "' "
        If txt(0).Text <> "" Then
            strSql = strSql & " And 单据编号='" & Trim(txt(0).Text) & "'"
        End If
        If txt(1).Text <> "" Then
            strSql = strSql & " And JOB_NO='" & Trim(txt(1).Text) & "'"
        End If
    End If
    '赋值到FRA(1)中 INIadoCon
    Fra(1).Caption = cmbCombo1(0).Text
    If Rs.State = adStateOpen Then Rs.Close
    If cmbCombo1(0).ListIndex = 0 Or cmbCombo1(0).ListIndex = 1 Or cmbCombo1(0).ListIndex = 2 Or cmbCombo1(0).ListIndex = 3 Then
        Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    Else
        Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    End If
    Fps(0).MaxRows = 0
    Set RsClone = Nothing
    If Not Rs.EOF Then
        Set RsClone = Rs.Clone '克隆一份数据到另一个数据集中，为后面使用
        With Fps(0)
            .MaxRows = 0
            Set .DataSource = Rs
            .MaxRows = Rs.RecordCount
        End With
    End If
    Rs.Close
    '特殊几个报表增加汇总栏位
    CalcTotal
    
End Sub
Private Sub CalcTotal()
'计算汇总
Dim i                   As Long
Dim dblTotal            As Double
Dim colTotal            As Integer
    
    If cmbCombo1(0).ListIndex <> 6 And cmbCombo1(0).ListIndex <> 7 And cmbCombo1(0).ListIndex <> 8 Then Exit Sub
    
    dblTotal = 0
    colTotal = 0
    With Fps(0)
        If .MaxRows <= 0 Then Exit Sub
        For i = 1 To .MaxRows
            .Row = i
            If cmbCombo1(0).ListIndex = 6 Or cmbCombo1(0).ListIndex = 8 Then
                colTotal = 8
                .Col = colTotal
            Else
                colTotal = 6
                .Col = colTotal
            End If
            dblTotal = dblTotal + Val(Trim$(.Text))
        Next
        If dblTotal > 0 Then '表示有数量
            .MaxRows = .MaxRows + 1
            .SetText colTotal, .MaxRows, dblTotal
        End If
        .DeleteCols FpsDetail.e_Choose, 1
        .MaxCols = .MaxCols - 1
    End With
    
End Sub

Private Sub cmdUpload_Click()
Dim strFilePath         As String
Dim strFileName         As String
Dim strSql              As String
Dim image_Data()        As Byte         '图片二进制
Dim Rs                  As New ADODB.Recordset
    '打开图片
    Com.Filter = "上传文件(*.xls,*.xlsx)|*.xls;*.xlsx"
    Com.ShowOpen '打开对话框
    strFilePath = Trim(Com.FileName)  '保存路径
    strFileName = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1) '文件名
    '开始保存到资料库
    '数据转换为流
    Open strFilePath For Binary As #1
    ReDim image_Data(LOF(1) - 1)
    Get #1, , image_Data()
    Close #1
    '查询是否保存过此图片
    strSql = "Select * From TblPMC_PicInfo Where FileName='" & Trim$(strFileName) & "' For Update"
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, Cnn, adOpenKeyset, adLockOptimistic
    If Not Rs.EOF Then
        Rs("FileName") = strFileName
        Rs("FilePath") = strFilePath
        Rs("FileComent") = image_Data()
        Rs("Flag") = 1
        Rs.Update
    Else
        Rs.AddNew
        Rs("FileName") = strFileName
        Rs("FilePath") = strFilePath
        Rs("FileComent") = image_Data()
        Rs("Flag") = 1
        '记得添加数据库中txt存放的路径和txt框
        Rs.Update
    End If
    Rs.Close
    
    MsgBox "上传成功", vbInformation, "提示"
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Fra(2).Move C_Left, Fra(2).Top, Me.ScaleWidth - C_Left, Fra(2).Height
    Fra(0).Move C_Left, Fra(0).Top, Fra(0).Width, Me.ScaleHeight - Fra(2).Height
    Fra(1).Move Fra(1).Left, Fra(1).Top, Me.ScaleWidth - Fra(0).Width - C_Top, Me.ScaleHeight - Fra(2).Height
    Fps(0).Move C_Left, Fps(0).Top, Fra(1).Width - C_Top, Me.ScaleHeight - Fra(2).Height - 3 * C_Top
End Sub
Private Sub Form_Load()
    If gUserName = "07885" Then
        cmdUpload.Visible = True
    End If
    DirShare = App.Path & "\NewSemtechReport"               '系统路径
    DirFileShare = App.Path & "\SemtechExcelReport"         '系统Excel文件路径
    DirInvRpt = "C:\37-InventoryRpt"                        '库存报表存放路径
    '判定文件夹是否存在,不存在就创建
    If Dir(DirShare, vbDirectory) = "" Then
        MkDir DirShare                                      '创建文件夹
    End If
    If Dir(DirFileShare, vbDirectory) = "" Then
        MkDir DirFileShare                                  '创建文件夹
    End If
    If Dir(DirInvRpt, vbDirectory) = "" Then
        MkDir DirInvRpt                                     '创建文件夹
    End If
    '初始化控件
    InitCtrl
    
End Sub
Private Sub CheckXls()
Dim strSql              As String
Dim Rs                  As New ADODB.Recordset
Dim image_filename      As String
Dim temp_image()        As Byte
Dim i                   As Integer
    '设定鼠标状态
    Screen.MousePointer = 0
    strSql = "Select * From TblPMC_PicInfo Where Flag=1 Order by Create_Date"
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not Rs.EOF Then
        For i = 1 To Rs.RecordCount
            '加载图片
            temp_image = Rs("FileComent")
            image_filename = DirShare & "\" & Rs("FileName")
            Open image_filename For Binary As #1
            Put #1, , temp_image()
            Close #1
            Rs.MoveNext
        Next
    End If
    Rs.Close

End Sub
'初始化控件
Private Sub InitCtrl()
Dim i                   As Integer
Dim strSql              As String
Dim Rs                  As New ADODB.Recordset
    
    strdjbh = ""
    '加载单据类型
    strSql = "select para_1 from tblsys_parameter where sysname='TSVSYS' and kind='Semtech报表' order by id "
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    cmbCombo1(0).Clear
    If Not Rs.EOF Then
        Do While Not Rs.EOF
            cmbCombo1(0).AddItem Trim$("" & Rs!para_1)
            Rs.MoveNext
        Loop
        cmbCombo1(0).ListIndex = 0
    End If
    Rs.Close
'    '加载仓库
'    cmbCombo1(1).Clear
'    cmbCombo1(1).AddItem "所有"
'    cmbCombo1(1).AddItem "1000"
'    cmbCombo1(1).AddItem "6000"
'    cmbCombo1(1).AddItem "7000"
'    cmbCombo1(1).AddItem "8000"
'    cmbCombo1(1).AddItem "Scrap"
'    cmbCombo1(1).ListIndex = 0
    '初始化FPS
    InitFps
    
   DTP(0).Value = Format(Now() - 1, "YYYY-MM-DD")
   DTP(1).Value = Format(Now(), "YYYY-MM-DD")
   '检查模版
   CheckXls
   
End Sub
'初始化FPS控件
Public Sub InitFps()
Dim i                   As Integer
    'Fps初始化
    With Fps(0)
        .ReDraw = False
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsBest
        .MaxRows = 0
        .ColsFrozen = 2
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        '设定列类型
        .Col = FpsDetail.e_Choose   '选择
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        '设定列宽
        .ColWidth(-1) = 10
        .ColWidth(FpsDetail.e_Choose) = 4
        .RowHeight(-1) = 10
        '设定是否排序
        .UserColAction = UserColActionSort
        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
        .ZOrder
        .ReDraw = True
    End With
End Sub
Private Sub Fps_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i           As Long
Dim j           As Integer
Dim strTmp      As String
    
    '几个报表特殊处理
    If cmbCombo1(0).ListIndex = 6 Or cmbCombo1(0).ListIndex = 7 Or cmbCombo1(0).ListIndex = 8 Then Exit Sub
    '点击把选择的单号都选上
    If Row < 1 Then Exit Sub
    If Col <> 1 Then Exit Sub
    With Fps(0)
'        .Col = FpsDetail.e_Choose
'        For i = 1 To .MaxRows
'            .Row = i
'            If i <> Row Then
'                .Col = FpsDetail.e_Choose
'                If Val(.Value) = 1 Then
''                    .Value = 0
'                    .Col = -1
'                    .ForeColor = vbBlack
'                End If
'            End If
'        Next

        .Col = FpsDetail.e_Choose
        .Row = Row
        .Value = Abs(Val(.Value) - 1)
'        strDJBH = ""
        If Val(.Value) = 1 Then
            '将所有一样的单号选择上
            .Col = FpsDetail.e_DJBH
            .Row = Row
            strTmp = Trim$(.Text)
'            strDJBH = Trim$(.Text) '共用的单据编号，在导出打印时会用到
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsDetail.e_DJBH
                If Trim$(.Text) = strTmp Then
                    .Col = FpsDetail.e_Choose
                    .Value = 1
                    .Col = -1
                    .ForeColor = &HFF8080
                End If
            Next
            
            order = strTmp & "'" & "," & "'" & order
            
        Else
            '将所有一样的单号选择上
            .Col = FpsDetail.e_DJBH
            .Row = Row
            strTmp = Trim$(.Text)
'            strDJBH = Trim$(.Text) '共用的单据编号，在导出打印时会用到
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsDetail.e_DJBH
                If Trim$(.Text) = strTmp Then
                    .Col = FpsDetail.e_Choose
                    .Value = 0
                    .Col = -1
                    .ForeColor = vbBlack
                End If
            Next
        End If
        
    End With
    
End Sub

'校验数据
Private Function CheckData() As Boolean
Dim i               As Integer
Dim intCount        As Integer
Dim strCust         As String

    CheckData = False
    
    strdjbh = ""     '--单据编号记录
    strCust = ""
    intCount = 0
    
    With Fps(0)
        If .MaxRows <= 0 Then
            MsgBox "没有任何资料,请先查询！", vbInformation, "提示"
            Exit Function
        End If
        '看是否有选择
        For i = 1 To .MaxRows
            .Row = i
            .Col = FpsDetail.e_Choose  '选择
            If .Value = 1 Then
                intCount = intCount + 1
                .Col = FpsDetail.e_DJBH '单据编号
                If InStr(strdjbh, Trim$(.Text)) <= 0 Then
                    strdjbh = strdjbh + Trim$(.Text) + ","
                    strdjbh1 = Mid(strdjbh, 2, Len(strdjbh)) + Trim$(.Text) + ","
                End If
            End If
        Next
    End With
    '去除单据编号最后一个逗号

    '--------------------------
    If intCount <= 0 Then
        MsgBox "没有选择任何资料！", vbInformation, "提示"
        Exit Function
    End If
    strdjbh = Left$(strdjbh, Len(strdjbh) - 1)
    strdjbh1 = Left$(strdjbh1, Len(strdjbh1) - 1)
    CheckData = True
End Function
'SEDC
Public Sub SEDCExportPrintExcel(ByVal Rs As ADODB.Recordset, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double '总金额
Dim strExtsion          As String '后缀名
Dim strNewFullPath      As String '新Excel文件

    
    If Rs.RecordCount <= 0 Then
        MsgBox "没有要导出的资料！", vbInformation, "提示！"
        Exit Sub
    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "初始化数据..."
    
    strFileName = DirShare & "\SEDC.xls" '要打开的文件
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '获取后缀名
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径
    Rs.MoveFirst    '数据集移动到第一个
    
    If Rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '是否显示
        
        Set wkbk = ExApp.Workbooks.open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        
        lngRows = 3
        IntInertRow = Rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = Rs.RecordCount
        
'        ClsP.ShowProgress 50, "正在导出..."
        
        Do While Not Rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(2).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & Rs.fields(3).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & Rs.fields(4).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & Rs.fields(5).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(6).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & Rs.fields(7).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & Rs.fields(8).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & Rs.fields(9).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & Rs.fields(10).Value)
            wkst.Cells(lngRows, 10) = Trim$("" & Rs.fields(11).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & Rs.fields(12).Value)
            wkst.Cells(lngRows, 12) = Trim$("" & Rs.fields(13).Value)
            wkst.Cells(lngRows, 13) = Trim$("" & Rs.fields(14).Value)
            
            lngRows = lngRows + 1
            Rs.MoveNext
        Loop
        
    Else
'        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Sub
    End If
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""楷体_GB2312,常?""&10報表名稱 ：   "  ' & Gsmc
'        .CenterHeader = "&""楷体_GB2312,常規""報表&""宋体,常規""" & Chr(10) & "&""楷体_GB2312,常?""&10日 期："
'        .RightHeader = "" & Chr(10) & "&""楷体_GB2312,常規""&10單位：滬士"
'        .LeftFooter = "&""楷体_GB2312,常規""&10制表人："
'        .CenterFooter = "&""楷体_GB2312,常規""&10制表日期："
        .RightFooter = "&20" & "第 &P 页，共 &N 页"
    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.Number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Sub
            End If
        End If
    End If
    wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "导出成功！"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.Description, vbInformation, "提示！"
    Exit Sub
End Sub
'Packing list
Public Sub InputPackinglistExportPrintExcel(ByVal Rs As ADODB.Recordset, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double '总金额
Dim strExtsion          As String '后缀名
Dim strNewFullPath      As String '新Excel文件
Dim strXH               As String '打印序号

    
    If Rs.RecordCount <= 0 Then
        MsgBox "没有要导出的资料！", vbInformation, "提示！"
        Exit Sub
    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "初始化数据..."
    
    strFileName = DirShare & "\output_packing_list.xls" '要打开的文件
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '获取后缀名
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径
    strXH = Mid$(strExName, InStrRev(strExName, "_"))    '获取序号
    Rs.MoveFirst    '数据集移动到第一个
    
    If Rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '是否显示
        
        Set wkbk = ExApp.Workbooks.open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        '赋值到Excel中，表头
        wkst.Cells(8, 12) = Format(Date, "YYYY/mm/DD")
        wkst.Cells(9, 12) = "HTKS-SEDC" & Format(Date, "YY-MMDD") & strXH
        
        lngRows = 17
        IntInertRow = Rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = Rs.RecordCount
        
'        ClsP.ShowProgress 50, "正在导出..."
        
        Do While Not Rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(2).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & Rs.fields(3).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & Rs.fields(4).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & Rs.fields(5).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(6).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & Rs.fields(7).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & Rs.fields(8).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & Rs.fields(9).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & Rs.fields(10).Value)
            
            'jiayun 修改 ，把0改为空值
            
'            wkst.Cells(lngRows, 10) = Trim$("" & Rs.fields(11).Value)
'            wkst.Cells(lngRows, 11) = Trim$("" & Rs.fields(12).Value)
            wkst.Cells(lngRows, 10) = ""
            wkst.Cells(lngRows, 11) = ""
            
            wkst.Cells(lngRows, 12) = Trim$("" & Rs.fields(13).Value)
            wkst.Cells(lngRows, 13) = Trim$("" & Rs.fields(14).Value)
            wkst.Cells(lngRows, 14) = Trim$("" & Rs.fields(15).Value)
            
            'jiayun add Bag#
            wkst.Cells(lngRows, 15) = Trim$("" & Rs.fields(16).Value)
            
            DblNum = DblNum + Val(Trim$("" & Rs.fields(10).Value))
            DblAmt = DblAmt + Val(Trim$("" & Rs.fields(12).Value))
            
            lngRows = lngRows + 1
            Rs.MoveNext
        Loop
        
        wkst.Cells(lngRows, 9) = DblNum
        'wkst.Cells(lngRows, 11) = DblAmt
        
        wkst.Cells(lngRows, 11) = ""
        
    Else
'        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Sub
    End If
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""楷体_GB2312,常?""&10報表名稱 ：   "  ' & Gsmc
'        .CenterHeader = "&""楷体_GB2312,常規""報表&""宋体,常規""" & Chr(10) & "&""楷体_GB2312,常?""&10日 期："
'        .RightHeader = "" & Chr(10) & "&""楷体_GB2312,常規""&10單位：滬士"
'        .LeftFooter = "&""楷体_GB2312,常規""&10制表人："
'        .CenterFooter = "&""楷体_GB2312,常規""&10制表日期："
        .RightFooter = "&20" & "第 &P 页，共 &N 页"
    End With
  '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.Number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Sub
            End If
        End If
    End If
    wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "导出成功！"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.Description, vbInformation, "提示！"
    Exit Sub
End Sub
'Invoice
Public Sub InputInvoiceExportPrintExcel(ByVal Rs As ADODB.Recordset, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double '总金额
Dim strExtsion          As String '后缀名
Dim strNewFullPath      As String '新Excel文件
Dim strXH               As String '序号
    
    If Rs.RecordCount <= 0 Then
        MsgBox "没有要导出的资料！", vbInformation, "提示！"
        Exit Sub
    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "初始化数据..."
    
    strFileName = DirShare & "\output_invoice.xls" '要打开的文件
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '获取后缀名
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径
    strXH = Mid$(strExName, InStrRev(strExName, "_"))    '获取序号
    Rs.MoveFirst    '数据集移动到第一个
    
    If Rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '是否显示
        
        Set wkbk = ExApp.Workbooks.open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        '赋值到Excel中，表头
        wkst.Cells(8, 11) = Format(Date, "YYYY/mm/DD")
        wkst.Cells(9, 11) = "HTKS-SEDC" & Format(Date, "YY-MMDD") & strXH
        
        lngRows = 17
        IntInertRow = Rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = Rs.RecordCount
        
'        ClsP.ShowProgress 50, "正在导出..."
        
        Do While Not Rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(2).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & Rs.fields(3).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & Rs.fields(4).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & Rs.fields(5).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(6).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & Rs.fields(7).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & Rs.fields(8).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & Rs.fields(9).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & Rs.fields(10).Value)
            wkst.Cells(lngRows, 10) = Trim$("" & Rs.fields(11).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & Rs.fields(12).Value)
            
            DblNum = DblNum + Val(Trim$("" & Rs.fields(10).Value))
            DblAmt = DblAmt + Val(Trim$("" & Rs.fields(12).Value))
            
            
            lngRows = lngRows + 1
            Rs.MoveNext
        Loop
        
        wkst.Cells(lngRows, 9) = DblNum
        wkst.Cells(lngRows, 11) = DblAmt
        
    Else
'        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Sub
    End If
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""楷体_GB2312,常?""&10報表名稱 ：   "  ' & Gsmc
'        .CenterHeader = "&""楷体_GB2312,常規""報表&""宋体,常規""" & Chr(10) & "&""楷体_GB2312,常?""&10日 期："
'        .RightHeader = "" & Chr(10) & "&""楷体_GB2312,常規""&10單位：滬士"
'        .LeftFooter = "&""楷体_GB2312,常規""&10制表人："
'        .CenterFooter = "&""楷体_GB2312,常規""&10制表日期："
        .RightFooter = "&20" & "第 &P 页，共 &N 页"
    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.Number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Sub
            End If
        End If
    End If
    wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "导出成功！"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.Description, vbInformation, "提示！"
    Exit Sub
End Sub

'Daily_inventory
Public Sub Daily_InvExportPrintExcel(ByVal Rs As ADODB.Recordset, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double '总金额
Dim strExtsion          As String '后缀名
Dim strNewFullPath      As String '新Excel文件

    
    If Rs.RecordCount <= 0 Then
        MsgBox "没有要导出的资料！", vbInformation, "提示！"
        Exit Sub
    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "初始化数据..."
    
    strFileName = DirShare & "\Daily_inventory_report.xls" '要打开的文件
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '获取后缀名
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径
    Rs.MoveFirst    '数据集移动到第一个
    
    If Rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '是否显示
        
        Set wkbk = ExApp.Workbooks.open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        
        lngRows = 3
        IntInertRow = Rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = Rs.RecordCount
        
'        ClsP.ShowProgress 50, "正在导出..."
        
        Do While Not Rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(1).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & Rs.fields(2).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & Rs.fields(3).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & Rs.fields(4).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(5).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & Rs.fields(6).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & Rs.fields(7).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & Rs.fields(8).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & Rs.fields(9).Value)
            wkst.Cells(lngRows, 10) = Trim$("" & Rs.fields(10).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & Rs.fields(11).Value)
            wkst.Cells(lngRows, 12) = Trim$("" & Rs.fields(12).Value)
            
            lngRows = lngRows + 1
            Rs.MoveNext
        Loop
        
    Else
'        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Sub
    End If
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""楷体_GB2312,常?""&10報表名稱 ：   "  ' & Gsmc
'        .CenterHeader = "&""楷体_GB2312,常規""報表&""宋体,常規""" & Chr(10) & "&""楷体_GB2312,常?""&10日 期："
'        .RightHeader = "" & Chr(10) & "&""楷体_GB2312,常規""&10單位：滬士"
'        .LeftFooter = "&""楷体_GB2312,常規""&10制表人："
'        .CenterFooter = "&""楷体_GB2312,常規""&10制表日期："
        .RightFooter = "&20" & "第 &P 页，共 &N 页"
    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.Number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Sub
            End If
        End If
    End If
    wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "导出成功！"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.Description, vbInformation, "提示！"
    Exit Sub
End Sub

'shipping Packing list
Public Sub ShippingPackinglistExportPrintExcel(ByVal Ordertemp As String, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double  '总金额
Dim intBoxNum           As Integer '箱数
Dim strPBigBox          As String  '前箱号
Dim strNBigBox          As String  '新箱号
Dim IntBMegerRow        As Integer
Dim IntEMegerRow        As Integer
Dim DblJZ               As Double   '净重
Dim DblMZ               As Double   '毛重
Dim DblJZ1               As Double   '净重
Dim DblMZ1               As Double   '毛重
Dim DblJZ2               As Double   '净重
Dim DblMZ2               As Double   '毛重
Dim intBegin            As Integer
Dim strdjTmp            As String
Dim SD                  As String
Dim SD1                  As String
Dim strTmp()            As String
Dim strExtsion          As String '后缀名
Dim strNewFullPath      As String '新Excel文件
Dim RsNew               As New ADODB.Recordset '记录大箱的个数，方便后面计算体积重
Dim Rs               As New ADODB.Recordset


    strPBigBox = ""
    strNBigBox = ""
    strdjTmp = ""
    intBoxNum = 1
'    If Rs.RecordCount <= 0 Then
'        MsgBox "没有要导出的资料！", vbInformation, "提示！"
'        Exit Sub
'    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "初始化数据..."
    
    strFileName = DirShare & "\shipping_packing_list.xlsx" '要打开的文件
    
    
    strSql = "SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & _
                 "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 ",箱号,料号,mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,净重,毛重,尺寸 " & _
                 " FROM Vw_InvShippedPLFor37 a  where 单据编号 in ('" & Ordertemp & "')  order by 箱号"

     Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     
'    '-------------RS判定和筛选--------------------------------------------
'    If InStr(strdjbh, ",") > 0 Then
'        strTmp = Split(strdjbh, ",")
'        For i = 0 To UBound(strTmp)
'            strdjTmp = strdjTmp + "单据编号='" + strTmp(i) + "' OR "
'        Next
'        strdjTmp = Mid$(strdjTmp, 1, Len(strdjTmp) - 5)
'    Else
'        strdjTmp = "单据编号='" & strdjbh & "'"
'    End If
'    Rs.Filter = Trim(strdjTmp)          '数据集筛选
'    Rs.Sort = "单据编号,箱号 ASC"       '数据集排序
   strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '获取后缀名
   strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径
'    Rs.MoveFirst    '数据集移动到第一个
'    '---------------------------------------------------------------------
    If Rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False   '是否显示
        
        Set wkbk = ExApp.Workbooks.open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblJZ = 0
        DblMZ = 0
    
        '赋值到Excel中，表头
        wkst.Cells(8, 2) = Trim$("" & Rs.fields(2).Value)
        wkst.Cells(8, 17) = Trim$("" & Rs.fields(3).Value)
        wkst.Cells(17, 2) = Trim$("" & Rs.fields(4).Value)
        wkst.Cells(18, 2) = Trim$("" & Rs.fields(5).Value)
        wkst.Cells(19, 2) = Trim$("" & Rs.fields(6).Value) & " " & Trim$("" & Rs.fields(7).Value)
        wkst.Cells(20, 2) = Trim$("" & Rs.fields(8).Value) & " " & Trim$("" & Rs.fields(9).Value) & " " & Trim$("" & Rs.fields(10).Value) & " " & Trim$("" & Rs.fields(11).Value)
        wkst.Cells(22, 2) = "Attn:" & Trim$("" & Rs.fields(12).Value) & " ,Tel:" & Trim$("" & Rs.fields(13).Value)
        wkst.Cells(23, 2) = ""
        wkst.Cells(17, 17) = Trim$("" & Rs.fields(11).Value) 'To
        wkst.Cells(25, 3) = Trim$("" & Rs.fields(14).Value)
        wkst.Cells(25, 6) = Trim$("" & Rs.fields(15).Value)
        
        lngRows = 28
        
        IntInertRow = Rs.RecordCount * 2
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
           wkst.Rows(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '边框无
        Next i
        IntMaxDetailRow = Rs.RecordCount
        
'        ClsP.ShowProgress 50, "正在导出..."
        
        IntBMegerRow = 27
        IntEMegerRow = 30
        intBegin = 1
        Dim QBX As String
        For i = 0 To Rs.RecordCount - 1
'            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(16).Value) '箱号

            strPBigBox = Trim$("" & Rs.fields(16).Value) '箱号
            'QBX = strPBigBox
            If strPBigBox <> strNBigBox Then
                strNBigBox = Trim$("" & Rs.fields(16).Value) '箱号
                '箱数
                intBoxNum = intBoxNum + 1
                wkst.Cells(lngRows, 1) = "K" & Trim(intBoxNum - 1) '箱号进行转换为客户所要资料
                
                IntBMegerRow = IntBMegerRow + intBegin
                intBegin = 1
            Else
'                '合并
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).Merge
'                '设定水平和竖直居中
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).HorizontalAlignment = xlCenter
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).VerticalAlignment = xlCenter
                '--------------------------
                intBegin = intBegin + 1
            End If
            
            If SD <> Trim$("" & Rs.fields(14).Value) Then
            SD = Trim$("" & Rs.fields(14).Value)
            SD1 = SD1 & SD & " "
           End If
          wkst.Cells(25, 3) = SD1
            
            wkst.Cells(lngRows, 3) = Trim$("" & Rs.fields(15).Value) 'PO
            wkst.Cells(lngRows, 4) = Trim$("" & Rs.fields(17).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(18).Value)
            wkst.Cells(lngRows, 7) = Val(Trim$("" & Rs.fields(19).Value)) / 1000 '数量改为已千为单位
            DblNum = DblNum + Val(Trim$("" & Rs.fields(19).Value))
            wkst.Cells(lngRows, 9) = "KPCS"
            wkst.Cells(lngRows, 10) = Trim$("" & Rs.fields(20).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & Rs.fields(21).Value) 'datacode
            wkst.Cells(lngRows, 12) = Trim$("" & Rs.fields(22).Value) 'lotno
            If strPBigBox <> QBX Then
            wkst.Cells(lngRows, 14) = Trim$("" & Rs.fields(24).Value) '净重
            wkst.Cells(lngRows, 15) = "KG"   '净重单位
            wkst.Cells(lngRows, 18) = "KG"   '毛重单位
            wkst.Cells(lngRows, 19) = Trim$("" & Rs.fields(26).Value)   '尺寸
            wkst.Cells(lngRows, 17) = Trim$("" & Rs.fields(25).Value)   '毛重
            End If
           
           DblJZ1 = Val(Trim$("" & Rs.fields(24).Value))
           If strPBigBox <> QBX Then
           DblJZ = DblJZ1 + DblJZ
           End If
            DblMZ1 = Val(Trim$("" & Rs.fields(25).Value))
           If strPBigBox <> QBX Then
            DblMZ = DblMZ + DblMZ1
            End If
            '
            
            
            
            lngRows = lngRows + 1
            
            wkst.Cells(lngRows, 4) = "CPN:"
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(23).Value)
            
            QBX = strPBigBox
            
            lngRows = lngRows + 1
            IntEMegerRow = lngRows
            Rs.MoveNext
        Next
        '计算汇总
        wkst.Cells(lngRows + 1, 7) = DblNum / 1000 '总数量改为已千为单位
        wkst.Cells(lngRows + 1, 9) = "KPCS" '单位
        wkst.Cells(lngRows + 1, 1) = Trim(intBoxNum - 1)    '箱数
        wkst.Cells(lngRows + 1, 14) = Format(DblJZ, "0.00") '净重
        wkst.Cells(lngRows + 1, 17) = DblMZ '毛重，记录它到后面进行对比
    Else
'        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Sub
    End If
    '查询箱号尺寸，计算体积重
    Dim strXHCC         As String       '箱数和尺寸
    Dim DblTJZ          As String       '体积重
    Dim order As String
    
    order = Replace(Ordertemp, "A", "")
    
    strXHCC = ""
    DblTJZ = 0
    'strdjbh1 = Mid(strdjbh, 2, Len(strdjbh) - 1)
    strSql = "SELECT COUNT(DISTINCT erpdata.dbo.f_getparent(b.箱号)) 箱数,c.尺寸 " & _
             " FROM erpdata..tblStockMove a " & _
             " INNER JOIN erpdata..tblStockMovesub b ON a.单据编号 = b.单据编号 AND a.序号=b.单据项次 " & _
             " INNER JOIN erpdata..tblStockNumTree c On c.箱号=erpdata.dbo.f_getparent(b.箱号) " & _
             " WHERE a.单据编号 IN ('" & order & "')" & _
             " GROUP BY c.尺寸"
    If RsNew.State = adStateOpen Then RsNew.Close
    RsNew.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If Not RsNew.EOF Then
        Do While Not RsNew.EOF
            '循环得到箱号的箱数和尺寸，进行拼接
            strXHCC = strXHCC & Trim$("" & RsNew!箱数) & "@" & Trim$("" & RsNew!尺寸) & "cm;"
            '对尺寸进行分割，计算体积重
            If Trim$("" & RsNew!尺寸) <> "" And InStr(Trim$("" & RsNew!尺寸), "*") > 0 Then
                strTmp = Split(Trim$("" & RsNew!尺寸), "*") '分割字符
                '计算体积重并汇总
                DblTJZ = DblTJZ + Val(Trim$("" & RsNew!箱数)) * strTmp(0) * strTmp(1) * strTmp(2) / 5000
            End If
            RsNew.MoveNext
        Loop
    End If
    RsNew.Close
    '赋值体积重
    wkst.Cells(lngRows + 3, 4) = Format(DblTJZ, "0.00")
    '比较体积重和毛重看哪个大,就取哪个
    If DblMZ > DblTJZ Then
        wkst.Cells(lngRows + 3, 11) = Format(DblMZ, "0.00")
    Else
        wkst.Cells(lngRows + 3, 11) = Format(DblTJZ, "0.00")
    End If
    '赋值到EXCEL箱数和尺寸
    wkst.Cells(lngRows + 4, 3) = strXHCC
'    ExApp.Columns.AutoFit '自适应列宽
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""楷体_GB2312,常?""&10報表名稱 ：   "  ' & Gsmc
'        .CenterHeader = "&""楷体_GB2312,常規""報表&""宋体,常規""" & Chr(10) & "&""楷体_GB2312,常?""&10日 期："
'        .RightHeader = "" & Chr(10) & "&""楷体_GB2312,常規""&10單位：滬士"
'        .LeftFooter = "&""楷体_GB2312,常規""&10制表人："
'        .CenterFooter = "&""楷体_GB2312,常規""&10制表日期："
       ' .RightFooter = "&20" & "第 &P 页，共 &N 页"
    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.Number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Sub
            End If
        End If
    End If
    wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "导出成功！"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.Description, vbInformation, "提示！"
    Exit Sub
End Sub

'shipping invoice
Public Sub ShippingInvoiceExportPrintExcel(ByVal Ordertemp As String, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double  '总金额
Dim intBoxNum           As Integer '箱数
Dim strPBigBox          As String  '前箱号
Dim strNBigBox          As String  '新箱号
Dim IntBMegerRow        As Integer
Dim IntEMegerRow        As Integer
Dim DblJZ               As Double   '净重
Dim DblMZ               As Double   '毛重
Dim intBegin            As Integer
Dim strdjTmp            As String
Dim strTmp()            As String
Dim SD                  As String
Dim SD1                  As String
Dim strExtsion          As String '后缀名
Dim strNewFullPath      As String '新Excel文件
Dim Rs               As New ADODB.Recordset
    strPBigBox = ""
    strNBigBox = ""
    strdjTmp = ""
    intBoxNum = 1
'
'    If Rs.RecordCount <= 0 Then
'        MsgBox "没有要导出的资料！", vbInformation, "提示！"
'        Exit Sub
'    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "初始化数据..."
    
    strFileName = DirShare & "\shipping_invoice.xlsx" '要打开的文件
    
    
                    
    strSql = " SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3" & _
                 " ,city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 " ,箱号,料号,mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,单价,AMount,customerPartNumber " & _
                 "  FROM Vw_InvShippedInvoiceFor37 a  where 单据编号 in ('" & Ordertemp & "')  order by 箱号  "

     Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     
'    '-------------RS判定和筛选--------------------------------------------
'    If InStr(strdjbh, ",") > 0 Then
'        strTmp = Split(strdjbh, ",")
'        For i = 0 To UBound(strTmp)
'            strdjTmp = strdjTmp + "单据编号='" + strTmp(i) + "' OR "
'        Next
'        strdjTmp = Mid$(strdjTmp, 1, Len(strdjTmp) - 5)
'    Else
'        strdjTmp = "单据编号='" & strdjbh & "'"
'    End If
'    Rs.Filter = Trim(strdjTmp)          '数据集筛选
'    Rs.Sort = "单据编号,箱号 ASC"       '数据集排序
   strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '获取后缀名
   strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径
'    Rs.MoveFirst    '数据集移动到第一个
'    '---------------------------------------------------------------------
    If Rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False   '是否显示
        
        Set wkbk = ExApp.Workbooks.open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
    
    
   
    
    
'    '-------------RS判定和筛选--------------------------------------------
'    If InStr(strdjbh, ",") > 0 Then
'        'strTmp = Split(strdjbh, ",")
'       ' For i = 0 To UBound(strTmp)
'            'strdjTmp = strdjTmp + "单据编号='" + strTmp(i) + "' AND "
'       ' Next
'       ' strdjTmp = Mid$(strdjTmp, 1, Len(strdjTmp) - 5)
'        MsgBox "无法同时导出多个单据号！"
'        Exit Sub
'    Else
'        strdjTmp = "单据编号='" & strdjbh & "'"
'    End If
'    Rs.Filter = Trim(strdjTmp)  '数据集筛选
'    Rs.Sort = "单据编号,箱号 ASC" '数据集排序
'    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '获取后缀名
'    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径
'    Rs.MoveFirst    '数据集移动到第一个
'    '---------------------------------------------------------------------
'    If Rs.RecordCount > 0 Then
''        ClsP.ShowProgress 30, "初始化Excel..."
'        Set ExApp = New Excel.Application
'        ExApp.Visible = False   '是否显示
'
'        Set wkbk = ExApp.Workbooks.open(strFileName)
'        Set wkst = wkbk.Worksheets(1)
''        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        DblJZ = 0
        DblMZ = 0
        '赋值到Excel中，表头
        wkst.Cells(13, 10) = Trim$("" & Rs.fields(2).Value)
        wkst.Cells(15, 10) = Trim$("" & Rs.fields(3).Value)
        wkst.Cells(18, 10) = Trim$("" & Rs.fields(11).Value) 'To
        
        wkst.Cells(13, 2) = Trim$("" & Rs.fields(4).Value)
        wkst.Cells(14, 2) = Trim$("" & Rs.fields(5).Value)
        wkst.Cells(15, 2) = Trim$("" & Rs.fields(6).Value) & " " & Trim$("" & Rs.fields(7).Value)
        wkst.Cells(16, 2) = Trim$("" & Rs.fields(8).Value) & " " & Trim$("" & Rs.fields(9).Value) & " " & Trim$("" & Rs.fields(10).Value) & " " & Trim$("" & Rs.fields(11).Value)
        
        wkst.Cells(18, 2) = "Attn:" & Trim$("" & Rs.fields(12).Value) & " ,Tel:" & Trim$("" & Rs.fields(13).Value)
        wkst.Cells(19, 2) = ""

        'wkst.Cells(23, 3) = Trim$("" & Rs.fields(14).Value)
        wkst.Cells(23, 5) = Trim$("" & Rs.fields(15).Value)
        
        lngRows = 27
        
        IntInertRow = Rs.RecordCount * 2
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
           wkst.Range(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '边框无
        Next i
        IntMaxDetailRow = Rs.RecordCount
        
'        ClsP.ShowProgress 50, "正在导出..."
        
        IntBMegerRow = 26
        IntEMegerRow = 29
        intBegin = 1
        Dim QBX As String
        
        For i = 0 To Rs.RecordCount - 1
'            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(16).Value) '箱号
            strPBigBox = Trim$("" & Rs.fields(16).Value) '箱号
            If strPBigBox <> strNBigBox Then
                strNBigBox = Trim$("" & Rs.fields(16).Value) '箱号
                '箱数
                intBoxNum = intBoxNum + 1
                wkst.Cells(lngRows, 1) = "K" & Trim(intBoxNum - 1) '箱号进行转换为客户所要资料
                QBX = "K" & Trim(intBoxNum - 1)
                
                IntBMegerRow = IntBMegerRow + intBegin
                intBegin = 1
            Else
'                '合并
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).Merge
'                '设定水平和竖直居中
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).HorizontalAlignment = xlCenter
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).VerticalAlignment = xlCenter
                '--------------------------
                intBegin = intBegin + 1
            End If
              If SD <> Trim$("" & Rs.fields(14).Value) Then
             SD = Trim$("" & Rs.fields(14).Value)
             SD1 = SD1 & SD & " "
             End If
            wkst.Cells(23, 3) = SD1
            
            wkst.Cells(lngRows, 3) = Trim$("" & Rs.fields(15).Value) 'PO
            wkst.Cells(lngRows, 4) = Trim$("" & Rs.fields(17).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(18).Value)
            wkst.Cells(lngRows, 7) = Val(Trim$("" & Rs.fields(19).Value)) / 1000 '数量都为除以1000的值
            DblNum = DblNum + Val(Trim$("" & Rs.fields(19).Value))
            wkst.Cells(lngRows, 9) = "KPCS"
            wkst.Cells(lngRows, 10) = Trim$("" & Rs.fields(20).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & Rs.fields(21).Value)
            wkst.Cells(lngRows, 12) = Trim$("" & Rs.fields(22).Value)
            wkst.Cells(lngRows, 13) = "US$"
            wkst.Cells(lngRows, 14) = Val(Trim$("" & Rs.fields(23).Value)) * 1000 '单价为乘以1000的值
            wkst.Cells(lngRows, 15) = "US$"
            wkst.Cells(lngRows, 16) = Trim$("" & Rs.fields(24).Value)
            DblAmt = DblAmt + Val(Trim$("" & Rs.fields(24).Value))
            lngRows = lngRows + 1
            
            wkst.Cells(lngRows, 4) = "CPN:"
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(25).Value)
            
            
            
            lngRows = lngRows + 1
            IntEMegerRow = lngRows
            Rs.MoveNext
        Next
        
        '计算汇总
        wkst.Cells(lngRows + 1, 7) = DblNum / 1000 '数量
        wkst.Cells(lngRows + 1, 9) = "KPCS" '单位
        wkst.Cells(lngRows + 1, 16) = DblAmt
        wkst.Cells(lngRows + 1, 1) = Trim(intBoxNum - 1)

        
    Else
'        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Sub
    End If
'    ExApp.Columns.AutoFit '自适应列宽
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""楷体_GB2312,常?""&10報表名稱 ：   "  ' & Gsmc
'        .CenterHeader = "&""楷体_GB2312,常規""報表&""宋体,常規""" & Chr(10) & "&""楷体_GB2312,常?""&10日 期："
'        .RightHeader = "" & Chr(10) & "&""楷体_GB2312,常規""&10單位：滬士"
'        .LeftFooter = "&""楷体_GB2312,常規""&10制表人："
'        .CenterFooter = "&""楷体_GB2312,常規""&10制表日期："
        '.RightFooter = "&20" & "第 &P 页，共 &N 页"
    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.Number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Sub
            End If
        End If
    End If
    wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "导出成功！"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.Description, vbInformation, "提示！"
    Exit Sub
End Sub

'根据Rs数据集语句导出Excel
Public Sub RsExporToExcel(Rs As ADODB.Recordset, RptName As String, ExcelFileName As String)
Dim Irowcount       As Long
Dim Icolcount       As Integer
Dim strFileName     As String

    Dim xlApp As New Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xlQuery As Excel.QueryTable
    Screen.MousePointer = 11
    With Rs
        If .RecordCount < 1 Then
            Screen.MousePointer = 0
            MsgBox ("没有可导出的资料")
            Exit Sub
        End If
        Irowcount = .RecordCount
        Icolcount = .fields.Count
    End With
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlBook = xlApp.Workbooks().Add
    Set xlSheet = xlBook.Worksheets("sheet1")
    xlApp.Visible = False

    Set xlQuery = xlSheet.QueryTables.Add(Rs, xlSheet.Range("a1"))
    
'    With xlQuery
'        .FieldNames = True
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .BackgroundQuery = True
'        .RefreshStyle = xlInsertDeleteCells
'        .SavePassword = True
'        .SaveData = True
'        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
'        .PreserveColumnInfo = True
'    End With
    xlSheet.Name = RptName
    xlQuery.FieldNames = True '
    xlQuery.Refresh
    
    With xlSheet
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Name = "宋体"
        '
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Bold = True
        '
        .Range(.Cells(1, 1), .Cells(Irowcount + 1, Icolcount)).Borders.LineStyle = xlContinuous
        '
'        .Range(.Cells(2, 1), .Cells(Irowcount + 1, Icolcount)).Font.Size = 9
    End With
    '另存文件
    strFileName = DirInvRpt + "\" + ExcelFileName
    xlBook.SaveAs strFileName, xlNormal, "", "", False, False
    xlBook.Saved = True
    
    Screen.MousePointer = 0
'    xlApp.Visible = True
    Set xlSheet = Nothing
    xlBook.Close
    Set xlBook = Nothing
    xlApp.Quit
    Set xlApp = Nothing
    
    
    
    

End Sub
