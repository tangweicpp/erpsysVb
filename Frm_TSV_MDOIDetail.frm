VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Frm_TSV_MDOIDetail 
   Caption         =   "市场部 来料明细报表查询"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0FF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9885
   ScaleWidth      =   17115
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtPT 
      Height          =   375
      Left            =   10200
      TabIndex        =   13
      Top             =   180
      Width           =   1695
   End
   Begin VB.ComboBox CmbFactory 
      Height          =   315
      ItemData        =   "Frm_TSV_MDOIDetail.frx":0000
      Left            =   1320
      List            =   "Frm_TSV_MDOIDetail.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "导出"
      Height          =   360
      Left            =   4440
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "查询"
      Height          =   360
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   7935
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   16695
      _Version        =   524288
      _ExtentX        =   29448
      _ExtentY        =   13996
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
      SpreadDesigner  =   "Frm_TSV_MDOIDetail.frx":0025
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin MSDataListLib.DataCombo CmbCustomer 
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   257556481
      CurrentDate     =   41424
   End
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   315
      Left            =   4440
      TabIndex        =   6
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   257556481
      CurrentDate     =   41424
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户机种："
      Height          =   195
      Left            =   9240
      TabIndex        =   12
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   6120
      TabIndex        =   11
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "事业部："
      Height          =   195
      Left            =   600
      TabIndex        =   9
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束时间："
      Height          =   195
      Left            =   3360
      TabIndex        =   8
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始时间："
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户代码："
      Height          =   195
      Left            =   3360
      TabIndex        =   4
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "Frm_TSV_MDOIDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0          'Detail

    E_id = 0                'id
    E_PO_NUM                'PO_NO
    E_SOURCE_BATCH_ID       'SOURCE_BATCH_ID
    E_MTRL_NUM               'WO_NO
    E_MPN_DESC               'CustomerDevice
    E_FAB_CONV_ID            'FAB_Device
    E_IMAGER_CUSTOMER_REV    'Version
    E_TEST_SITE              'Ship_To
    E_CREATED_DATE            'Date
    E_SHIP_SITE               'Supplier
    E_QTECH_CREATED_DATE       '上传日期
    E_CUSTOMERSHORTNAME        '客户代码
    E_LotID                     'LotID
    E_SUBSTRATEID               'WaferID
    E_PASSBINCOUNT               'GoodDieQty
    E_FAILBINCOUNT               'NGDieQty
    E_ProductId                  'MarkingLotID
    
    E_End
    
End Enum

Dim rs         As New ADODB.Recordset

Dim mainItemRS As New ADODB.Recordset

Dim bomRS2     As New ADODB.Recordset

Private Sub CmdOut_Click()

    Dim beginTime   As String

    Dim endTime     As String

    Dim woTemp      As String

    Dim productTemp As String

    Dim sqlTemp     As String

    Dim sql1        As String

    Dim sql2        As String

    Dim sql3        As String

    sql1 = ""
    sql2 = ""
    sql3 = ""

    woTemp = UCase(Trim(txtWO.Text))
    productTemp = UCase(Trim(TxtProduct.Text))
    beginTime = Format(DTP1.Value, "YYYY/MM/DD")
    endTime = Format(DTP2.Value, "YYYY/MM/DD")

    sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
          
    sql3 = " order by a.ordername,b.waferid  "
  
    If productTemp <> "" Then
  
        sql2 = " and a.product='" + productTemp + "'"
  
    End If
  
    If woTemp <> "" Then
  
        sql2 = sql2 & " and a.ORDERNAME='" + woTemp + "'"
  
    End If
  
    If Trim(sql2) <> "" Then
  
        sqlTemp = sql1 & sql2 & sql3
  
    Else
  
        sql2 = " and a.erpcreatedate>=to_date('" + beginTime + "','YYYY/MM/DD') and a.erpcreatedate<=to_date('" + endTime + "','YYYY/MM/DD')"
  
        sqlTemp = sql1 & sql2 & sql3
  
    End If
  
    ExporToExcel (sqlTemp)

End Sub

Private Sub Command1_Click()

    Dim waferIdTemp As String

    Dim endTime     As String

    Dim woTemp      As String

    Dim productTemp As String

    Dim sqlTemp     As String

    Dim sql1        As String

    Dim sql2        As String

    Dim sql3        As String

    sql1 = ""
    sql2 = ""
    sql3 = ""
    waferIdTemp = UCase(Trim(txtWaferID.Text))

    'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
    '          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
    '
          
    '  sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & _
    '"  CASE ORDERTYPE  WHEN '1' THEN '一般工单'  WHEN '5' THEN '再加工工单'   WHEN '7' THEN '委外工单'   WHEN '8' THEN '重工委外工单' " & _
    '" WHEN '11' THEN '拆件式工单'    WHEN '13' THEN '预测工单'   WHEN '15' THEN '试产工单' Else '其他' END as ORDERTYPE ," & _
    '"  PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & _
    '" PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & _
    '" from erpintegration2.ib_wohistory a where  modifyflag='1' "

    sql1 = " select b.wafersequence, b.ordername,b.waferid,b.completeflag,b.dieqty, b.FGDieQty , b.WaferLot, b.MarkingCode " & "from erpintegration2.ib_wohistory a, ib_waferlist b  where a.modifyflag='1' and b.ordername=a.ordername and b.waferid='" + waferIdTemp + "' "
          
    sql3 = " order by a.ordername ,b.WaferLot,b.waferid "
  
    If Trim(sql2) <> "" Then
  
        sqlTemp = sql1 & sql2 & sql3
  
    Else
  
        sql2 = ""
  
        sqlTemp = sql1 & sql2 & sql3
  
    End If

    Set mainItemRS = GetPMCWOLine(sqlTemp)

    With fps(1)
        .MaxRows = 0

        If mainItemRS.RecordCount > 0 Then
            Set .DataSource = mainItemRS
       
        End If

    End With

End Sub

Private Sub CmbCustomer_Change()

    Dim customerTemp As String

    customerTemp = CmbCustomer.Text

    Label5.Caption = GetCustomerNameSqlServer(customerTemp)

End Sub

'Private Sub CmbCustomer_Click(Area As Integer)
'
'End Sub

'Private Sub CmbFactory_Change()
'Dim factoryTemp As String
'factoryTemp = CmbFactory.Text
'
'IniCustomerName (factoryTemp)
'
'End Sub

Private Sub CmbFactory_Click()

    CmbCustomer.Text = ""

    Dim factoryTemp As String

    factoryTemp = CmbFactory.Text

    IniCustomerName (factoryTemp)

End Sub

Private Sub cmdQuery_Click()

    Dim sqlTemp   As String

    Dim date1Temp As String

    Dim date2Temp As String

    Dim custTemp  As String

    Dim ptTemp    As String

    Dim sqlTemp2  As String

    date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
    date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")
     
    sqlTemp = " select a.PO_NUM , a.SOURCE_BATCH_ID,   a.MTRL_NUM , a.MPN_DESC,  a.FAB_CONV_ID , a.IMAGER_CUSTOMER_REV,  a.TEST_SITE , b.qtech_created_date  ,a.SHIP_SITE , " & " b.qtech_created_date , a.CustomershortName, b.lotid, b.SubstrateId, b.PassBinCount, b.FailBinCount, b.ProductId " & " from customeroitbl_test a,mappingdatatest b " & " Where a.flag='Y' and b.flag='Y' and  b.qtech_created_date>=to_date('" + date1Temp + "','YYYY-MM-DD') " & " and b.qtech_created_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD') and b.filename=to_char(a.id)"

    If CmbCustomer.Text <> "" Then

        custTemp = UCase(Trim(CmbCustomer.Text))

        If custTemp = "56" Then
            sqlTemp = " select a.PO_NUM , a.SOURCE_BATCH_ID,   a.MTRL_NUM , a.MPN_DESC,  a.FAB_CONV_ID , a.IMAGER_CUSTOMER_REV,  a.TEST_SITE , b.qtech_created_date  ,a.SHIP_SITE , " & " b.qtech_created_date , a.CustomershortName, b.lotid, b.SubstrateId, b.PassBinCount, b.FailBinCount, b.ProductId " & " from customeroitbl_test a,mappingdatatest b " & " Where a.flag='Y' and b.flag='Y' and  b.qtech_created_date>=to_date('" + date1Temp + "','YYYY-MM-DD') " & " and b.qtech_created_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD')  and a.source_batch_id = b.lotid"
        Else
            sqlTemp = " select a.PO_NUM , a.SOURCE_BATCH_ID,   a.MTRL_NUM , a.MPN_DESC,  a.FAB_CONV_ID , a.IMAGER_CUSTOMER_REV,  a.TEST_SITE , b.qtech_created_date  ,a.SHIP_SITE , " & " b.qtech_created_date , a.CustomershortName, b.lotid, b.SubstrateId, b.PassBinCount, b.FailBinCount, b.ProductId " & " from customeroitbl_test a,mappingdatatest b " & " Where a.flag='Y' and b.flag='Y' and  b.qtech_created_date>=to_date('" + date1Temp + "','YYYY-MM-DD') " & " and b.qtech_created_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD') and b.filename=to_char(a.id)"
        
        End If

    End If

    If CmbCustomer.Text <> "" Then

        custTemp = UCase(Trim(CmbCustomer.Text))

        sqlTemp = sqlTemp + " and a.customershortname='" + custTemp + "'  and b.customershortname='" + custTemp + "'  "

    End If

    If TxtPT.Text <> "" Then

        ptTemp = UCase(Trim(TxtPT.Text))

        sqlTemp = sqlTemp + " and a.MPN_DESC='" + ptTemp + "'   "

    End If

    sqlTemp2 = "  order by a.CustomershortName,a.PO_NUM, a.MTRL_NUM "

    sqlTemp = sqlTemp + sqlTemp2

    Set mainItemRS = GetMDPODetail(sqlTemp)
           
    With fps(0)
        .MaxRows = 0

        If mainItemRS.RecordCount > 0 Then
            Set .DataSource = mainItemRS
       
        End If

    End With

    MsgBox "查询成功！", vbInformation, "友情提示"

End Sub

Private Sub Command2_Click()

    Dim sqlTemp   As String

    Dim date1Temp As String

    Dim date2Temp As String

    Dim custTemp  As String

    Dim ptTemp    As String

    Dim sqlTemp2  As String

    date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
    date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")
     
    sqlTemp = " select a.PO_NUM , a.SOURCE_BATCH_ID,   a.MTRL_NUM as WO_NO , a.MPN_DESC as CustomerDevice,  a.FAB_CONV_ID as FAB_Device , a.IMAGER_CUSTOMER_REV as Version ,  a.TEST_SITE as Ship_To , a.CREATED_DATE  ,a.SHIP_SITE as Supplier , " & " a.qtech_created_date , a.CustomershortName, b.lotid, b.SubstrateId as WaferID, b.PassBinCount, b.FailBinCount, b.ProductId as MarkingLotID " & " from customeroitbl_test a,mappingdatatest b " & " Where a.flag='Y' and b.flag='Y' and  b.qtech_created_date>=to_date('" + date1Temp + "','YYYY-MM-DD') " & " and b.qtech_created_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD') and b.filename=to_char(a.id)"

    If CmbCustomer.Text <> "" Then

        custTemp = UCase(Trim(CmbCustomer.Text))

        sqlTemp = sqlTemp + " and a.customershortname='" + custTemp + "'  and b.customershortname='" + custTemp + "'  "

    End If

    If TxtPT.Text <> "" Then

        ptTemp = UCase(Trim(TxtPT.Text))

        sqlTemp = sqlTemp + " and a.MPN_DESC='" + ptTemp + "'   "

    End If

    sqlTemp2 = "  order by a.CustomershortName,a.PO_NUM, a.MTRL_NUM "

    sqlTemp = sqlTemp + sqlTemp2
           
    ExporToExcel (sqlTemp)

End Sub

Private Sub ComOutLine_Click()

    Dim beginTime   As String

    Dim endTime     As String

    Dim woTemp      As String

    Dim productTemp As String

    Dim sqlTemp     As String

    Dim sql1        As String

    Dim sql2        As String

    Dim sql3        As String

    sql1 = ""
    sql2 = ""
    sql3 = ""

    woTemp = UCase(Trim(txtWO.Text))
    productTemp = UCase(Trim(TxtProduct.Text))
    beginTime = Format(DTP1.Value, "YYYY/MM/DD")
    endTime = Format(DTP2.Value, "YYYY/MM/DD")

    'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
    '          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
    '
          
    sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & "  CASE ORDERTYPE  WHEN '1' THEN '一般工单'  WHEN '5' THEN '再加工工单'   WHEN '7' THEN '委外工单'   WHEN '8' THEN '重工委外工单' " & " WHEN '11' THEN '拆件式工单'    WHEN '13' THEN '预测工单'   WHEN '15' THEN '试产工单' Else '其他' END as ORDERTYPE ," & "  PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & " PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & " from erpintegration2.ib_wohistory a where  modifyflag='1' "
          
    sql3 = " order by a.ordername  "
  
    If productTemp <> "" Then
  
        sql2 = " and a.product='" + productTemp + "'"
  
    End If
  
    If woTemp <> "" Then
  
        sql2 = sql2 & " and a.ORDERNAME='" + woTemp + "'"
  
    End If
  
    If Trim(sql2) <> "" Then
  
        sqlTemp = sql1 & sql2 & sql3
  
    Else
  
        sql2 = " and a.erpcreatedate>=to_date('" + beginTime + "','YYYY/MM/DD') and a.erpcreatedate<to_date('" + endTime + "','YYYY/MM/DD')+1"
  
        sqlTemp = sql1 & sql2 & sql3
  
    End If
  
    ExporToExcel (sqlTemp)

End Sub

Private Sub ComQueryHead_Click()
    'HEAD

    Dim beginTime   As String

    Dim endTime     As String

    Dim woTemp      As String

    Dim productTemp As String

    Dim sqlTemp     As String

    Dim sql1        As String

    Dim sql2        As String

    Dim sql3        As String

    sql1 = ""
    sql2 = ""
    sql3 = ""

    woTemp = UCase(Trim(txtWO.Text))
    productTemp = UCase(Trim(TxtProduct.Text))
    beginTime = Format(DTP1.Value, "YYYY/MM/DD")
    endTime = Format(DTP2.Value, "YYYY/MM/DD")

    'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
    '          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
    '
          
    sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & "  CASE ORDERTYPE  WHEN '1' THEN '一般工单'  WHEN '5' THEN '再加工工单'   WHEN '7' THEN '委外工单'   WHEN '8' THEN '重工委外工单' " & " WHEN '11' THEN '拆件式工单'    WHEN '13' THEN '预测工单'   WHEN '15' THEN '试产工单' Else '其他' END as ORDERTYPE ," & "  PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & " PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & " from erpintegration2.ib_wohistory a where  modifyflag='1' "
          
    sql3 = " order by a.ordername  "
  
    If productTemp <> "" Then
  
        sql2 = " and a.product='" + productTemp + "'"
  
    End If
  
    If woTemp <> "" Then
  
        sql2 = sql2 & " and a.ORDERNAME='" + woTemp + "'"
  
    End If
  
    If Trim(sql2) <> "" Then
  
        sqlTemp = sql1 & sql2 & sql3
  
    Else
  
        sql2 = " and a.erpcreatedate>=to_date('" + beginTime + "','YYYY/MM/DD') and a.erpcreatedate<to_date('" + endTime + "','YYYY/MM/DD')+1"
  
        sqlTemp = sql1 & sql2 & sql3
  
    End If

    Set reportRS = GetPMCWOHeader(sqlTemp)

    With fps(0)
        .MaxRows = 0

        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If

    End With

End Sub

Private Sub ComQueryLine_Click()
    'Line

    Dim beginTime   As String

    Dim endTime     As String

    Dim woTemp      As String

    Dim productTemp As String

    Dim sqlTemp     As String

    Dim sql1        As String

    Dim sql2        As String

    Dim sql3        As String

    sql1 = ""
    sql2 = ""
    sql3 = ""

    woTemp = UCase(Trim(txtWO.Text))
    productTemp = UCase(Trim(TxtProduct.Text))
    beginTime = Format(DTP1.Value, "YYYY/MM/DD")
    endTime = Format(DTP2.Value, "YYYY/MM/DD")

    'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
    '          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
    '
          
    '  sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & _
    '"  CASE ORDERTYPE  WHEN '1' THEN '一般工单'  WHEN '5' THEN '再加工工单'   WHEN '7' THEN '委外工单'   WHEN '8' THEN '重工委外工单' " & _
    '" WHEN '11' THEN '拆件式工单'    WHEN '13' THEN '预测工单'   WHEN '15' THEN '试产工单' Else '其他' END as ORDERTYPE ," & _
    '"  PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & _
    '" PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & _
    '" from erpintegration2.ib_wohistory a where  modifyflag='1' "

    sql1 = " select b.wafersequence, b.ordername,b.waferid,b.completeflag,b.dieqty, b.FGDieQty , b.WaferLot, b.MarkingCode " & "from erpintegration2.ib_wohistory a, ib_waferlist b  where a.modifyflag='1' and b.ordername=a.ordername"
          
    sql3 = " order by a.ordername ,b.WaferLot,b.waferid "
  
    If productTemp <> "" Then
  
        sql2 = " and a.product='" + productTemp + "'"
  
    End If
  
    If woTemp <> "" Then
  
        sql2 = sql2 & " and a.ORDERNAME='" + woTemp + "'"
  
    End If
  
    If Trim(sql2) <> "" Then
  
        sqlTemp = sql1 & sql2 & sql3
  
    Else
  
        sql2 = " and a.erpcreatedate>=to_date('" + beginTime + "','YYYY/MM/DD') and a.erpcreatedate<to_date('" + endTime + "','YYYY/MM/DD')+1"
  
        sqlTemp = sql1 & sql2 & sql3
  
    End If

    Set mainItemRS = GetPMCWOLine(sqlTemp)

    With fps(1)
        .MaxRows = 0

        If mainItemRS.RecordCount > 0 Then
            Set .DataSource = mainItemRS
       
        End If

    End With

End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub Form_Activate()
    'CmbLine.Text = "TSV"

    IniFpsHeader1
    'IniFpsHeader2

End Sub

'Private Sub IniProduct()
'Set mainItemRS = GetProduct()
'Set Text3.RowSource = mainItemRS
'Text3.ListField = mainItemRS("productname").Name
'Text3.BoundColumn = mainItemRS("PID").Name
'
'End Sub

Private Sub IniCustomerName(factoryTemp As String)
    Set mainItemRS = GetOracleCustomerName(factoryTemp)
    Set CmbCustomer.RowSource = mainItemRS
    CmbCustomer.ListField = mainItemRS("PID").Name
    CmbCustomer.BoundColumn = mainItemRS("PID").Name

End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub IniFpsHeader1()

    With fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
        .MaxRows = 0
        
        '砞竚Α
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = False
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        .SetText E_FPS0.E_id, 0, "序号"
        .SetText E_FPS0.E_PO_NUM, 0, "PO_NO"
        .SetText E_FPS0.E_SOURCE_BATCH_ID, 0, "SOURCE_BATCH_ID"
        .SetText E_FPS0.E_MTRL_NUM, 0, "WO_NO"
        .SetText E_FPS0.E_MPN_DESC, 0, "CustomerDevice"
        .SetText E_FPS0.E_FAB_CONV_ID, 0, "FAB_Device"
        .SetText E_FPS0.E_IMAGER_CUSTOMER_REV, 0, "Version"
        .SetText E_FPS0.E_TEST_SITE, 0, "Ship_To"
        .SetText E_FPS0.E_CREATED_DATE, 0, "Date"
        .SetText E_FPS0.E_SHIP_SITE, 0, "Supplier"

        .SetText E_FPS0.E_QTECH_CREATED_DATE, 0, "上传日期"
        .SetText E_FPS0.E_CUSTOMERSHORTNAME, 0, "客户代码"
        .SetText E_FPS0.E_LotID, 0, "LotID"
        .SetText E_FPS0.E_SUBSTRATEID, 0, "WaferID"
        .SetText E_FPS0.E_PASSBINCOUNT, 0, "GoodDieQty"
        .SetText E_FPS0.E_FAILBINCOUNT, 0, "NGDieQty"
        .SetText E_FPS0.E_ProductId, 0, "MarkingLotID"

        .ColWidth(E_FPS0.E_id) = 5
        .ColWidth(E_FPS0.E_PO_NUM) = 12
        .ColWidth(E_FPS0.E_SOURCE_BATCH_ID) = 8
        .ColWidth(E_FPS0.E_MTRL_NUM) = 13
        
        .ColWidth(E_FPS0.E_MPN_DESC) = 8
        .ColWidth(E_FPS0.E_FAB_CONV_ID) = 8
        .ColWidth(E_FPS0.E_IMAGER_CUSTOMER_REV) = 8
        .ColWidth(E_FPS0.E_TEST_SITE) = 8
        .ColWidth(E_FPS0.E_CREATED_DATE) = 8
        .ColWidth(E_FPS0.E_SHIP_SITE) = 8
        .ColWidth(E_FPS0.E_QTECH_CREATED_DATE) = 8
        .ColWidth(E_FPS0.E_CUSTOMERSHORTNAME) = 8
        .ColWidth(E_FPS0.E_LotID) = 9
        .ColWidth(E_FPS0.E_SUBSTRATEID) = 12
        .ColWidth(E_FPS0.E_PASSBINCOUNT) = 8
        .ColWidth(E_FPS0.E_FAILBINCOUNT) = 8
        .ColWidth(E_FPS0.E_ProductId) = 8

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .ReDraw = True

    End With

End Sub

Private Sub IniFpsHeader2()

    With fps(1)
        .ReDraw = False
        .MaxCols = E_FPS1.E_End - 1
        .MaxRows = 0
        
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = False
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
          
        .SetText E_FPS1.E_id, 0, "序号"
        .SetText E_FPS1.E_Wo, 0, "工单号"
        .SetText E_FPS1.e_WaferID, 0, "WaferId"
        .SetText E_FPS1.E_CompleteFlag, 0, "完成标志"
        .SetText E_FPS1.E_TotalDie, 0, "TotalDie数量"
        .SetText E_FPS1.E_GoodDie, 0, "GoodDie数量"
        .SetText E_FPS1.E_WaferLot, 0, "WaferLot"
        .SetText E_FPS1.E_MarkingCode, 0, "MarkingCode"
        
        .ColWidth(E_FPS1.E_id) = 10
        .ColWidth(E_FPS1.E_Wo) = 10
        .ColWidth(E_FPS1.e_WaferID) = 10
        .ColWidth(E_FPS1.E_CompleteFlag) = 10
        .ColWidth(E_FPS1.E_TotalDie) = 10
        .ColWidth(E_FPS1.E_GoodDie) = 10
        .ColWidth(E_FPS1.E_WaferLot) = 10
        .ColWidth(E_FPS1.E_MarkingCode) = 10

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .ReDraw = True

    End With

End Sub

Private Sub Form_Load()
    'IniProduct

    'IniCustomerName

    DTP1.Value = Now - 1

    DTP2.Value = Now

End Sub

Private Sub Label9_Click()

End Sub

