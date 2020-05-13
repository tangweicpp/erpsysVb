VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_TSV_AAMPNData 
   Caption         =   "入库信息查询"
   ClientHeight    =   9975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17535
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
   ScaleHeight     =   9975
   ScaleWidth      =   17535
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdStockExport 
      Caption         =   "仓库导出专用"
      Height          =   480
      Left            =   14520
      TabIndex        =   24
      Top             =   3240
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTP4 
      Height          =   375
      Left            =   3120
      TabIndex        =   23
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111214594
      CurrentDate     =   43825.5
   End
   Begin MSComCtl2.DTPicker DTP3 
      Height          =   375
      Left            =   3120
      TabIndex        =   22
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111214594
      CurrentDate     =   43825.625
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0080FFFF&
      Caption         =   "退出"
      Height          =   480
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "导出"
      Height          =   480
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton CmdQuery 
      BackColor       =   &H00FFFF00&
      Caption         =   "查询"
      Height          =   480
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "客户类型"
      Height          =   1215
      Left            =   480
      TabIndex        =   11
      Top             =   1920
      Width           =   14055
      Begin VB.CheckBox chkStockOnly 
         Caption         =   "只查库存"
         Height          =   255
         Left            =   8760
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton OptCIS 
         Caption         =   "CIS"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton OptBUMPING 
         Caption         =   "BUMPING"
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton OptAA 
         Caption         =   "AA客户"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblLabel4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码"
         Height          =   195
         Left            =   5280
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   720
      End
      Begin MSForms.ComboBox cbCusCode 
         Height          =   375
         Left            =   6360
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3836;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblAddtion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Addtion: 入库单查询请先选好客户"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   2640
      End
   End
   Begin VB.TextBox TxtInBill 
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   1440
      Width           =   3855
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      ItemData        =   "Frm_TSV_AAMPNData.frx":0000
      Left            =   7320
      List            =   "Frm_TSV_AAMPNData.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox TxtBillNo 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111214593
      CurrentDate     =   41424
   End
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111214593
      CurrentDate     =   41424
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   6015
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   3960
      Width           =   21495
      _Version        =   524288
      _ExtentX        =   37915
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
      SpreadDesigner  =   "Frm_TSV_AAMPNData.frx":0022
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "入库单据编号："
      Height          =   195
      Left            =   5880
      TabIndex        =   10
      Top             =   1560
      Width           =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "类型："
      Height          =   195
      Left            =   6720
      TabIndex        =   7
      Top             =   360
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发货单据编号："
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始时间："
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束时间："
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   900
   End
End
Attribute VB_Name = "Frm_TSV_AAMPNData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
    e_ID = 0                'id
    E_ContainerName
    E_QboxName           '箱号
    E_LOTID              'lot
    E_qty                '数量
    E_MPNSeq             'seq
    E_NewLotid           'new lotid
    E_PRODUCT            'product
    E_CustomerProuduct   'CustomerProuduct
    E_BigQbox             'BigBox
    E_QtyInStock          '库存数量
    E_END
    
    
End Enum

Private Enum E_FPS1
    e_NO = 0
    e_date
    e_OrderName
    E_SmallBoxID
    E_name
    E_partno
    E_BatchPlant
    E_LOTID
    E_WAFERID
    E_GrossDie
    E_GoodDie
    E_Pieces
    E_INsiteNgDie
    E_BigBoxID
    E_SecondCode
    E_QtyInStock          '库存数量
    E_END
End Enum
        
Private Enum E_FPS_GC
    e_NO = 0
    e_date '入库时间
    E_Bond '保税或非保
    E_Mpndesc '客户机种
    E_name '品名
    E_partno '料号
    E_BatchPlant 'LOT号后缀
    E_DieByPcs '单片数量=客户设计GoodDie/进厂片数
    E_Pieces '进厂片数
    E_GrossDie '客户设计GoodDie
    E_SecondCode '二级代码
    E_ProductType '形式
    e_ShipNo '出货单号
    E_ShipTo '出货地
    E_BigBoxID '大箱号
    E_SmallBoxID '小箱号
    e_OrderName '工单号
    E_LOTID 'LOT号
    E_WAFERID 'WAFERLIST
    E_GoodDie 'GOODDIE数量
    E_INsiteNgDie '厂内NG
    E_QtyInStock          '库存数量
    E_END
End Enum


Dim reportRS As New ADODB.Recordset
Dim mainItemRS As New ADODB.Recordset
Dim bomRS2        As New ADODB.Recordset




Private Sub CmdOut_Click()
Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String


Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""


woTemp = UCase(Trim(txtWO.text))
productTemp = UCase(Trim(TxtProduct.text))
beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")

sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
          
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
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""
waferIdTemp = UCase(Trim(txtWaferID.text))


'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
'          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
'
          
'  sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & _
'"  CASE ORDERTYPE  WHEN '1' THEN '一般工单'  WHEN '5' THEN '再加工工单'   WHEN '7' THEN '委外工单'   WHEN '8' THEN '重工委外工单' " & _
'" WHEN '11' THEN '拆件式工单'    WHEN '13' THEN '预测工单'   WHEN '15' THEN '试产工单' Else '其他' END as ORDERTYPE ," & _
'"  PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & _
'" PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & _
'" from erpintegration2.ib_wohistory a where  modifyflag='1' "


  sql1 = " select b.wafersequence, b.ordername,b.waferid,b.completeflag,b.dieqty, b.FGDieQty , b.WaferLot, b.MarkingCode " & _
"from erpintegration2.ib_wohistory a, ib_waferlist b  where a.modifyflag='1' and b.ordername=a.ordername and b.waferid='" + waferIdTemp + "' "

          
          
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



Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdquery_Click()

    Dim beginTime   As String

    Dim endTime     As String

    Dim woTemp      As String

    Dim sqlTemp     As String

    Dim sql1        As String

    Dim billNoTemp  As String

    Dim lotIdTemp   As String

    Dim bigQboxTemp As String

    Dim productTemp As String

    Dim date1Temp   As String

    Dim date2Temp   As String

    Dim sql2        As String

    Dim sql3        As String

    Dim sql4        As String

    sql1 = ""
    sql2 = ""
    sql3 = ""
    sqlTemp = ""
    sql4 = ""

    '1. 发货
    If Trim(TxtBillNo.text) <> "" Then
        Call IniFpsHeader1
        billNoTemp = UCase(Trim(TxtBillNo.text))
           
        sqlTemp = " select X.箱号 , X.qboxnumber, X.工单号, Sum(X.数量),  substring(X.箱号,2,9)  as MPN_SEQ, X.newlotid, X.料号, X.MPN, X.箱号2 from ( " & " SELECT distinct b.箱号,   C.qboxnumber, B.工单号,B.流程卡编号,B.数量, C.MPN_SEQ,C.newlotid ,B.料号,d.mpn,f.箱号 as 箱号2 " & " FROM   tblStockMove A ,tblStockMovesub B ,TblQBOXNUMBER_TSVMPN C,dbo.TblQBOXNUMBER_TSV d  ,tblPackTreeInf e ,tblPackTreeInf f " & " Where A.产线标记=1 AND A.客户代码='AA' AND A.单据类型=1 and A.单据编号='" + billNoTemp + "' " & " and b.单据编号=a.单据编号 and b.工单号=a.工单号 and C.qboxnumber=B.箱号  and d.QBOXNUMBER=C.qboxnumber and e.箱号=C.qboxnumber and f.序号=e.上级序号) X " & " group by X.箱号,  X.qboxnumber, X.工单号,X.newlotid ,X.料号,X.mpn,X.箱号2 order by X.箱号2, X.箱号,X.工单号 "
           
        Set mainItemRS = GetAAMPNDataSQL(sqlTemp)
        
        With fps(0)
            .MaxRows = 0

            If mainItemRS.RecordCount > 0 Then
                Set .DataSource = mainItemRS
       
            End If

        End With
        
        Exit Sub

    End If

    ' 2. 入库
  '  If Trim(TxtInBill.Text) <> "" Or OptBUMPING.Value = True Then
        date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
        date2Temp = Format(DTP2.Value, "YYYY-MM-DD")
        
        billNoTemp = UCase(Trim(TxtInBill.text))
        Dim s1 As String
        s1 = """;""CUSTOMER_PN"""

        If OptAA.Value = True Then  ' AA
            If Trim(TxtInBill.text) = "" Then
                MsgBox "请输入入库单号!", vbInformation, "提示"
                Exit Sub
            End If
            Call IniFpsHeader1

             sqlTemp = "select X.入库单编号 , X.箱号, X.WAFERNUMBER, Sum(X.NDPW),  substring(X.箱号,2,9) as MPN_SEQ, X.newlotid, X.productName, X.MPN, X.箱号2  FROM ( " & _
            " SELECT distinct a.入库单编号,c.箱号 ,d.WAFERNUMBER,d.WAFERSCRIBENUMBER,d.NDPW,e.MPN_SEQ,e.newlotid,d.PRODUCTNAME,d.MPN ,b.箱号 as 箱号2 " & _
            " FROM dbo.tblPackToHouseSub a ,tblPackTreeInf b ,tblPackTreeInf c ,dbo.TblQBOXNUMBER_TSV d,TblQBOXNUMBER_TSVMPN e " & _
            " WHERE a.入库单编号='" & billNoTemp & "' and b.箱号=a.箱号 and c.上级序号=b.序号 and d.QBOXNUMBER=c.箱号 and e.qboxnumber=c.箱号 ) X " & _
            " group by X.入库单编号,X.箱号 ,X.WAFERNUMBER,X.newlotid,X.PRODUCTNAME,X.MPN ,X.箱号2 " & _
            " UNION  " & _
            " SELECT xx.入库单编号,xx.箱号 ,xx.工单号 ,xx.qty ,xx.AA_Q ,SUBSTRING(yy.Content,CHARINDEX('CUSTOMER_LOT_COMPLEX"",""',yy.Content) + 23,10 ),xx.PRODUCT,xx.MPN  ,xx.BBOX FROM ( " & _
            " SELECT  x.入库单编号,x.箱号,x.工单号,x.qty,x.AA_Q,x.KEYID ,X.PRODUCT,  X.MPN,X.BBOX, MAX(y.Createdate) AS createdate FROM ( " & _
            "  SELECT a.入库单编号,b.箱号,b.工单号,SUM(b.数量) AS qty , SUBSTRING(b.箱号,2,9)  AS AA_Q  ,c.KEYID ,e.PRODUCT,  e.MPN,aa.箱号 AS BBOX " & _
            "  FROM erpdata..tblPackTreeInf a  ,erpdata..tblPackMainInfSub b ,erpdata..tblErpInStockDetailInfo c ,erpdata..tblTSVworkorder e ,erpdata..tblPackTreeInf aa " & _
            "   WHERE a.入库单编号 = '" & billNoTemp & "'  AND b.箱号 = a.箱号 AND c.KEY_VALUE = b.箱号  AND e.ORDERNAME = b.大工单  AND aa.序号 = a.上级序号 " & _
            "  GROUP BY a.入库单编号,b.箱号,b.工单号,c.KEYID,e.PRODUCT,  e.MPN,aa.箱号 ) X  left JOIN   erpdata..tblME_PrintInfo y ON  y.EVENT_ID = x.KEYID  AND y.LABEL_ID = 'AAMPN4' " & _
            "  GROUP BY x.入库单编号,x.箱号,x.工单号,x.qty,x.AA_Q,x.KEYID,X.PRODUCT,  X.MPN,X.BBOX ) xx " & _
            "  left JOIN erpdata..tblME_PrintInfo yy ON  yy.EVENT_ID = xx.KEYID  AND yy.LABEL_ID = 'AAMPN4' AND yy.Createdate = xx.createdate  ORDER BY X.PRODUCTNAME "
             
 
            Set mainItemRS = GetAAMPNDataSQL(sqlTemp)
            
            With fps(0)
                .MaxRows = 0

                If mainItemRS.RecordCount > 0 Then
                    Set .DataSource = mainItemRS
       
                End If

            End With
            
            Exit Sub

        ElseIf OptCIS.Value = True Then ' CIS
            date1Temp = Format(DTP1.Value, "YYYY-MM-DD") & " " & Format(DTP3.Value, "hh:mm:ss")
            date2Temp = Format(DTP2.Value, "YYYY-MM-DD") & " " & Format(DTP4.Value, "hh:mm:ss")

            Call delQboxNoTw
            Call InsertQboxTemp(billNoTemp)
            If UCase(Trim(cbCusCode.text)) = "GC" Then ' GC
                Call IniFpsHeader_GC
                sqlTemp = " SELECT x.创建时间 AS 入库时间,CASE LEFT(x.大工单,1) WHEN 'A'THEN '保税' ELSE '非保税'END AS 保税或非保,x.Mpn_desc AS 客户机种,x.QTECHPTNO AS 品名 ,x.料号 AS 料号," & _
                        " case CHARINDEX(',', x.WaferId)  when  0 then Rtrim(x.工单号)+'.' +x.WaferId  else Rtrim(x.工单号) +'.' + left(x.WaferId,CHARINDEX(',', x.WaferId)-1) End AS LOT号后缀, " & _
                         " SUM(x.gross) /sum(x.price) AS 单片数量 , sum(x.price) AS 进厂片数,SUM(x.gross) as '客户设计GoodDie', x.IMAGER_CUSTOMER_REV AS 二级代码 , isnull(z.制程,'') AS 形式,'' AS 出货单号,'' AS 出货地,x.大箱 AS 大箱号,x.箱号 AS 小箱号, x.大工单 AS 工单号 ,x.工单号 AS LOT号, x.WaferId AS WAFERLIST  FROM ( " & _
                        " SELECT CONVERT(VARCHAR(20), c.创建时间, 23) AS 创建时间, b.大工单,f.Mpn_desc,  b.箱号, g.QTECHPTNO, b.料号, SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12) AS SFC, b.工单号, " & _
                        "  WaferId = (STUFF((SELECT ',' + SUBSTRING(REPLACE(b1.流程卡编号, '+', ''), len(REPLACE(b1.流程卡编号, '+', '')) - 1, 2)  FROM erpdata .. tblPackMainInfSub b1 " & _
                        "  WHERE b.箱号 = b1.箱号 and  b.大工单 = b1.大工单 order by b1.流程卡编号 FOR XML PATH('')), 1,  1, '')),  e.PASSBINCOUNT + e.FAILBINCOUNT AS gross,  SUM(CONVERT(INT, h1.KEY_VALUE)) AS good_die, " & _
                        "  COUNT(DISTINCT b.流程卡编号) AS price,  SUM(CONVERT(INT, h2.KEY_VALUE)) AS ng_die, a1.箱号 AS 大箱,  f.IMAGER_CUSTOMER_REV, b.流程卡编号 ,isnull(k.数量,'') as 库存 FROM erpdata .. tblPackTreeInf a " & _
                        " INNER JOIN erpdata .. tblPackMainInfSub b  ON b.箱号 = a.箱号 " & _
                        " INNER JOIN ERPBASE .. tblmappingData e ON e.SUBSTRATEID = b.流程卡编号 " & _
                        " INNER JOIN ERPBASE .. tblCustomerOI f  ON  convert(varchar(50),convert(int,f.ID)) = e.FILENAME and f.SOURCE_BATCH_ID = e.LOTID " & _
                        " INNER JOIN erptemp .. tbltsvnpiproduct g ON g.QTECHPTNO2 = b.料号  AND g.CUSTOMERPTNO1 = f.MPN_DESC  AND g.CUSTOMERSHORTNAME = f.CUSTOMERSHORTNAME " & _
                        " INNER JOIN erpdata .. tblErpInStockDetailInfo h  ON h.KEY_VALUE = a.箱号  " & _
                        " INNER JOIN erpdata .. tblErpInStockRelation j ON j.BOX_ID = h.BOX_ID  AND SUBSTRING(REPLACE(j.WAFER_ID, j.SFC_ID, ''),  2, CHARINDEX('::', REPLACE(j.WAFER_ID, j.SFC_ID, '')) - 2) =  b.流程卡编号 " & _
                        " LEFT join erpdata .. tblErpInStockDetailInfo h1  ON h1.BOX_ID = h.BOX_ID  AND h1.KEY_NAME = 'GOOD_DIE' AND h1.KEY_TYPE = 'WAFER' AND H1.KEYID = J.WAFER_ID AND h1.KEY_VALUE <> '0' " & _
                        " LEFT join erpdata .. tblPackTreeInf a1  ON a1.序号 = a.上级序号 " & _
                        " INNER JOIN erpdata .. tblPackMainInf c  ON c.箱号 = a1.箱号 " & _
                        " LEFT join erpdata .. tblErpInStockDetailInfo h2   ON h2.BOX_ID = h.BOX_ID AND h2.KEY_NAME IN ('BAD1_DIE', 'BAD2_DIE') AND h2.KEY_TYPE = 'WAFER' AND H2.KEYID = J.WAFER_ID  AND h2.KEY_VALUE <> '0' "
                        
                        If chkStockOnly.Value = 1 Then '；只查库存
                            sqlTemp = sqlTemp & "  inner JOIN erpdata..Tblstocknumsub k ON b.流程卡编号 =k.流程卡编号   where k.库房编号<>72 "
                            If Trim(TxtInBill.text) <> "" Then
                                sqlTemp = sqlTemp & "  and  f.CUSTOMERSHORTNAME='GC'  and  a.入库单编号 = '" & billNoTemp & " ' "
                            Else
                                sqlTemp = sqlTemp & " and  f.CUSTOMERSHORTNAME='GC'  and  c.创建时间 >= '" & date1Temp & "' and c.创建时间 <= '" & date2Temp & " '  AND isnull(a.入库单编号,'')<>'' "
                            End If
                        Else '所有的
                            sqlTemp = sqlTemp & "  left JOIN erpdata..Tblstocknumsub k ON b.流程卡编号 =k.流程卡编号 and b.箱号 =k.箱号 "
                            If Trim(TxtInBill.text) <> "" Then
                                sqlTemp = sqlTemp & "  where  f.CUSTOMERSHORTNAME='GC'  and  a.入库单编号 = '" & billNoTemp & " ' "
                            Else
                                sqlTemp = sqlTemp & " where  f.CUSTOMERSHORTNAME='GC'  and  c.创建时间 >= '" & date1Temp & "' and c.创建时间 <= '" & date2Temp & " '  AND isnull(a.入库单编号,'')<>'' "
                            End If
                        End If

                        
                        sqlTemp = sqlTemp & " GROUP BY CONVERT(VARCHAR(20), c.创建时间, 23), b.大工单, b. 箱号, g.QTECHPTNO, b.料号,b.流程卡编号,  SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12), b.工单号, e.PASSBINCOUNT + e.FAILBINCOUNT,  a1.箱号,  f.IMAGER_CUSTOMER_REV,f.Mpn_desc,k.数量 ) x " & _
                        " LEFT JOIN  erptemp..GcCode_Reference  z ON  x.Mpn_desc=z.客户机种名 AND x.料号=z.成品料号  AND  (SUBSTRING(x.IMAGER_CUSTOMER_REV,3,1)=z.二级代码 or SUBSTRING(x.IMAGER_CUSTOMER_REV,3,1)=z.分bin二级代码 ) " & _
                        "  GROUP BY  x.创建时间,x.大工单,x.箱号,x.QTECHPTNO,x.料号,x.SFC,x.工单号,x.WaferId ,x.大箱,x.IMAGER_CUSTOMER_REV,x.Mpn_desc,z.制程 ORDER BY x.QTECHPTNO "
', SUM(x.good_die) AS GOODDIE数量 ,SUM(x.ng_die) AS 厂内NG, x.库存
',x.库存

            Else
                Call IniFpsHeader2
               ' sqlTemp = "select waredate,workordername,qboxnumber,alternatename,productname,firstname,wafernumber,TO_CHAR(wmsys.wm_concat(substr(waferscribenumber, -2, 2))) As WaferList," & "sum(gross_dies),sum(gross_dies) - sum(ng_dies),count(waferscribenumber),sum(ng_dies),outpack,imager_customer_rev from cis_in_report group by waredate,workordername," & "qboxnumber,alternatename,productname,firstname,wafernumber,outpack,imager_customer_rev order by wafernumber"
                  
               sqlTemp = " SELECT x.创建时间,x.大工单,x.箱号,x.QTECHPTNO,x.料号," & _
                        " case CHARINDEX(',', x.WaferId)  when  0 then Rtrim(x.工单号)+'.' +x.WaferId  else Rtrim(x.工单号) +'.' + left(x.WaferId,CHARINDEX(',', x.WaferId)-1) End ,x.工单号, " & _
                        " x.WaferId,SUM(x.gross), SUM(x.good_die), sum(x.price),SUM(x.ng_die),x.大箱,x.IMAGER_CUSTOMER_REV  FROM ( " & _
                        " SELECT CONVERT(VARCHAR(20), c.创建时间, 23) AS 创建时间, b.大工单,  b.箱号, g.QTECHPTNO, b.料号, SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12) AS SFC, b.工单号, " & _
                        "  WaferId = (STUFF((SELECT ',' + SUBSTRING(REPLACE(b1.流程卡编号, '+', ''), len(REPLACE(b1.流程卡编号, '+', '')) - 1, 2)  FROM erpdata .. tblPackMainInfSub b1 " & _
                        "  WHERE b.箱号 = b1.箱号 and  b.大工单 = b1.大工单 order by b1.流程卡编号 FOR XML PATH('')), 1,  1, '')),  e.PASSBINCOUNT + e.FAILBINCOUNT AS gross,  SUM(CONVERT(INT, h1.KEY_VALUE)) AS good_die, " & _
                        "  COUNT(DISTINCT b.流程卡编号) AS price,  SUM(CONVERT(INT, h2.KEY_VALUE)) AS ng_die, a1.箱号 AS 大箱,  f.IMAGER_CUSTOMER_REV, b.流程卡编号  , k.数量 as 库存  FROM erpdata .. tblPackTreeInf a " & _
                        " INNER JOIN erpdata .. tblPackMainInfSub b  ON b.箱号 = a.箱号 INNER JOIN erpdata .. tblPackMainInf c  ON c.箱号 = b.箱号 INNER JOIN ERPBASE .. tblmappingData e ON e.SUBSTRATEID = b.流程卡编号 " & _
                        " INNER JOIN ERPBASE .. tblCustomerOI f  ON  convert(varchar(50),convert(int,f.ID)) = e.FILENAME and f.SOURCE_BATCH_ID = e.LOTID INNER JOIN erptemp .. tbltsvnpiproduct g ON g.QTECHPTNO2 = b.料号 " & _
                        "   AND g.CUSTOMERPTNO1 = f.MPN_DESC  AND g.CUSTOMERSHORTNAME = f.CUSTOMERSHORTNAME INNER JOIN erpdata .. tblErpInStockDetailInfo h  ON h.KEY_VALUE = a.箱号 INNER JOIN erpdata .. tblErpInStockRelation j " & _
                        "  ON j.BOX_ID = h.BOX_ID  AND SUBSTRING(REPLACE(j.WAFER_ID, j.SFC_ID, ''),  2, CHARINDEX('::', REPLACE(j.WAFER_ID, j.SFC_ID, '')) - 2) =  b.流程卡编号  LEFT join erpdata .. tblErpInStockDetailInfo h1 " & _
                        "  ON h1.BOX_ID = h.BOX_ID  AND h1.KEY_NAME = 'GOOD_DIE' AND h1.KEY_TYPE = 'WAFER' AND H1.KEYID = J.WAFER_ID AND h1.KEY_VALUE <> '0'  LEFT join erpdata .. tblPackTreeInf a1  ON a1.序号 = a.上级序号 " & _
                        "  LEFT join erpdata .. tblErpInStockDetailInfo h2   ON h2.BOX_ID = h.BOX_ID AND h2.KEY_NAME IN ('BAD1_DIE', 'BAD2_DIE') AND h2.KEY_TYPE = 'WAFER' AND H2.KEYID = J.WAFER_ID  AND h2.KEY_VALUE <> '0' "
                If chkStockOnly.Value = 1 Then '；只查库存
                    sqlTemp = sqlTemp & "  inner JOIN erpdata..Tblstocknumsub k ON b.流程卡编号 =k.流程卡编号  where k.库房编号<>72 "
                    If Trim(TxtInBill.text) <> "" Then
                        sqlTemp = sqlTemp & "  and   a.入库单编号 = '" & billNoTemp & " ' "
                    Else
                        sqlTemp = sqlTemp & "  and  c.创建时间 >= '" & date1Temp & "' and c.创建时间 <= '" & date2Temp & " '  AND isnull(a.入库单编号,'')<>'' "
                    End If
                Else '所有的
                    sqlTemp = sqlTemp & "  left JOIN erpdata..Tblstocknumsub k ON b.流程卡编号 =k.流程卡编号 and  b.箱号 =k.箱号 "
                    If Trim(TxtInBill.text) <> "" Then
                        sqlTemp = sqlTemp & "  Where  a.入库单编号 = '" & billNoTemp & " ' "
                    Else
                        sqlTemp = sqlTemp & "  Where  c.创建时间 >= '" & date1Temp & "' and c.创建时间 <= '" & date2Temp & " '  AND isnull(a.入库单编号,'')<>'' "
                    End If
                                        
                End If

                If Trim(cbCusCode.text) <> "" Then
                    sqlTemp = sqlTemp & " and f.CUSTOMERSHORTNAME='" & Trim(cbCusCode.text) & "' "
                End If
                sqlTemp = sqlTemp & " GROUP BY CONVERT(VARCHAR(20), c.创建时间, 23), b.大工单, b. 箱号, g.QTECHPTNO, b.料号,b.流程卡编号,  SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12), b.工单号, e.PASSBINCOUNT + e.FAILBINCOUNT,  a1.箱号,  f.IMAGER_CUSTOMER_REV , k.数量) x " & _
                        "  GROUP BY  x.创建时间,x.大工单,x.箱号,x.QTECHPTNO,x.料号,x.SFC,x.工单号,x.WaferId ,x.大箱,x.IMAGER_CUSTOMER_REV  order by x.QTECHPTNO"
            End If
                                      
            Set mainItemRS = getSqlStr(sqlTemp)
            
            With fps(0)
                .MaxRows = 0

                If mainItemRS.RecordCount > 0 Then
                    Set .DataSource = mainItemRS
       
                End If

            End With
            
            Exit Sub

        ElseIf OptBUMPING.Value = True Then ' BUMPING
            date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
            date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")

            If billNoTemp = "" Then
                sqlTemp = "select X.入库时间 ,X.入库单编号,X.ORDERNAME as 工单号,X.箱号,X.SALESORDER as 机种,X.PRODUCT as 料号,X.SFC as 生产批号,X.工单号 as 客户LOT,X.流程卡编号 as WAFER_ID,CONVERT(INT, X.DIEQTY)  - CONVERT(INT, Z.KEY_VALUE) as 来料良品,X.KEY_VALUE as 制程良品,Y.KEY_VALUE as 制程不良品,Z.KEY_VALUE as 来料不良" & _
                   ",X.MARKINGCODE as 打标码   from (select CONVERT(varchar(100),aa.入库时间,23) as 入库时间,a.入库单编号,g.ORDERNAME,a.箱号,g.SALESORDER,g.PRODUCT,replace(d.SFC_ID,'SFCBO:1020,','') as SFC,b.工单号 " & _
                   ",b.流程卡编号,e.KEY_VALUE,f.DIEQTY,F.MARKINGCODE,e.KEYID from erpdata..tblPackTreeInf a ,erpdata..tblPackToHouse aa ,erpdata..tblPackMainInfSub b  ,erpdata..tblErpInStockDetailInfo c ,erpdata..tblErpInStockRelation d " & _
                   ",erpdata..tblErpInStockDetailInfo e,erpdata..tblTSVwaferlist f,erpdata..tblTSVworkorder g where aa.入库时间 >= '" & date1Temp & "' and aa.入库时间 <= '" & date2Temp & "' and a.上级序号 <> '0' and aa.入库单编号 = a.入库单编号 and b.箱号 = a.箱号 " & _
                   "and c.KEY_VALUE = b.箱号  and d.BOX_ID = c.BOX_ID and SUBSTRING(replace(d.WAFER_ID,'SFCBO:1020,',''), CHARINDEX(',',replace(d.WAFER_ID,'SFCBO:1020,',''))+1,CHARINDEX('::',replace(d.WAFER_ID,'SFCBO:1020,',''))-CHARINDEX(',',replace(d.WAFER_ID,'SFCBO:1020,',''))-1) = b.流程卡编号 " & _
                   "and e.KEYID = d.WAFER_ID and e.KEY_NAME in ('GOOD_DIE' ) and f.WAFERID = b.流程卡编号 and g.ORDERNAME = f.ORDERNAME and g.ORDERNAME = aa.大工单 ) x,erpdata..tblErpInStockDetailInfo y,erpdata..tblErpInStockDetailInfo z where y.KEYID = x.KEYID " & _
                   "and y.KEY_NAME = 'BAD1_DIE' AND Z.KEYID = Y.KEYID AND Z.KEY_NAME = 'BAD2_DIE' order by x.入库时间,x.入库单编号, x.流程卡编号"
            Else
                Dim rkdID As String
                rkdID = billNoTemp
                 sqlTemp = " SELECT X.入库时间 ,X.入库单编号,X.ORDERNAME AS 工单号,X.箱号, X.HT_DEVICE AS 品名,X.PRODUCT AS 料号,X.SFC AS 厂内批号,X.工单号 AS LOT_ID, " & _
                        " X.流程卡编号 AS WAFER_ID,CONVERT(INT, X.DIEQTY) AS 客户设计GOODDIE,X.KEY_VALUE AS GOODDIE数量,1 AS 良品进站片数, Y.KEY_VALUE AS 厂内NGDIE数量, " & _
                        " Z.KEY_VALUE AS 客户NGDIE数量,X.MARKINGCODE AS 镭射码内容,X.CUSTOMER AS 客户代码 FROM (SELECT CONVERT(VARCHAR(100), aa.入库时间, 23) AS 入库时间, " & _
                        " a.入库单编号, g.ORDERNAME, a.箱号,h.HT_DEVICE,g.PRODUCT,REPLACE(d.SFC_ID, 'SFCBO:1020,', '') AS SFC,b.工单号, b.流程卡编号,e.KEY_VALUE,f.DIEQTY, " & _
                        " F.MARKINGCODE,e.KEYID,g.CUSTOMER FROM erpdata .. tblPackTreeInf a,erpdata .. tblPackToHouse aa,erpdata .. tblPackMainInfSub b, erpdata .. tblErpInStockDetailInfo c, " & _
                        " erpdata .. tblErpInStockRelation d,erpdata .. tblErpInStockDetailInfo e,erpdata .. tblTSVwaferlist f,erpdata .. tblTSVworkorder g,erpdata .. SHOP_ORDER h " & _
                        " WHERE a.上级序号 <> '0'  AND aa.入库单编号 = a.入库单编号 AND b.箱号 = a.箱号 and a.入库单编号 = '" & billNoTemp & "' AND c.KEY_VALUE = b.箱号 AND h.SHOP_ORDER = g.ORDERNAME " & _
                        " AND d.BOX_ID = c.BOX_ID AND SUBSTRING(REPLACE(d.WAFER_ID, 'SFCBO:1020,', ''), CHARINDEX(',',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) + 1, CHARINDEX('::',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) - " & _
                        " CHARINDEX(',',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) - 1) =b.流程卡编号 AND e.KEYID = d.WAFER_ID AND e.KEY_NAME IN ('GOOD_DIE') AND f.WAFERID = b.流程卡编号 AND g.ORDERNAME = f.ORDERNAME) x, " & _
                        " erpdata .. tblErpInStockDetailInfo y, erpdata .. tblErpInStockDetailInfo z WHERE y.KEYID = x.KEYID AND y.KEY_NAME = 'BAD1_DIE' AND Z.KEYID = Y.KEYID AND Z.KEY_NAME = 'BAD2_DIE' "

            End If

            Set mainItemRS = GetAAMPNDataSQL(sqlTemp)
            
            With fps(0)
                .MaxRows = 0

                If mainItemRS.RecordCount > 0 Then
                    Set .DataSource = mainItemRS
       
                End If

            End With
            
            Exit Sub

        Else
            Call IniFpsHeader1
            Exit Sub

        End If
    
   ' End If
    
    ' 3.按多个LOTID合批
    If Trim(TxtInBill.text) = "" And Trim(TxtBillNo.text) = "" Then
        Call IniFpsHeader1
        productTemp = UCase(Trim(CmbType.text))
     
        date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
        date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")
     
        If productTemp = "多个LOTID合批" Then
     
            sqlTemp = "select  X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,sum(X.NDPW), substr(X.QBOXNUMBER,2,9)  as MPN_SEQ,X.NEWLOTID ,X.productname, X.mpn ,''  from ( " & " select  distinct a.CONTAINERNAME,a.QBOXNUMBER,a.WAFERNUMBER,a.WAFERSCRIBENUMBER,a.NDPW,b.MPN_SEQ,b.NEWLOTID ,a.productname, a.mpn ,'' " & " from tsv_qboxnumber_details a ,tsv_qboxnumber_mpn b " & " where a.customername='AA' and a.containername like '%-A%' and a.QBOXNUMBER like 'QM%' " & " and a.create_date>=to_date('" + date1Temp + "','YYYY-MM-DD') and a.create_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD')" & " and b.containername=a.containername ) X " & " group by X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,X.NEWLOTID ,X.productname, X.mpn order by X.containername,X.wafernumber"
     
        Else
     
            sqlTemp = " select  X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,sum(X.NDPW),substr(X.QBOXNUMBER,2,9) as MPN_SEQ,X.NEWLOTID ,X.productname, X.mpn ,''  from (" & " select  distinct a.CONTAINERNAME,a.QBOXNUMBER,a.WAFERNUMBER,a.WAFERSCRIBENUMBER,a.NDPW,b.MPN_SEQ,b.NEWLOTID, a.productname, a.mpn,'' " & " from tsv_qboxnumber_details a ,tsv_qboxnumber_mpn b " & " where a.customername='AA' and a.containername like '%-A%' " & " and a.create_date>=to_date('" + date1Temp + "','YYYY-MM-DD') and a.create_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD')" & " and b.containername=a.containername ) X " & " group by X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,X.NEWLOTID ,X.productname, X.mpn order by X.containername,X.wafernumber"

        End If

        Set mainItemRS = GetAAMPNData(sqlTemp)
        
        With fps(0)
            .MaxRows = 0

            If mainItemRS.RecordCount > 0 Then
                Set .DataSource = mainItemRS
       
            End If

        End With

    End If

End Sub



Private Sub InitCuscode()
Dim rs As New ADODB.Recordset, i As Integer

Set rs.ActiveConnection = SqlConnect
rs.Source = "select distinct 客户代码 from tblxcustomer"
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
cbCusCode.Clear
If rs.RecordCount > 0 Then
    rs.MoveFirst

    For i = 1 To rs.RecordCount
        cbCusCode.AddItem Trim(rs("客户代码"))
        rs.MoveNext
    Next i

End If

End Sub

Private Sub cmdStockExport_Click()
'    SqlServer2ExporToExcel ("select AA.大工单,CC.QTY as 工单数,BB.入库数, AA.出货数 from(select 大工单,SUM(数量) as 出货数 from ERPDATA.. TBLSTOCKMOVESUB where 大工单 is not null group by 大工单)AA, " & _
'" (select 大工单,SUM(入库数) as 入库数 from dbo.tblPackToHouse  group by 大工单) BB, " & _
'" dbo.[tblTSVworkorder] CC Where aa.大工单 = BB.大工单 and AA.大工单 = CC.ORDERNAME and CC.CUSTOMER = '95' and AA.大工单 not like '%-16%'")

    Dim beginTime   As String

    Dim endTime     As String

    Dim woTemp      As String

    Dim sqlTemp     As String

    Dim sql1        As String

    Dim billNoTemp  As String

    Dim lotIdTemp   As String

    Dim bigQboxTemp As String

    Dim productTemp As String

    Dim date1Temp   As String

    Dim date2Temp   As String

    Dim sql2        As String

    Dim sql3        As String

    sql1 = ""
    sql2 = ""
    sql3 = ""
    sqlTemp = ""

    If Trim(TxtBillNo.text) <> "" Then   '发货

        billNoTemp = UCase(Trim(TxtBillNo.text))

        sqlTemp = " select X.箱号 , X.qboxnumber, X.工单号, Sum(X.数量) as 数量, substring(X.箱号,2,9) as MPN_SEQ, X.newlotid, X.料号, X.MPN as OPN, X.箱号2 as 大箱号 from ( " & " SELECT distinct b.箱号,   C.qboxnumber, B.工单号,B.流程卡编号,B.数量, C.MPN_SEQ,C.newlotid ,B.料号,d.mpn,f.箱号 as 箱号2 " & " FROM   tblStockMove A ,tblStockMovesub B ,TblQBOXNUMBER_TSVMPN C,dbo.TblQBOXNUMBER_TSV d  ,tblPackTreeInf e ,tblPackTreeInf f " & " Where A.产线标记=1 AND A.客户代码='AA' AND A.单据类型=1 and A.单据编号='" + billNoTemp + "' " & " and b.单据编号=a.单据编号 and b.工单号=a.工单号 and C.qboxnumber=B.箱号  and d.QBOXNUMBER=C.qboxnumber and e.箱号=C.qboxnumber and f.序号=e.上级序号) X " & " group by X.箱号,  X.qboxnumber, X.工单号,X.newlotid ,X.料号,X.mpn,X.箱号2 order by X.箱号2, X.箱号,X.工单号 "

        SqlServer2ExporToExcel_Trim (sqlTemp)
        Exit Sub

    End If
    If OptCIS.Value = True Then   ' CIS 改为直接由FPS输出EXCEL
        FpsToExcel
        Exit Sub
    End If
    

    If Trim(TxtInBill.text) <> "" Or OptBUMPING.Value = True Then

        billNoTemp = UCase(Trim(TxtInBill.text))

        If OptAA.Value = True Then ' AA
 
             sqlTemp = "select X.入库单编号 , X.箱号, X.WAFERNUMBER, Sum(X.NDPW),  substring(X.箱号,2,9) as MPN_SEQ, X.newlotid, X.productName, X.MPN,x.单重 as  单重 , X.箱号2,  '' as 箱序号,'' as 数量 ,'' as 净重  ,x.重量 as 毛重 ,x.尺寸 as 体积,'' as 日期 FROM ( " & _
            " SELECT distinct a.入库单编号,c.箱号 ,d.WAFERNUMBER,d.WAFERSCRIBENUMBER,d.NDPW,e.MPN_SEQ,e.newlotid,d.PRODUCTNAME,d.MPN ,b.箱号 as 箱号2 ,f.尺寸,f.重量,g.单重 " & _
            " FROM dbo.tblPackToHouseSub a ,tblPackTreeInf b ,tblPackTreeInf c ,dbo.TblQBOXNUMBER_TSV d,TblQBOXNUMBER_TSVMPN e ,erpdata..tblStockNumTree f,erpdata..tblWeight_AA g " & _
            " WHERE a.入库单编号='" & billNoTemp & "' and b.箱号=a.箱号 and c.上级序号=b.序号 and d.QBOXNUMBER=c.箱号 and e.qboxnumber=c.箱号 and  f.箱号=b.箱号 and g.料号=d.productName ) X " & _
            " group by X.入库单编号,X.箱号 ,X.WAFERNUMBER,X.newlotid,X.PRODUCTNAME,X.MPN ,X.箱号2  ,x.尺寸,x.重量,x.单重  " & _
            " UNION  " & _
            " SELECT xx.入库单编号,xx.箱号 ,xx.工单号 ,xx.qty ,xx.AA_Q ,SUBSTRING(yy.Content,CHARINDEX('CUSTOMER_LOT_COMPLEX"",""',yy.Content) + 23,10),xx.PRODUCT,xx.MPN,mm.单重 as  单重   ,xx.BBOX , '' as 箱序号,'' as 数量 ,'' as 净重  ,zz.重量 as 毛重,  zz.尺寸  as 体积,'' as 日期 FROM ( " & _
            " SELECT  x.入库单编号,x.箱号,x.工单号,x.qty,x.AA_Q,x.KEYID ,X.PRODUCT,  X.MPN,X.BBOX, MAX(y.Createdate) AS createdate FROM ( " & _
            "  SELECT a.入库单编号,b.箱号,b.工单号,SUM(b.数量) AS qty , SUBSTRING(b.箱号,2,9)  AS AA_Q  ,c.KEYID ,e.PRODUCT,  e.MPN,aa.箱号 AS BBOX " & _
            "  FROM erpdata..tblPackTreeInf a  ,erpdata..tblPackMainInfSub b ,erpdata..tblErpInStockDetailInfo c ,erpdata..tblTSVworkorder e ,erpdata..tblPackTreeInf aa " & _
            "   WHERE a.入库单编号 = '" & billNoTemp & "'  AND b.箱号 = a.箱号 AND c.KEY_VALUE = b.箱号  AND e.ORDERNAME = b.大工单  AND aa.序号 = a.上级序号 " & _
            "  GROUP BY a.入库单编号,b.箱号,b.工单号,c.KEYID,e.PRODUCT,  e.MPN,aa.箱号 ) X  left JOIN   erpdata..tblME_PrintInfo y ON  y.EVENT_ID = x.KEYID  AND y.LABEL_ID = 'AAMPN4' " & _
            "  GROUP BY x.入库单编号,x.箱号,x.工单号,x.qty,x.AA_Q,x.KEYID,X.PRODUCT,  X.MPN,X.BBOX ) xx " & _
            "  LEFT JOIN  erpdata..tblStockNumTree zz ON xx.BBOX=zz.箱号 " & _
            "  LEFT JOIN  erpdata..tblWeight_AA mm ON xx.PRODUCT=mm.料号 " & _
            "  left JOIN erpdata..tblME_PrintInfo yy ON  yy.EVENT_ID = xx.KEYID  AND yy.LABEL_ID = 'AAMPN4' AND yy.Createdate = xx.createdate  ORDER BY X.PRODUCTNAME , X.箱号2"
             

 
            SqlServer2ExporToExcel_Trim_AA (sqlTemp)

        ' ElseIf OptCIS.Value = True Then   ' CIS
                   
            ' sqlTemp = " SELECT x.创建时间,x.大工单,x.箱号,x.QTECHPTNO,x.料号, " & _
                    ' " case CHARINDEX(',', x.WaferId)  when  0 then Rtrim(x.工单号)+'.' +x.WaferId  else Rtrim(x.工单号) +'.' + left(x.WaferId,CHARINDEX(',', x.WaferId)-1) End ,x.工单号, " & _
                    ' "x.WaferId,SUM(x.gross), SUM(x.good_die), sum(x.price),SUM(x.ng_die),x.大箱,x.IMAGER_CUSTOMER_REV FROM ( " & _
                    ' " SELECT CONVERT(VARCHAR(20), c.创建时间, 23) AS 创建时间, b.大工单,  b.箱号, g.QTECHPTNO, b.料号, SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12) AS SFC, b.工单号, " & _
                    ' "  WaferId = (STUFF((SELECT ',' + SUBSTRING(REPLACE(b1.流程卡编号, '+', ''), len(REPLACE(b1.流程卡编号, '+', '')) - 1, 2)  FROM erpdata .. tblPackMainInfSub b1 " & _
                    ' "  WHERE b.箱号 = b1.箱号 and  b.大工单 = b1.大工单 order by b1.流程卡编号 FOR XML PATH('')), 1,  1, '')),  e.PASSBINCOUNT + e.FAILBINCOUNT AS gross,  SUM(CONVERT(INT, h1.KEY_VALUE)) AS good_die, " & _
                    ' "  COUNT(DISTINCT b.流程卡编号) AS price,  SUM(CONVERT(INT, h2.KEY_VALUE)) AS ng_die, a1.箱号 AS 大箱,  f.IMAGER_CUSTOMER_REV, b.流程卡编号  FROM erpdata .. tblPackTreeInf a " & _
                    ' " INNER JOIN erpdata .. tblPackMainInfSub b  ON b.箱号 = a.箱号 INNER JOIN erpdata .. tblPackMainInf c  ON c.箱号 = b.箱号 INNER JOIN ERPBASE .. tblmappingData e ON e.SUBSTRATEID = b.流程卡编号 " & _
                    ' " INNER JOIN ERPBASE .. tblCustomerOI f  ON  convert(varchar(50),convert(int,f.ID)) = e.FILENAME and f.SOURCE_BATCH_ID = e.LOTID INNER JOIN erptemp .. tbltsvnpiproduct g ON g.QTECHPTNO2 = b.料号 " & _
                    ' "   AND g.CUSTOMERPTNO1 = f.MPN_DESC  AND g.CUSTOMERSHORTNAME = f.CUSTOMERSHORTNAME INNER JOIN erpdata .. tblErpInStockDetailInfo h  ON h.KEY_VALUE = a.箱号 INNER JOIN erpdata .. tblErpInStockRelation j " & _
                    ' "  ON j.BOX_ID = h.BOX_ID  AND SUBSTRING(REPLACE(j.WAFER_ID, j.SFC_ID, ''),  2, CHARINDEX('::', REPLACE(j.WAFER_ID, j.SFC_ID, '')) - 2) =  b.流程卡编号  LEFT join erpdata .. tblErpInStockDetailInfo h1 " & _
                    ' "  ON h1.BOX_ID = h.BOX_ID  AND h1.KEY_NAME = 'GOOD_DIE' AND h1.KEY_TYPE = 'WAFER' AND H1.KEYID = J.WAFER_ID AND h1.KEY_VALUE <> '0'  LEFT join erpdata .. tblPackTreeInf a1  ON a1.序号 = a.上级序号 " & _
                    ' "  LEFT join erpdata .. tblErpInStockDetailInfo h2   ON h2.BOX_ID = h.BOX_ID AND h2.KEY_NAME IN ('BAD1_DIE', 'BAD2_DIE') AND h2.KEY_TYPE = 'WAFER' AND H2.KEYID = J.WAFER_ID  AND h2.KEY_VALUE <> '0' " & _
                    ' "  WHERE a.入库单编号 = '" & billNoTemp & " ' " & _
                    ' " GROUP BY CONVERT(VARCHAR(20), c.创建时间, 23), b.大工单, b. 箱号, g.QTECHPTNO, b.料号,b.流程卡编号,  SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12), b.工单号, e.PASSBINCOUNT + e.FAILBINCOUNT,  a1.箱号,  f.IMAGER_CUSTOMER_REV ) x " & _
                    ' "  GROUP BY  x.创建时间,x.大工单,x.箱号,x.QTECHPTNO,x.料号,x.SFC,x.工单号,x.WaferId ,x.大箱,x.IMAGER_CUSTOMER_REV"
                                  
           ' ' ExporToExcel (sqlTemp)
           ' SqlServer2ExporToExcel_Trim (sqlTemp)

        
        ElseIf OptBUMPING.Value = True Then ' BUMPING
    
            date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
            date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")

            If TxtInBill.text = "" Then
                sqlTemp = "select X.入库时间 ,X.入库单编号,X.ORDERNAME as 工单号,X.箱号,X.SALESORDER as 机种,X.PRODUCT as 料号,X.SFC as 生产批号,X.工单号 as 客户LOT,X.流程卡编号 as WAFER_ID,CONVERT(INT, X.DIEQTY)  - CONVERT(INT, Z.KEY_VALUE) as 来料良品,X.KEY_VALUE as 制程良品,Y.KEY_VALUE as 制程不良品,Z.KEY_VALUE as 来料不良" & _
                   ",X.MARKINGCODE as 打标码   from (select CONVERT(varchar(100),aa.入库时间,23) as 入库时间,a.入库单编号,g.ORDERNAME,a.箱号,g.SALESORDER,g.PRODUCT,replace(d.SFC_ID,'SFCBO:1020,','') as SFC,b.工单号 " & _
                   ",b.流程卡编号,e.KEY_VALUE,f.DIEQTY,F.MARKINGCODE,e.KEYID from erpdata..tblPackTreeInf a ,erpdata..tblPackToHouse aa ,erpdata..tblPackMainInfSub b  ,erpdata..tblErpInStockDetailInfo c ,erpdata..tblErpInStockRelation d " & _
                   ",erpdata..tblErpInStockDetailInfo e,erpdata..tblTSVwaferlist f,erpdata..tblTSVworkorder g where aa.入库时间 >= '" & date1Temp & "' and aa.入库时间 <= '" & date2Temp & "' and a.上级序号 <> '0' and aa.入库单编号 = a.入库单编号 and b.箱号 = a.箱号 " & _
                   "and c.KEY_VALUE = b.箱号  and d.BOX_ID = c.BOX_ID and SUBSTRING(replace(d.WAFER_ID,'SFCBO:1020,',''), CHARINDEX(',',replace(d.WAFER_ID,'SFCBO:1020,',''))+1,CHARINDEX('::',replace(d.WAFER_ID,'SFCBO:1020,',''))-CHARINDEX(',',replace(d.WAFER_ID,'SFCBO:1020,',''))-1) = b.流程卡编号 " & _
                   "and e.KEYID = d.WAFER_ID and e.KEY_NAME in ('GOOD_DIE' ) and f.WAFERID = b.流程卡编号 and g.ORDERNAME = f.ORDERNAME and g.ORDERNAME = aa.大工单 ) x,erpdata..tblErpInStockDetailInfo y,erpdata..tblErpInStockDetailInfo z where y.KEYID = x.KEYID " & _
                   "and y.KEY_NAME = 'BAD1_DIE' AND Z.KEYID = Y.KEYID AND Z.KEY_NAME = 'BAD2_DIE' order by x.入库时间,x.入库单编号, x.流程卡编号"
            Else
                Dim rkdID As String
                    
                rkdID = Trim$(TxtInBill.text)
            
                sqlTemp = " SELECT X.入库时间 ,X.入库单编号,X.ORDERNAME AS 工单号,X.箱号, X.HT_DEVICE AS 品名,X.PRODUCT AS 料号,X.SFC AS 厂内批号,X.工单号 AS LOT_ID, " & _
                        " X.流程卡编号 AS WAFER_ID,CONVERT(INT, X.DIEQTY) AS 客户设计GOODDIE,X.KEY_VALUE AS GOODDIE数量,1 AS 良品进站片数, Y.KEY_VALUE AS 厂内NGDIE数量, " & _
                        " Z.KEY_VALUE AS 客户NGDIE数量,X.MARKINGCODE AS 镭射码内容,X.CUSTOMER AS 客户代码 FROM (SELECT CONVERT(VARCHAR(100), aa.入库时间, 23) AS 入库时间, " & _
                        " a.入库单编号, g.ORDERNAME, a.箱号,h.HT_DEVICE,g.PRODUCT,REPLACE(d.SFC_ID, 'SFCBO:1020,', '') AS SFC,b.工单号, b.流程卡编号,e.KEY_VALUE,f.DIEQTY, " & _
                        " F.MARKINGCODE,e.KEYID,g.CUSTOMER FROM erpdata .. tblPackTreeInf a,erpdata .. tblPackToHouse aa,erpdata .. tblPackMainInfSub b, erpdata .. tblErpInStockDetailInfo c, " & _
                        " erpdata .. tblErpInStockRelation d,erpdata .. tblErpInStockDetailInfo e,erpdata .. tblTSVwaferlist f,erpdata .. tblTSVworkorder g,erpdata .. SHOP_ORDER h " & _
                        " WHERE a.上级序号 <> '0'  AND aa.入库单编号 = a.入库单编号 AND b.箱号 = a.箱号 and a.入库单编号 = '" & billNoTemp & "' AND c.KEY_VALUE = b.箱号 AND h.SHOP_ORDER = g.ORDERNAME " & _
                        " AND d.BOX_ID = c.BOX_ID AND SUBSTRING(REPLACE(d.WAFER_ID, 'SFCBO:1020,', ''), CHARINDEX(',',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) + 1, CHARINDEX('::',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) - " & _
                        " CHARINDEX(',',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) - 1) =b.流程卡编号 AND e.KEYID = d.WAFER_ID AND e.KEY_NAME IN ('GOOD_DIE') AND f.WAFERID = b.流程卡编号 AND g.ORDERNAME = f.ORDERNAME) x, " & _
                        " erpdata .. tblErpInStockDetailInfo y, erpdata .. tblErpInStockDetailInfo z WHERE y.KEYID = x.KEYID AND y.KEY_NAME = 'BAD1_DIE' AND Z.KEYID = Y.KEYID AND Z.KEY_NAME = 'BAD2_DIE' "

            End If

            SqlServer2ExporToExcel_Trim (sqlTemp)
        
        Else
        
        End If
  
    End If
    
    If Trim(TxtBillNo.text) = "" And Trim(TxtInBill.text) = "" Then

        productTemp = UCase(Trim(CmbType.text))
     
        date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
        date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")
     
        If productTemp = "多个LOTID合批" Then
     
            sqlTemp = "select  X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,sum(X.NDPW) as 数量,substr(X.QBOXNUMBER,2,9) as MPN_SEQ,X.NEWLOTID ,X.productname, X.mpn  as OPN ,'' as 大箱号  from ( " & " select  distinct a.CONTAINERNAME,a.QBOXNUMBER,a.WAFERNUMBER,a.WAFERSCRIBENUMBER,a.NDPW,b.MPN_SEQ,b.NEWLOTID ,a.productname, a.mpn ,'' " & " from tsv_qboxnumber_details a ,tsv_qboxnumber_mpn b " & " where a.customername='AA' and a.containername like '%-A%' and a.QBOXNUMBER like 'QM%' " & " and a.create_date>=to_date('" + date1Temp + "','YYYY-MM-DD') and a.create_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD')" & " and b.containername=a.containername ) X " & " group by X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,X.NEWLOTID ,X.productname, X.mpn order by X.containername,X.wafernumber"
     
        Else

            sqlTemp = " select  X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,sum(X.NDPW) as 数量,substr(X.QBOXNUMBER,2,9) as MPN_SEQ,X.NEWLOTID ,X.productname, X.mpn as OPN ,'' as 大箱号  from (" & " select  distinct a.CONTAINERNAME,a.QBOXNUMBER,a.WAFERNUMBER,a.WAFERSCRIBENUMBER,a.NDPW,b.MPN_SEQ,b.NEWLOTID, a.productname, a.mpn,'' " & " from tsv_qboxnumber_details a ,tsv_qboxnumber_mpn b " & " where a.customername='AA' and a.containername like '%-A%' " & " and a.create_date>=to_date('" + date1Temp + "','YYYY-MM-DD') and a.create_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD')" & " and b.containername=a.containername ) X " & " group by X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,X.NEWLOTID ,X.productname, X.mpn order by X.containername,X.wafernumber"
          
        End If
           
        ExporToExcel (sqlTemp)

    End If

End Sub

Private Sub Command2_Click()

'    SqlServer2ExporToExcel ("select AA.大工单,CC.QTY as 工单数,BB.入库数, AA.出货数 from(select 大工单,SUM(数量) as 出货数 from ERPDATA.. TBLSTOCKMOVESUB where 大工单 is not null group by 大工单)AA, " & _
'" (select 大工单,SUM(入库数) as 入库数 from dbo.tblPackToHouse  group by 大工单) BB, " & _
'" dbo.[tblTSVworkorder] CC Where aa.大工单 = BB.大工单 and AA.大工单 = CC.ORDERNAME and CC.CUSTOMER = '95' and AA.大工单 not like '%-16%'")

    Dim beginTime   As String

    Dim endTime     As String

    Dim woTemp      As String

    Dim sqlTemp     As String

    Dim sql1        As String

    Dim billNoTemp  As String

    Dim lotIdTemp   As String

    Dim bigQboxTemp As String

    Dim productTemp As String

    Dim date1Temp   As String

    Dim date2Temp   As String

    Dim sql2        As String

    Dim sql3        As String

    sql1 = ""
    sql2 = ""
    sql3 = ""
    sqlTemp = ""

    If Trim(TxtBillNo.text) <> "" Then   '发货

        billNoTemp = UCase(Trim(TxtBillNo.text))

        sqlTemp = " select X.箱号 , X.qboxnumber, X.工单号, Sum(X.数量) as 数量, substring(X.箱号,2,9) as MPN_SEQ, X.newlotid, X.料号, X.MPN as OPN, X.箱号2 as 大箱号 from ( " & " SELECT distinct b.箱号,   C.qboxnumber, B.工单号,B.流程卡编号,B.数量, C.MPN_SEQ,C.newlotid ,B.料号,d.mpn,f.箱号 as 箱号2 " & " FROM   tblStockMove A ,tblStockMovesub B ,TblQBOXNUMBER_TSVMPN C,dbo.TblQBOXNUMBER_TSV d  ,tblPackTreeInf e ,tblPackTreeInf f " & " Where A.产线标记=1 AND A.客户代码='AA' AND A.单据类型=1 and A.单据编号='" + billNoTemp + "' " & " and b.单据编号=a.单据编号 and b.工单号=a.工单号 and C.qboxnumber=B.箱号  and d.QBOXNUMBER=C.qboxnumber and e.箱号=C.qboxnumber and f.序号=e.上级序号) X " & " group by X.箱号,  X.qboxnumber, X.工单号,X.newlotid ,X.料号,X.mpn,X.箱号2 order by X.箱号2, X.箱号,X.工单号 "

        SqlServer2ExporToExcel_Trim (sqlTemp)
        Exit Sub

    End If
    If OptCIS.Value = True Then   ' CIS 改为直接由FPS输出EXCEL
        FpsToExcel
        Exit Sub
    End If
    

    If Trim(TxtInBill.text) <> "" Or OptBUMPING.Value = True Then

        billNoTemp = UCase(Trim(TxtInBill.text))

        If OptAA.Value = True Then ' AA
 
             sqlTemp = "select X.入库单编号 , X.箱号, X.WAFERNUMBER, Sum(X.NDPW),  substring(X.箱号,2,9) as MPN_SEQ, X.newlotid, X.productName, X.MPN, X.箱号2  FROM ( " & _
            " SELECT distinct a.入库单编号,c.箱号 ,d.WAFERNUMBER,d.WAFERSCRIBENUMBER,d.NDPW,e.MPN_SEQ,e.newlotid,d.PRODUCTNAME,d.MPN ,b.箱号 as 箱号2 " & _
            " FROM dbo.tblPackToHouseSub a ,tblPackTreeInf b ,tblPackTreeInf c ,dbo.TblQBOXNUMBER_TSV d,TblQBOXNUMBER_TSVMPN e " & _
            " WHERE a.入库单编号='" & billNoTemp & "' and b.箱号=a.箱号 and c.上级序号=b.序号 and d.QBOXNUMBER=c.箱号 and e.qboxnumber=c.箱号 ) X " & _
            " group by X.入库单编号,X.箱号 ,X.WAFERNUMBER,X.newlotid,X.PRODUCTNAME,X.MPN ,X.箱号2 " & _
            " UNION  " & _
            " SELECT xx.入库单编号,xx.箱号 ,xx.工单号 ,xx.qty ,xx.AA_Q ,SUBSTRING(yy.Content,CHARINDEX('CUSTOMER_LOT_COMPLEX"",""',yy.Content) + 23,10),xx.PRODUCT,xx.MPN  ,xx.BBOX FROM ( " & _
            " SELECT  x.入库单编号,x.箱号,x.工单号,x.qty,x.AA_Q,x.KEYID ,X.PRODUCT,  X.MPN,X.BBOX, MAX(y.Createdate) AS createdate FROM ( " & _
            "  SELECT a.入库单编号,b.箱号,b.工单号,SUM(b.数量) AS qty , SUBSTRING(b.箱号,2,9)  AS AA_Q  ,c.KEYID ,e.PRODUCT,  e.MPN,aa.箱号 AS BBOX " & _
            "  FROM erpdata..tblPackTreeInf a  ,erpdata..tblPackMainInfSub b ,erpdata..tblErpInStockDetailInfo c ,erpdata..tblTSVworkorder e ,erpdata..tblPackTreeInf aa " & _
            "   WHERE a.入库单编号 = '" & billNoTemp & "'  AND b.箱号 = a.箱号 AND c.KEY_VALUE = b.箱号  AND e.ORDERNAME = b.大工单  AND aa.序号 = a.上级序号 " & _
            "  GROUP BY a.入库单编号,b.箱号,b.工单号,c.KEYID,e.PRODUCT,  e.MPN,aa.箱号 ) X  left JOIN   erpdata..tblME_PrintInfo y ON  y.EVENT_ID = x.KEYID  AND y.LABEL_ID = 'AAMPN4' " & _
            "  GROUP BY x.入库单编号,x.箱号,x.工单号,x.qty,x.AA_Q,x.KEYID,X.PRODUCT,  X.MPN,X.BBOX ) xx " & _
            "  left JOIN erpdata..tblME_PrintInfo yy ON  yy.EVENT_ID = xx.KEYID  AND yy.LABEL_ID = 'AAMPN4' AND yy.Createdate = xx.createdate  ORDER BY X.PRODUCTNAME "
             

 
            SqlServer2ExporToExcel_Trim (sqlTemp)

        ' ElseIf OptCIS.Value = True Then   ' CIS
                   
            ' sqlTemp = " SELECT x.创建时间,x.大工单,x.箱号,x.QTECHPTNO,x.料号, " & _
                    ' " case CHARINDEX(',', x.WaferId)  when  0 then Rtrim(x.工单号)+'.' +x.WaferId  else Rtrim(x.工单号) +'.' + left(x.WaferId,CHARINDEX(',', x.WaferId)-1) End ,x.工单号, " & _
                    ' "x.WaferId,SUM(x.gross), SUM(x.good_die), sum(x.price),SUM(x.ng_die),x.大箱,x.IMAGER_CUSTOMER_REV FROM ( " & _
                    ' " SELECT CONVERT(VARCHAR(20), c.创建时间, 23) AS 创建时间, b.大工单,  b.箱号, g.QTECHPTNO, b.料号, SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12) AS SFC, b.工单号, " & _
                    ' "  WaferId = (STUFF((SELECT ',' + SUBSTRING(REPLACE(b1.流程卡编号, '+', ''), len(REPLACE(b1.流程卡编号, '+', '')) - 1, 2)  FROM erpdata .. tblPackMainInfSub b1 " & _
                    ' "  WHERE b.箱号 = b1.箱号 and  b.大工单 = b1.大工单 order by b1.流程卡编号 FOR XML PATH('')), 1,  1, '')),  e.PASSBINCOUNT + e.FAILBINCOUNT AS gross,  SUM(CONVERT(INT, h1.KEY_VALUE)) AS good_die, " & _
                    ' "  COUNT(DISTINCT b.流程卡编号) AS price,  SUM(CONVERT(INT, h2.KEY_VALUE)) AS ng_die, a1.箱号 AS 大箱,  f.IMAGER_CUSTOMER_REV, b.流程卡编号  FROM erpdata .. tblPackTreeInf a " & _
                    ' " INNER JOIN erpdata .. tblPackMainInfSub b  ON b.箱号 = a.箱号 INNER JOIN erpdata .. tblPackMainInf c  ON c.箱号 = b.箱号 INNER JOIN ERPBASE .. tblmappingData e ON e.SUBSTRATEID = b.流程卡编号 " & _
                    ' " INNER JOIN ERPBASE .. tblCustomerOI f  ON  convert(varchar(50),convert(int,f.ID)) = e.FILENAME and f.SOURCE_BATCH_ID = e.LOTID INNER JOIN erptemp .. tbltsvnpiproduct g ON g.QTECHPTNO2 = b.料号 " & _
                    ' "   AND g.CUSTOMERPTNO1 = f.MPN_DESC  AND g.CUSTOMERSHORTNAME = f.CUSTOMERSHORTNAME INNER JOIN erpdata .. tblErpInStockDetailInfo h  ON h.KEY_VALUE = a.箱号 INNER JOIN erpdata .. tblErpInStockRelation j " & _
                    ' "  ON j.BOX_ID = h.BOX_ID  AND SUBSTRING(REPLACE(j.WAFER_ID, j.SFC_ID, ''),  2, CHARINDEX('::', REPLACE(j.WAFER_ID, j.SFC_ID, '')) - 2) =  b.流程卡编号  LEFT join erpdata .. tblErpInStockDetailInfo h1 " & _
                    ' "  ON h1.BOX_ID = h.BOX_ID  AND h1.KEY_NAME = 'GOOD_DIE' AND h1.KEY_TYPE = 'WAFER' AND H1.KEYID = J.WAFER_ID AND h1.KEY_VALUE <> '0'  LEFT join erpdata .. tblPackTreeInf a1  ON a1.序号 = a.上级序号 " & _
                    ' "  LEFT join erpdata .. tblErpInStockDetailInfo h2   ON h2.BOX_ID = h.BOX_ID AND h2.KEY_NAME IN ('BAD1_DIE', 'BAD2_DIE') AND h2.KEY_TYPE = 'WAFER' AND H2.KEYID = J.WAFER_ID  AND h2.KEY_VALUE <> '0' " & _
                    ' "  WHERE a.入库单编号 = '" & billNoTemp & " ' " & _
                    ' " GROUP BY CONVERT(VARCHAR(20), c.创建时间, 23), b.大工单, b. 箱号, g.QTECHPTNO, b.料号,b.流程卡编号,  SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12), b.工单号, e.PASSBINCOUNT + e.FAILBINCOUNT,  a1.箱号,  f.IMAGER_CUSTOMER_REV ) x " & _
                    ' "  GROUP BY  x.创建时间,x.大工单,x.箱号,x.QTECHPTNO,x.料号,x.SFC,x.工单号,x.WaferId ,x.大箱,x.IMAGER_CUSTOMER_REV"
                                  
           ' ' ExporToExcel (sqlTemp)
           ' SqlServer2ExporToExcel_Trim (sqlTemp)

        
        ElseIf OptBUMPING.Value = True Then ' BUMPING
    
            date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
            date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")

            If TxtInBill.text = "" Then
                sqlTemp = "select X.入库时间 ,X.入库单编号,X.ORDERNAME as 工单号,X.箱号,X.SALESORDER as 机种,X.PRODUCT as 料号,X.SFC as 生产批号,X.工单号 as 客户LOT,X.流程卡编号 as WAFER_ID,CONVERT(INT, X.DIEQTY)  - CONVERT(INT, Z.KEY_VALUE) as 来料良品,X.KEY_VALUE as 制程良品,Y.KEY_VALUE as 制程不良品,Z.KEY_VALUE as 来料不良" & _
                   ",X.MARKINGCODE as 打标码   from (select CONVERT(varchar(100),aa.入库时间,23) as 入库时间,a.入库单编号,g.ORDERNAME,a.箱号,g.SALESORDER,g.PRODUCT,replace(d.SFC_ID,'SFCBO:1020,','') as SFC,b.工单号 " & _
                   ",b.流程卡编号,e.KEY_VALUE,f.DIEQTY,F.MARKINGCODE,e.KEYID from erpdata..tblPackTreeInf a ,erpdata..tblPackToHouse aa ,erpdata..tblPackMainInfSub b  ,erpdata..tblErpInStockDetailInfo c ,erpdata..tblErpInStockRelation d " & _
                   ",erpdata..tblErpInStockDetailInfo e,erpdata..tblTSVwaferlist f,erpdata..tblTSVworkorder g where aa.入库时间 >= '" & date1Temp & "' and aa.入库时间 <= '" & date2Temp & "' and a.上级序号 <> '0' and aa.入库单编号 = a.入库单编号 and b.箱号 = a.箱号 " & _
                   "and c.KEY_VALUE = b.箱号  and d.BOX_ID = c.BOX_ID and SUBSTRING(replace(d.WAFER_ID,'SFCBO:1020,',''), CHARINDEX(',',replace(d.WAFER_ID,'SFCBO:1020,',''))+1,CHARINDEX('::',replace(d.WAFER_ID,'SFCBO:1020,',''))-CHARINDEX(',',replace(d.WAFER_ID,'SFCBO:1020,',''))-1) = b.流程卡编号 " & _
                   "and e.KEYID = d.WAFER_ID and e.KEY_NAME in ('GOOD_DIE' ) and f.WAFERID = b.流程卡编号 and g.ORDERNAME = f.ORDERNAME and g.ORDERNAME = aa.大工单 ) x,erpdata..tblErpInStockDetailInfo y,erpdata..tblErpInStockDetailInfo z where y.KEYID = x.KEYID " & _
                   "and y.KEY_NAME = 'BAD1_DIE' AND Z.KEYID = Y.KEYID AND Z.KEY_NAME = 'BAD2_DIE' order by x.入库时间,x.入库单编号, x.流程卡编号"
            Else
                Dim rkdID As String
                    
                rkdID = Trim$(TxtInBill.text)
            
                sqlTemp = " SELECT X.入库时间 ,X.入库单编号,X.ORDERNAME AS 工单号,X.箱号, X.HT_DEVICE AS 品名,X.PRODUCT AS 料号,X.SFC AS 厂内批号,X.工单号 AS LOT_ID, " & _
                        " X.流程卡编号 AS WAFER_ID,CONVERT(INT, X.DIEQTY) AS 客户设计GOODDIE,X.KEY_VALUE AS GOODDIE数量,1 AS 良品进站片数, Y.KEY_VALUE AS 厂内NGDIE数量, " & _
                        " Z.KEY_VALUE AS 客户NGDIE数量,X.MARKINGCODE AS 镭射码内容,X.CUSTOMER AS 客户代码 FROM (SELECT CONVERT(VARCHAR(100), aa.入库时间, 23) AS 入库时间, " & _
                        " a.入库单编号, g.ORDERNAME, a.箱号,h.HT_DEVICE,g.PRODUCT,REPLACE(d.SFC_ID, 'SFCBO:1020,', '') AS SFC,b.工单号, b.流程卡编号,e.KEY_VALUE,f.DIEQTY, " & _
                        " F.MARKINGCODE,e.KEYID,g.CUSTOMER FROM erpdata .. tblPackTreeInf a,erpdata .. tblPackToHouse aa,erpdata .. tblPackMainInfSub b, erpdata .. tblErpInStockDetailInfo c, " & _
                        " erpdata .. tblErpInStockRelation d,erpdata .. tblErpInStockDetailInfo e,erpdata .. tblTSVwaferlist f,erpdata .. tblTSVworkorder g,erpdata .. SHOP_ORDER h " & _
                        " WHERE a.上级序号 <> '0'  AND aa.入库单编号 = a.入库单编号 AND b.箱号 = a.箱号 and a.入库单编号 = '" & billNoTemp & "' AND c.KEY_VALUE = b.箱号 AND h.SHOP_ORDER = g.ORDERNAME " & _
                        " AND d.BOX_ID = c.BOX_ID AND SUBSTRING(REPLACE(d.WAFER_ID, 'SFCBO:1020,', ''), CHARINDEX(',',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) + 1, CHARINDEX('::',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) - " & _
                        " CHARINDEX(',',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) - 1) =b.流程卡编号 AND e.KEYID = d.WAFER_ID AND e.KEY_NAME IN ('GOOD_DIE') AND f.WAFERID = b.流程卡编号 AND g.ORDERNAME = f.ORDERNAME) x, " & _
                        " erpdata .. tblErpInStockDetailInfo y, erpdata .. tblErpInStockDetailInfo z WHERE y.KEYID = x.KEYID AND y.KEY_NAME = 'BAD1_DIE' AND Z.KEYID = Y.KEYID AND Z.KEY_NAME = 'BAD2_DIE' "

            End If

            SqlServer2ExporToExcel_Trim (sqlTemp)
        
        Else
        
        End If
  
    End If
    
    If Trim(TxtBillNo.text) = "" And Trim(TxtInBill.text) = "" Then

        productTemp = UCase(Trim(CmbType.text))
     
        date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
        date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")
     
        If productTemp = "多个LOTID合批" Then
     
            sqlTemp = "select  X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,sum(X.NDPW) as 数量,substr(X.QBOXNUMBER,2,9) as MPN_SEQ,X.NEWLOTID ,X.productname, X.mpn  as OPN ,'' as 大箱号  from ( " & " select  distinct a.CONTAINERNAME,a.QBOXNUMBER,a.WAFERNUMBER,a.WAFERSCRIBENUMBER,a.NDPW,b.MPN_SEQ,b.NEWLOTID ,a.productname, a.mpn ,'' " & " from tsv_qboxnumber_details a ,tsv_qboxnumber_mpn b " & " where a.customername='AA' and a.containername like '%-A%' and a.QBOXNUMBER like 'QM%' " & " and a.create_date>=to_date('" + date1Temp + "','YYYY-MM-DD') and a.create_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD')" & " and b.containername=a.containername ) X " & " group by X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,X.NEWLOTID ,X.productname, X.mpn order by X.containername,X.wafernumber"
     
        Else

            sqlTemp = " select  X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,sum(X.NDPW) as 数量,substr(X.QBOXNUMBER,2,9) as MPN_SEQ,X.NEWLOTID ,X.productname, X.mpn as OPN ,'' as 大箱号  from (" & " select  distinct a.CONTAINERNAME,a.QBOXNUMBER,a.WAFERNUMBER,a.WAFERSCRIBENUMBER,a.NDPW,b.MPN_SEQ,b.NEWLOTID, a.productname, a.mpn,'' " & " from tsv_qboxnumber_details a ,tsv_qboxnumber_mpn b " & " where a.customername='AA' and a.containername like '%-A%' " & " and a.create_date>=to_date('" + date1Temp + "','YYYY-MM-DD') and a.create_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD')" & " and b.containername=a.containername ) X " & " group by X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,X.NEWLOTID ,X.productname, X.mpn order by X.containername,X.wafernumber"
          
        End If
           
        ExporToExcel (sqlTemp)

    End If

End Sub

Private Sub ComOutLine_Click()

Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""


woTemp = UCase(Trim(txtWO.text))
productTemp = UCase(Trim(TxtProduct.text))
beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")

'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
'          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
'
          
  sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & _
"  CASE ORDERTYPE  WHEN '1' THEN '一般工单'  WHEN '5' THEN '再加工工单'   WHEN '7' THEN '委外工单'   WHEN '8' THEN '重工委外工单' " & _
" WHEN '11' THEN '拆件式工单'    WHEN '13' THEN '预测工单'   WHEN '15' THEN '试产工单' Else '其他' END as ORDERTYPE ," & _
"  PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & _
" PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & _
" from erpintegration2.ib_wohistory a where  modifyflag='1' "
          
          
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

Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""


woTemp = UCase(Trim(txtWO.text))
productTemp = UCase(Trim(TxtProduct.text))
beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")

'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
'          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
'
          
  sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & _
"  CASE ORDERTYPE  WHEN '1' THEN '一般工单'  WHEN '5' THEN '再加工工单'   WHEN '7' THEN '委外工单'   WHEN '8' THEN '重工委外工单' " & _
" WHEN '11' THEN '拆件式工单'    WHEN '13' THEN '预测工单'   WHEN '15' THEN '试产工单' Else '其他' END as ORDERTYPE ," & _
"  PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & _
" PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & _
" from erpintegration2.ib_wohistory a where  modifyflag='1' "
          
          
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


Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""


woTemp = UCase(Trim(txtWO.text))
productTemp = UCase(Trim(TxtProduct.text))
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


  sql1 = " select b.wafersequence, b.ordername,b.waferid,b.completeflag,b.dieqty, b.FGDieQty , b.WaferLot, b.MarkingCode " & _
"from erpintegration2.ib_wohistory a, ib_waferlist b  where a.modifyflag='1' and b.ordername=a.ordername"

          
          
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

Private Sub Form_Activate()
'CmbLine.Text = "TSV"

TxtBillNo.SetFocus

OptAA.Value = True
DTP1.Value = Now - 1

DTP2.Value = Now

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


Private Sub TabStrip1_Click()

End Sub

Private Sub IniFpsHeader1()
    With fps(0)
        .TypeMaxEditLen = 500
        .ReDraw = False
        .MaxCols = E_FPS0.E_END - 1
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
        

        .SetText E_FPS0.e_ID, 0, "序号"
        .SetText E_FPS0.E_ContainerName, 0, "主批号"
        .SetText E_FPS0.E_QboxName, 0, "小箱号"
        .SetText E_FPS0.E_LOTID, 0, "LotID"
        .SetText E_FPS0.E_qty, 0, "数量"
        .SetText E_FPS0.E_MPNSeq, 0, "SerialNumber"
        .SetText E_FPS0.E_NewLotid, 0, "新LotID"
        .SetText E_FPS0.E_PRODUCT, 0, "厂内料号"
        .SetText E_FPS0.E_CustomerProuduct, 0, "客户料号"
        .SetText E_FPS0.E_BigQbox, 0, "大箱号"


        .ColWidth(E_FPS0.e_ID) = 5
        .ColWidth(E_FPS0.E_ContainerName) = 10
        .ColWidth(E_FPS0.E_QboxName) = 10
        .ColWidth(E_FPS0.E_LOTID) = 10
        .ColWidth(E_FPS0.E_qty) = 5
        .ColWidth(E_FPS0.E_MPNSeq) = 10
        .ColWidth(E_FPS0.E_NewLotid) = 10
        
        .ColWidth(E_FPS0.E_PRODUCT) = 11
        .ColWidth(E_FPS0.E_CustomerProuduct) = 17
        .ColWidth(E_FPS0.E_BigQbox) = 10
        
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
    

End Sub

Private Sub IniFpsHeader2()
    With fps(0)
        .TypeMaxEditLen = 500
        .ReDraw = False
        .MaxCols = E_FPS1.E_END - 1
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
        
        .SetText E_FPS1.e_NO, 0, "序号"
        .SetText E_FPS1.e_date, 0, "入库时间"
        .SetText E_FPS1.e_OrderName, 0, "工单号"
        .SetText E_FPS1.E_SmallBoxID, 0, "小箱号"
        .SetText E_FPS1.E_name, 0, "品名"
        .SetText E_FPS1.E_partno, 0, "料号"
       ' .SetText E_FPS1.E_BatchPlant, 0, "厂内批号"
        .SetText E_FPS1.E_BatchPlant, 0, "LOT号后缀"
        .SetText E_FPS1.E_LOTID, 0, "LOT号"
        .SetText E_FPS1.E_WAFERID, 0, "WAFERLIST"
        .SetText E_FPS1.E_GrossDie, 0, "客户设计GoodDie"
        .SetText E_FPS1.E_GoodDie, 0, "GOODDIE数量"
        .SetText E_FPS1.E_Pieces, 0, "进厂片数"
        .SetText E_FPS1.E_INsiteNgDie, 0, "厂内NG"
        .SetText E_FPS1.E_BigBoxID, 0, "大箱号"
        .SetText E_FPS1.E_SecondCode, 0, "二级代码"
        .SetText E_FPS1.E_QtyInStock, 0, "库存数量"
        
        
        .ColWidth(E_FPS1.e_NO) = 10
        .ColWidth(E_FPS1.e_date) = 10
        .ColWidth(E_FPS1.e_OrderName) = 10
        .ColWidth(E_FPS1.E_SmallBoxID) = 10
        .ColWidth(E_FPS1.E_name) = 10
        .ColWidth(E_FPS1.E_partno) = 10
        .ColWidth(E_FPS1.E_BatchPlant) = 10
        .ColWidth(E_FPS1.E_LOTID) = 10
        .ColWidth(E_FPS1.E_WAFERID) = 50
        .ColWidth(E_FPS1.E_GrossDie) = 10
        .ColWidth(E_FPS1.E_GoodDie) = 10
        .ColWidth(E_FPS1.E_Pieces) = 10
        .ColWidth(E_FPS1.E_INsiteNgDie) = 10
        .ColWidth(E_FPS1.E_BigBoxID) = 10
        .ColWidth(E_FPS1.E_SecondCode) = 10
        .ColWidth(E_FPS1.E_QtyInStock) = 10

        
        .RowHeight(0) = 20
        .RowHeight(-1) = 15
             
        .ReDraw = True
    End With
    
    
    

End Sub

Private Sub IniFpsHeader_GC()
    With fps(0)
        .TypeMaxEditLen = 500
        .ReDraw = False
        .MaxCols = E_FPS_GC.E_END - 1
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
        
        .SetText E_FPS_GC.e_NO, 0, "序号"
        .SetText E_FPS_GC.e_date, 0, "入库时间"
        .SetText E_FPS_GC.e_OrderName, 0, "工单号"
        .SetText E_FPS_GC.E_SmallBoxID, 0, "小箱号"
        .SetText E_FPS_GC.E_name, 0, "品名"
        .SetText E_FPS_GC.E_partno, 0, "料号"
       ' .SetText E_FPS1.E_BatchPlant, 0, "厂内批号"
        .SetText E_FPS_GC.E_BatchPlant, 0, "LOT号+后缀"
        .SetText E_FPS_GC.E_LOTID, 0, "LOT号"
        .SetText E_FPS_GC.E_WAFERID, 0, "WAFER号"
        .SetText E_FPS_GC.E_GrossDie, 0, "客户设计GoodDie"
        .SetText E_FPS_GC.E_GoodDie, 0, "GOODDIE数量"
        .SetText E_FPS_GC.E_Pieces, 0, "进厂片数"
        .SetText E_FPS_GC.E_INsiteNgDie, 0, "厂内NG"
        .SetText E_FPS_GC.E_BigBoxID, 0, "大箱号"
        .SetText E_FPS_GC.E_SecondCode, 0, "二级代码"
        .SetText E_FPS_GC.E_Mpndesc, 0, "客户机种"
        .SetText E_FPS_GC.E_Bond, 0, "保税或非保"
        .SetText E_FPS_GC.E_ProductType, 0, "形式"
        .SetText E_FPS_GC.E_QtyInStock, 0, "库存数量"
        .SetText E_FPS_GC.E_DieByPcs, 0, "单片数量"
        .SetText E_FPS_GC.e_ShipNo, 0, "出货单号"
        .SetText E_FPS_GC.E_ShipTo, 0, "出货地"
        
        
        
        .ColWidth(E_FPS_GC.e_NO) = 10
        .ColWidth(E_FPS_GC.e_date) = 10
        .ColWidth(E_FPS_GC.e_OrderName) = 10
        .ColWidth(E_FPS_GC.E_SmallBoxID) = 10
        .ColWidth(E_FPS_GC.E_name) = 10
        .ColWidth(E_FPS_GC.E_partno) = 15
        .ColWidth(E_FPS_GC.E_BatchPlant) = 10
        .ColWidth(E_FPS_GC.E_LOTID) = 10
        .ColWidth(E_FPS_GC.E_WAFERID) = 50
        .ColWidth(E_FPS_GC.E_GrossDie) = 10
        .ColWidth(E_FPS_GC.E_GoodDie) = 10
        .ColWidth(E_FPS_GC.E_Pieces) = 10
        .ColWidth(E_FPS_GC.E_INsiteNgDie) = 10
        .ColWidth(E_FPS_GC.E_BigBoxID) = 10
        .ColWidth(E_FPS_GC.E_SecondCode) = 10
        .ColWidth(E_FPS_GC.E_Mpndesc) = 10
        .ColWidth(E_FPS_GC.E_Bond) = 10
        .ColWidth(E_FPS_GC.E_ProductType) = 10
        .ColWidth(E_FPS_GC.E_QtyInStock) = 10
        
        .ColWidth(E_FPS_GC.E_DieByPcs) = 10
        .ColWidth(E_FPS_GC.e_ShipNo) = 10
        .ColWidth(E_FPS_GC.E_ShipTo) = 10
        
        
        
        
        .RowHeight(0) = 20
        .RowHeight(-1) = 15
             
        .ReDraw = True
    End With
    
    
    

End Sub
Private Sub IniFpsHeader3()
    With fps(0)
        .TypeMaxEditLen = 500
        .ReDraw = False
        .MaxCols = E_FPS1.E_END - 1
        .MaxRows = 0
        
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = False
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        .RowHeight(0) = 20
        .RowHeight(-1) = 15
             
        .ReDraw = True
    End With
    
End Sub



Private Sub Form_Load()
'IniProduct
OptCIS.Value = True


End Sub


Private Sub OptAA_Click()
lblLabel4.Visible = False
cbCusCode.Visible = False
chkStockOnly.Visible = False

IniFpsHeader1
End Sub

Private Sub OptCIS_Click()
lblLabel4.Visible = True
cbCusCode.Visible = True
chkStockOnly.Visible = True
InitCuscode
'IniFpsHeader2

End Sub


Private Sub OptBUMPING_Click()
lblLabel4.Visible = False
cbCusCode.Visible = False
chkStockOnly.Visible = False
IniFpsHeader3

End Sub

Private Function SqlServer2ExporToExcel_Trim(strOpen As String)
'增加去前后空格的功能
'*********************************************************
'* 嘿ExporToExcel
'* 旧计誹EXCEL
'* ノ猭ExporToExcel(sql琩高才﹃)
'*********************************************************
' 导出SqlServer Excel

    Dim Rs_Data As New ADODB.Recordset
    Dim Irowcount As Long
    Dim Icolcount As Integer
    Dim i As Integer
    Dim j As Integer
    
    Dim xlApp As New Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xlQuery As Excel.QueryTable
    
'      If Cnn.State = 0 Then
'        ConOracle
'      End If
      
    If INIadoCon2.State = 0 Then
        INIConnectSTART2
    End If


    
    
   ' CONN_TO_ORACLE_DATABASE2
    
    With Rs_Data
        If .State = adStateOpen Then
            .Close
        End If
        .ActiveConnection = INIadoCon2
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Source = strOpen
        .Open
    End With
    With Rs_Data
        If .RecordCount < 1 Then
            MsgBox ("查询不到数据!")
            Exit Function
        End If
        '癘魁羆计
        Irowcount = .RecordCount
        '琿羆计
        Icolcount = .Fields.count
    End With
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlBook = xlApp.Workbooks().Add
    Set xlSheet = xlBook.Worksheets("sheet1")
    xlApp.Visible = False
    
    '睰琩高粂旧EXCEL计沮誹
    Set xlQuery = xlSheet.QueryTables.Add(Rs_Data, xlSheet.Range("a1"))
    
    With xlQuery
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = True
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
    End With
    
    xlQuery.FieldNames = True '陪ボ琿
    xlQuery.Refresh
    
    With xlSheet
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.name = "堵蔨"
        '砞竚堵蔨
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Bold = True
        '砞竚蔨彩
        .Range(.Cells(1, 1), .Cells(Irowcount + 1, Icolcount)).Borders.LineStyle = xlContinuous
        '砞竚娩妓Α

        For i = 2 To Irowcount + 1
            For j = 1 To Icolcount
                .Cells(i, j).Value = Replace(Trim(.Cells(i, j).Value), Chr(10), "")
            Next
        Next
    End With
    With xlSheet.Range("2:" & Irowcount + 1)
        .horizontalAlignment = xlLeft
    End With
    xlSheet.Range("A1").Select
    xlApp.Columns.AutoFit
    
    
    
    
'    With xlSheet.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""发蔨_GB2312,盽?""&10そ嘿"   ' & Gsmc
'        .CenterHeader = "&""发蔨_GB2312,盽砏""蹦潦快龟兜&""Ш蔨,盽砏""" & Chr(10) & "&""发蔨_GB2312,盽?""&10ら 戳"
'        .RightHeader = "" & Chr(10) & "&""发蔨_GB2312,盽砏""&10虫涵"
'        .LeftFooter = "&""发蔨_GB2312,盽砏""&10"
'        .CenterFooter = "&""发蔨_GB2312,盽砏""&10ら戳"
'        .RightFooter = "&""发蔨_GB2312,盽?""&10材&P&N"
'    End With
    
    xlApp.Application.Visible = True
    Set xlApp = Nothing  '"ユ临北倒Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing

End Function


Private Function SqlServer2ExporToExcel_Trim_AA(strOpen As String)
'增加去前后空格的功能
'*********************************************************
'* 嘿ExporToExcel
'* 旧计誹EXCEL
'* ノ猭ExporToExcel(sql琩高才﹃)
'*********************************************************
' 导出SqlServer Excel

    Dim Rs_Data As New ADODB.Recordset
    Dim Irowcount As Long
    Dim Icolcount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Bigboxid As Integer
    Dim QtyinBigbox As Long
    Dim WeightinBigbox As Double
    Dim xlApp As New Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xlQuery As Excel.QueryTable
    
'      If Cnn.State = 0 Then
'        ConOracle
'      End If
      
    If INIadoCon2.State = 0 Then
        INIConnectSTART2
    End If


    
    
   ' CONN_TO_ORACLE_DATABASE2
    
    With Rs_Data
        If .State = adStateOpen Then
            .Close
        End If
        .ActiveConnection = INIadoCon2
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Source = strOpen
        .Open
    End With
    With Rs_Data
        If .RecordCount < 1 Then
            MsgBox ("查询不到数据!")
            Exit Function
        End If
        '癘魁羆计
        Irowcount = .RecordCount
        '琿羆计
        Icolcount = .Fields.count
    End With
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlBook = xlApp.Workbooks().Add
    Set xlSheet = xlBook.Worksheets("sheet1")
    xlApp.Visible = False
    
    '睰琩高粂旧EXCEL计沮誹
    Set xlQuery = xlSheet.QueryTables.Add(Rs_Data, xlSheet.Range("a1"))
    
    With xlQuery
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = True
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
    End With
    
    xlQuery.FieldNames = True '陪ボ琿
    xlQuery.Refresh
    
    With xlSheet
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.name = "堵蔨"
        '砞竚堵蔨
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Bold = True
        '砞竚蔨彩
        .Range(.Cells(1, 1), .Cells(Irowcount + 1, Icolcount)).Borders.LineStyle = xlContinuous
        '砞竚娩妓Α

        For i = 2 To Irowcount + 1
            For j = 1 To Icolcount
                .Cells(i, j).Value = Replace(Trim(.Cells(i, j).Value), Chr(10), "")
            Next
        Next
        '合并单元格
        Bigboxid = 0
        QtyinBigbox = 0
        For i = 2 To Irowcount + 1
           Bigboxid = Bigboxid + 1
           QtyinBigbox = Val(.Cells(i, 4).Value)
           WeightinBigbox = Val(.Cells(i, 4).Value) * Val(.Cells(i, 9).Value)
            For j = i + 1 To Irowcount + 1
               If .Cells(j, 10).Value <> .Cells(i, 10).Value Then
                   Exit For
               End If
               QtyinBigbox = QtyinBigbox + Val(.Cells(j, 4).Value)
               WeightinBigbox = WeightinBigbox + Val(.Cells(j, 4).Value) * Val(.Cells(j, 9).Value)
            Next
            .Cells(i, 11).Value = Bigboxid '箱序号
            .Cells(i, 12).Value = QtyinBigbox '数量
            .Cells(i, 13).Value = WeightinBigbox '净重
            .Cells(i, 16).Value = Format(Now(), "yyyy/mm/dd") '日期
            xlApp.Application.DisplayAlerts = False '不弹出提示框
            xlSheet.Range(Chr(10 + 64) & i & ":" & Chr(10 + 64) & j - 1).Merge
            xlSheet.Range(Chr(11 + 64) & i & ":" & Chr(11 + 64) & j - 1).Merge
            xlSheet.Range(Chr(12 + 64) & i & ":" & Chr(12 + 64) & j - 1).Merge
            xlSheet.Range(Chr(13 + 64) & i & ":" & Chr(13 + 64) & j - 1).Merge
            xlSheet.Range(Chr(14 + 64) & i & ":" & Chr(14 + 64) & j - 1).Merge
            xlSheet.Range(Chr(15 + 64) & i & ":" & Chr(15 + 64) & j - 1).Merge
            xlSheet.Range(Chr(16 + 64) & i & ":" & Chr(16 + 64) & j - 1).Merge
            i = j - 1
        Next
        .Cells(Irowcount + 2, 4).Value = WorksheetFunction.Sum(Range(Chr(4 + 64) & "2:" & Chr(4 + 64) & Irowcount + 1))
        .Cells(Irowcount + 2, 11).Value = Bigboxid
        .Cells(Irowcount + 2, 12).Value = WorksheetFunction.Sum(Range(Chr(12 + 64) & "2:" & Chr(12 + 64) & Irowcount + 1))
        .Cells(Irowcount + 2, 13).Value = WorksheetFunction.Sum(Range(Chr(13 + 64) & "2:" & Chr(13 + 64) & Irowcount + 1))
        .Cells(Irowcount + 2, 14).Value = WorksheetFunction.Sum(Range(Chr(14 + 64) & "2:" & Chr(14 + 64) & Irowcount + 1))
      
    End With
    With xlSheet.Range("2:" & Irowcount + 2)
        .horizontalAlignment = xlLeft
    End With
    xlSheet.Range("A1").Select
    xlApp.Columns.AutoFit


    
'    With xlSheet.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""发蔨_GB2312,盽?""&10そ嘿"   ' & Gsmc
'        .CenterHeader = "&""发蔨_GB2312,盽砏""蹦潦快龟兜&""Ш蔨,盽砏""" & Chr(10) & "&""发蔨_GB2312,盽?""&10ら 戳"
'        .RightHeader = "" & Chr(10) & "&""发蔨_GB2312,盽砏""&10虫涵"
'        .LeftFooter = "&""发蔨_GB2312,盽砏""&10"
'        .CenterFooter = "&""发蔨_GB2312,盽砏""&10ら戳"
'        .RightFooter = "&""发蔨_GB2312,盽?""&10材&P&N"
'    End With
    
    xlApp.Application.Visible = True
    Set xlApp = Nothing  '"ユ临北倒Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing

End Function


Private Sub FpsToExcel()
    If fps(0).MaxRows = 0 Then
        MsgBox "没有数据可以导出", vbInformation, "提示"
        Exit Sub
    End If

    Dim i As Integer
    Dim j As Integer
    
    Dim xlApp      As Excel.Application
    Dim xlBook     As Excel.Workbook
    Dim xlSheet    As Excel.Worksheet
    

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)

    With xlApp
        .Rows(1).Font.Bold = True
    End With
    
 On Error GoTo Ert
    With fps(0)

        For i = 0 To .MaxRows
            For j = 1 To .MaxCols
                .Col = j
                .Row = i
                xlSheet.Cells(i + 1, j) = Trim$(("'" & .text))
            Next j
       
        Next i
        
    End With

    '数字列格式调整
    For j = 1 To fps(0).MaxCols
        If Trim(xlSheet.Cells(1, j)) = "客户设计GoodDie" Or Trim(xlSheet.Cells(1, j)) = "单片数量" Or Trim(xlSheet.Cells(1, j)) = "GOODDIE数量" Or Trim(xlSheet.Cells(1, j)) = "进厂片数" Or Trim(xlSheet.Cells(1, j)) = "厂内NG" Or Trim(xlSheet.Cells(1, j)) = "库存数量" Then
            For i = 2 To fps(0).MaxRows + 1
                xlSheet.Cells(i, j) = Replace(xlSheet.Cells(i, j), "'", "")
            Next
        End If
    Next
    With xlSheet.Range("2:" & fps(0).MaxRows + 1)
        .horizontalAlignment = xlLeft
    End With
    xlSheet.Range("A1").Select
    xlApp.Columns.AutoFit
    
    xlApp.Application.Visible = True
    
    
    Set xlApp = Nothing  '"ユ临北倒Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing
Ert:
    If Not (xlApp Is Nothing) Then
        
    Set xlApp = Nothing  '"ユ临北倒Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing
    End If
    
    
End Sub






























