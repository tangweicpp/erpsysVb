VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Frm_TSV_ERPQbox 
   Caption         =   "GC不良品查询"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16170
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
   ScaleWidth      =   16170
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "导出"
      Height          =   360
      Left            =   7080
      TabIndex        =   14
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "查询"
      Height          =   360
      Left            =   4440
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox TxtBillNo 
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   240
      Width           =   3975
   End
   Begin VB.TextBox TxtBigQbox 
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox TxtWo 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   122159105
      CurrentDate     =   41424
   End
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   121962497
      CurrentDate     =   41424
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   7215
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   15135
      _Version        =   524288
      _ExtentX        =   26696
      _ExtentY        =   12726
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
      SpreadDesigner  =   "Frm_TSV_ERPQbox.frx":0000
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin MSDataListLib.DataCombo Text3 
      Height          =   315
      Left            =   7080
      TabIndex        =   11
      Top             =   780
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label9 
      Caption         =   "入库信息"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9720
      TabIndex        =   16
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "发货信息"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   15
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "成品料号："
      Height          =   195
      Left            =   6120
      TabIndex        =   12
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发货单据编号："
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   360
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始时间："
      Height          =   195
      Left            =   840
      TabIndex        =   7
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "大箱号："
      Height          =   195
      Left            =   6240
      TabIndex        =   5
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LotID："
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束时间："
      Height          =   195
      Left            =   3480
      TabIndex        =   1
      Top             =   840
      Width           =   900
   End
End
Attribute VB_Name = "Frm_TSV_ERPQbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
    E_id = 0                'id
    E_Bill
    E_InDate             '入库时间
    E_BigBox              '大箱号
     E_SmallBox              '大箱号
    E_PRODUCT              '成品料号
    E_LotID                'LotID
    E_WaferId
    E_ALLQty               'TotalDie
    E_AQty              'A数量
    E_EQty              'E数量
   
    E_End
    
    
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


woTemp = UCase(Trim(txtWO.Text))
productTemp = UCase(Trim(TxtProduct.Text))
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
waferIdTemp = UCase(Trim(TxtWaferID.Text))


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

Private Sub cmdQuery_Click()


Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim billNoTemp As String
Dim lotIdTemp As String
Dim bigQboxTemp As String
Dim productTemp As String
Dim date1Temp As String
Dim date2Temp As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""
sqlTemp = ""

If Trim(TxtBillNo.Text) <> "" Then

billNoTemp = UCase(Trim(TxtBillNo.Text))

sqlTemp = " SELECT distinct  b.单据编号,d.入库时间,f.箱号 as 大箱号,b.箱号,b.料号,b.工单号,b.流程卡编号,b.数量 ," & _
" [dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as goodDie,b.数量-[dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as ngDie " & _
" FROM  tblStockMove a, tblStockMovesub b , dbo.tblPackToHouseRec c,dbo.tblPackToHouse d, tblPackTreeInf e,tblPackTreeInf f " & _
" WHERE b.单据编号=a.单据编号 and a.单据编号='" + billNoTemp + "' " & _
" and c.流程卡编号=b.流程卡编号 and d.入库单编号=c.入库单编号 and e.箱号=b.箱号 and f.序号=e.上级序号 " & _
" order by f.箱号,b.箱号,b.工单号,b.流程卡编号"
  
 ElseIf Trim(Text3.Text) <> "" Then
     productTemp = UCase(Trim(Text3.Text))
     date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
     date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")
 
 sqlTemp = " SELECT distinct  a.入库单编号,a.入库时间,f.箱号 as 大箱号,c.QBOXNUMBER,a.料号, b.工单号,b.流程卡编号,b.入库数, " & _
           "  [dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as goodDie,b.入库数-[dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as ngDie " & _
           "  FROM dbo.tblPackToHouse a , dbo.tblPackToHouseRec b,dbo.TblQBOXNUMBER_TSV c,tblPackTreeInf e,tblPackTreeInf f " & _
           " Where a.料号 like '" + productTemp + "%' and a.创建时间>='" + date1Temp + "' and a.创建时间<='" + date2Temp + "'  and  b.入库单编号 = a.入库单编号 And b.工单号 = a.工单号   " & _
           " and c.WAFERSCRIBENUMBER=b.流程卡编号 and c.WAFERNUMBER=b.工单号 and e.箱号=c.QBOXNUMBER and f.序号=e.上级序号 " & _
           "  order by f.箱号,c.QBOXNUMBER,b.工单号,b.流程卡编号 "
End If
 
 
 If Trim(txtWO.Text) <> "" Then
        lotIdTemp = UCase(Trim(txtWO.Text))
  sqlTemp = " SELECT distinct  a.入库单编号,a.入库时间,f.箱号 as 大箱号,c.QBOXNUMBER,a.料号, b.工单号,b.流程卡编号,b.入库数, " & _
           "  [dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as goodDie,b.入库数-[dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as ngDie " & _
           "  FROM dbo.tblPackToHouse a , dbo.tblPackToHouseRec b,dbo.TblQBOXNUMBER_TSV c,tblPackTreeInf e,tblPackTreeInf f " & _
           " Where a.工单号 = '" + lotIdTemp + "' and  b.入库单编号 = a.入库单编号 And b.工单号 = a.工单号   " & _
           " and c.WAFERSCRIBENUMBER=b.流程卡编号 and c.WAFERNUMBER=b.工单号 and e.箱号=c.QBOXNUMBER and f.序号=e.上级序号 " & _
           "  order by f.箱号,c.QBOXNUMBER,b.工单号,b.流程卡编号 "
           
End If
 
If Trim(TxtBigQbox.Text) <> "" Then
       bigQboxTemp = UCase(Trim(TxtBigQbox.Text))
 sqlTemp = " SELECT distinct  a.入库单编号,a.入库时间,f.箱号 as 大箱号,c.QBOXNUMBER,a.料号, b.工单号,b.流程卡编号,b.入库数, " & _
           "  [dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as goodDie,b.入库数-[dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as ngDie " & _
           "  FROM dbo.tblPackToHouse a , dbo.tblPackToHouseRec b,dbo.TblQBOXNUMBER_TSV c,tblPackTreeInf e,tblPackTreeInf f " & _
           " Where b.入库单编号 = a.入库单编号 And b.工单号 = a.工单号   " & _
           " and c.WAFERSCRIBENUMBER=b.流程卡编号 and c.WAFERNUMBER=b.工单号 and e.箱号=c.QBOXNUMBER and f.序号=e.上级序号  and  f.箱号 = '" + bigQboxTemp + "' " & _
           "  order by f.箱号,c.QBOXNUMBER,b.工单号,b.流程卡编号 "
 
End If
       
       
 
 
 If Trim(TxtBillNo.Text) = "" And Trim(Text3.Text) = "" And Trim(txtWO.Text) = "" And Trim(TxtBigQbox.Text) = "" Then
 
     date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
     date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")
 
 sqlTemp = " SELECT distinct  a.入库单编号,a.入库时间,f.箱号 as 大箱号,c.QBOXNUMBER,a.料号, b.工单号,b.流程卡编号,b.入库数, " & _
           "  [dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as goodDie,b.入库数-[dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as ngDie " & _
           "  FROM dbo.tblPackToHouse a , dbo.tblPackToHouseRec b,dbo.TblQBOXNUMBER_TSV c,tblPackTreeInf e,tblPackTreeInf f " & _
           " Where a.客户代码='GC' and a.创建时间>='" + date1Temp + "' and a.创建时间<='" + date2Temp + "'  and  b.入库单编号 = a.入库单编号 And b.工单号 = a.工单号   " & _
           " and c.WAFERSCRIBENUMBER=b.流程卡编号 and c.WAFERNUMBER=b.工单号 and e.箱号=c.QBOXNUMBER and f.序号=e.上级序号 " & _
           "  order by f.箱号,c.QBOXNUMBER,b.工单号,b.流程卡编号 "
 
 End If


Set mainItemRS = GetGCNGRpt(sqlTemp)

With fps(0)
        .MaxRows = 0
        If mainItemRS.RecordCount > 0 Then
            Set .DataSource = mainItemRS
       
        End If
End With



End Sub

Private Sub Command2_Click()


Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim billNoTemp As String
Dim lotIdTemp As String
Dim bigQboxTemp As String
Dim productTemp As String
Dim date1Temp As String
Dim date2Temp As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""
sqlTemp = ""

If Trim(TxtBillNo.Text) <> "" Then

billNoTemp = UCase(Trim(TxtBillNo.Text))

sqlTemp = " SELECT distinct  b.单据编号,d.入库时间,f.箱号 as 大箱号,b.箱号,b.料号,b.工单号,b.流程卡编号,b.数量 ," & _
" [dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as goodDie,b.数量-[dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as ngDie " & _
" FROM  tblStockMove a, tblStockMovesub b , dbo.tblPackToHouseRec c,dbo.tblPackToHouse d, tblPackTreeInf e,tblPackTreeInf f " & _
" WHERE b.单据编号=a.单据编号 and a.单据编号='" + billNoTemp + "' " & _
" and c.流程卡编号=b.流程卡编号 and d.入库单编号=c.入库单编号 and e.箱号=b.箱号 and f.序号=e.上级序号 " & _
" order by f.箱号,b.箱号,b.工单号,b.流程卡编号"
  
 ElseIf Trim(Text3.Text) <> "" Then
     productTemp = UCase(Trim(Text3.Text))
     date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
     date2Temp = Format(DTP2.Value, "YYYY-MM-DD")
 
sqlTemp = " SELECT distinct  a.入库单编号,a.入库时间,f.箱号 as 大箱号,c.QBOXNUMBER,a.料号, b.工单号,b.流程卡编号,b.入库数, " & _
           "  [dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as goodDie,b.入库数-[dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as ngDie " & _
           "  FROM dbo.tblPackToHouse a , dbo.tblPackToHouseRec b,dbo.TblQBOXNUMBER_TSV c,tblPackTreeInf e,tblPackTreeInf f " & _
           " Where a.料号 like '" + productTemp + "%' and a.创建时间>='" + date1Temp + "' and a.创建时间<='" + date2Temp + "'  and  b.入库单编号 = a.入库单编号 And b.工单号 = a.工单号   " & _
           " and c.WAFERSCRIBENUMBER=b.流程卡编号 and c.WAFERNUMBER=b.工单号 and e.箱号=c.QBOXNUMBER and f.序号=e.上级序号 " & _
           "  order by f.箱号,c.QBOXNUMBER,b.工单号,b.流程卡编号 "
End If
 
 
 If Trim(txtWO.Text) <> "" Then
        lotIdTemp = UCase(Trim(txtWO.Text))
  sqlTemp = " SELECT distinct  a.入库单编号,a.入库时间,f.箱号 as 大箱号,c.QBOXNUMBER,a.料号, b.工单号,b.流程卡编号,b.入库数, " & _
           "  [dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as goodDie,b.入库数-[dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as ngDie " & _
           "  FROM dbo.tblPackToHouse a , dbo.tblPackToHouseRec b,dbo.TblQBOXNUMBER_TSV c,tblPackTreeInf e,tblPackTreeInf f " & _
           " Where a.工单号 = '" + lotIdTemp + "' and  b.入库单编号 = a.入库单编号 And b.工单号 = a.工单号   " & _
           " and c.WAFERSCRIBENUMBER=b.流程卡编号 and c.WAFERNUMBER=b.工单号 and e.箱号=c.QBOXNUMBER and f.序号=e.上级序号 " & _
           "  order by f.箱号,c.QBOXNUMBER,b.工单号,b.流程卡编号 "
           
End If
 
If Trim(TxtBigQbox.Text) <> "" Then
       bigQboxTemp = UCase(Trim(TxtBigQbox.Text))
 sqlTemp = " SELECT distinct  a.入库单编号,a.入库时间,f.箱号 as 大箱号,c.QBOXNUMBER,a.料号, b.工单号,b.流程卡编号,b.入库数, " & _
           "  [dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as goodDie,b.入库数-[dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as ngDie " & _
           "  FROM dbo.tblPackToHouse a , dbo.tblPackToHouseRec b,dbo.TblQBOXNUMBER_TSV c,tblPackTreeInf e,tblPackTreeInf f " & _
           " Where b.入库单编号 = a.入库单编号 And b.工单号 = a.工单号   " & _
           " and c.WAFERSCRIBENUMBER=b.流程卡编号 and c.WAFERNUMBER=b.工单号 and e.箱号=c.QBOXNUMBER and f.序号=e.上级序号  and  f.箱号 = '" + bigQboxTemp + "' " & _
           "  order by f.箱号,c.QBOXNUMBER,b.工单号,b.流程卡编号 "
 
End If

 If Trim(TxtBillNo.Text) = "" And Trim(Text3.Text) = "" And Trim(txtWO.Text) = "" And Trim(TxtBigQbox.Text) = "" Then
 
     date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
     date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")
 
 sqlTemp = " SELECT distinct  a.入库单编号,a.入库时间,f.箱号 as 大箱号,c.QBOXNUMBER,a.料号, b.工单号,b.流程卡编号,b.入库数, " & _
           "  [dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as goodDie,b.入库数-[dbo].[Get_TSV_GC_WaferGDie](b.流程卡编号) as ngDie " & _
           "  FROM dbo.tblPackToHouse a , dbo.tblPackToHouseRec b,dbo.TblQBOXNUMBER_TSV c,tblPackTreeInf e,tblPackTreeInf f " & _
           " Where a.客户代码='GC' and a.创建时间>='" + date1Temp + "' and a.创建时间<='" + date2Temp + "'  and  b.入库单编号 = a.入库单编号 And b.工单号 = a.工单号   " & _
           " and c.WAFERSCRIBENUMBER=b.流程卡编号 and c.WAFERNUMBER=b.工单号 and e.箱号=c.QBOXNUMBER and f.序号=e.上级序号 " & _
           "  order by f.箱号,c.QBOXNUMBER,b.工单号,b.流程卡编号 "
 
 End If
  
  
  
     SqlServer2ExporToExcel (sqlTemp)






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


woTemp = UCase(Trim(txtWO.Text))
productTemp = UCase(Trim(TxtProduct.Text))
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


woTemp = UCase(Trim(txtWO.Text))
productTemp = UCase(Trim(TxtProduct.Text))
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


DTP1.Value = Now - 1

DTP2.Value = Now

IniFpsHeader1
'IniFpsHeader2

End Sub
Private Sub IniProduct()
Set mainItemRS = GetProduct()
Set Text3.RowSource = mainItemRS
Text3.ListField = mainItemRS("productname").Name
Text3.BoundColumn = mainItemRS("PID").Name

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
.SetText E_FPS0.E_Bill, 0, "单据编号"
.SetText E_FPS0.E_InDate, 0, "入库时间"
.SetText E_FPS0.E_BigBox, 0, "大箱号"
.SetText E_FPS0.E_SmallBox, 0, "小箱号"
.SetText E_FPS0.E_PRODUCT, 0, "成品料号"
.SetText E_FPS0.E_LotID, 0, "LotID"
.SetText E_FPS0.E_WaferId, 0, "WaferID"
.SetText E_FPS0.E_ALLQty, 0, "TotalDie"
.SetText E_FPS0.E_AQty, 0, "良品数"
.SetText E_FPS0.E_EQty, 0, "制程不良数"
    


        .ColWidth(E_FPS0.E_id) = 7
         .ColWidth(E_FPS0.E_Bill) = 10
        .ColWidth(E_FPS0.E_InDate) = 15
        .ColWidth(E_FPS0.E_BigBox) = 10
         .ColWidth(E_FPS0.E_SmallBox) = 10
        .ColWidth(E_FPS0.E_PRODUCT) = 15
        .ColWidth(E_FPS0.E_LotID) = 10
        .ColWidth(E_FPS0.E_WaferId) = 10
        .ColWidth(E_FPS0.E_ALLQty) = 10

        .ColWidth(E_FPS0.E_AQty) = 10
        .ColWidth(E_FPS0.E_EQty) = 10


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
        
          
        .SetText E_FPS1.E_id, 0, "序号"
        .SetText E_FPS1.E_Wo, 0, "工单号"
        .SetText E_FPS1.E_WaferId, 0, "WaferId"
        .SetText E_FPS1.E_CompleteFlag, 0, "完成标志"
        .SetText E_FPS1.E_TotalDie, 0, "TotalDie数量"
        .SetText E_FPS1.E_GoodDie, 0, "GoodDie数量"
        .SetText E_FPS1.E_WaferLot, 0, "WaferLot"
        .SetText E_FPS1.E_MarkingCode, 0, "MarkingCode"
        
        
        .ColWidth(E_FPS1.E_id) = 10
        .ColWidth(E_FPS1.E_Wo) = 10
        .ColWidth(E_FPS1.E_WaferId) = 10
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
IniProduct
End Sub
