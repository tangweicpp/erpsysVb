VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form Frm_TSV_MDOI 
   Caption         =   "市场部 来料汇总报表查询"
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
   Begin VB.CommandButton Command2 
      Caption         =   "导出"
      Height          =   360
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "查询"
      Height          =   360
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   9015
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   16695
      _Version        =   196613
      _ExtentX        =   29448
      _ExtentY        =   15901
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
      SpreadDesigner  =   "Frm_TSV_MDOI.frx":0000
      TextTip         =   2
   End
End
Attribute VB_Name = "Frm_TSV_MDOI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
    E_ID = 0                'id
    E_Factory            '事业部
    E_Customer           '客户代码
    E_ToDayQty            '当天片数
    E_Mon1                '一月
    E_Mon2                '二月
    E_Mon3                '三月
    E_Mon4                '四月
    E_Mon5                '五月
    E_Mon6                '六月
    E_Mon7                '7月
    E_Mon8                '8月
    E_Mon9                '9月
    E_Mon10                '10月
    E_Mon11                '11月
    E_Mon12                '12月
    E_sum                  '汇总
    

    E_End
    
    
End Enum



Dim Rs As New ADODB.Recordset
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


woTemp = UCase(Trim(TxtWo.Text))
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


Dim waferidTemp As String
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
waferidTemp = UCase(Trim(TxtWaferID.Text))


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
"from erpintegration2.ib_wohistory a, ib_waferlist b  where a.modifyflag='1' and b.ordername=a.ordername and b.waferid='" + waferidTemp + "' "

          
          
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

Private Sub CmdQuery_Click()
Dim strSql As String
Dim m As Integer
Dim i As Integer
Dim factoryIdTemp As Integer
Dim custIDTemp As Integer
Dim custNameTemp As String

Dim j As Integer
Dim qtyMonTemp As Long




strSql = "  select a.id,b.id as cusID, a.workname,  b.customerid  from  TSV_MDRpt_Work_type a,TSV_MDRpt_Customer_type b " & _
         " Where b.workid = a.id and b.flag='Y' and a.flag='Y' order by a.id,b.id"

 If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

      i = 0
      With fps(0)
            .MaxRows = Rs.RecordCount + 1
            
              For m = 1 To Rs.RecordCount
                factoryIdTemp = CInt(Rs("id").Value)
                custIDTemp = CInt(Rs("cusID").Value)
                custNameTemp = Rs("customerid").Value
                
                
                i = i + 1
                .Row = i
                .Col = 1
                .Text = Rs("workname").Value
                
                .Row = i
                .Col = 2
                .Text = custNameTemp
                
                .Row = i
                .Col = 3
                .Text = GetQtyMDDay(factoryIdTemp, custIDTemp, j, custNameTemp)
                

                 
                 For j = 1 To 12
                 
                 qtyMonTemp = GetQtyMDMonth(factoryIdTemp, custIDTemp, j, custNameTemp)
                 
                 
                .Row = i
                .Col = 3 + j
                .Text = qtyMonTemp
                 
                 Next j
                 
                
                
               Rs.MoveNext
               
              Next m
             
         End With
         
'汇总

'汇总最后一行：

'Dim iclo As Integer
'Dim irow As Integer
'Dim qtyAll As Long
'Dim qty01Temp As Long
'
'
'
'
'With fps(0)
' .Row = .MaxRows
'.Col = 1
'.Text = "Total："
'
'
'For iclo = 3 To 15
'    qtyAll = 0
'
'     For irow = 1 To .MaxRows - 1
'        .Row = irow
'        .Col = iclo
'        qty01Temp = CLng(.Text)
'
'        qtyAll = qtyAll + qty01Temp
'
'     Next irow
'
'      .Row = .MaxRows
'      .Col = iclo
'      .Text = qtyAll
'
'Next iclo
'
'End With



'汇总

'汇总最后一行：

Dim iclo As Integer
Dim irow As Integer
Dim qtyAll As Long
Dim qty01Temp As Long

Dim TSVQtyALL As Long
Dim BumpingQtyALL As Long
Dim T12QtyALL As Long

Dim TSVQtyRID As Integer
Dim BumpingQtyRID As Integer
Dim T12QtyRID As Integer


TSVQtyALL = 0
BumpingQtyALL = 0
T12QtyALL = 0

TSVQtyRID = 0
BumpingQtyRID = 0
T12QtyRID = 0

With fps(0)

.MaxRows = .MaxRows + 3



  For irow = 1 To .MaxRows - 4
        .Row = irow
        .Col = 1
        
        If .Text = "TSV" Then
            TSVQtyRID = irow
        End If
        
        If .Text = "Bumping" Then
            BumpingQtyRID = irow
        End If
        
        If .Text = "12寸" Then
            T12QtyRID = irow
        End If
        
  

  Next irow

End With



With fps(0)
 .Row = .MaxRows - 3
.Col = 1
.Text = "TSV汇总："

 .Row = .MaxRows - 2
.Col = 1
.Text = "Bumping汇总："

 .Row = .MaxRows - 1
.Col = 1
.Text = "12寸汇总："

 .Row = .MaxRows
.Col = 1
.Text = "所有汇总："



For iclo = 3 To 15
    qtyAll = 0
    
    TSVQtyALL = 0
    BumpingQtyALL = 0
    T12QtyALL = 0
    

     For irow = 1 To TSVQtyRID
        .Row = irow
        .Col = iclo
        qty01Temp = CLng(.Text)
        
        TSVQtyALL = TSVQtyALL + qty01Temp
        
     Next irow
     
     
       For irow = TSVQtyRID + 1 To BumpingQtyRID
        .Row = irow
        .Col = iclo
        qty01Temp = CLng(.Text)
        
        BumpingQtyALL = BumpingQtyALL + qty01Temp
        
      Next irow
      
    For irow = BumpingQtyRID + 1 To .MaxRows - 4
    
        .Row = irow
        .Col = iclo
        qty01Temp = CLng(.Text)
        
        T12QtyALL = T12QtyALL + qty01Temp
        
      Next irow
     
     
     
     
     
      .Row = .MaxRows - 3
      .Col = iclo
      .Text = TSVQtyALL
      
      .Row = .MaxRows - 2
      .Col = iclo
      .Text = BumpingQtyALL
      
      .Row = .MaxRows - 1
      .Col = iclo
      .Text = T12QtyALL
      
       .Row = .MaxRows
      .Col = iclo
      .Text = TSVQtyALL + BumpingQtyALL + T12QtyALL
      

Next iclo

End With


       
        
'汇总最后一列：



With fps(0)

For iclo = 1 To .MaxRows
    qtyAll = 0

     For irow = 4 To 15
        .Row = iclo
        .Col = irow
        qty01Temp = CLng(.Text)

        qtyAll = qtyAll + qty01Temp

     Next irow

      .Row = iclo
      .Col = 16
      .Text = qtyAll


Next iclo

End With


 MsgBox "查询成功", vbInformation, "友情提示"


End Sub

Private Sub Command2_Click()




  Call ExportFpspreadToExcel(fps(0))






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


woTemp = UCase(Trim(TxtWo.Text))
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


woTemp = UCase(Trim(TxtWo.Text))
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


woTemp = UCase(Trim(TxtWo.Text))
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
        
        
.SetText E_FPS0.E_ID, 0, "序号"
.SetText E_FPS0.E_Factory, 0, "事业部"
.SetText E_FPS0.E_Customer, 0, "客户代码"
.SetText E_FPS0.E_ToDayQty, 0, "当天来料片数"
.SetText E_FPS0.E_Mon1, 0, "1月份"
.SetText E_FPS0.E_Mon2, 0, "2月份"
.SetText E_FPS0.E_Mon3, 0, "3月份"
.SetText E_FPS0.E_Mon4, 0, "4月份"
.SetText E_FPS0.E_Mon5, 0, "5月份"
.SetText E_FPS0.E_Mon6, 0, "6月份"

.SetText E_FPS0.E_Mon7, 0, "7月份"
.SetText E_FPS0.E_Mon8, 0, "8月份"
.SetText E_FPS0.E_Mon9, 0, "9月份"
.SetText E_FPS0.E_Mon10, 0, "10月份"
.SetText E_FPS0.E_Mon11, 0, "11月份"
.SetText E_FPS0.E_Mon12, 0, "12月份"
.SetText E_FPS0.E_sum, 0, "Total"


        .ColWidth(E_FPS0.E_ID) = 5
        .ColWidth(E_FPS0.E_Factory) = 8
        .ColWidth(E_FPS0.E_Customer) = 8
        .ColWidth(E_FPS0.E_ToDayQty) = 8
        
        .ColWidth(E_FPS0.E_Mon1) = 8
        .ColWidth(E_FPS0.E_Mon2) = 8
        .ColWidth(E_FPS0.E_Mon3) = 8
        .ColWidth(E_FPS0.E_Mon4) = 8
        .ColWidth(E_FPS0.E_Mon5) = 8
        .ColWidth(E_FPS0.E_Mon6) = 8
        .ColWidth(E_FPS0.E_Mon7) = 8
        .ColWidth(E_FPS0.E_Mon8) = 8
        .ColWidth(E_FPS0.E_Mon9) = 8
        .ColWidth(E_FPS0.E_Mon10) = 8
        .ColWidth(E_FPS0.E_Mon11) = 8
        .ColWidth(E_FPS0.E_Mon12) = 8
        .ColWidth(E_FPS0.E_sum) = 8
        

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
        
          
        .SetText E_FPS1.E_ID, 0, "序号"
        .SetText E_FPS1.E_Wo, 0, "工单号"
        .SetText E_FPS1.E_WaferID, 0, "WaferId"
        .SetText E_FPS1.E_CompleteFlag, 0, "完成标志"
        .SetText E_FPS1.E_TotalDie, 0, "TotalDie数量"
        .SetText E_FPS1.E_GoodDie, 0, "GoodDie数量"
        .SetText E_FPS1.E_WaferLot, 0, "WaferLot"
        .SetText E_FPS1.E_MarkingCode, 0, "MarkingCode"
        
        
        .ColWidth(E_FPS1.E_ID) = 10
        .ColWidth(E_FPS1.E_Wo) = 10
        .ColWidth(E_FPS1.E_WaferID) = 10
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
End Sub

Private Sub Label9_Click()

End Sub

