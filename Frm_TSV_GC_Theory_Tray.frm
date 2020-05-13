VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form Frm_TSV_GC_Theory_Tray 
   Caption         =   "GC Tray理论用量报表"
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
      Width           =   19815
      _Version        =   196613
      _ExtentX        =   34951
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
      SpreadDesigner  =   "Frm_TSV_GC_Theory_Tray.frx":0000
      TextTip         =   2
   End
End
Attribute VB_Name = "Frm_TSV_GC_Theory_Tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
    E_ID = 0                'id
    E_Factory            '事业部
    E_CustomerPT           '客户机种
    E_HTPT                  '厂内机种
    E_RateNormalGood       '用量比率Good(Normal)
    E_RateNormalNG         '用量比率NG(Normal)
    E_RateWLAGood       '用量比率Good(WLA)
    E_RateWLANG         '用量比率NG(WLA)
    
    
    E_InvQtyGoodQty    'Good Tray(本周库存量)
    E_InvQtyNGQty      'NG Tray(本周库存量)
    
    E_InQtyGoodQty    'Good Tray(本周来料)
    E_InQtyNGQty      'NG Tray(本周来料)
    
    E_UseQtyGoodQty    'Good Tray(本周使用量)
    E_UseQtyNGQty      'NG Tray(本周使用量)
    
    
    E_DiffQtyGoodQty    '下周Good Tray
    E_DiffQtyNGQty      '下周NG Tray
    
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

Dim wlaFlagTemp As String
Dim productTemp As String
Dim rowIdTemp As Integer
Dim deptTemp As String

Dim cusPTTemp As String
Dim htPTTemp As String

Dim rate_NormalGood As Integer
Dim rate_NormalNG As Integer

Dim rate_WlaGood As Integer
Dim rate_WlaNG As Integer

Dim InvGoodQty As Long
Dim InvNGQty As Long

Dim InGoodQty As Long
Dim InNGQty As Long

Dim UseGoodQty As Long
Dim UseNGQty As Long

Dim NextGoodQty As Long
Dim NextNGQty As Long
Dim beginDate As Date
Dim endDate As Date

                
 endDate = Format(CDate(Now - Weekday(Now) + 2), "YYYY-MM-DD")
 beginDate = endDate - 7
 
      

Dim j As Integer
Dim qtyMonTemp As Long


'strSql = "  select wlaflag, productname,id,deptname,gcpt,htpt,rate_normalgood,rate_normalng,rate_wlagood,rate_wlang,goodtrayweekqty,ngtrayweekqty  " & _
'         "  From TSV_GCTrayRptSet where flag='Y' order by id "
'
         
         
strSql = "select wlaflag, productname,id,deptname,gcpt,htpt,rate_normalgood,rate_normalng,rate_wlagood,rate_wlang, " & _
         " GOODTRAYWEEKQTY , NGTRAYWEEKQTY, GOODTRAYInQTY, NGTRAYInQTY, GOODTRAYUseQTY, NGTRAYUseQTY, GOODTRAYNextWEEKQTY, NGTRAYNextWEEKQTY " & _
         " From TSV_GCTRAYLiLunRPT where flag='Y' order by id"
          
 If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

      i = 0
      With fps(0)
            .MaxRows = Rs.RecordCount + 1
            
              For m = 1 To Rs.RecordCount
              
                wlaFlagTemp = Rs("wlaflag").Value
                productTemp = Rs("productname").Value
                
                rowIdTemp = CInt(Rs("id").Value)
                deptTemp = Rs("deptname").Value
                cusPTTemp = Rs("gcpt").Value
                htPTTemp = Rs("htpt").Value
                
                rate_NormalGood = CInt(Rs("rate_normalgood").Value)
                rate_NormalNG = CInt(Rs("rate_normalng").Value)
                rate_WlaGood = CInt(Rs("rate_wlagood").Value)
                rate_WlaNG = CInt(Rs("rate_wlang").Value)
                
                
               InvGoodQty = CLng(Rs("GOODTRAYWEEKQTY").Value)
               
               InvGoodQty = GetGCTrayThInvQty(cusPTTemp, htPTTemp, "GD")
               
               InvNGQty = CLng(Rs("NGTRAYWEEKQTY").Value)
               
              InvNGQty = GetGCTrayThInvQty(cusPTTemp, htPTTemp, "NG")
               
               
               
               InGoodQty = CLng(Rs("GOODTRAYInQTY").Value)
               '2015-11-04 jiayun add 抓取erp
               ' 厂内机种,上周一 日期，本周一日期
               
               InGoodQty = GetGCTrayERPInQty(htPTTemp, beginDate, endDate, "GD")
               
               
               InNGQty = CLng(Rs("NGTRAYInQTY").Value)
               InNGQty = GetGCTrayERPInQty(htPTTemp, beginDate, endDate, "NG")
                  
               
               UseGoodQty = CLng(Rs("GOODTRAYUseQTY").Value)
               
               '2015-11-16 jiayun add  本周使用量
               ' 应该是 ( 上周一WIP+上周wafer来料-本周一WIP）*BOM，含PP之前
               
               UseGoodQty = GetGCTrayERPWeekUseQty(cusPTTemp, beginDate, endDate, "GD", rate_NormalGood, rate_WlaGood)
               
               UseNGQty = CLng(Rs("NGTRAYUseQTY").Value)
               UseNGQty = GetGCTrayERPWeekUseQty(cusPTTemp, beginDate, endDate, "GD", rate_NormalNG, rate_WlaNG)
               
               
               NextGoodQty = CLng(Rs("GOODTRAYNextWEEKQTY").Value)
               NextGoodQty = InvGoodQty + InGoodQty - UseGoodQty

               NextNGQty = CLng(Rs("NGTRAYNextWEEKQTY").Value)
               NextNGQty = InvNGQty + InNGQty - UseNGQty
                  
                i = i + 1
                .Row = i
                .Col = 1
                .Text = deptTemp
                
                .Row = i
                .Col = 2
                .Text = cusPTTemp
                
                .Row = i
                .Col = 3
                .Text = htPTTemp
                
                
                .Row = i
                .Col = 4
                .Text = rate_NormalGood
                
                .Row = i
                .Col = 5
                .Text = rate_NormalNG
                
                .Row = i
                .Col = 6
                .Text = rate_WlaGood
                
                .Row = i
                .Col = 7
                .Text = rate_WlaNG
                
                
                .Row = i
                .Col = 8
                .Text = InvGoodQty
                
                
                .Row = i
                .Col = 9
                .Text = InvNGQty
                
                
                .Row = i
                .Col = 10
                .Text = InGoodQty
                
                .Row = i
                .Col = 11
                .Text = InNGQty
                
       ' ------------
                 .Row = i
                .Col = 12
                .Text = UseGoodQty
                
                
                .Row = i
                .Col = 13
                .Text = UseNGQty
                
                
                .Row = i
                .Col = 14
                .Text = NextGoodQty
                
                .Row = i
                .Col = 15
                .Text = NextNGQty
                
                
'
'                '查询Wafer数量
'
'                Dim waferNormalQty As Long
'                Dim waferWlaQty As Long
'
'                Dim TrayGoodWlaQty As Long
'                Dim TrayNGWlaQty As Long
'                Dim normalSetQty As Long
'                Dim tempNormail1 As Long
'
'
'                waferNormalQty = 0
'                waferWlaQty = 0
'
'                TrayGoodWlaQty = 0
'                TrayNGWlaQty = 0
'
'                normalSetQty = 0
'                tempNormail1 = 0
'
'
''                InvGoodQty = 0
''                InvNGQty = 0
'
'
'
  '             waferNormalQty = GetQtyTotalWafer(cusPTTemp)
'
'                If wlaFlagTemp = "Y" Then
'                   '查询normail,wla
'                   waferWlaQty = 0
'
''                   TrayGoodWlaQty = rate_WlaGood * waferWlaQty
''                   TrayNGWlaQty = rate_WlaNG * waferWlaQty
'
'
'                   normalSetQty = GetQtySETNormalWafer(cusPTTemp)
'
'                   tempNormail1 = normalSetQty
'
'                .Row = i
'                .Col = 8
'                .Text = tempNormail1
'
'                .Row = i
'                .Col = 9
'                .Text = waferNormalQty - tempNormail1
'
'
'                waferWlaQty = waferNormalQty - tempNormail1
'
'                waferNormalQty = tempNormail1
'
'                   TrayGoodWlaQty = rate_WlaGood * waferWlaQty
'                   TrayNGWlaQty = rate_WlaNG * waferWlaQty
'
'
'
'                 Else
'
'
'                .Row = i
'                .Col = 8
'                .Text = waferNormalQty
'
'                .Row = i
'                .Col = 9
'                .Text = waferWlaQty
'
'
'
'                End If
'
'
'
'
'                .Row = i
'                .Col = 10
'                .Text = CStr(rate_NormalGood * waferNormalQty)
'
'                .Row = i
'                .Col = 11
'                .Text = CStr(rate_NormalNG * waferNormalQty)
'
'
'                 .Row = i
'                .Col = 12
'                .Text = TrayGoodWlaQty
'
'                .Row = i
'                .Col = 13
'                .Text = TrayNGWlaQty
'
'                ' 取本周库存量库存量
''                  InvGoodQty = 0
''                  InvNGQty = 0
'
'                .Row = i
'                .Col = 14
'                .Text = InvGoodQty
'
'                .Row = i
'                .Col = 15
'                .Text = InvNGQty
'
'                .Row = i
'                .Col = 16
'                .Text = CStr(InvGoodQty - rate_NormalGood * waferNormalQty - TrayGoodWlaQty)
'
'                .Row = i
'                .Col = 17
'                .Text = CStr(InvNGQty - rate_NormalNG * waferNormalQty - TrayNGWlaQty)
                
              
              

'
'
'                 For j = 1 To 12
'
'                 qtyMonTemp = GetQtyMDMonth(factoryIdTemp, custIDTemp, j, custNameTemp)
'
'
'                .Row = i
'                .Col = 3 + j
'                .Text = qtyMonTemp
'
'                 Next j
                 
                
                
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

'Dim iclo As Integer
'Dim irow As Integer
'Dim qtyAll As Long
'Dim qty01Temp As Long
'
'Dim TSVQtyALL As Long
'Dim BumpingQtyALL As Long
'Dim T12QtyALL As Long
'
'Dim TSVQtyRID As Integer
'Dim BumpingQtyRID As Integer
'Dim T12QtyRID As Integer
'
'
'TSVQtyALL = 0
'BumpingQtyALL = 0
'T12QtyALL = 0
'
'TSVQtyRID = 0
'BumpingQtyRID = 0
'T12QtyRID = 0
'
'With fps(0)
'
'.MaxRows = .MaxRows + 3
'
'
'
'  For irow = 1 To .MaxRows - 4
'        .Row = irow
'        .Col = 1
'
'        If .Text = "TSV" Then
'            TSVQtyRID = irow
'        End If
'
'        If .Text = "Bumping" Then
'            BumpingQtyRID = irow
'        End If
'
'        If .Text = "12寸" Then
'            T12QtyRID = irow
'        End If
'
'
'
'  Next irow
'
'End With
'
'
'
'With fps(0)
' .Row = .MaxRows - 3
'.Col = 1
'.Text = "TSV汇总："
'
' .Row = .MaxRows - 2
'.Col = 1
'.Text = "Bumping汇总："
'
' .Row = .MaxRows - 1
'.Col = 1
'.Text = "12寸汇总："
'
' .Row = .MaxRows
'.Col = 1
'.Text = "所有汇总："
'
'
'
'For iclo = 3 To 15
'    qtyAll = 0
'
'    TSVQtyALL = 0
'    BumpingQtyALL = 0
'    T12QtyALL = 0
'
'
'     For irow = 1 To TSVQtyRID
'        .Row = irow
'        .Col = iclo
'        qty01Temp = CLng(.Text)
'
'        TSVQtyALL = TSVQtyALL + qty01Temp
'
'     Next irow
'
'
'       For irow = TSVQtyRID + 1 To BumpingQtyRID
'        .Row = irow
'        .Col = iclo
'        qty01Temp = CLng(.Text)
'
'        BumpingQtyALL = BumpingQtyALL + qty01Temp
'
'      Next irow
'
'    For irow = BumpingQtyRID + 1 To .MaxRows - 4
'
'        .Row = irow
'        .Col = iclo
'        qty01Temp = CLng(.Text)
'
'        T12QtyALL = T12QtyALL + qty01Temp
'
'      Next irow
'
'
'
'
'
'      .Row = .MaxRows - 3
'      .Col = iclo
'      .Text = TSVQtyALL
'
'      .Row = .MaxRows - 2
'      .Col = iclo
'      .Text = BumpingQtyALL
'
'      .Row = .MaxRows - 1
'      .Col = iclo
'      .Text = T12QtyALL
'
'       .Row = .MaxRows
'      .Col = iclo
'      .Text = TSVQtyALL + BumpingQtyALL + T12QtyALL
'
'
'Next iclo
'
'End With
'
'
'
'
''汇总最后一列：
'
'
'
'With fps(0)
'
'For iclo = 1 To .MaxRows
'    qtyAll = 0
'
'     For irow = 4 To 15
'        .Row = iclo
'        .Col = irow
'        qty01Temp = CLng(.Text)
'
'        qtyAll = qtyAll + qty01Temp
'
'     Next irow
'
'      .Row = iclo
'      .Col = 16
'      .Text = qtyAll
'
'
'Next iclo
'
'End With


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
.SetText E_FPS0.E_Factory, 0, "事业部产品"
.SetText E_FPS0.E_CustomerPT, 0, "格科型号"
.SetText E_FPS0.E_HTPT, 0, "华天型号"

.SetText E_FPS0.E_RateNormalGood, 0, "用量比率Good(Normal)"
.SetText E_FPS0.E_RateNormalNG, 0, "用量比率NG(Normal)"
.SetText E_FPS0.E_RateWLAGood, 0, "用量比率Good(WLA)"
.SetText E_FPS0.E_RateWLANG, 0, "用量比率NG(WLA)"


.SetText E_FPS0.E_InvQtyGoodQty, 0, "Good Tray(本周库存量)"
.SetText E_FPS0.E_InvQtyNGQty, 0, "NG Tray(本周库存量)"

.SetText E_FPS0.E_InQtyGoodQty, 0, "本周GoodTray来料"
.SetText E_FPS0.E_InQtyNGQty, 0, "本周NGTray来料"

.SetText E_FPS0.E_UseQtyGoodQty, 0, "GoodTray本周使用量"
.SetText E_FPS0.E_UseQtyNGQty, 0, "NGTray本周使用量"

.SetText E_FPS0.E_DiffQtyGoodQty, 0, "下周GoodTray"
.SetText E_FPS0.E_DiffQtyNGQty, 0, "下周NGTray"
    


        .ColWidth(E_FPS0.E_ID) = 5
        .ColWidth(E_FPS0.E_Factory) = 9
        .ColWidth(E_FPS0.E_CustomerPT) = 9
        .ColWidth(E_FPS0.E_HTPT) = 9
        
        .ColWidth(E_FPS0.E_RateNormalGood) = 9
        .ColWidth(E_FPS0.E_RateNormalNG) = 9
        .ColWidth(E_FPS0.E_RateWLAGood) = 9
        .ColWidth(E_FPS0.E_RateWLANG) = 9
        
        .ColWidth(E_FPS0.E_InvQtyGoodQty) = 9
        .ColWidth(E_FPS0.E_InvQtyNGQty) = 9
        
        
        .ColWidth(E_FPS0.E_InQtyGoodQty) = 10
        .ColWidth(E_FPS0.E_InQtyNGQty) = 10
        .ColWidth(E_FPS0.E_UseQtyGoodQty) = 10
        .ColWidth(E_FPS0.E_UseQtyNGQty) = 10
        
        
        .ColWidth(E_FPS0.E_DiffQtyGoodQty) = 10
        .ColWidth(E_FPS0.E_DiffQtyNGQty) = 10
        

        .RowHeight(0) = 30
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

