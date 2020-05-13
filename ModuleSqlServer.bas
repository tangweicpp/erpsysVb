Attribute VB_Name = "ModuleSqlServer"




Public INIadoCon As New ADODB.Connection
Public INIadoCon2 As New ADODB.Connection
Public cmd As New ADODB.Command
Public rs As New ADODB.Recordset

Public Sub INIConnectSTART()
   Const strSrvName As String = "10.160.1.13"   '数据服务器名称
   Const strDbName As String = "erpbase"   '数据库名称
   Const strUID As String = "sa"      '登录用户名称
   Const strPSWD As String = "ksxtDB"     '登录密码"
   Dim INIstrCnn As String '连接字符串
   

    Set INIadoCon = New ADODB.Connection
    INIstrCnn = "driver={SQL Server};server=" & strSrvName & ";UID=" & strUID & "; " & _
    "pwd=" & strPSWD & ";database=" & strDbName & ""
    INIadoCon.CursorLocation = adUseClient
    INIadoCon.ConnectionTimeout = 100
    INIadoCon.CommandTimeout = 100
    INIadoCon.Open INIstrCnn
    If INIadoCon.State = 0 Then
       MsgBox "错误:" & Err.DESCRIPTION & vbCrLf & " 解决方法请寻求有关帮助。", vbExclamation, "系统"
       Exit Sub
    End If
End Sub



Public Sub INIConnectSTART2()
   Const strSrvName As String = "10.160.1.13"   '数据服务器名称
   Const strDbName As String = "erpdata"   '数据库名称
   Const strUID As String = "sa"      '登录用户名称
   Const strPSWD As String = "ksxtDB"     '登录密码"
   Dim INIstrCnn As String '连接字符串
   

    Set INIadoCon2 = New ADODB.Connection
    INIstrCnn = "driver={SQL Server};server=" & strSrvName & ";UID=" & strUID & "; " & _
    "pwd=" & strPSWD & ";database=" & strDbName & ""
    INIadoCon2.CursorLocation = adUseClient
    INIadoCon2.ConnectionTimeout = 0
    INIadoCon2.CommandTimeout = 100
    INIadoCon2.Open INIstrCnn
    If INIadoCon2.State = 0 Then
       MsgBox "错误:" & Err.DESCRIPTION & vbCrLf & " 解决方法请寻求有关帮助。", vbExclamation, "系统"
       Exit Sub
    End If
End Sub


'Insert into DB
Public Function AddSql2(cmdStr As String) As Long
If INIadoCon.State = 0 Then
INIConnectSTART
End If

cmd.ActiveConnection = INIadoCon
cmd.CommandText = cmdStr
cmd.CommandType = adCmdText
cmd.Execute SD
AddSql2 = SD

End Function

Public Function GetProductChildBom(productNameTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select  材料规范编号,物料编号  from  [erpdata].[dbo].[TSVtblSetMRule] where 物料编号='" + productNameTemp + "'"
       
Set RSResult = getSqlStr(cmdStr)

Set GetProductChildBom = RSResult
End Function




Public Function GetDiaoBoList(inqboxtemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
      

cmdStr = "select b.trayqboxnumber from  [erpdata].[dbo].[TblTSV_INBOX_DETAILS] a ,[erpdata].[dbo].[TblTSV_Tray_details] b " & _
" where  a.labeltype='INQbox' and a.containername in ('" + productNameTemp + "') " & _
" and b.trayqboxnumber=a.SUBCONTAINERNAME"

         
Set RSResult = getSqlStr(cmdStr)
Set GetDiaoBoList = RSResult
End Function





Public Function GetMDSOCustomer(fiertPtTemp As String, typaName As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'cmdStr = "select  材料规范编号,物料编号  from  [erpdata].[dbo].[TSVtblSetMRule] where 物料编号='" + productNameTemp + "'"


If typaName = "TSV" Then

cmdStr = " SELECT  distinct rtrim(a.客户代码) as custid FROM   [erpdata].[dbo].[tblStockMove] a ,[erpdata].[dbo].[tblStockMovesub] b " & _
" where a.操作日期>='2015-01-01' and a.产线标记=1 and a.单据类型=1 and a.实发良品数>0 " & _
" and b.单据编号=a.单据编号 and b.工单号=a.工单号 and left(b.料号,2)='" + fiertPtTemp + "' " & _
" and  right(RTRIM(b.料号),2) in ('CF','CP','CM') order by 1 "

Else

cmdStr = " SELECT  distinct rtrim(a.客户代码) as custid FROM   [erpdata].[dbo].[tblStockMove] a ,[erpdata].[dbo].[tblStockMovesub] b " & _
" where a.操作日期>='2015-01-01' and a.产线标记=1 and a.单据类型=1 and a.实发良品数>0 " & _
" and b.单据编号=a.单据编号 and b.工单号=a.工单号 and left(b.料号,2)='" + fiertPtTemp + "' " & _
" and  right(RTRIM(b.料号),2) in ('BR','BS','BL') order by 1 "

End If

   
Set RSResult = getSqlStr(cmdStr)

Set GetMDSOCustomer = RSResult
End Function

Public Function GetGCNGRpt(billNoTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = billNoTemp



Set RSResult = getSqlStr2(cmdStr)

Set GetGCNGRpt = RSResult
End Function


Public Function GetAAMPNDataSQL(billNoTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = billNoTemp



Set RSResult = getSqlStr2(cmdStr)

Set GetAAMPNDataSQL = RSResult
End Function
Public Function GetShelfInfoSQL(sql As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

Set RSResult = getSqlStr2(sql)

Set GetShelfInfoSQL = RSResult
End Function

Public Function GetOrderName(billNoTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = billNoTemp



Set RSResult = getSqlStr2(cmdStr)

Set GetOrderName = RSResult
End Function





Public Function GetProductChildBomAdd(productNameTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select b.物料编号, b.物料名称,b.规格型号,b.计量单位名称,b.型号  from dbo.tblSmainM2 b where b.料号='" + productNameTemp + "'"
 
Set RSResult = getSqlStr(cmdStr)

Set GetProductChildBomAdd = RSResult
End Function




Public Function GetGCCTSqlData(beginDateTemp As String, endDateTemp As String, custernameTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'cmdStr = " select b.物料编号, b.物料名称,b.规格型号,b.计量单位名称,b.型号  from dbo.tblSmainM2 b where b.料号='" + productNameTemp + "'"



'cmdStr = "  SELECT   distinct  RTRIM(ltrim(b.流程卡编号))+ '-A' as containername, RTRIM(ltrim(b.流程卡编号)) as waferid ,a.操作日期 as cdate  FROM   [erpdata].[dbo].[tblStockMove] a ,[erpdata].[dbo].[tblStockMovesub] b " & _
'" where a.客户代码='" + custernameTemp + "' and a.操作日期>='" + beginDateTemp + "' " & _
'" and a.操作日期<'" + endDateTemp + "' and b.单据编号 like 'F%' " & _
'" and b.单据编号=a.单据编号 and b.工单号=a.工单号 and a.实发良品数>0 "




cmdStr = "  select containername,waferid,cdate,MIN(操作日期) as 操作日期 from ( " & _
 " SELECT   distinct  RTRIM(ltrim(b.流程卡编号))+ '-A' as containername, RTRIM(ltrim(b.流程卡编号)) as waferid ,a.操作日期 as cdate ,c.操作日期 " & _
" FROM   [erpdata].[dbo].[tblStockMove] a ,[erpdata].[dbo].[tblStockMovesub] b , erpbase.dbo.tblStoEntrybill  c " & _
" where a.客户代码='" + custernameTemp + "' and a.操作日期>='" + beginDateTemp + "'  and a.操作日期<'" + endDateTemp + "'   and b.单据编号 like 'F%' " & _
" and b.单据编号=a.单据编号 and b.工单号=a.工单号 and a.实发良品数>0  and c.批号=b.工单号) X group by containername,waferid,cdate "







Set RSResult = getSqlStr(cmdStr)

Set GetGCCTSqlData = RSResult
End Function




Public Function GetCOGBaseData(billTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


  cmdStr = " select distinct a.CartonNO ,rtrim(ltrim(c.箱号)) as qboxTemp from [erpdata].[dbo].[GR_GC_DetailHistory] a ,[erpdata].[dbo].[tblPackTreeInf] b ,[erpdata].[dbo].[tblPackTreeInf] c " & _
          " where a.单据编号='" + billTemp + "'  and b.箱号=a.CartonNO and c.上级序号=b.序号 order by 2 "
   
Set RSResult = getSqlStr(cmdStr)

Set GetCOGBaseData = RSResult
End Function
Public Function GetCOGBaseData2(billTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset



cmdStr = " SELECT distinct rtrim(ltrim(uU.箱号)) FROM erpdata..tblStockMoveSUB uU WHERE UU.单据编号='" + billTemp + "'"

 
Set RSResult = getSqlStr(cmdStr)

Set GetCOGBaseData2 = RSResult
End Function



Public Sub delCOGRptInt01()
Dim cmdStr As String
Dim cmdStr2 As String

cmdStr = "   delete from  [erpdata].[dbo].[GR_COG_LV_Data]   "
                                                  
AddSql2 (cmdStr)

End Sub





Public Function GetSqlServerStr(cmdStr As String) As String
    Dim resut As New ADODB.Recordset
    
If INIadoCon.State = 0 Then
INIConnectSTART
End If

    resut.Open cmdStr, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If resut.RecordCount > 0 Then
        GetSqlServerStr = resut.Fields(0).Value
    Else
        GetSqlServerStr = ""
    End If
End Function



Public Function GetGCRptFileNo(cmdStr As String) As String
    Dim resut As New ADODB.Recordset
    
If INIadoCon.State = 0 Then
INIConnectSTART
End If

    resut.Open cmdStr, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If resut.RecordCount > 0 Then
        GetGCRptFileNo = resut.Fields(0).Value
    Else
        GetGCRptFileNo = ""
    End If
End Function


'Public Function GetGCRptFileNo_pj(cmdStr As String) As ADODB.Recordset
'    Dim resut As New ADODB.Recordset
'
'If INIadoCon.State = 0 Then
'INIConnectSTART
'End If
'
'    resut.open cmdStr, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
''    If resut.RecordCount > 0 Then
''        GetGCRptFileNo = resut.fields(0).Value
''    Else
''        GetGCRptFileNo = ""
''    End If
'GetGCRptFileNo_pj = resut
'End Function



'Public Function Get37InQboxTxt(lotid As String, qboxnumberTemp As String, containerTemp As String, seqTemp As String) As ADODB.Recordset
'
''取工单后四位
'Dim cmdStr As String
'Dim RSResult As New ADODB.Recordset
'
''            cmdStr = " SELECT workorderattr3||'-2.5' || ','||wafernumber||',' ||imager_customer_rev||','  " & _
''         " ||'SH'||','||count(waferscribenumber1)||','|| LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ';')), ';') ||','||wafernumber || '.' || min(waferscribenumber1) ||','|| '" & qboxnumberTemp & "' as txt " & _
''         " FROM (select  f.workorderattr3 ,sum(b.qty) as QTY1,a.wafernumber,'SH' as comp_code,g.imager_customer_rev,substr(a.waferscribenumber, -2, 2) waferscribenumber1," & _
''         " ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN " & _
''         " from a_lotwafers a,container b,a_lotattributes c,product d,productbase e,mfgorder f,customeroitbl_test g ,  mappingdatatest h  " & _
''         " Where a.containerid = b.containerid And b.containerid = c.containerid And d.productbaseid = e.productbaseid " & _
''         " and f.mfgordername = a.workordername and b.productid = d.productid and b.containername in ('" & lotid & "') and c.customername = 'GC'  and h.lotid=g.source_batch_id and h.filename=g.id  " & _
''         " and g.source_batch_id=a.wafernumber and h.substrateid=a.waferscribenumber GROUP BY f.workorderattr3 ,a.wafernumber,g.comp_code,g.imager_customer_rev ,a.waferscribenumber) " & _
''         " START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
''         " GROUP BY workorderattr3,QTY1,wafernumber,comp_code,imager_customer_rev  order by wafernumber"
'
'
'cmdStr = " select a.customerpt +','+a.customerlotid +','+'1T'+a.customerlotid +','+a.customerpt +','+'1P'+a.customerpt +','+min(a.podatecode) +','+min(a.podatecode)+',' " & _
'" +max(a.htlotid)+'" & seqTemp & "'+','+'S'+max(a.htlotid)+'" & seqTemp & "' +','+rtrim(sum(qty)) +','+'Q'+rtrim(sum(qty)) +','" & _
'" +min(a.htdatecode) +','+min(a.htdatecode) " & _
'" from  [erpdata].[dbo].[TblTSV_Tray_details] a where trayqboxnumber in ('" & lotid & "') " & _
'" group by a.customerpt ,a.customerlotid ,a.customerlotid ,a.customerpt ,a.podatecode ,a.htlotid,a.htdatecode"
'
'
'
'
''RSResult = GetGCRptFileNo(cmdStr)
''RSResult =
'Get37InQboxTxt = GetGCRptFileNo_pj(cmdStr)
'End Function


Public Function GetServerSeq(cmdStr As String) As Long
    Dim resut As New ADODB.Recordset
    
If INIadoCon.State = 0 Then
INIConnectSTART
End If

    resut.Open cmdStr, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    
    If resut.RecordCount > 0 Then
    
    'GetServerSeq = CLng(resut.fields("ID").Value)
     GetServerSeq = "" & CStr(IIf(IsNull(resut.Fields("ID").Value), "0", resut.Fields("ID").Value))
    
    Else
    GetServerSeq = 0
    
    End If
    
End Function



Public Function GetQtyMDSOMonth(custNameTemp As String, ptFirstTemp As String, monThTemp As String, typeTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long


 cmdStr = "select ID= [erpdata].[dbo].[Get_TSV_SO_PieceQty] ('" & custNameTemp & "','" & ptFirstTemp & "','" & monThTemp & "','" & typeTemp & "') "


RSResult = GetServerSeq(cmdStr)
GetQtyMDSOMonth = RSResult
End Function



Public Function Get37BigQboxID(qboxtemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long


 cmdStr = "select 序号 as ID from [erpdata].[dbo].[tblStockNumTree] a where a.箱号='" & qboxtemp & "' and Memo='37' "


RSResult = GetServerSeq(cmdStr)
Get37BigQboxID = RSResult
End Function


Public Function Get37BigQboxIDV1(qboxtemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long


 cmdStr = "select 序号 as ID from [erpdata].[dbo].[tblPackTreeInf] a where a.箱号='" & qboxtemp & "' and Memo='37' "

RSResult = GetServerSeq(cmdStr)
Get37BigQboxIDV1 = RSResult
End Function


Public Function Get37BigQboxQty(qboxtemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long


 'cmdStr = " select  top 1 Qty  as ID  from  [erpdata].[dbo].[TblTSV_OutBOX_DETAILS] where CONTAINERNAME='" & qboxTemp & "' "
 cmdStr = " select  SUM(Qty)  as ID  from  [erpdata].[dbo].[TblTSV_OutBOX_DETAILS] where CONTAINERNAME='" & qboxtemp & "'  GROUP BY CONTAINERNAME"
 
RSResult = GetServerSeq(cmdStr)
Get37BigQboxQty = RSResult
End Function







Public Function GetGCTrayERPInQty(ptFirstTemp As String, beginDate As Date, endDate As Date, typeTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long


' cmdStr = "select ID= [erpdata].[dbo].[Get_TSV_SO_PieceQty] ('" & custNameTemp & "','" & ptFirstTemp & "','" & monThTemp & "','" & typeTemp & "') "



 cmdStr = " SELECT ID=isnull(SUM(a.实收),0) FROM erpbase.dbo.tblStoEntrybill a ,dbo.tblSmainM2 b,[ERPBASE].[dbo].[tblTSVGCTraySetPtNo] c Where b.物料编号 = a.物料编号 " & _
" and a.操作日期>='" & beginDate & "' and a.操作日期<'" & endDate & "' " & _
" and a.仓库编号='21' and b.料号=c.traypt and  c.htpt='" & ptFirstTemp & "'  and c.flag='Y'  and c.traytype='" & typeTemp & "'  "




RSResult = GetServerSeq(cmdStr)
GetGCTrayERPInQty = RSResult
End Function



Public Function GetGCTrayERPWeekUseQty(ptFirstTemp As String, beginDate As Date, endDate As Date, typeTemp As String, goodBomRate As Integer, goodWlaRate) As Long
Dim cmdStr As String
Dim RSResult As Long
Dim lastWeekQty As Long
Dim lastWeekWoQty As Long
Dim thisWeekQty As Long
Dim waferQtyTemp As Long
Dim wlaFlag As Boolean

'-------------wla
Dim lastWeekQty_WLA As Long
Dim lastWeekQty_Normal As Long
Dim lastWeekQty_All As Long


Dim thisWeekQty_WLA As Long
Dim thisWeekQty_Normal As Long
Dim thisWeekQty_All As Long





'-02 来料
Dim lastWeekWoQty_WLA As Long
Dim lastWeekWoQty_Normal As Long
Dim lastWeekWoQty_All As Long



lastWeekQty = 0
lastWeekWoQty = 0
thisWeekQty = 0
waferQtyTemp = 0



RSResult = 0
lastWeekQty_WLA = 0
lastWeekQty_Normal = 0
lastWeekQty_All = 0


thisWeekQty_WLA = 0
thisWeekQty_Normal = 0
thisWeekQty_All = 0

lastWeekWoQty_WLA = 0
lastWeekWoQty_Normal = 0
lastWeekWoQty_All = 0


wlaFlag = False


'查询理论用量中的，本周使用量

' 应该是 ( 上周一WIP+上周wafer来料-本周一WIP）*BOM，含PP之前
'ptFirstTemp 为客户机种号



' cmdStr = "select ID= [erpdata].[dbo].[Get_TSV_SO_PieceQty] ('" & custNameTemp & "','" & ptFirstTemp & "','" & monThTemp & "','" & typeTemp & "') "



' cmdStr = " SELECT ID=isnull(SUM(a.实收),0) FROM erpbase.dbo.tblStoEntrybill a ,dbo.tblSmainM2 b,[ERPBASE].[dbo].[tblTSVGCTraySetPtNo] c Where b.物料编号 = a.物料编号 " & _
'" and a.操作日期>='" & beginDate & "' and a.操作日期<'" & endDate & "' " & _
'" and a.仓库编号='21' and b.料号=c.traypt and  c.htpt='" & ptFirstTemp & "'  and c.flag='Y'  and c.traytype='" & typeTemp & "'  "

'2015-11-18 jiayun add 根据客户机种号，看它分不分WLA

If ptFirstTemp = "GC0310" Or ptFirstTemp = "GC0312" Or ptFirstTemp = "GC6123" Then
    wlaFlag = True
End If

If wlaFlag = False Then
        If typeTemp = "GD" Then
        
            lastWeekQty = GetGCTrayThLastWeekQty(ptFirstTemp, beginDate)
            lastWeekWoQty = GetGCTrayThLastWeekWoQty(ptFirstTemp, beginDate, endDate)
            
            thisWeekQty = GetGCTrayThLastWeekQty(ptFirstTemp, endDate)
            
            waferQtyTemp = lastWeekQty + lastWeekWoQty - thisWeekQty
            
            RSResult = waferQtyTemp * goodBomRate
        
        End If

Else
   '要分WLA   ( 上周一WIP+上周wafer来料-本周一WIP）*BOM，含PP之前    WLA+Normail
    '查出总数，再查出Normal数，减为WlA数
    
    '01 上周
    lastWeekQty_All = GetGCTrayThLastWeekQty(ptFirstTemp, beginDate)
    lastWeekQty_Normal = GetGCTrayThLastWeekQty_Normal(ptFirstTemp, beginDate)
    lastWeekQty_WLA = lastWeekQty_All - lastWeekQty_Normal
    
    
    '02 来料
    lastWeekWoQty_All = GetGCTrayThLastWeekWoQty(ptFirstTemp, beginDate, endDate)
    lastWeekWoQty_Normal = GetGCTrayThLastWeekWoQty_Normal(ptFirstTemp, beginDate, endDate)
    lastWeekWoQty_WLA = lastWeekWoQty_All - lastWeekWoQty_Normal
    

    '03 本周
    thisWeekQty_All = GetGCTrayThLastWeekQty(ptFirstTemp, endDate)
    thisWeekQty_Normal = GetGCTrayThLastWeekQty_Normal(ptFirstTemp, endDate)
    thisWeekQty_WLA = thisWeekQty_All - thisWeekQty_Normal
    
    RSResult = ((lastWeekQty_WLA + lastWeekWoQty_WLA - thisWeekQty_WLA) * goodWlaRate) + ((lastWeekQty_Normal + lastWeekWoQty_Normal - thisWeekQty_Normal) * goodBomRate)
   
End If


'RSResult = GetServerSeq(cmdStr)
GetGCTrayERPWeekUseQty = RSResult
End Function







Public Function GetQtyMDSOMonthDay(custNameTemp As String, ptFirstTemp As String, typeTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long


 cmdStr = "select ID= [erpdata].[dbo].[Get_TSV_SO_PieceDayQty] ('" & custNameTemp & "','" & ptFirstTemp & "','" & typeTemp & "') "


RSResult = GetServerSeq(cmdStr)
GetQtyMDSOMonthDay = RSResult
End Function




Public Function GetQty37Baofei(lotIdTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long



 
cmdStr = "select COUNT(waferid) as ID from  [erpdata].[dbo].[TSVtblBAOFEI] a where a.LOTID='" & lotIdTemp & "' "
 

RSResult = GetServerSeq(cmdStr)
GetQty37Baofei = RSResult
End Function



Public Function GetLotOutTimes(lotIdTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long



 
cmdStr = "select COUNT(distinct 单据编号 ) as ID from   [erpdata].[dbo].[tblStockMovesub] A WHERE a.工单号='" & lotIdTemp & "' "



RSResult = GetServerSeq(cmdStr)
GetLotOutTimes = RSResult
End Function





Public Function GetQty37OIAllQty(lotIdTemp As String, potemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long



 
cmdStr = "select  die_qty  as ID  from  [ERPBASE].[dbo].[tblCustomerOI] where SOURCE_BATCH_ID='" & lotIdTemp & "' and PO_NUM='" & potemp & "' "



 

RSResult = GetServerSeq(cmdStr)
GetQty37OIAllQty = RSResult
End Function








Public Function GetQty37OutQty(lotIdTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long



 
'cmdStr = "select COUNT(waferid) as ID from  [erpdata].[dbo].[TSVtblBAOFEI] a where a.LOTID='" & lotIDTemp & "' "

cmdStr = " SELECT COUNT(流程卡编号) as ID FROM [erpdata].[dbo].[tblStockMovesub]  WHERE 工单号='" & lotIdTemp & "'  and 单据编号 like 'F%' "
 

RSResult = GetServerSeq(cmdStr)
GetQty37OutQty = RSResult
End Function








Public Function GetQty37OutTimes(lotIdTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

'cmdStr = "select COUNT(waferid) as ID from  [erpdata].[dbo].[TSVtblBAOFEI] a where a.LOTID='" & lotIDTemp & "' "

'cmdStr = " SELECT COUNT(流程卡编号) as ID FROM [erpdata].[dbo].[tblStockMovesub]  WHERE 工单号='" & lotIDTemp & "'  and 单据编号 like 'F%' "

cmdStr = " select COUNT(*) as ID from [erpdata].[dbo].[tblStockMove]  a where a.客户代码='37' and 单据编号 like 'F%'  and a.工单号='" & lotIdTemp & "' " & _
 " and Convert(varchar(10),a.操作日期,120)=Convert(varchar(10),getdate(),120)"


RSResult = GetServerSeq(cmdStr)
GetQty37OutTimes = RSResult
End Function


Public Function GetQty37OutMaxBill(lotIdTemp As String) As String
'耞琌
Dim cmdStr As String
Dim slectResult As String
slectResult = False

'cmdStr = "SELECT COUNT(*)+1  From [erpdata].[dbo].[GR_GC_SendHistory] Where flag='Y' and [SendTime] = Convert(VarChar(10), Getdate(),111)"


cmdStr = " select max(a.单据编号) from [erpdata].[dbo].[tblStockMove]  a where a.客户代码='37' and 单据编号 like 'F%'  and a.工单号='" & lotIdTemp & "' " & _
 " and Convert(varchar(10),a.操作日期,120)=Convert(varchar(10),getdate(),120)"


slectResult = GetGCRptFileNo(cmdStr)
GetQty37OutMaxBill = slectResult
End Function






Public Function GetServerSeqDouble(cmdStr As String) As Double
    Dim resut As New ADODB.Recordset
    
If INIadoCon.State = 0 Then
INIConnectSTART
End If

    resut.Open cmdStr, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    GetServerSeqDouble = CDbl(resut.Fields("ID").Value)
End Function




Public Function getSqlStr(cmdStr As String) As ADODB.Recordset
    Dim resut As New ADODB.Recordset
    
    'INIadoCon.Close

If INIadoCon.State = 0 Then
INIConnectSTART
End If
    resut.Open cmdStr, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Set getSqlStr = resut
    
End Function



Public Function getSqlStr2(cmdStr As String) As ADODB.Recordset
    Dim resut As New ADODB.Recordset
    
    INIadoCon2.Close

If INIadoCon2.State = 0 Then
INIConnectSTART2
End If
    resut.Open cmdStr, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    Set getSqlStr2 = resut
End Function


Public Function JudgeGRBillNo(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select *  from [erpdata].[dbo].[GRdetailHistory] where 单据编号='" + billNoTemp + "'"
      
slectResult = SqlServerQueryStr(cmdStr)
JudgeGRBillNo = slectResult
End Function


Public Function Judge37TrayIn(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "SELECT * FROM  [erpdata].[dbo].[tblstocknumsub] WHERE  箱号= '" + billNoTemp + "'  "

slectResult = SqlServerQueryStr(cmdStr)
Judge37TrayIn = slectResult
End Function




Public Function Judge37InvType(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "SELECT * FROM  [erpdata].[dbo].[tblstocknumsub] WHERE  箱号= '" + billNoTemp + "' and 库房编号 in (36,37) "

slectResult = SqlServerQueryStr(cmdStr)
Judge37InvType = slectResult
End Function




Public Function Judge37InBoxIn(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "SELECT * From [erpdata].[dbo].[TblTSV_INBOX_DETAILS] where containername='" + billNoTemp + "'"
  

slectResult = SqlServerQueryStr(cmdStr)
Judge37InBoxIn = slectResult
End Function




Public Function Judge37ExistInBox(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "SELECT *  FROM [erpdata].[dbo].[TblTSV_OutBOX_DETAILS] where SUBCONTAINERNAME=  '" + billNoTemp + "' "

slectResult = SqlServerQueryStr(cmdStr)
Judge37ExistInBox = slectResult
End Function

Public Function Judge37ExistInBox1(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "SELECT *  FROM [erpdata].[dbo].[TblTSV_INBOX_DETAILS] where SUBCONTAINERNAME=  '" + billNoTemp + "' "

slectResult = SqlServerQueryStr(cmdStr)
Judge37ExistInBox1 = slectResult
End Function

Public Function Judge37ExistInBox2(inboxtemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "SELECT *  FROM [erpdata].[dbo].[TblTSV_INBOX_DETAILS] where CONTAINERNAME =  '" + inboxtemp + "' "

slectResult = SqlServerQueryStr(cmdStr)
Judge37ExistInBox2 = slectResult
End Function



Public Function Get37OutQboxTxt(LOTID As String, qboxnumberTemp As String, containerTemp As String, dnTemp As String) As String

Dim cmdStr As String
Dim RSResult As String

'cmdStr = " select a.customerpt ||','||a.customerlotid ||','||'1T'||a.customerlotid ||','||a.customerpt ||','||'1P'||a.customerpt ||','||min(a.podatecode) ||','||min(a.podatecode)||',' " & _
'" ||max(a.htlotid)||get_37_LableID('INQbox','" & containerTemp & "',max(a.htlotid))||','||'S'||max(a.htlotid)||get_37_LableID('INQbox','" & containerTemp & "',max(a.htlotid)) ||','||sum(qty) ||','||'Q'||sum(qty) ||','" & _
'" ||min(a.htdatecode) ||','||min(a.htdatecode) " & _
'" from  TSV_Tray_details a where trayqboxnumber in ('" & lotid & "') " & _
'" group by a.customerpt ,a.customerlotid ,a.customerlotid ,a.customerpt ,a.podatecode ,a.htlotid,a.htdatecode"
'


'cmdStr = "select ship.shiptoname+','+ship.shiptostreet1+','+ship.shiptostreet2+','+ship.shiptostreet3+','+ship.city+' '+ship.state+' '+ship.postalcode +','+" & _
'" ship.countrykey+','+ship.contactname+','+ship.phone+','+ship.delivery+','+'I'+ship.delivery +','+ship.purchasingdocno+','+'K'+ship.purchasingdocno +','+ " & _
'" ship.customerpartnumber+','+'P'+ship.customerpartnumber +','+a.customerpt+','+'Z'+a.customerpt+','+rtrim(sum(a.qty))+','+'Q'+rtrim(sum(a.qty)) +','+ " & _
'" ship.freightforwarder+','+'' +','+'' +','+'' +','+'COO:CHINA' +','+'CHINA'  " & _
'" from [ERPBASE].[dbo].[tblCustomerShippingUp] ship ,[erpdata].[dbo].[TblTSV_INBOX_DETAILS] a " & _
'" where a.labeltype='INQbox' and a.containername in ('" & lotid & "') and ship.batchnumber=a.customerlotid " & _
'" Group By ship.shiptoname,ship.shiptostreet1,ship.shiptostreet2,ship.shiptostreet3,ship.city,ship.state,ship.postalcode , " & _
'" ship.countrykey,ship.contactname,ship.phone,ship.delivery,'I'+ship.delivery ,ship.purchasingdocno, ship.customerpartnumber,a.customerpt,ship.freightforwarder+"" "
'
     
     cmdStr = "select left(ship.shiptoname,33) +','+ship.shiptostreet1+','+ship.shiptostreet2+','+ship.shiptostreet3+','+ship.city+' '+ship.state+' '+ship.postalcode +','+" & _
" ship.countrykey+','+ship.contactname+','+ship.phone+','+ship.delivery+','+'I'+ship.delivery +','+ left(ship.purchasingdocno,10) +','+'K'+left(ship.purchasingdocno,10) +','+ " & _
" left(ship.customerpartnumber,11) +','+'P'+ left(ship.customerpartnumber,11) +','+a.customerpt+','+'Z'+a.customerpt+','+rtrim(sum(c.qty))+','+'Q'+rtrim(sum(c.qty)) +','+ " & _
" ship.freightforwarder +','+'' +','+'' +','+'' +','+'COO:CHINA' +','+'CHINA'  " & _
" from [ERPBASE].[dbo].[tblCustomerShippingUp] ship ,[erpdata].[dbo].[TblTSV_INBOX_DETAILS] a ,[erpdata].[dbo].[TblTSV_Tray_details]  c  " & _
" where a.labeltype='INQbox' and a.containername in ('" & LOTID & "') and ship.batchnumber=c.customerlotid   and c.TRAYQBOXNUMBER=a.SUBCONTAINERNAME   and ship.delivery = '" & dnTemp & " '" & _
" Group By ship.shiptoname,ship.shiptostreet1,ship.shiptostreet2,ship.shiptostreet3,ship.city,ship.state,ship.postalcode , " & _
" ship.countrykey,ship.contactname,ship.phone,ship.delivery,'I'+ship.delivery ,ship.purchasingdocno, ship.customerpartnumber,a.customerpt,ship.freightforwarder  "

    
RSResult = GetGCRptFileNo(cmdStr)
Get37OutQboxTxt = RSResult
End Function

  



Public Function Judge37DnNom(billNoTemp As String, dnTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False


 cmdStr = " select ship.* from  [ERPBASE].[dbo].[tblCustomerShippingUp] ship ,[erpdata].[dbo].[TblTSV_INBOX_DETAILS] a " & _
" where a.labeltype='INQbox' and a.containername in ('" + billNoTemp + "') and ship.batchnumber=a.customerlotid " & _
" and ship.Delivery='" + dnTemp + "' "



    
slectResult = SqlServerQueryStr(cmdStr)
Judge37DnNom = slectResult
End Function








Public Function JudgeMPSBillNo(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select *  from [erpdata].[dbo].[tblStockMove] where 单据编号='" + billNoTemp + "'  and  客户代码 in ('68','70')  "
    
slectResult = SqlServerQueryStr(cmdStr)
JudgeMPSBillNo = slectResult
End Function



Public Function JudgeBomProduct(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select *  from [erpdata].[dbo].[tblSmainM2] where 料号='" + billNoTemp + "'"


slectResult = SqlServerQueryStr(cmdStr)
JudgeBomProduct = slectResult
End Function

Public Function JudgeBomPT(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select *  from [erpdata].[dbo].[tblSmainM2] where 料号='" + billNoTemp + "'"
      
slectResult = SqlServerQueryStr(cmdStr)
JudgeBomPT = slectResult
End Function


Public Function JudgeBomHeaderStaus(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select *  from [erpdata].[dbo].[TSVtblSetMRule] where 物料编号='" + billNoTemp + "'"

slectResult = SqlServerQueryStr(cmdStr)

JudgeBomHeaderStaus = slectResult
End Function




Public Function JudgeTSVBomAdd(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select *   FROM [erpdata].[dbo].[TSVtblBomSetup] where ProductName='" + billNoTemp + "' "

      
slectResult = SqlServerQueryStr(cmdStr)
JudgeTSVBomAdd = slectResult
End Function




Public Function JudgeHDWaferStatus(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select * from  [erpbase].[dbo].[tbltorec_wafer] a  where a.晶圆ID  ='" + billNoTemp + "'"
     
slectResult = SqlServerQueryStr(cmdStr)
JudgeHDWaferStatus = slectResult
End Function




Public Function JudgeSpecialGRBillNo(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select * from  [erpdata].[dbo].[SpecialGRdetailHistory] where Previous_Batch_ID='" + billNoTemp + "'"
   
slectResult = SqlServerQueryStr(cmdStr)
JudgeSpecialGRBillNo = slectResult
End Function

Public Function JudgeSpecialGRBillNoOI(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select * from  [ERPBASE].[dbo].[tblCustomerOI]  where source_batch_id='" + billNoTemp + "'"

   
slectResult = SqlServerQueryStr(cmdStr)
JudgeSpecialGRBillNoOI = slectResult
End Function


Public Function JudgeGRBillNo2(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select *  from [erpdata].[dbo].[GRdetailSetUp] where 单据编号='" + billNoTemp + "'"
      
slectResult = SqlServerQueryStr(cmdStr)
JudgeGRBillNo2 = slectResult
End Function

Public Function GetGC_FileNo(custNameTemp As String) As String
'耞琌
Dim cmdStr As String
Dim slectResult As String
slectResult = False

'cmdStr = "SELECT COUNT(*)+1  From [erpdata].[dbo].[GR_GC_SendHistory] Where flag='Y' and [SendTime] = Convert(VarChar(10), Getdate(),111)"


cmdStr = "SELECT right('0'+CAST((COUNT(*)+1) AS varchar(2)),2)  From [erpdata].[dbo].[GR_GC_SendHistory] Where flag='Y' and customername='" + custNameTemp + "' and [SendTime] = Convert(VarChar(10), Getdate(),111)"


slectResult = GetGCRptFileNo(cmdStr)
GetGC_FileNo = slectResult
End Function


Public Function GetMPS_OutDate(custNameTemp As String) As String
'耞琌
Dim cmdStr As String
Dim slectResult As String
slectResult = False

'cmdStr = "SELECT COUNT(*)+1  From [erpdata].[dbo].[GR_GC_SendHistory] Where flag='Y' and [SendTime] = Convert(VarChar(10), Getdate(),111)"


'cmdStr = "SELECT right('0'+CAST((COUNT(*)+1) AS varchar(2)),2)  From [erpdata].[dbo].[GR_GC_SendHistory] Where flag='Y' and customername='" + custNameTemp + "' and [SendTime] = Convert(VarChar(10), Getdate(),111)"

cmdStr = " SELECT Convert(varchar(10),操作日期,120) as outdate From erpdata.dbo.tblStockMove where 客户代码 in ('68','70') and 单据类型='1' and 单据编号='" + custNameTemp + "' "



slectResult = GetGCRptFileNo(cmdStr)
GetMPS_OutDate = slectResult
End Function




Public Function GetGC_FileNoNew(custNameTemp As String) As String
'耞琌
Dim cmdStr As String
Dim slectResult As String
slectResult = False

'cmdStr = "SELECT COUNT(*)+1  From [erpdata].[dbo].[GR_GC_SendHistory] Where flag='Y' and [SendTime] = Convert(VarChar(10), Getdate(),111)"


cmdStr = "SELECT right('00'+CAST((COUNT(*)+1) AS varchar(3)),3)  From [erpdata].[dbo].[GR_GC_SendHistory] Where flag='Y' and customername='" + custNameTemp + "' and [SendTime] = Convert(VarChar(10), Getdate(),111)"


slectResult = GetGCRptFileNo(cmdStr)
GetGC_FileNoNew = slectResult
End Function





Public Function GetGWoDeptID(custNameTemp As String) As String
'耞琌
Dim cmdStr As String
Dim slectResult As String
slectResult = False

cmdStr = "select FNumber  from AIS20141114094336.dbo.t_Department where  FName='" + custNameTemp + "' "

slectResult = GetGCRptFileNo(cmdStr)
GetGWoDeptID = slectResult
End Function



Public Function JudgeGRBillNoGC(billNoTemp As String, custNameTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

'cmdStr = "select *  from [erpdata].[dbo].[GR_GC_DetailHistory] where 单据编号='" + billNoTemp + "'"

cmdStr = "select a.*  from [erpdata].[dbo].[GR_GC_DetailHistory] a, [erpdata].[dbo].[tblStockMove] b where a.单据编号='" + billNoTemp + "' and b.单据编号=a.单据编号 and b.客户代码='" + custNameTemp + "' order by a.CartonNO "


slectResult = SqlServerQueryStr(cmdStr)
JudgeGRBillNoGC = slectResult
End Function


Public Function JudgeSemtechBillNo(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

'cmdStr = "select *  from [erpdata].[dbo].[GR_GC_DetailHistory] where 单据编号='" + billNoTemp + "'"

cmdStr = "SELECT *  FROM   [erpdata].[dbo].[tblStockMove] a where a.单据编号='" + billNoTemp + "' and 客户代码='37' "


slectResult = SqlServerQueryStr(cmdStr)
JudgeSemtechBillNo = slectResult
End Function




Public Function JudgeGRBillNoGCCodeLen(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

'cmdStr = "select *  from [erpdata].[dbo].[GR_GC_DetailHistory] where 单据编号='" + billNoTemp + "'"

'cmdStr = "select a.*  from [erpdata].[dbo].[GR_GC_DetailHistory] a, [erpdata].[dbo].[tblStockMove] b where a.单据编号='" + billNoTemp + "' and b.单据编号=a.单据编号 and b.客户代码='" + custNameTemp + "' "



cmdStr = " select * From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + billNoTemp + "' and  len( RTRIM( LTRIM([GC_Version])))<>3"

slectResult = SqlServerQueryStr(cmdStr)
JudgeGRBillNoGCCodeLen = slectResult
End Function



Public Function JudgeGRBillNoGCWlt(billNoTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

'cmdStr = "select *  from [erpdata].[dbo].[GR_GC_DetailHistory] where 单据编号='" + billNoTemp + "'"

cmdStr = "select a.*  from [erpdata].[dbo].[GR_GC_DetailHistory] a, [erpdata].[dbo].[tblStockMove] b where a.单据编号='" + billNoTemp + "' and b.单据编号=a.单据编号  "


slectResult = SqlServerQueryStr(cmdStr)
JudgeGRBillNoGCWlt = slectResult
End Function


Public Function SqlServerQueryStr(cmdStr As String) As Boolean
    Dim resut As Boolean
    resut = False
    
    If INIadoCon.State = 0 Then
    INIConnectSTART
    End If

    If rs.State = 1 Then
        rs.Close
    End If
    
    rs.Open cmdStr, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
        resut = True
    End If
    SqlServerQueryStr = resut
End Function

Public Function GetJDCustomerName() As ADODB.Recordset
'客户代码
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'cmdStr = "select distinct b.productname PID,b.productname from product a ,PRODUCTBASE b" & _
'" where a.productbaseid=b.productbaseid and b.objectcategory='PN' and a.objecttype='PN' and b.productname like '18%' order by b.productname"
'

'cmdStr = "SELECT 客户代码 as PID,客户代码 as productname FROM dbo.tblXCustomer union  select 'JX117' as PID,'JX117' as productname union  select 'AA(ON)' as PID,'AA(ON)' as productname union  select '37(ICI)' as PID,'37(ICI)' as productname  order by 客户代码 "
cmdStr = "SELECT 客户代码 as PID,客户代码 as productname FROM dbo.tblXCustomer union  select 'JX117' as PID,'JX117' as productname union  select 'AA(ON)' as PID,'AA(ON)' as productname union  select '37(ICI)' as PID,'37(ICI)' as productname  union  select 'AB18(2)' as PID,'AB18(2)' as productname union select 'YZ22(2)' as PID,'YZ22(2)' as productname order by 客户代码 "


Set RSResult = getSqlStr(cmdStr)
Set GetJDCustomerName = RSResult
End Function



Public Function GetSingelPrice(bj As String) As ADODB.Recordset
Dim cmdSql  As String

cmdSql = "select WAFER_PRICE, DIE_PRICE from erptemp..tblBB_QUOTATION where QUOTATION = '" & bj & "'"

Set GetSingelPrice = getSqlStr(cmdSql)

End Function



Public Function GetBomProductName() As ADODB.Recordset
'成品料号
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'cmdStr = "select distinct b.productname PID,b.productname from product a ,PRODUCTBASE b" & _
'" where a.productbaseid=b.productbaseid and b.objectcategory='PN' and a.objecttype='PN' and b.productname like '18%' order by b.productname"
'

cmdStr = "select rtrim(ltrim(物料编号)) as PID, rtrim(ltrim(物料编号)) as productname from  [erpdata].[dbo].[TSVtblSetMRule] Where 是否共用标记 = 0 order by 物料编号"


Set RSResult = getSqlStr(cmdStr)
Set GetBomProductName = RSResult
End Function



Public Function GetBomTYName() As ADODB.Recordset
'通用Bom
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'cmdStr = "select distinct b.productname PID,b.productname from product a ,PRODUCTBASE b" & _
'" where a.productbaseid=b.productbaseid and b.objectcategory='PN' and a.objecttype='PN' and b.productname like '18%' order by b.productname"
'

cmdStr = "select 物料编号 as PID, rtrim(ltrim(材料规范编号)) +'/' + rtrim(ltrim(物料编号)) as productname from  [erpdata].[dbo].[TSVtblSetMRule] Where 是否共用标记 <> 0 order by 物料编号"


Set RSResult = getSqlStr(cmdStr)
Set GetBomTYName = RSResult
End Function


Public Function GetSpecialGRBefor(lotIdTemp As String) As ADODB.Recordset
'查询以前的dataCode与测试版本号
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select MAX(Date_code) Date_code,max(Tst_Program_rev) Tst_Program_rev from  [erpdata].[dbo].[GRdetailHistory] " & _
 " where Previous_Batch_ID='" + lotIdTemp + "' "
 
Set RSResult = getSqlStr(cmdStr)
Set GetSpecialGRBefor = RSResult
End Function



Public Function GetLotSendTimes(lotIdTemp As String) As ADODB.Recordset
'发货次数
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


 cmdStr = " SELECT distinct 单据编号,CONVERT(char(10), 操作日期, 120) as 操作日期 FROM   [erpdata].[dbo].[tblStockMove]  where 客户代码='68'  and   工单号 ='" & lotIdTemp & "' "
 
Set RSResult = getSqlStr(cmdStr)
Set GetLotSendTimes = RSResult
End Function





Public Function GetGCNeedIn() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
         
'cmdStr = " select a.WORKORDERNAME,a.WAFERNUMBER, a.WAFERSCRIBENUMBER,a.NDPW,a.QBOXNUMBER,a.CONTAINERNAME   from  erpdata.dbo.TblQBOXNUMBER_TSV a " & _
'" where a.CUSTOMERNAME='GC' and a.CREATE_DATE>GETDATE()-1 " & _
'" and a.WAFERSCRIBENUMBER not in (SELECT 流程卡编号 FROM erpdata.dbo.tblPackToHouseRec) " & _
'" order by a.WORKORDERNAME,a.WAFERNUMBER, a.WAFERSCRIBENUMBER,a.QBOXNUMBER,a.CONTAINERNAME "


         
cmdStr = " select a.WORKORDERNAME,a.WAFERNUMBER, a.WAFERSCRIBENUMBER,a.NDPW,a.QBOXNUMBER,a.CONTAINERNAME   from  erpdata.dbo.TblQBOXNUMBER_TSV a " & _
" where a.CUSTOMERNAME='GC' and a.CREATE_DATE>GETDATE()-1 " & _
" order by a.WORKORDERNAME,a.WAFERNUMBER, a.WAFERSCRIBENUMBER,a.QBOXNUMBER,a.CONTAINERNAME "

            
Set RSResult = getSqlStr(cmdStr)
Set GetGCNeedIn = RSResult

End Function



Public Function GetSqlServerFpsCloseWo() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select 工单号,PRODUCT,Qty,FGQty,Qty-FGQty,rtrim(round(cast((FGQty/Qty)* 100 as decimal(10,2)),2))+'%',BomStatus,'' from ( " & _
" select x.工单号,x.PRODUCT,x.QTY,  erpdata.dbo.Get_TSV_WO_DieQty(x.工单号) as FGQty,erpdata.dbo.Get_TSV_WO_BomFinish(x.工单号) as BomStatus  from ( " & _
" select distinct e.工单号,f.PRODUCT ,f.QTY from  [erpbase].[dbo].[tblllplan] e , [erpdata].[dbo].[tblTSVworkorder] f " & _
" where    f.ORDERNAME=e.工单号 and  e.产线标记=1  and e.完工标记=0 ) X)Y "
        
        
Set RSResult = getSqlStr(cmdStr)

Set GetSqlServerFpsCloseWo = RSResult
End Function




Public Function GetFps37POComplete(billNoTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


 
cmdStr = " SELECT  'COMPLETE' as Event ,c.PO_NUM,c.PO_ITEM,c.MPN,'' as OrderClose,c.CURRENT_WAFER_QTY,COUNT(b.流程卡编号) as Quantity,'' as ScrapQuantity, " & _
" a.工单号 FROM  [ERPdata].[dbo].[tblStockMove] a, [ERPdata].[dbo].[tblStockMovesub] b ,[ERPBASE].[dbo].[tblCustomerOI] c " & _
" where a.单据编号='" + billNoTemp + "' and a.客户代码='37' " & _
" and b.单据编号=a.单据编号 and b.工单号=a.工单号 " & _
" and c.CUSTOMERSHORTNAME='37' and c.SOURCE_BATCH_ID=a.工单号 and c.PO_NUM <>'' " & _
" group by   c.PO_NUM,c.PO_ITEM,c.MPN,c.CURRENT_WAFER_QTY, a.工单号 "


Set RSResult = getSqlStr(cmdStr)

Set GetFps37POComplete = RSResult
End Function


Public Function GetFpsMPSCt() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


 
'cmdStr = " SELECT  'COMPLETE' as Event ,c.PO_NUM,c.PO_ITEM,c.MPN,'' as OrderClose,c.CURRENT_WAFER_QTY,COUNT(b.流程卡编号) as Quantity,'' as ScrapQuantity, " & _
'" a.工单号 FROM  [ERPdata].[dbo].[tblStockMove] a, [ERPdata].[dbo].[tblStockMovesub] b ,[ERPBASE].[dbo].[tblCustomerOI] c " & _
'" where a.单据编号='" + billNoTemp + "' and a.客户代码='37' " & _
'" and b.单据编号=a.单据编号 and b.工单号=a.工单号 " & _
'" and c.CUSTOMERSHORTNAME='37' and c.SOURCE_BATCH_ID=a.工单号 and c.PO_NUM <>'' " & _
'" group by   c.PO_NUM,c.PO_ITEM,c.MPN,c.CURRENT_WAFER_QTY, a.工单号 "



cmdStr = "  SELECT a.客户代码, a.单据编号,CONVERT(char(10), a.操作日期, 120) as 操作日期,b.料号,b.工单号,count(b.流程卡编号) as qty FROM" & _
" [ERPdata].[dbo].[tblStockMove] a ,[ERPdata].[dbo].[tblStockMovesub] b " & _
" where a.客户代码 in ('68','70') and a.单据编号 like 'F%'  " & _
" and b.单据编号=a.单据编号 and b.工单号=a.工单号 " & _
" group by a.客户代码,a.单据编号,b.料号,a.操作日期,b.工单号 order by a.操作日期 desc  "

Set RSResult = getSqlStr(cmdStr)

Set GetFpsMPSCt = RSResult
End Function




Public Function GetFps37POShip(billNoTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset




cmdStr = "  SELECT  'SHIP' as Event ,c.PO_NUM,c.PO_ITEM,c.MPN," & _
" substring(ltrim(c.offshore_asm_company),1,4) as MPlant,substring(ltrim(c.offshore_test_company),1,4) as SPlant," & _
" c.MPN_DESC,c.source_mtrl_sloc,'' as OrderClose,COUNT(b.流程卡编号) as Quantity," & _
" CONVERT(char(8), a.操作日期, 112) as sdate,'CN' AS COrigin,CONVERT(varchar(100), d.ERPCREATEDATE, 23) as datecode ,a.工单号 " & _
" FROM  [erpdata].[dbo].[tblStockMove] a, [erpdata].[dbo].[tblStockMovesub] b ,[ERPBASE].[dbo].[tblCustomerOI] c ,[erpdata].[dbo].[tblTSVworkorder] d,[erpdata].[dbo].[tblTSVwaferlist] e " & _
" where a.单据编号='" + billNoTemp + "'and a.客户代码='37' and b.单据编号=a.单据编号 and b.工单号=a.工单号 " & _
" and c.CUSTOMERSHORTNAME='37' and c.SOURCE_BATCH_ID=a.工单号 and c.PO_NUM <>'' " & _
" and d.ORDERNAME=e.ORDERNAME and e.WAFERLOT=a.工单号 and e.WAFERID=b.流程卡编号 " & _
" group by   c.PO_NUM,c.PO_ITEM,c.MPN,c.CURRENT_WAFER_QTY, a.工单号,c.offshore_asm_company,c.offshore_test_company," & _
" c.MPN_DESC , c.source_mtrl_sloc, CONVERT(Char(8), a.操作日期, 112), d.ERPCreateDate"


Set RSResult = getSqlStr(cmdStr)

Set GetFps37POShip = RSResult
End Function

Public Function GetSqlServerFpsCloseWo1() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'cmdStr = " select 工单号,PRODUCT,Qty,FGQty,Qty-FGQty,rtrim(round(cast((FGQty/Qty)* 100 as decimal(10,2)),2))+'%',BomStatus,'' from ( " & _
'" select x.工单号,x.PRODUCT,x.QTY,  erpdata.dbo.Get_TSV_WO_DieQty(x.工单号) as FGQty,erpdata.dbo.Get_TSV_WO_BomFinish(x.工单号) as BomStatus  from ( " & _
'" select distinct e.工单号,f.PRODUCT ,f.QTY from  [erpbase].[dbo].[tblllplan] e , [erpdata].[dbo].[tblTSVworkorder] f " & _
'" where    f.ORDERNAME=e.工单号 and  e.产线标记=1  and e.完工标记=0 ) X)Y "
'
'cmdStr = "SELECT  orderName,PRODUCT,woQty,invQty,wipQty,finishRate,BomStatus,flag FROM [erpdata].[dbo].[Vw_TSV_CloseWo] where wipQty<1 and BomStatus='是'"

cmdStr = "SELECT a.orderName,a.PRODUCT, a.woQty, a.invQty, a.wipQty, a.finishRate, a.BomStatus, CONVERT(varchar(100), b.ERPCREATEDATE, 23), DATEDIFF(day,b.ERPCREATEDATE,GETDATE()),a.flag " & _
"FROM [erpdata].[dbo].[Vw_TSV_CloseWo] a left join [erpdata].[dbo].[tblTSVworkorder] b on b.ORDERNAME = a.ORDERNAME where a.wipQty < 1 and a.BomStatus = '是' "

Set RSResult = getSqlStr(cmdStr)

Set GetSqlServerFpsCloseWo1 = RSResult
End Function



Public Function GetSqlServerFpsCloseWo2() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'cmdStr = " select 工单号,PRODUCT,Qty,FGQty,Qty-FGQty,rtrim(round(cast((FGQty/Qty)* 100 as decimal(10,2)),2))+'%',BomStatus,'' from ( " & _
'" select x.工单号,x.PRODUCT,x.QTY,  erpdata.dbo.Get_TSV_WO_DieQty(x.工单号) as FGQty,erpdata.dbo.Get_TSV_WO_BomFinish(x.工单号) as BomStatus  from ( " & _
'" select distinct e.工单号,f.PRODUCT ,f.QTY from  [erpbase].[dbo].[tblllplan] e , [erpdata].[dbo].[tblTSVworkorder] f " & _
''" where    f.ORDERNAME=e.工单号 and  e.产线标记=1  and e.完工标记=0 ) X)Y "
'
'cmdStr = "SELECT  orderName,PRODUCT,woQty,invQty,wipQty,finishRate,BomStatus,flag FROM [erpdata].[dbo].[Vw_TSV_CloseWo] where wipQty>0 or BomStatus='否'"
'
'cmdStr = "SELECT a.orderName,a.PRODUCT, a.woQty, a.invQty, a.wipQty, a.finishRate, a.BomStatus, CONVERT(varchar(100), b.ERPCREATEDATE, 23), DATEDIFF(day,b.ERPCREATEDATE,GETDATE()),a.flag " & _
'"FROM [erpdata].[dbo].[Vw_TSV_CloseWo] a left join [erpdata].[dbo].[tblTSVworkorder] b on b.ORDERNAME = a.ORDERNAME where a.wipQty > 0 and a.BomStatus = '否' "
cmdStr = "SELECT a.orderName,a.PRODUCT, a.woQty, a.invQty, a.wipQty, a.finishRate, a.BomStatus, CONVERT(varchar(100), b.ERPCREATEDATE, 23), DATEDIFF(day,b.ERPCREATEDATE,GETDATE()),a.flag " & _
"FROM [erpdata].[dbo].[Vw_TSV_CloseWo] a left join [erpdata].[dbo].[tblTSVworkorder] b on b.ORDERNAME = a.ORDERNAME where a.wipQty > 0 "

Set RSResult = getSqlStr(cmdStr)

Set GetSqlServerFpsCloseWo2 = RSResult
End Function

Public Function GetCustomerNameSqlServer(custNameTemp As String) As String
'耞琌
Dim cmdStr As String
Dim slectResult As String
slectResult = False

cmdStr = "SELECT RTRIM( ltrim(客户名称)) FROM dbo.tblXCustomer where 客户代码='" + custNameTemp + "' "


slectResult = GetGCRptFileNo(cmdStr)
GetCustomerNameSqlServer = slectResult
End Function

Public Function GetCustomerNameSqlServer1(custNameTemp As String) As String

Dim cmdStr As String
Dim slectResult As String
slectResult = False

cmdStr = "SELECT RTRIM( ltrim(客户编号)) FROM erptemp..tblbb_customer where 客户代码='" + custNameTemp + "' "


slectResult = GetGCRptFileNo(cmdStr)
GetCustomerNameSqlServer1 = slectResult
End Function

Public Function GetCustomerNameSqlServer2(custNameTemp As String) As String
'耞琌
Dim cmdStr As String
Dim slectResult As String
slectResult = False

cmdStr = "SELECT RTRIM( ltrim(客户区域)) FROM erptemp..tblbb_customer where 客户代码='" + custNameTemp + "' "


slectResult = GetGCRptFileNo(cmdStr)
GetCustomerNameSqlServer2 = slectResult
End Function

