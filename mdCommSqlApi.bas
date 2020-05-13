Attribute VB_Name = "mdCommSqlApi"
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Const mc_strIniFileName As String = "Config.ini"

Public strGReason As String

Public Enum E_TBL_EVENT

    E_INSERT
    E_DELETE
    E_UPDATE
    E_QUERY

End Enum

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim rsCommon      As New ADODB.Recordset

Public cmd        As New ADODB.Command

' 数据库连接
Public OraConnect As New ADODB.Connection

Public SqlConnect As ADODB.Connection

Public SqlConnect2 As ADODB.Connection

Public strMediaDir  As String

Public strTestPath As String

Public strFlagPath As String, str37BCIDPath As String, str37CartonPath As String

Public strSSBoxPath As String, strSSReelPath As String, strSSCartonPath As String

Public strHTQCartonPath As String, strHWBoxPath As String, strHWReelPath As String

Public Type WO_MOD

    strWaferID As String
    strLotID As String
    strpo As String
    strGoodDies As String
    strBadDies As String
    strVERSION As String
    strCUSDEVICE As String
    strCusCode As String
    strPRODUCTID As String
    strJobID As String

End Type

Public Type DN_DETAILS

    ID As Long
    Delivery As String
    ItemNo As String
    DeliveryCreationDate As String
    Plant As String
    SalesDocument As String
    SOItemNo As String
    Material As String
    MarketingPN As String
    MaterialDescription As String
    PlannedGIdate As String
    CustomerPartnumber As String
    ShiptoName As String
    ShiptoCustomer As String
    PurchasingDocNo As String
    DateCodeRestrictions As String
    LabelRequirement As String
    ReLabelInstructions As String
    ShipToStreet1   As String
    ShipToStreet2   As String
    ShipToStreet3   As String
    City    As String
    State As String
    PostalCode   As String
    CountryKey  As String
    ContactName As String
    Phone As String
    Fax As String
    FreightForwarder    As String
    ShippingInstruction As String
    AdditionalComments  As String
    StorageLocation As String
    BatchNumber As String
    Quantity As String
    VolumeWeight    As String
    GrossWeight As String
    netweight   As String
    UoMForWeight   As String
    NoOfCartons As String
    VendorLotNumber As String
    ShelfLocation   As String
    BOLOrAirwayBillNo As String
    ActualShippingDate As String
    PackagingDetails As String
    PackingStatus As String
    PickingStatus As String
    CustomerCalendar As String
    FatherBatch As String
    MotherBatch As String
    FatherBatchDateCode As String
    MotherBatchDateCode As String
    TransferOrderStatus As String
    DATECODE As String
    FatherBatchQty As String
    ShippingPoint As String
    ShipmentNumber As String
    FabSite As String
    AssemblySite As String
    TestSite As String

End Type

Public Type WORKORDER_DATA

    ID As String
    WORK_ORDER_ID As String
    Lot_id As String
    Wafer_id As String
    MARKING_CODE As String
    TOTALDIE As Long
    gooddie As Long

End Type

Public Type LBL_WAFER_INFO

    strWaferID As String
    strSecCode As String
    strCodePP As String
    strCusDev As String
    lGrossDiesQty As Long
    lBin1DiesQty As Long
    lBin2DiesQty As Long
    lBin3DiesQty As Long
    bChecked As Boolean

End Type

Public Type WOBACKUP

    CUSTOMER As String
    Wafer_id As String
    lot As String
    good_die As String
    BAD_DIE As String
    Customer_Device As String
    PO_NUM As String
    Fab_Device As String
    SHIP_TO As String
    SEC_CODE As String
    CUST_SEC_CODE As String
    Create_date As String
    create_by As String
    LASTUPDATE As String
    LASTUPDATE_BY As String
    EVENT As String

End Type

Public Type ReproductionWaferData

    SUBSTRATEID As String
    LOTID As String
    PRODUCTID As String
    PASSBINCOUNT As Long
    GROSSBINCOUNT As Long
    FLAG As String
    QTECH_CREATED_BY As String
    QTECH_CREATED_DATE As String
    Wafer_id As String
    CUSTOMERSHORTNAME As String
    CUSTOMERDEVICE As String
    JOBNO As String
    ID As String

End Type

Public Type tScan

    sKey As String
    sVal As String
    sSel As String

End Type

'''''''''''''''''''''''''''''''''''''
' 外箱标签数据
Public Type tOutPkgLblData

    sInvoice As String
    sPurchaseOrder As String
    sCustomerPartNo As String
    sMfgPartNo As String
    sLotNo As String
    sQty As String

End Type

' 外箱标签校验状态
Public Type tOutPkgLblStatus

    bST As Boolean
    bSS As Boolean

End Type

' 内箱标签数据
Public Type tInPkgLblData

    sJobNo As String
    sMFG As String
    sLotNumber As String
    sQty As String
    sStrInfo As String    ' 条码串
    sPartNo As String

End Type

' LVS
Public Type tLVSData

    sDN As String
    sPO As String
    sCPN As String
    sDev As String
    sJobNo As String
    sJobQty As String

End Type

Public Type tLVS

    sLot As String
    sQty As Long
    sFlag As Boolean

End Type

Public Type tSTData

    TRAYID As String
    INBOX_NUM As String
    OUTBOX_NUM As String
    DN_NUM As String
    JOB_ID As String
    QTY As Long
    Customer_Device As String
    Create_date As String
    create_by As String
    PRINT_FLAG As String
    FLAG As String
    carton As String
    REEL_ID As String
    SEQ As Long
    KID As String
    DC As String
    C_ID As String
    B_ID As String
    PSN As String

End Type

Public Type tLbl

    dn As String
    PO As String
    CPN As String
    DEV As String
    JOB As String
    lot As String
    QTY As String
    ADATE As String
    TDATE As String

End Type

Public Type CusReel

    PN As String
    lot As String
    DEV As String
    QTY As String
    TRAYID As String

End Type

Public Type CusBox

    DEV As String
    PN As String
    QTY As String

End Type

Public Type HWBox

    CPN As String
    MPN As String
    PODATE As String
    lot As String
    QTY As String
    PSN As String

End Type

Public Type STBox

    JOB As String
    DEV As String
    lot As String
    QTY As String
    DATECODE As String
    testdateCode As String

End Type

Public Type STCarton

    JOB As String
    DEV As String
    lot As String
    QTY As String
    DATECODE As String
    testdateCode As String

End Type

Public Type CUSCARTON

    dn As String
    PO As String
    CPN As String
    MPN As String
    JOB As String
    QTY As String
    KID As String
    DATECODE As String

End Type

Public Type M_CUS_CARTON

    dn As String
    PO As String
    MPN As String
    CPN As String
    lot As String
    QTY As String

End Type

Public Type M_SEM_BOX

    JOB As String
    MPN As String
    lot As String
    QTY As String

End Type

Public Type M_CUS_REEL

    PN As String
    lot As String
    QTY As String

End Type

Public Type M_CUS_BOX

    PN As String
    QTY As String

End Type

Rem: For new LPS

    Public Type ST_TR_SEQ

        sDN As String
        sReelID As String
        sJob As String
        sLot As String
        sDev As String
        lQty As Long
        lSeq As Long

    End Type

    Rem:--------------

    Public Type LPSM1

        dn As String
        PO As String
        CPN As String
        MPN As String
        lot As String
        QTY As String

    End Type
    
Public Function GetRsData(rs As ADODB.Recordset, Para As String)

    Dim sRtn As String

    If IsNull(rs.Fields(Para)) = True Then
        sRtn = ""
    Else
        sRtn = Trim(rs.Fields(Para))

    End If

    GetRsData = sRtn

End Function

Public Sub addOrderLogTxt(woTemp As String, msgTxt As String)

    '判断txt文件是否存在，如不存在，则建立
    Dim fileNameTemp As String

    Dim dirNameTemp  As String

    Dim fileTemp     As String

    dirNameTemp = "C:\Program Files\HT_OrderLog\"
    fileNameTemp = woTemp & ".txt"
    fileTemp = dirNameTemp & fileNameTemp

    Open fileTemp For Append As #1   '文件存在就追加，不存在就自动创建
    Print #1, CStr(Now) + msgTxt
    Close #1

End Sub

'Insert into DB
Public Sub Exec_Ora(cmdStr As String)

    If Cnn.State = 0 Then
        ConOracle

    End If

    cmd.ActiveConnection = Cnn
    cmd.CommandText = cmdStr
    cmd.CommandType = adCmdText
    cmd.Execute
    
End Sub

Public Sub Exec_Sql(cmdStr As String)

    If INIadoCon.State = 0 Then
        INIConnectSTART

    End If

    cmd.ActiveConnection = INIadoCon
    cmd.CommandText = cmdStr
    cmd.CommandType = adCmdText
    cmd.Execute
    
End Sub

Public Sub SaveWOLot_37(strlot As String, strWOID As String, strCustPN As String)
Dim strSql As String
Dim iCnt As Integer

strSql = "select count(1)+1 from cust_lot_num where to_char(sysdate + 1, 'YYWW') = to_char(create_date + 1, 'YYWW')"
iCnt = Get_OracleNo(strSql)

strSql = "select * from cust_lot_num  where to_char(sysdate + 1, 'YYWW') = to_char(create_date + 1, 'YYWW') and LOT = '" & strlot & "'"
If Get_OracleCnt(strSql) > 0 Then
    'update
    strSql = "update cust_lot_num set order_num = order_num ||';'|| '" & strWOID & "',last_create_date = sysdate  where lot = '" & strlot & "' and to_char(sysdate + 1, 'YYWW') = to_char(create_date + 1, 'YYWW') "
Else
    'insert
    strSql = "INSERT INTO  cust_lot_num (customer,cust_part,lot,lot_num,order_num,create_date,REMARK1 ) " & _
        " values ('37','" & strCustPN & "','" & strlot & "','" & iCnt & "','" & strWOID & "',sysdate,to_char(sysdate + 1, 'YYWW')) "

End If

AddSql (strSql)

End Sub

Public Function Get_OracleRs(cmdStr As String) As ADODB.Recordset

    Dim rs As New ADODB.Recordset

    If Cnn.State = 0 Then
        ConOracle

    End If
    
    rs.Open cmdStr, Cnn, adOpenStatic, adLockReadOnly, adCmdText

    Set Get_OracleRs = rs

End Function


Public Function Get_SqlserveRs(cmdStr As String) As ADODB.Recordset
    
    Dim rs As New ADODB.Recordset

    If INIadoCon.State = 0 Then
        INIConnectSTART

    End If
    
    rs.Open cmdStr, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

    Set Get_SqlserveRs = rs
    
End Function

Public Function Get_OracleStr(cmdStr As String) As String

    Dim rs As New ADODB.Recordset
    
    If Cnn.State = 0 Then
        ConOracle

    End If
    
    rs.Open cmdStr, Cnn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Get_OracleStr = Trim(IIf(IsNull(rs.Fields(0).Value), "", Trim("" & rs.Fields(0).Value)))
    Else
        Get_OracleStr = ""

    End If

    rs.Close

End Function

Public Function Get_SqlStr(cmdStr As String) As String

    Dim rs As New ADODB.Recordset

    If INIadoCon.State = 0 Then
        INIConnectSTART

    End If

    rs.Open cmdStr, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Get_SqlStr = IIf(IsNull(rs.Fields(0).Value), "", Trim("" & rs.Fields(0).Value))
    Else
        Get_SqlStr = ""

    End If

    rs.Close
    Set rs = Nothing

End Function

Public Function Get_SqlStr2(cmdStr As String) As String

    Dim rs As New ADODB.Recordset

    If SqlConnect2.State = 0 Then
        ConnectSql2

    End If

    rs.Open cmdStr, SqlConnect2, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Get_SqlStr2 = IIf(IsNull(rs.Fields(0).Value), "", rs.Fields(0).Value)
    Else
        Get_SqlStr2 = ""

    End If

    rs.Close
    Set rs = Nothing

End Function

Public Function Get_OracleNo(cmdStr As String) As Long

    Dim rs As New ADODB.Recordset
    
    If Cnn.State = 0 Then
        ConOracle

    End If
    
    rs.Open cmdStr, Cnn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Get_OracleNo = CLng(IIf(IsNull(rs.Fields(0).Value), 0, rs.Fields(0).Value))
    Else
        Get_OracleNo = 0

    End If

End Function

Public Function Get_OracleCnt(cmdStr As String) As Long
    
    Dim rs As New ADODB.Recordset
    
    If Cnn.State = 0 Then
        ConOracle

    End If
    
    rs.Open cmdStr, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
        Get_OracleCnt = rs.RecordCount
    Else
        Get_OracleCnt = 0
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function Get_SqlserverCnt(cmdStr As String) As Long

    Dim rs As New ADODB.Recordset

    If INIadoCon.State = 0 Then
        INIConnectSTART

    End If
    
    rs.Open cmdStr, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
        Get_SqlserverCnt = rs.RecordCount
    Else
        Get_SqlserverCnt = 0
    End If

    rs.Close
    Set rs = Nothing

End Function

Public Function Get_SqlserverNo(cmdStr As String) As Long

    Dim rs As New ADODB.Recordset

    If INIadoCon.State = 0 Then
        INIConnectSTART

    End If
    
    rs.Open cmdStr, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Get_SqlserverNo = IIf(IsNull(rs.Fields(0).Value), 0, rs.Fields(0).Value)
    Else
        Get_SqlserverNo = 0

    End If

End Function

Public Function Get_CusCode() As ADODB.Recordset

    Dim sSql As String

    Dim rs   As New ADODB.Recordset

    sSql = "SELECT 客户代码 as PID,客户代码 as productname FROM erpdata.dbo.tblXCustomer union  select 'JX117' as PID,'JX117' as productname union  select 'AA(ON)' as PID,'AA(ON)' as productname union  select '37(ICI)' as PID,'37(ICI)' as productname  union  select 'AB18(2)' as PID,'AB18(2)' as productname union select 'YZ22(2)' as PID,'YZ22(2)' as productname order by 客户代码 "
    sSql = "SELECT 客户代码 as PID,客户代码 as productname FROM erpdata.dbo.tblXCustomer " & " union  select 'JX117' as PID,'JX117' as productname " & " union  select 'AA(ON)' as PID,'AA(ON)' as productname " & " union  select '37(ICI)' as PID,'37(ICI)' as productname " & " union  select 'AB18(2)' as PID,'AB18(2)' as productname " & " union  select 'BUMPINGDM' as PID,'BUMPINGDM' as productname " & " union select 'YZ22(2)' as PID,'YZ22(2)' as productname order by 客户代码"
    Set rs = Get_SqlserveRs(sSql)
    Set Get_CusCode = rs

End Function

Public Function Get_CusDevice(sCusCode As String) As ADODB.Recordset

    Dim sOra As String

    Dim rs   As New ADODB.Recordset

    ' 根据客户代码筛选出客户机种
    sOra = "select distinct CUSTOMERPTNO1 from tbltsvnpiproduct where customershortname = '" & sCusCode & "' "

    Set rs = Get_OracleRs(sOra)
    Set Get_CusDevice = rs

End Function

Public Function Get_ProductNo(sCusCode As String, sCusDevice As String) As ADODB.Recordset

    Dim sOra As String

    Dim rs   As New ADODB.Recordset

    If sCusDevice <> "" Then
        ' 根据客户代码和客户机种筛选出产品料号
        sOra = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & sCusCode & "' and customerptno1 = '" & sCusDevice & "' "
    Else
        sOra = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & sCusCode & "'"

    End If

    Set rs = Get_OracleRs(sOra)
    Set Get_ProductNo = rs

End Function

Public Function Get_CusDeviceP(sProductNo As String) As String

    Dim sOra As String

    sOra = "select distinct customerptno1 from tbltsvnpiproduct where qtechptno2 = '" & sProductNo & "' "

    Get_CusDeviceP = Get_OracleStr(sOra)

End Function

Public Function Get_PlantDevice(sProductNo As String) As String

    Dim sOra As String

    sOra = "select distinct qtechptno from tbltsvnpiproduct where qtechptno2 = '" & sProductNo & "' "

    Get_PlantDevice = Get_OracleStr(sOra)

End Function

Public Function Insert_WoTbl(sOrderType As String, _
                             sCusCode As String, _
                             sProductNo As String, _
                             iWaferQty As Long, _
                             sCusDev As String) As String

    Dim sOra         As String

    Dim sSql         As String

    Dim sLotTmp      As String

    Dim sLotType     As String

    Dim sCusDevice   As String

    Dim sPlantDevice As String

    Dim sLotWaferId  As String

    Dim iPassDies    As Long

    Dim iNgDies      As Long

    Dim ID           As Long

    Dim sMarkCode    As String

    Dim iWafer       As Integer

    iPassDies = 0
    iNgDies = 0

    Select Case sOrderType

        Case "Dummy工单"
            sLotType = "D"

        Case "玻璃工单"
            sLotType = "G"

        Case "硅基工单"
            sLotType = "SI"

        Case "FO_CSP工单"
            sLotType = "SI"

        Case Else
            MsgBox "未知工单"
            Exit Function

    End Select

    ID = GetMaxID()

    sOra = "select SPECIALLOT.nextval ID from dual"
    sLotTmp = sLotType & Right(Year(Now), 2) & Right(("0" & Month(Now)), 2) & Right(("0" & Day(Now)), 2) & Right("00" & Get_OracleStr(sOra), 3)

'    If sOrderType <> "Dummy工单" Then
'        sOra = "select *  from tbltsvnpiproduct where customershortname = '" & sCusCode & "' and qtechptno2 = '" & sProductNo & "'"
'        Set rsCommon = Get_OracleRs(sOra)
'        sCusDevice = rsCommon.Fields("customerptno1").Value
'        sPlantDevice = rsCommon.Fields("qtechptno").Value
'        iPassDies = rsCommon.Fields("customerdieqty").Value
'        iNgDies = 0
'    Else
'        sCusDevice = sCusDev
'
'    End If

    sOra = "select *  from tbltsvnpiproduct where customershortname = '" & sCusCode & "' and qtechptno2 = '" & sProductNo & "' and customerptno1 = '" & sCusDev & "'"
    Set rsCommon = Get_OracleRs(sOra)
    
    sCusDevice = sCusDev
    sPlantDevice = rsCommon.Fields("qtechptno").Value
    iPassDies = rsCommon.Fields("customerdieqty").Value
    iNgDies = 0

    ' 插入WO头表
    sOra = "insert into ORDER_DATA_TEMP_HEADER(id,source_batch_id,SHIP_SITE,mpn_desc,mtrl_num,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date) " & " values ('" & ID & "','" & sLotTmp & "', '" & sCusCode & "', '" & sCusDevice & "', '" & sPlantDevice & "', '" & sCusCode & "','P', '" & gUserName & "', sysdate) "

    sSql = "insert into erpdata.dbo.ORDER_DATA_TEMP_HEADER(id,source_batch_id,SHIP_SITE,mpn_desc,mtrl_num,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date) " & " values ('" & ID & "','" & sLotTmp & "', '" & sCusCode & "', '" & sCusDevice & "', '" & sPlantDevice & "', '" & sCusCode & "','P', '" & gUserName & "', GETDATE()) "

    Call Get_OracleRs(sOra)
    Call Get_SqlserveRs(sSql)

    ' 插入WO子表
    For iWafer = 1 To iWaferQty
        sLotWaferId = sLotTmp & Right$("0" & iWafer, 2)
    
        sMarkCode = Right$(sLotWaferId, 6)
    
        sOra = "insert into ORDER_DATA_TEMP_DETAILS(id,substrateid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename, productid)" & " values( mappingData_SEQ.Nextval,'" & sLotWaferId & "','" & sLotTmp & "','" & iWafer & "','" & iPassDies & "','" & iNgDies & "','" & sCusCode & "','P','" & gUserName & "',sysdate,'" & ID & "', '" & sMarkCode & "')"
                                                    
        sSql = "insert into erpdata.dbo.ORDER_DATA_TEMP_DETAILS(substrateid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename,productid)" & " values('" & sLotWaferId & "','" & sLotTmp & "','" & iWafer & "','" & iPassDies & "','" & iNgDies & "','" & sCusCode & "','P','" & gUserName & "',GETDATE(),'" & ID & "', '" & sMarkCode & "')"

        Call Get_OracleRs(sOra)
        Call Get_SqlserveRs(sSql)
    Next

    Insert_WoTbl = sLotTmp

End Function

Public Function Get_OrderDetailsFps(sqlTemp As String, _
                                    customerTemp As String, _
                                    sOrderType As String, _
                                    sJob As String) As ADODB.Recordset

    Dim cmdStr   As String

    Dim JOB      As String

    Dim RSResult As New ADODB.Recordset

    If sJob = "" Then
    
        If sOrderType = "Dummy工单" Or sOrderType = "玻璃工单" Or sOrderType = "硅基工单" Or sOrderType = "FO_CSP工单" Then
            
            cmdStr = " select wafer_id, substrateid,'',passbincount+failbincount,passbincount,lotid,productid " & "  from ORDER_DATA_TEMP_DETAILS a where lotid in ('" & sqlTemp & "' )  and flag='P' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by substrateid  "
        
        ElseIf sOrderType = "重工工单" Then
      
            cmdStr = " select wafer_id, substrateid,'','','',lotid,replace(productid,'_','') " & "  from mappingDataTest a,customeroitbl_test b where lotid in ('" & sqlTemp & "' ) and to_char(b.id) = a.filename "
        Else
            
            cmdStr = " select wafer_id, substrateid,'',passbincount+failbincount,passbincount,lotid,replace(productid,'_','') " & "  from mappingDataTest a where lotid in ('" & sqlTemp & "' )  and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by substrateid  "
        End If

    Else

        If sOrderType = "Dummy工单" Or sOrderType = "玻璃工单" Or sOrderType = "硅基工单" Or sOrderType = "FO_CSP工单" Then
            cmdStr = " select wafer_id, substrateid,'',passbincount+failbincount,passbincount,lotid,productid " & "  from ORDER_DATA_TEMP_DETAILS a where lotid in ('" & sqlTemp & "' )  and flag='P' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by substrateid  "
        ElseIf sOrderType = "重工工单" Then
            cmdStr = " select '', substrateid,'','','',lotid,replace(productid,'_','') " & "  from mappingDataTest a where lotid in ('" & sqlTemp & "' )  and flag='Y'"
    
            cmdStr = " select wafer_id, substrateid,'','','',lotid,replace(productid,'_','') " & "  from mappingDataTest a,customeroitbl_test b where lotid in ('" & sqlTemp & "' ) and to_char(b.id) = a.filename and b.test_mtrl_desc =  '" & sJob & "'   "
        Else
            cmdStr = " select wafer_id, substrateid,'',passbincount+failbincount,passbincount,lotid,replace(productid,'_','') " & "  from mappingDataTest a where lotid in ('" & sqlTemp & "' )  and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by substrateid  "

        End If
    
    End If

    Set RSResult = Get_OracleRs(cmdStr)
    Set Get_OrderDetailsFps = RSResult

End Function

Public Function Insert_OrderToDb(dataHeader As BillHeader, _
                                 dataDetails() As BillDetail, _
                                 j As Integer) As Boolean

    Dim sOra         As String

    Dim sSql         As String

    Dim sOrder       As String

    Dim aLot()       As String

    Dim i            As Integer

    Dim qtyWaferTemp As Long

    Dim sDepCode     As String

    Insert_OrderToDb = False

    ' step0: 数据初始化
    sOrder = dataHeader.ORDERNAME
    qtyWaferTemp = 0

    On Error GoTo DealError

'    Cnn.BeginTrans
'    INIadoCon.BeginTrans

    '------------------------------------------------------------------------------老工单接口-------------------------------------------------------------
    
    ' step1: 插入头表: ib_workorder
    sOra = "insert into ib_workorder (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
       " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
       "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PARA10,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
       " Values (" & dataHeader.ID & ",'" & dataHeader.ORDERNAME & "','" & dataHeader.ORDERTYPE & "' ,'CREATED','" & dataHeader.ERPUSER & "','" & dataHeader.product & "'," & dataHeader.QTY & ",sysdate,to_date('" & dataHeader.PLANSTARTDATE & "','yyyy-mm-dd'),to_date('" & dataHeader.PLANENDDATE & "','yyyy-mm-dd')," & _
       " '" & dataHeader.CUSTOMER & "','" & dataHeader.SALESORDER & "','" & dataHeader.CustomerERPN & "','" & dataHeader.FABFACILITY & "','" & dataHeader.IMAGERREV & "','" & dataHeader.DESIGNID & "','" & dataHeader.MLEVEL235 & "','" & dataHeader.MLEVEL260 & "','" & dataHeader.NGFLAG & "','" & dataHeader.PARA1 & "'," & _
       "  '" & dataHeader.PARA2 & "','" & dataHeader.PARA3 & "','" & dataHeader.PARA4 & "','" & dataHeader.PARA5 & "','" & dataHeader.PARA6 & "','" & dataHeader.RequestDate & "','" & dataHeader.PARA8 & "','" & dataHeader.PARA10 & "','" & dataHeader.PROTECTIVE_FILM_APLD & "','" & dataHeader.Lot_Stauts & "'," & _
       " '" & dataHeader.MPN & "')"

    sDepCode = Right(dataHeader.PARA8, Len(dataHeader.PARA8) - InStr(dataHeader.PARA8, "_"))
    
    sSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
       " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
       "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PARA10,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
       " Values (" & dataHeader.ID & ",'" & dataHeader.ORDERNAME & "','" & dataHeader.ORDERTYPE & "' ,'CREATED','" & dataHeader.ERPUSER & "','" & dataHeader.product & "'," & dataHeader.QTY & ",convert(datetime,'" & dataHeader.ERPCREATEDATE & "'),convert(datetime,'" & dataHeader.PLANSTARTDATE & "'),convert(datetime,'" & dataHeader.PLANENDDATE & "')," & _
       " '" & dataHeader.CUSTOMER & "','" & dataHeader.SALESORDER & "','" & dataHeader.CustomerERPN & "','" & dataHeader.FABFACILITY & "','" & dataHeader.IMAGERREV & "','" & dataHeader.DESIGNID & "','" & dataHeader.MLEVEL235 & "','" & dataHeader.MLEVEL260 & "','" & dataHeader.NGFLAG & "','" & dataHeader.PARA1 & "'," & _
       "  '" & dataHeader.PARA2 & "','" & dataHeader.PARA3 & "','" & dataHeader.PARA4 & "','" & dataHeader.PARA5 & "','" & dataHeader.PARA6 & "','" & dataHeader.RequestDate & "','" & sDepCode & "', '" & dataHeader.PARA10 & "','" & dataHeader.PROTECTIVE_FILM_APLD & "','" & dataHeader.Lot_Stauts & "'," & _
       " '" & dataHeader.MPN & "')"
         
    AddSql (sOra)
    AddSql2 (sSql)
    
    Call addLogTxt(sOrder, " 插入表:ib_workorder,tblTSVworkorder ")

    ' step2: 插入wafer明细表: ib_waferlist
    For i = 0 To j - 1
        sOra = "insert into ib_waferlist(ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & dataDetails(i).ORDERNAME & "'," & " '" & dataDetails(i).waferid & "'," & dataDetails(i).DIEQTY & "," & dataDetails(i).FGDIEQTY & ",'" & dataDetails(i).WAFERLOT & "',100,'" & dataDetails(i).MARKINGCODE & "')"
    
        sSql = "insert into [erpdata].[dbo].[tblTSVwaferlist](ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & dataDetails(i).ORDERNAME & "'," & " '" & dataDetails(i).waferid & "'," & dataDetails(i).DIEQTY & "," & dataDetails(i).FGDIEQTY & ",'" & dataDetails(i).WAFERLOT & "',100,'" & dataDetails(i).MARKINGCODE & "')"
  
        AddSql (sOra)
        AddSql2 (sSql)
    
        Call addLogTxt(sOrder, " 插入表:ib_waferlist,tblTSVwaferlist " & dataDetails(i).waferid)
    Next

    qtyWaferTemp = j

    ' Step3: 插入SqlServer表: tblllplan
    If InStr(1, UpLotId, ",") > 0 Then
        aLot = Split(UpLotId, ",")
        
        For i = 1 To UBound(aLot)

            If dataHeader.CUSTOMER = "AA" Or dataHeader.CUSTOMER = "AA(ON)" Then
                Call ONLotIDClose(CStr(aLot(i)))
            Else
                Call updateHeaderDateForGC(CStr(aLot(i)), "Y", 0, dataHeader.SALESORDER)

            End If

        Next

    End If

    Call addLogTxt(sOrder, "插入SqlServer表:tblllplan " & "料号：" & dataHeader.product)
    
    qtyDieTemp = dataHeader.QTY

    If dataDetails(i).WAFERLOT = "95FPC" And dataHeader.CUSTOMER = "95" Then

        bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) " & " SELECT distinct  '" + sOrder + "',X.物料编号,'1','主选材料', " & " CAST( (CAST(X.用量 AS DECIMAL(18,8)) * " & qtyDieTemp & " ) AS  DECIMAL(18,3))  ,1 " & " from ( SELECT b.材料规范编号, b.物料编号,sum(b.每只用量) as 用量 " & " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & " Where a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & dataHeader.product & "' " & " group by b.材料规范编号, b.物料编号 )  X "
    Else

        If Frm_WORK_ORDER.cbOrderType = "重工工单" Or Frm_WORK_ORDER.cbOrderType = "Dummy工单" Then
            bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记)" & "select b.ORDERNAME ,a.物料编号,'1','主选材料',SUM(convert(int, c.DIEQTY)),'1' " & "from erpdata..tblSmainM2 a ,erpdata..tblTSVworkorder b, erpdata .. tblTSVwaferlist c " & "where b.ORDERNAME = '" & sOrder & "'  and  b.PRODUCT = a.料号 and c.ORDERNAME = b.ORDERNAME " & "group by b.ORDERNAME ,a.物料编号"
        
        ElseIf Frm_WORK_ORDER.cbOrderType = "FO_CSP工单" Then
            bomStrTemp = "INSERT INTO [erpbase].[dbo].[tblllplan](工单号, 物料编号, 序组, 材料, 用量, 产线标记)SELECT a.ORDERNAME,c.物料编号,'1',  '主选材料',  case when c.物料编号 like '07.01.03.01%' then  CEILING(b.DIEQTY * c.每只用量) else  b.DIEQTY * c.每只用量 end," & " '1' FROM erpdata .. tblTSVworkorder a,  erpdata..tblTSVwaferlist b, erpdata..TSVtblMRuleData c  where a.ORDERNAME = '" & sOrder & "' and b.ORDERNAME = a.ORDERNAME and c.工序号 = a.PRODUCT"

        Else
            bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) " & " SELECT distinct  '" & sOrder & "',X.物料编号,'1','主选材料', " & " CAST( (CAST(X.用量 AS DECIMAL(18,8)) * " & qtyWaferTemp & " ) AS  DECIMAL(18,3))  ,1 " & " from ( SELECT b.材料规范编号, b.物料编号,sum(b.每只用量) as 用量 " & " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & " Where a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & dataHeader.product & "' " & " group by b.材料规范编号, b.物料编号 )  X "

        End If

    End If

    AddSql2 (bomStrTemp)

    Call addLogTxt(sOrder, " 插入SqlServer表:tblllplan OK")

    ' Step4: 插入PJ_WO_PRI
    sOra = " insert into PJ_WO_PRI values('" & sOrder & "','" & dataHeader.sPri & "',to_char(sysdate,'YYYY-MM-DD'),'" & dataHeader.sLotType & "', '" & gUserName & "')"

    AddSql (sOra)
    
    Insert_OrderToDb = True

    Cnn.CommitTrans
    INIadoCon.CommitTrans

    Exit Function

DealError:

    Cnn.RollbackTrans
    INIadoCon.RollbackTrans

    Call addLogTxt(sOrder, "保存工单失败!")

    MsgBox "老工单接口执行失败！" + Chr(13) + "原因:" + Err.DESCRIPTION, vbInformation, "警告"

End Function

Public Function Insert_Shop_Order(sOrdername As String, _
                             sUserName As String, _
                             sCusDevice As String, _
                             sProductNo As String, _
                             sType As String) As Boolean

    Dim sOra               As String

    Dim sSql               As String

    Dim sLog               As String

    Dim sSelCustomerDevice As String

    Dim sSelProductID      As String
    
    Insert_Shop_Order = False

    sSelCustomerDevice = ""
    sSelProductID = ""


    On Error GoTo DealError

    Cnn.BeginTrans
    INIadoCon.BeginTrans

    ' ---------------------------------------------------------------------------------------------新工单接口------------------------------------------------------------------------
    
    Call addOrderLogTxt(sOrdername, " 准备插入新工单接口 :")

    Call addOrderLogTxt(sOrdername, " 客户机种: '" & sCusDevice & "', 料号: '" & sProductNo & "'")

    ' 表1: shop_order_detail
    sOra = "insert into shop_order_detail(SHOP_ORDER,CUST_LOT_ID,WAFER_ID,GROSS_DIE_QTY,GOOD_DIE_QTY, MARK_CODE) select a.ordername   as SHOP_ORDER," & "a.waferlot as CUST_LOT_ID,a.waferid as WAFER_ID,a.dieqty as GROSS_DIE_QTY,a.fgdieqty as GOOD_DIE_QTY,a.markingcode as MARK_CODE " & "  from ib_waferlist a, ib_workorder b, MAPPINGDATATEST C, PJ_WO_PRI e where b.ordername = a.ordername and b.ordername = '" & sOrdername & "' AND C.SUBSTRATEID = a.waferid and e.wo = b.ordername"
    AddSql (sOra)

    sOra = "select shop_order from shop_order_detail where shop_order = '" & sOrdername & "'"
    sLog = Get_OracleStr(sOra)

    If sLog <> sOrdername Then
        Call addOrderLogTxt(sOrdername, " 插入新工单shop_order_detail: 无数据")
    Else
        Call addOrderLogTxt(sOrdername, " 插入新工单shop_order_detail: 成功")

    End If

    ' 表2: SHOP_ORDER_PROPERTY
    sOra = "select SHOP_ORDER_PROPERTY_PKG.SHOP_ORDER_PROPERTY('" & sOrdername & "')  from dual"
    AddSql (sOra)

    sOra = "select shop_order from SHOP_ORDER_PROPERTY where shop_order = '" & sOrdername & "'"
    sLog = Get_OracleStr(sOra)

    If sLog <> sOrdername Then
        Call addOrderLogTxt(sOrdername, " 插入新工单SHOP_ORDER_PROPERTY: 无数据")
    Else
        Call addOrderLogTxt(sOrdername, " 插入新工单SHOP_ORDER_PROPERTY: 成功")

    End If

    ' 表3: shop_order
    sOra = "insert into shop_order(SHOP_ORDER,PRD_ID, PRD_VER,ERP_ROUTING, ORDER_QTY, CUST_LOT_QTY, PLAN_STAR_DATE, PLAN_END_DATE, MANF_DEPT, MANF_DEPT_DESC, LOT_TYPE, PRIORITY, PKG, CUST_ID,ERP_CREATE_DATE,CREATOR,flag,ht_device,RELEASE_TYPE)" & _
       "select a.ordername as SHOP_ORDER, b.product as PRD_ID, 'A' as PRD_VER, '' as ERP_ROUTING, COUNT(distinct A.WAFERID) as ORDER_QTY, COUNT(distinct A.WAFERLOT) as CUST_LOT_QTY, B.PLANSTARTDATE AS PLAN_STAR_DATE, B.PLANENDDATE AS PLAN_END_DATE, B.PARA8 AS MANF_DEPT, g.manf_dept_desc AS MANF_DEPT_DESC, e.lot_type as LOT_TYPE, decode(e.pri, 'Hot Lot', 1,'Super Hot Lot',1,4) as PRIORITY, f.pkg_type as PKG, shop_order_property_pkg.SHOP_ORDER_CUSTOMER(b.customer) AS CUST_ID, b.erpcreatedate as ERP_CREATE_DATE, '" + gUserName + "' as CREATOR, '0' as flag, f.qtechptno as ht_device, '1' as RELEASE_TYPE " & _
       "from ib_waferlist a, ib_workorder b, MAPPINGDATATEST C, PJ_WO_PRI e, tbltsvnpiproduct f, MES_DEPT g " & _
       "where b.ordername = a.ordername and b.ordername = '" & sOrdername & "' AND C.SUBSTRATEID = a.waferid and e.wo = b.ordername and f.qtechptno2 = b.product and f.customerptno1 = '" & sCusDevice & "' and g.manf_dept = substr(b.para8,1,instr(b.para8,'_')-1) group by a.ordername,b.product,B.PLANSTARTDATE,B.PLANENDDATE,B.PARA8,e.lot_type,e.pri,f.pkg_type,shop_order_property_pkg.SHOP_ORDER_CUSTOMER(b.customer),b.erpcreatedate,e.creat_by,f.qtechptno,g.manf_dept_desc"
    AddSql (sOra)

    sOra = "select shop_order from shop_order where shop_order = '" & sOrdername & "'"
    sLog = Get_OracleStr(sOra)

    If sLog <> sOrdername Then
        Call addOrderLogTxt(sOrdername, " 插入新工单shop_order: 无数据")
    Else
        Call addOrderLogTxt(sOrdername, " 插入新工单shop_order: 成功")

    End If
    
    ' 表4: Dummy工单发料信息
    If sType = "Dummy工单" Or sType = "重工工单" Or sType = "FO_CSP工单" Then
        
        sOra = "insert into ERPBASE..TblERPFLToME (STOCK_TYPE,STOCK_ID,PRD_ID,PRD_VER,QTY,PRD_DATE,EFF_DATE,SHOP_ORDER,SupSN,Flag,Memo,CreateDate,FStauts,HeaderID) " & "select 'W',b.ORDERNAME + c.WAFERLOT,e.料号,'A',COUNT(*),GETDATE() - 1,GETDATE() + 300,b.ORDERNAME,c.WAFERLOT,0,'',GETDATE(),'','' from erpdata .. tblTSVworkorder b,erpdata .. tblTSVwaferlist c," & "erpdata..tblSmainM2 e where c.ORDERNAME = b.ORDERNAME and b.ORDERNAME in ( '" & sOrdername & "' ) and b.PRODUCT = e.料号 group by b.PRODUCT, b.ORDERNAME,e.料号, c.WAFERLOT"
    
    Else
        
        sOra = "insert into ERPBASE..TblERPFLToME (STOCK_TYPE,STOCK_ID,PRD_ID,PRD_VER,QTY,PRD_DATE,EFF_DATE,SHOP_ORDER,SupSN,Flag,Memo,CreateDate,FStauts,HeaderID) " & "select 'W',b.ORDERNAME +c.WAFERLOT,e.料号,'A',COUNT(*),GETDATE() - 1,GETDATE() + 300,b.ORDERNAME ,c.WAFERLOT,0,'',GETDATE(),'','' from erpdata .. tblTSVworkorder b, " & "erpdata .. tblTSVwaferlist c,ERPBASE .. tblllplan d,erpdata..tblSmainM2 e where c.ORDERNAME = b.ORDERNAME and b.ORDERNAME in ('" & sOrdername & "') and d.工单号 = c.ORDERNAME " & "and (d.物料编号 like '01.01.01%' or d.物料编号 like '03.06.02%') and e.物料编号 = d.物料编号 group by b.PRODUCT, b.ORDERNAME,e.料号, c.WAFERLOT "

    End If
    
    AddSql2 (sOra)
    
    Insert_Shop_Order = True
    
    Cnn.CommitTrans
    INIadoCon.CommitTrans
    
    MsgBox "成功保存工单:" & sOrdername, vbInformation, "提示"
    
    Exit Function

DealError:

    Cnn.RollbackTrans
    INIadoCon.RollbackTrans

    Call addOrderLogTxt(sOrdername, " 插入新工单失败")
    MsgBox "新工单接口执行失败！" + Chr(13) + "原因:" + Err.DESCRIPTION, vbInformation, "警告"

End Function

Public Function Trans_LotData(sOrderType As String) As Boolean

On Error GoTo ERRON

    Cnn.BeginTrans
    INIadoCon.BeginTrans

Trans_LotData = True
    Dim sOra As String

    Dim sSql As String

    ' 插入头表
    sOra = "insert into customeroitbl_test(id,source_batch_id,SHIP_SITE,mpn_desc,mtrl_num,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date) select *  from ORDER_DATA_TEMP_HEADER"
    sSql = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,source_batch_id,SHIP_SITE,mpn_desc,mtrl_num,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date) select * from erpdata.dbo.ORDER_DATA_TEMP_HEADER"

    AddSql (sOra)
    AddSql2 (sSql)

    ' 插入子表
    sOra = "insert into mappingdatatest(id,substrateid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename,PRODUCTID) select * from ORDER_DATA_TEMP_DETAILS"
    sSql = "insert into [ERPBASE].[dbo].[tblmappingData](substrateid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename,PRODUCTID) select * from erpdata.dbo.ORDER_DATA_TEMP_DETAILS"

    AddSql (sOra)
    AddSql2 (sSql)

    ' 插入2次WO
    If sOrderType = "硅基工单" Or sOrderType = "FO_CSP工单" Then
        sOra = "insert into mappingdatatest(id,substrateid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename,PRODUCTID) select id,substrateid || '+',lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename, productid from ORDER_DATA_TEMP_DETAILS"
        sSql = "insert into [ERPBASE].[dbo].[tblmappingData](substrateid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename,PRODUCTID) select substrateid+'+',lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename,productid from erpdata.dbo.ORDER_DATA_TEMP_DETAILS"

        AddSql (sOra)
        AddSql2 (sSql)

    End If
    
    Exit Function
    
ERRON:
    Trans_LotData = False

End Function

Public Function Add_WOPRI(woTemp As String, _
                          pritemp As String, _
                          LotType As String, _
                          player As String) As Boolean

    Dim cmdStr As String

    On Error GoTo DealError
        
    Cnn.BeginTrans

    cmdStr = " insert into PJ_WO_PRI values('" & woTemp & "','" & pritemp & "',to_char(sysdate,'YYYY-MM-DD'),'" & LotType & "', '" & player & "')"
                            
    Exec_Ora (cmdStr)

    Cnn.CommitTrans

    Exit Function

DealError:
    MsgBox "插入PJ_WO_PRI错误"
    Cnn.RollbackTrans
    Add_WOPRI = False

End Function

Public Function Insert_to_repDetails(tRepData As ReproductionWaferData) As Boolean
    Insert_to_repDetails = True
On Error GoTo CANCELTHIS

    Dim sOra     As String

    Dim sSql     As String

    Dim lFailBin As Long

    lFailBin = tRepData.GROSSBINCOUNT - tRepData.PASSBINCOUNT

    sOra = "insert into mappingdatatest(substrateid, lotid, productid, passbincount, failbincount, flag, qtech_created_by, qtech_created_date, wafer_id, customershortname, filename) values('" & tRepData.SUBSTRATEID & "', '" & tRepData.LOTID & "', '" & tRepData.PRODUCTID & "', '" & tRepData.PASSBINCOUNT & "','" & lFailBin & "', 'Y', '" & gUserName & "', sysdate,'" & tRepData.Wafer_id & "', '" & tRepData.CUSTOMERSHORTNAME & "', '" & tRepData.ID & "' )"

    sSql = "insert into [ERPBASE].[dbo].[tblmappingData](substrateid, lotid, productid, passbincount, failbincount, flag, qtech_created_by, qtech_created_date, wafer_id, customershortname, filename) values('" & tRepData.SUBSTRATEID & "', '" & tRepData.LOTID & "', '" & tRepData.PRODUCTID & "', '" & tRepData.PASSBINCOUNT & "','" & lFailBin & "', 'Y', '" & gUserName & "', GETDATE(),'" & tRepData.Wafer_id & "', '" & tRepData.CUSTOMERSHORTNAME & "', '" & tRepData.ID & "' )"

    AddSql (sOra)
    AddSql2 (sSql)
    
    Exit Function

CANCELTHIS:
    Insert_to_repDetails = False

End Function

Public Function Insert_to_repHeader(tRepData As ReproductionWaferData) As Boolean
Insert_to_repHeader = True
On Error GoTo CANCELTHIS
    
    Cnn.BeginTrans
    INIadoCon.BeginTrans
    
    Dim sOra As String

    Dim sSql As String

    sOra = "insert into customeroitbl_test(ID, source_batch_id, mpn_desc, test_mtrl_desc,SOURCE_MTRL_SLOC, flag,qtech_created_by, qtech_created_date, customershortname, Invflag ) values('" & tRepData.ID & "', '" & tRepData.LOTID & "', '" & tRepData.CUSTOMERDEVICE & "', '" & tRepData.JOBNO & "','" & tRepData.JOBNO & "', 'Y', '" & gUserName & "', sysdate, '" & tRepData.CUSTOMERSHORTNAME & "', '0')"

    sSql = "insert into [ERPBASE].[dbo].[tblCustomerOI](ID, source_batch_id, mpn_desc, test_mtrl_desc, SOURCE_MTRL_SLOC, flag,qtech_created_by, qtech_created_date, customershortname ) values('" & tRepData.ID & "', '" & tRepData.LOTID & "', '" & tRepData.CUSTOMERDEVICE & "', '" & tRepData.JOBNO & "','" & tRepData.JOBNO & "', 'Y', '" & gUserName & "', GETDATE(), '" & tRepData.CUSTOMERSHORTNAME & "')"

    AddSql (sOra)
    AddSql2 (sSql)

    Exit Function
    
CANCELTHIS:
    Insert_to_repHeader = False
    
End Function

Public Function CheckGrossDie(iGrossDies As Long, sLotId As String) As Boolean

    Dim sOra         As String

    Dim sOra_diesQty As String

    CheckGrossDie = True

    sOra = "select  max(passbincount + failbincount) from mappingdatatest where lotid = '" & sLotId & "' "

    sOra_diesQty = Get_OracleStr(sOra)

    If iGrossDies > CLng(sOra_diesQty) Then
        CheckGrossDie = False

    End If

End Function

Public Function DistributeMaterial(sOrder As String)

    Dim sOra As String

    sOra = "insert into ERPBASE..TblERPFLToME (STOCK_TYPE,STOCK_ID,PRD_ID,PRD_VER,QTY,PRD_DATE,EFF_DATE,SHOP_ORDER,SupSN,Flag,Memo,CreateDate,FStauts,HeaderID) " & "select 'W',b.ORDERNAME + c.WAFERLOT,e.料号,'A',COUNT(*),GETDATE() - 1,GETDATE() + 300,b.ORDERNAME,c.WAFERLOT,0,'',GETDATE(),'','' from erpdata .. tblTSVworkorder b,erpdata .. tblTSVwaferlist c," & "erpdata..tblSmainM2 e where c.ORDERNAME = b.ORDERNAME and b.ORDERNAME in ( '" & sOrder & "' ) and b.PRODUCT = e.料号 group by b.PRODUCT, b.ORDERNAME,e.料号, c.WAFERLOT"
 
    AddSql (sOra)

End Function

Public Function PrintLabelTxt(filename As String, msgTxt As String, dirtemp As String)

    '判断txt文件是否存在，如不存在，则建立
    Dim fileNameTemp As String

    Dim dirNameTemp  As String

    Dim fileTemp     As String

    dirNameTemp = dirtemp
    fileNameTemp = Replace(filename, "'", "") & ".txt"
    fileTemp = dirNameTemp & fileNameTemp
    
    Open fileTemp For Output As #1   '直接覆盖
    Print #1, msgTxt
    Close #1

End Function

Public Function PrintInbox(sInbox As String, sDN As String)

    Dim txtStr              As String

    Dim dirtemp             As String

    Dim cmdStr2             As String

    Dim fileNameTemp        As String

    Dim msgTxtTemp          As String

    Dim msgTxtTemp2         As String

    Dim qboxNoTemp          As String

    Dim qboxNoContainerTemp As String

    Dim inBoxContainerTemp  As String

    Dim qboxNoSeqTemp       As String

    Dim qboxNoSeqTemp1      As String

    Dim inboxnum            As String

    Dim stqtpj              As String

    Dim sqlDB               As String

    Dim sqlDBRS             As New ADODB.Recordset

    Dim lotStr              As String

    Dim bid

    Dim sOra      As String

    Dim sSql      As String

    Dim rs        As New ADODB.Recordset

    Dim rsIP      As New ADODB.Recordset

    Dim pp        As Integer

    Dim strRQ     As String

    Dim strLSM    As String

    Dim strqbnum  As String

    Dim strqbnum1 As String

    Dim finame    As String

    Dim qbnum     As String

    Dim j         As Integer

    pp = 0

    If sInbox = "" Then
        ' MsgBox "打印完毕", vbInformation
        Exit Function

    End If

    msgTxtTemp = Replace(sInbox, vbCrLf, "','")
    msgTxtTemp2 = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & "," & Right(msgTxtTemp, Len(msgTxtTemp) - InStr(msgTxtTemp, ","))

    bid = Split(msgTxtTemp, "','")

    ' 检查卷盘标签
    For i = 0 To UBound(bid) - 1

        lotStr = bid(i)
   
        If lotStr <> "" Then
    
            ' 先判断是否在仓库
            If Not Judge37TrayIn(lotStr) Then
                '     MsgBox "此卷：" & lotStr & " 不存在于ERP仓库中，不可以合内箱，请确认!", vbInformation, "友情提示"
                '     Exit Function
            Else

                ' 再判断是不是在60000仓与60001仓库
                If Judge37InvType(lotStr) Then

                    'MsgBox "此卷：" & lotStr & " 存在于ERP 6000或6001仓中，不可以合内箱，请确认!", vbInformation, "友情提示"
                    'Exit Function
                End If

            End If
    
            ' 判断有没有装过
            If Judge37ExistInBox1(lotStr) Then

                'MsgBox "此卷：" & lotStr & " 已装过内箱，不可以重复装，请确认!", vbInformation, "友情提示"
                'Exit Function
            End If
        
        End If

    Next i

    ' 1. 插合内盒数据
    strRQ = "NH" + Format(Now(), "YYMMDD")
    sOra = "SELECT MAX(NHBox) NHBox FROM erpdata..TblTSV_INBOX_DETAILS WHERE NHBox LIKE '" & strRQ & "%'"

    Set rs = Get_SqlserveRs(sOra)

    If Trim("" & rs!NHBox) = "" Then
        strLSM = strRQ + "0001"
    Else
        strLSM = strRQ + Right("0000" + Trim$(Val(Right(Trim$("" & rs!NHBox), 4)) + 1), 4)

    End If

    ' 2. 打印Semtech内盒分Lot标签
    sOra = "select b.htlotid as firname from TSV_Tray_details b where b.trayqboxnumber in ('" & msgTxtTemp2 & "') group by b.htlotid"
    Set rsIP = Get_OracleRs(sOra)

    If Not rsIP.EOF Then

        Do While Not rsIP.EOF
    
            finame = rsIP!firname
    
            sOra = "select  '-B'|| substr('00'||(nvl(max(a.seqtxt),0)+1),-2)  from TSV_QBOXTBL_37SEQ a where a.firtname = '" & finame & "' group by a.firtname"
    
            Set rs = Get_OracleRs(sOra)

            If Not rs.EOF Then
                qbnum = rs.Fields(0).Value
            Else
                qbnum = "-B01"

            End If
     
            ' 标签生成
        
            ' 1.Semtech内盒标签(外箱)
        
            sSql = " select replace(a.customerpt,'.P2','') +','+ replace(a.customerlotid,'M','') +','+'1T'+ replace(a.customerlotid,'M','') +','+ replace(a.customerpt,'.P2','') +','+'1P'+ replace(a.customerpt,'.P2','') +','+min(a.podatecode) +','+min(a.podatecode)+',' " & " +max(a.htlotid)+'" & qbnum & "'+','+'S'+max(a.htlotid)+'" & qbnum & "' +','+rtrim(sum(qty)) +','+'Q'+rtrim(sum(qty)) +','" & " +min(a.htdatecode) +','+min(a.htdatecode) " & " from  [erpdata].[dbo].[TblTSV_Tray_details] a where trayqboxnumber in ('" & msgTxtTemp2 & "')  and a.htlotid = '" & finame & "'  " & " group by a.customerpt ,a.customerlotid ,a.customerlotid ,a.customerpt ,a.podatecode ,a.htlotid,a.htdatecode"

            Set rs = Get_SqlserveRs(sSql)

            pp = pp + 1

            fileNameTemp = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & Format(Now(), "YYYYMMDDHHmmSS") & Trim(pp)
            dirtemp = "\\10.160.1.14\BarCode\37\37内箱\"

            Call addLabelTxt(fileNameTemp, rs.Fields(0).Value, dirtemp)
        
            '        ' 2.三星内盒汇总标签
            '        sSql = " select replace(a.customerpt,'.P2','') +','+ '" & strLSM & "' +','+'1T'+ '" & strLSM & "' +','+ '" & strLSM & "' +','+'1P'+ '" & strLSM & "' +','+min(a.podatecode) +','+min(a.podatecode)+',' " & _
            '                " +'" & strLSM & "'+','+'S'+'" & strLSM & "' +','+'" & strLSM & "' +','+'Q'+'" & strLSM & "' +','" & _
            '                " +min(a.htdatecode) +','+min(a.htdatecode) " & _
            '                " from  [erpdata].[dbo].[TblTSV_Tray_details] a where trayqboxnumber in ('" & msgTxtTemp2 & "')  and a.htlotid = '" & finame & "'  " & _
            '                " group by a.customerpt ,a.customerlotid ,a.customerlotid ,a.customerpt ,a.podatecode ,a.htlotid,a.htdatecode"
            '
            '        Set rs = Get_SqlserveRs(sSql)
            '
            '        Call addLabelTxt(fileNameTemp, rs.fields(0).Value, dirtemp)

            For i = 0 To UBound(bid) - 1
                lotStr = bid(i)
    
                stqtpj = Mid(qbnum, 3, 2)
   
                If Mid(lotStr, 2, InStr(lotStr, "-") - 2) = finame Then
            
                    sSql = " insert into [erpdata].[dbo].[TblTSV_INBOX_DETAILS](id,Containername,Subcontainername,Labeltype,customerpt,customerlotid,htlotid,podatecode,htdatecode,qty,flag,created_by,created_date,NHBox)" & " select   1, 'S'+inboxno,trayboxno,typename,mpn_desc,jobnumber,htlotid,date_code,htdatecode,qty,'Y','" & useridTemp & "',getdate(),'" & strLSM & "' from ( " & " select  max(a.htlotid) +'" & qbnum & "'  as inboxno,  '" & lotStr & "' as trayboxno , 'INQbox' as typename," & " a.customerpt as mpn_desc,a.customerlotid as jobnumber,max(a.htlotid) as htlotid ,min(a.podatecode) as date_code,min(a.htdatecode) as htdatecode,sum(qty) as qty" & " from  [erpdata].[dbo].[TblTSV_Tray_details] a where trayqboxnumber in ('" & lotStr & "' )  " & " group by a.customerpt ,a.customerlotid ,a.customerlotid ,a.customerpt ,a.podatecode ,a.htdatecode ) X "
 
                    sOra = " insert into TSV_QBOXTBL_37SEQ(typename,createdate,seqtxt,containername,Firtname) " & "  values ('INQbox',sysdate,'" & stqtpj & "','',substr( '" & lotStr & "',2,instr( '" & lotStr & "','-R')-2))"

                    Exec_Sql (sSql)
                    Exec_Ora (sOra)
        
                End If

            Next i
    
            rsIP.MoveNext
        Loop
       
        ' 2.打印三星标签
        Sleep (5000)

        FrmSemtech_LablePrint.Show
        FrmSemtech_LablePrint.Hide
        FrmSemtech_LablePrint.cmbDN.text = sDN
        ' 1)内盒
        FrmSemtech_LablePrint.Opt(1).Value = True
        ' 查询
        Call FrmSemtech_LablePrint.cmd_Click(0)
    
        With FrmSemtech_LablePrint.fpS(0)
        
            For j = 1 To .MaxRows
                .Row = j
                .Col = 2

                If .text = strLSM Then
                    .Col = 1
                    .text = 1

                End If

            Next

        End With
    
        ' 打印
        Sleep (2000)
        Call FrmSemtech_LablePrint.cmd_Click(2)
    
        ' 2)卷盘
        FrmSemtech_LablePrint.Opt(0).Value = True
        ' 查询
        Call FrmSemtech_LablePrint.cmd_Click(0)
    
        For i = 0 To UBound(bid) - 1
            lotStr = bid(i)

            If lotStr <> "" Then

                With FrmSemtech_LablePrint.fpS(0)

                    For j = 1 To .MaxRows
                        .Row = j
                        .Col = 2

                        If .text = lotStr Then
                            .Col = 1
                            .text = 1

                        End If

                    Next

                End With

            End If

        Next i
        
        ' 打印
        Sleep (2000)
        Call FrmSemtech_LablePrint.cmd_Click(2)
    
    Else
        qbnum = "-B01"
    
        sqlDB = " select replace(a.customerpt,'.P2','') +','+ replace(a.customerlotid,'M','') +','+'1T'+ replace(a.customerlotid,'M','') +','+ replace(a.customerpt,'.P2','') +','+'1P'+ replace(a.customerpt,'.P2','') +','+min(a.podatecode) +','+min(a.podatecode)+',' " & " +max(a.htlotid)+'" & qbnum & "'+','+'S'+max(a.htlotid)+'" & qbnum & "' +','+rtrim(sum(qty)) +','+'Q'+rtrim(sum(qty)) +','" & " +min(a.htdatecode) +','+min(a.htdatecode) " & " from  [erpdata].[dbo].[TblTSV_Tray_details] a where trayqboxnumber in ('" & msgTxtTemp2 & "') " & " group by a.customerpt ,a.customerlotid ,a.customerlotid ,a.customerpt ,a.podatecode ,a.htlotid,a.htdatecode"

        If sqlDBRS.State = adStateOpen Then
            sqlDBRS.Close

        End If
    
        sqlDBRS.Open sqlDB, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

        Dim dd As Integer

        dd = 0
    
        If Not sqlDBRS.EOF Then
    
            Do While Not sqlDBRS.EOF
                pp = pp + 1
                
                fileNameTemp = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & Format(Now(), "YYYYMMDDHHmmSS") & Trim(dd)
                dirtemp = TxtDirInQbox.text

                Call addLabelTxt(fileNameTemp, sqlDBRS.Fields(0).Value, dirtemp)
            
                sqlDBRS.MoveNext
            Loop
    
        End If
    
        'add 把 内箱号，几个Tray号，保存到一个表里
        For i = 0 To iPos - 1
            lotStr = bid(i)
    
            stqtpj = Mid(qbnum, 3, 2)
    
            If lotStr <> "" Then
        
                cmdStr2 = " insert into [erpdata].[dbo].[TblTSV_INBOX_DETAILS](id,Containername,Subcontainername,Labeltype,customerpt,customerlotid,htlotid,podatecode,htdatecode,qty,flag,created_by,created_date,NHBox)" & " select   1, 'S'+inboxno,trayboxno,typename,mpn_desc,jobnumber,htlotid,date_code,htdatecode,qty,'Y','" & useridTemp & "',getdate(),'" & strLSM & "' from ( " & " select  max(a.htlotid) +'" & qbnum & "'  as inboxno,  '" & lotStr & "' as trayboxno , 'INQbox' as typename," & " a.customerpt as mpn_desc,a.customerlotid as jobnumber,max(a.htlotid) as htlotid ,min(a.podatecode) as date_code,min(a.htdatecode) as htdatecode,sum(qty) as qty" & " from  [erpdata].[dbo].[TblTSV_Tray_details] a where trayqboxnumber in ('" & bid(i) & "' ) " & " group by a.customerpt ,a.customerlotid ,a.customerlotid ,a.customerpt ,a.podatecode ,a.htdatecode ) X "
 
                cmdStpj = " insert into TSV_QBOXTBL_37SEQ(typename,createdate,seqtxt,containername,Firtname) " & "  values ('INQbox',sysdate,'" & stqtpj & "','" & lotStr & "',substr( '" & lotStr & "',2,instr( '" & lotStr & "','-R')-2))"
                    
                AddSql2 (cmdStr2)
                AddSql (cmdStpj)

            End If
  
        Next i

    End If

    Dim sSub As String

    Sleep (10000)

    ' 检查卷盘标签
    For i = 0 To UBound(bid) - 1

        lotStr = bid(i)
   
        If lotStr <> "" Then
            ' 回插
            sSql = "select CONTAINERNAME from [erpdata].[dbo].[TblTSV_INBOX_DETAILS] where SUBCONTAINERNAME = '" & lotStr & "' "
            Set rs = Get_SqlserveRs(sSql)
            
            sSub = rs.Fields("CONTAINERNAME")
            
            sOra = "update PACKING_DETAILED set inbox_num =  '" & sSub & "' where trayid = '" & lotStr & "'"
        
            Exec_Ora (sOra)

        End If

    Next i

End Function

' 检查Oracle是否有记录
Public Function IsOraRecord(sOra As String) As Boolean

    Dim bRtn As Boolean

    bRtn = False

    If Get_OracleCnt(sOra) = 0 Then
        bRtn = False    ' 不存在
    Else
        bRtn = True ' 存在

    End If

    IsOraRecord = bRtn

End Function

Public Function IsSqlRecord(sSql As String) As Boolean

    Dim bRtn As Boolean

    bRtn = False

    If Get_SqlserverCnt(sSql) = 0 Then
        bRtn = False    ' 不存在
    Else
        bRtn = True ' 存在

    End If

    IsSqlRecord = bRtn

End Function

Public Function CheckKID(sCartonID As String, sDN As String) As Boolean

    Dim sOra As String

    sOra = "select * from PACKING_DETAILED where dn_num = '" & sDN & "' and kid = '" & sCartonID & "'"
    CheckKID = IsOraRecord(sOra)

End Function

Public Function GetDNType(sDN As String) As String

    Dim sOra  As String

    Dim sType As String

    sOra = "select UPPER(labelrequirement) as type from CUSTOMERSHIPPINGUPTBL where delivery = '" & sDN & "'"
    sType = Get_OracleStr(sOra)

    If InStr(sType, "SAMSUNG") Then
        GetDNType = "SSE2"
        
    End If

    If InStr(sType, "HUAWEI") Then
        GetDNType = "HW"

    End If

    If InStr(sType, "SEMTECH") Then
        GetDNType = "ST"

    End If
    
    If InStr(sType, "SHORT") Then
        GetDNType = "SHORT"
    End If

End Function

'************************************************************************************************

' 连接ORACLE:
Public Sub ConnectOracle()

    On Error Resume Next

    While OraConnect.State = 0

        OraConnect.Open "Provider=OraOLEDB.Oracle.1;Password=KsMesDB_ht89;User ID=insiteqt2;Data Source=testmes;Persist Security Info=True"
    
        If OraConnect.State <> 0 Then
            GoTo EndCon

        End If

    Wend

EndCon:

End Sub

' 连接SQLSERVER:
Public Sub ConnectSql()

    Const strSrvName As String = "10.160.1.13"   '数据服务器名称10.160.1.13

    Const strDbName  As String = "erpbase"   '数据库名称

    Const strUID     As String = "sa"      '登录用户名称

    Const strPSWD    As String = "ksxtDB"     '登录密码

    Dim INIstrCnn    As String '连接字符串

    Set SqlConnect = New ADODB.Connection
    INIstrCnn = "driver={SQL Server};server=" & strSrvName & ";UID=" & strUID & "; " & "pwd=" & strPSWD & ";database=" & strDbName & ""
    SqlConnect.CursorLocation = adUseClient
    SqlConnect.ConnectionTimeout = 120
    SqlConnect.Open INIstrCnn

    If SqlConnect.State = 0 Then
        MsgBox "错误:" & Err.DESCRIPTION & vbCrLf & " 解决方法请寻求有关帮助。", vbExclamation, "系统"
        Exit Sub

    End If

End Sub

' 连接SQLSERVER:
Public Sub ConnectSql2()

    Const strSrvName As String = "10.160.8.15"   '数据服务器名称10.160.1.13

    Const strDbName  As String = "XTW"   '数据库名称

    Const strUID     As String = "sa"      '登录用户名称

    Const strPSWD    As String = "123"     '登录密码

    Dim INIstrCnn    As String '连接字符串

    Set SqlConnect2 = New ADODB.Connection
    INIstrCnn = "driver={SQL Server};server=" & strSrvName & ";UID=" & strUID & "; " & "pwd=" & strPSWD & ";database=" & strDbName & ""
    SqlConnect2.CursorLocation = adUseClient
    SqlConnect2.ConnectionTimeout = 120
    SqlConnect2.Open INIstrCnn

    If SqlConnect2.State = 0 Then
        MsgBox "错误:" & Err.DESCRIPTION & vbCrLf & " 解决方法请寻求有关帮助。", vbExclamation, "系统"
        Exit Sub

    End If

End Sub


Public Function GetBinQty(strWaferID As String, strBin As String) As Long
Dim lRtn As Long
Dim strSql As String

lRtn = 0

strSql = "select isnull(SUM(A.qty), 0) from erptemp..WAFER_BIN_LIST A " & _
" inner join erpdata..tblErpInStockRelation B on A.SFC = B.SFC_ID and CHARINDEX(A.WAFER_ID, B.WAFER_ID) <> 0 and A.WAFER_ID = '" & strWaferID & "' and A.GRADES = '" & strBin & "' "

lRtn = Get_SqlserverNo(strSql)

GetBinQty = lRtn

End Function

Public Function getMediaUrl(strVicCmd As String) As String

    getMediaUrl = strMediaDir & strVicCmd & ".wav"

End Function

Public Function chkDNToHW(strDN As String) As Boolean
Dim strSql As String
chkDNToHW = False

strSql = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "'"
If Get_OracleCnt(strSql) = 0 Then
    Frm_37LblPrint.m1.url = getMediaUrl("D N不正确 或挑料信息没有上传")
    Exit Function
End If

'strSql = "select * from PRINT_37FLAG where dn = '" & strDN & "' and PRINTED = '1' "
'If Get_OracleCnt(strSql) > 0 Then
'    Frm_37LblPrint.m1.url = getMediaUrl("该D N已经出过标签, 请勿再次打印")
'    Exit Function
'End If

Frm_37LblPrint.m1.url = getMediaUrl("D N已获取,请依次扫描挑料卷盘")
chkDNToHW = True
End Function

Public Function chkDNToHW_ONELOT(strDN As String) As Boolean
Dim strSql As String
chkDNToHW_ONELOT = False

strSql = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "'"
If Get_OracleCnt(strSql) = 0 Then
    Frm_37LblPrint_ONELOT.m1.url = getMediaUrl("D N不正确 或挑料信息没有上传")
    Exit Function
End If

'strSql = "select * from PRINT_37FLAG where dn = '" & strDN & "' and PRINTED = '1' "
'If Get_OracleCnt(strSql) > 0 Then
'    Frm_37LblPrint_ONELOT.m1.url = getMediaUrl("该D N已经出过标签, 请勿再次打印")
'    Exit Function
'End If

Frm_37LblPrint_ONELOT.m1.url = getMediaUrl("D N已获取,请依次扫描挑料卷盘")
chkDNToHW_ONELOT = True
End Function

Public Function chkReelID(strReelID As String, strDN As String, strJob As String, strMPN As String) As Boolean
Dim strSql As String

Dim lLastMpnQty As Long, lLastJobQty As Long, lMaxJobQty As Long, lMaxMpnQty As Long
Dim strLastMPN As String, strLastJob As String
Dim rs As New ADODB.Recordset

chkReelID = False

strSql = "select * from ST_TR_SEQ where reelid = '" & strReelID & "'"
If Get_OracleCnt(strSql) > 0 Then
    Frm_37LblPrint.m1.url = getMediaUrl("该卷盘已经扫描过, 请勿重复扫描")
    Exit Function
End If

strSql = "SELECT * FROM [erpdata].[dbo].[tblstocknumsub] where 箱号 = '" & strReelID & "'"
If Get_SqlserverCnt(strSql) = 0 Then
    Frm_37LblPrint.m1.url = getMediaUrl("卷盘号有误,或者已经出货")
    Exit Function
End If

strSql = "SELECT * FROM [erpdata].[dbo].[tblstocknumsub] WHERE 箱号= '" & strReelID & "' and 库房编号 in (36,37)"
If Get_SqlserverCnt(strSql) > 0 Then
    Frm_37LblPrint.m1.url = getMediaUrl("该卷盘在6000,6001仓, 不可合箱")
    Exit Function
End If

strSql = "select * from packing_detailed where trayid = '" & strReelID & "'"
If Get_OracleCnt(strSql) > 0 Then
    Frm_37LblPrint.m1.url = getMediaUrl("该卷盘已经合箱打印,请确认是否有误")
    Exit Function
End If

strSql = "select * from CUSTOMERSHIPPINGUPTBL where batchnumber = '" & strJob & "' and delivery = '" & strDN & "'"
If Get_OracleCnt(strSql) = 0 Then
    Frm_37LblPrint.m1.url = getMediaUrl("该卷盘的JOB号不是本次D N的JOB, 卷盘挑料出错")
    Exit Function
End If

strSql = "select job,dev from ST_TR_SEQ where dn = '" & strDN & "' order by seq desc"
Set rs = Get_OracleRs(strSql)
If rs.RecordCount > 0 Then
    strLastJob = "" & rs!JOB
    strLastMPN = "" & rs!DEV
End If

If strLastJob <> "" And strLastMPN <> "" Then
    ' Job
    strSql = "select sum(qty) qty from ST_TR_SEQ where dn = '" & strDN & "' and job = '" & strLastJob & "'"
    lLastJobQty = Get_OracleNo(strSql)
    
    strSql = "select sum(quantity) qty from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "' and batchnumber = '" & strLastJob & "'"
    lMaxJobQty = Get_OracleNo(strSql)
    
    If lMaxJobQty = lLastJobQty And strJob = strLastJob Then
        Frm_37LblPrint.m1.url = getMediaUrl("该JOB已扫完,请勿再次扫描该JOB的卷盘")
        MsgBox "该JOB: " & strJob & "已扫完,请勿再次扫描该JOB的卷盘", vbExclamation, "警告"
        
        Exit Function
    End If
    
    If lMaxJobQty > lLastJobQty And strJob <> strLastJob Then
        Frm_37LblPrint.m1.url = getMediaUrl("上一个JOB没有扫完,请勿扫描其他JOB的卷盘")
        MsgBox "上一个JOB:" & strLastJob & vbCrLf & "没有扫完,请勿扫描其他JOB的卷盘", vbExclamation, "警告"
        
        Exit Function
    End If
    
    If lMaxJobQty < lLastJobQty Then
        Frm_37LblPrint.m1.url = getMediaUrl("上个Job数量超出, 挑料出错")
        MsgBox "上一个JOB:" & strLastJob & vbCrLf & "数量超出, 挑料出错", vbExclamation, "警告"
        
        Exit Function
    End If
    
    ' M.P.N
    strSql = "select sum(qty) qty from ST_TR_SEQ where dn = '" & strDN & "' and dev = '" & strLastMPN & "'"
    lLastMpnQty = Get_OracleNo(strSql)
    
    strSql = "select sum(quantity) qty from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "' and marketingpn = '" & strLastMPN & "'"
    lMaxMpnQty = Get_OracleNo(strSql)
    
    If lMaxMpnQty = lLastMpnQty And strMPN = strLastMPN Then
        Frm_37LblPrint.m1.url = getMediaUrl("该机种已扫完,请勿再次扫描该机种的卷盘")
        MsgBox "该机种: " & strMPN & "已扫完,请勿再次扫描该机种的卷盘", vbExclamation, "警告"
        
        Exit Function
    End If
    
    If lMaxMpnQty > lLastMpnQty And strMPN <> strLastMPN Then
        Frm_37LblPrint.m1.url = getMediaUrl("上一个机种没有扫完,请勿扫描其他机种的卷盘")
        MsgBox "上一个机种: " & strLastMPN & vbCrLf & "没有扫完,请勿扫描其他机种的卷盘", vbExclamation, "警告"
        
        Exit Function
    End If
    
    If lMaxMpnQty < lLastMpnQty Then
        Frm_37LblPrint.m1.url = getMediaUrl("上个机种数量超出, 挑料出错")
        MsgBox "上一个机种: " & strLastMPN & vbCrLf & "数量超出,挑料出错", vbExclamation, "警告"
        
        Exit Function
    End If

End If

Frm_37LblPrint.m1.url = getMediaUrl("卷盘号正确")
chkReelID = True

End Function
Public Function chkReelID_ONELOT(strReelID As String, strDN As String, strJob As String, strMPN As String) As Boolean
Dim strSql As String

Dim lLastMpnQty As Long, lLastJobQty As Long, lMaxJobQty As Long, lMaxMpnQty As Long
Dim strLastMPN As String, strLastJob As String
Dim rs As New ADODB.Recordset

chkReelID_ONELOT = False

strSql = "select * from packing_detailed where trayid = '" & strReelID & "' and dn_num = '" & strDN & "'"
If Get_OracleCnt(strSql) > 0 Then
    Frm_37LblPrint_ONELOT.m1.url = getMediaUrl("该卷盘已经扫描过, 请勿重复扫描")
    Exit Function
End If

strSql = "SELECT * FROM [erpdata].[dbo].[tblstocknumsub] where 箱号 = '" & strReelID & "'"
If Get_SqlserverCnt(strSql) = 0 Then
    Frm_37LblPrint_ONELOT.m1.url = getMediaUrl("卷盘号有误,或者已经出货")
    Exit Function
End If

strSql = "SELECT * FROM [erpdata].[dbo].[tblstocknumsub] WHERE 箱号= '" & strReelID & "' and 库房编号 in (36,37)"
If Get_SqlserverCnt(strSql) > 0 Then
    Frm_37LblPrint_ONELOT.m1.url = getMediaUrl("该卷盘在6000,6001仓, 不可合箱")

End If

strSql = "select job_id,customer_device from packing_detailed where dn_num = '" & strDN & "' order by seq desc"
Set rs = Get_OracleRs(strSql)
If rs.RecordCount > 0 Then
    strLastJob = "" & rs!JOB_ID
    strLastMPN = "" & rs!Customer_Device
End If

If strLastJob <> "" And strLastMPN <> "" Then
    ' Job
    strSql = "select sum(qty) qty from packing_detailed where dn_num = '" & strDN & "' and job_id = '" & strLastJob & "'"
    lLastJobQty = Get_OracleNo(strSql)
    
    strSql = "select sum(quantity) qty from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "' and batchnumber = '" & strLastJob & "'"
    lMaxJobQty = Get_OracleNo(strSql)
    
    If lMaxJobQty = lLastJobQty And strJob = strLastJob Then
        Frm_37LblPrint_ONELOT.m1.url = getMediaUrl("该JOB已扫完,请勿再次扫描该JOB的卷盘")
        MsgBox "该JOB: " & strJob & "已扫完,请勿再次扫描该JOB的卷盘", vbExclamation, "警告"
        
        Exit Function
    End If
    
    If lMaxJobQty > lLastJobQty And strJob <> strLastJob Then
        Frm_37LblPrint_ONELOT.m1.url = getMediaUrl("上一个JOB没有扫完,请勿扫描其他JOB的卷盘")
        MsgBox "上一个JOB:" & strLastJob & vbCrLf & "没有扫完,请勿扫描其他JOB的卷盘", vbExclamation, "警告"
        
        Exit Function
    End If
    
    If lMaxJobQty < lLastJobQty Then
        Frm_37LblPrint_ONELOT.m1.url = getMediaUrl("上个Job数量超出, 挑料出错")
        MsgBox "上一个JOB:" & strLastJob & vbCrLf & "数量超出, 挑料出错", vbExclamation, "警告"
        
        Exit Function
    End If
    
    ' M.P.N
    strSql = "select sum(qty) qty from packing_detailed where dn_num = '" & strDN & "' and customer_device = '" & strLastMPN & "'"
    lLastMpnQty = Get_OracleNo(strSql)
    
    strSql = "select sum(quantity) qty from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "' and marketingpn = '" & strLastMPN & "'"
    lMaxMpnQty = Get_OracleNo(strSql)
    
    If lMaxMpnQty = lLastMpnQty And strMPN = strLastMPN Then
        Frm_37LblPrint_ONELOT.m1.url = getMediaUrl("该机种已扫完,请勿再次扫描该机种的卷盘")
        MsgBox "该机种: " & strMPN & "已扫完,请勿再次扫描该机种的卷盘", vbExclamation, "警告"
        
        Exit Function
    End If
    
    If lMaxMpnQty > lLastMpnQty And strMPN <> strLastMPN Then
        Frm_37LblPrint_ONELOT.m1.url = getMediaUrl("上一个机种没有扫完,请勿扫描其他机种的卷盘")
        MsgBox "上一个机种: " & strLastMPN & vbCrLf & "没有扫完,请勿扫描其他机种的卷盘", vbExclamation, "警告"
        
        Exit Function
    End If
    
    If lMaxMpnQty < lLastMpnQty Then
        Frm_37LblPrint_ONELOT.m1.url = getMediaUrl("上个机种数量超出, 挑料出错")
        MsgBox "上一个机种: " & strLastMPN & vbCrLf & "数量超出,挑料出错", vbExclamation, "警告"
        
        Exit Function
    End If

End If

Frm_37LblPrint_ONELOT.m1.url = getMediaUrl("卷盘号正确")
chkReelID_ONELOT = True

End Function

Public Function insertReelID(strReelID As String, strDN As String, strJob As String, strMPN As String, strQty As String, lMaxQty As Long) As String
Dim strSql As String
Dim strseq As String
Dim strHtLotID As String
Dim lQty As Long

Cnn.BeginTrans

strSql = "select nvl(max(seq)+1, 1) from ST_TR_SEQ where dn = '" & strDN & "'  "
strseq = Get_OracleStr(strSql)

strHtLotID = Left(strReelID, InStr(strReelID, "-") - 1)

strSql = "insert into ST_TR_SEQ(dn, job, dev, qty, Seqtime, lotid,Reelid, seq) values('" & strDN & "', '" & strJob & "', '" & strMPN & "', '" & strQty & "',sysdate,'" & strHtLotID & "', '" & strReelID & "', '" & strseq & "')"
AddSql (strSql)

strSql = "select sum(qty) from ST_TR_SEQ where dn = '" & strDN & "'"
lQty = Get_OracleNo(strSql)
insertReelID = lQty

If lQty > lMaxQty Then
    Exit Function
End If

Cnn.CommitTrans
End Function

Public Function insertReelID_ONELOT(strReelID As String, _
                                    strDN As String, _
                                    strJob As String, _
                                    strMPN As String, _
                                    lReelQty As Long) As String
Dim strSql     As String
Dim strseq     As String
Dim strHtLotID As String
Dim lCurQty       As Long
Dim rs         As New ADODB.Recordset
Dim strMaxOP   As String
Dim strMaxIP   As String
Dim tD         As tSTData
Dim strLastMPN     As String
Dim strLastJob     As String
Dim strLastReelCnt As String
Dim strLastBoxCnt  As String
Dim lDNMaxQty As Long

Cnn.BeginTrans
strSql = "select nvl(max(seq)+1, 1) from packing_detailed where dn_num = '" & strDN & "'  "
strseq = Get_OracleStr(strSql)
strHtLotID = Left(strReelID, InStr(strReelID, "-") - 1)
strMaxOP = Get_OracleStr("select nvl(max(OUTBOX_NUM),1) from PACKING_DETAILED where dn_num = '" & strDN & "'")
strMaxIP = Get_OracleStr("select nvl(max(INBOX_NUM),1) from PACKING_DETAILED where dn_num = '" & strDN & "' and OUTBOX_NUM = '" & strMaxOP & "' ")
strLastMPN = Get_OracleStr("select CUSTOMER_DEVICE from packing_DETAILED where dn_num = '" & strDN & "' order by seq desc")
strLastJob = Get_OracleStr("select JOB_ID from packing_DETAILED where dn_num = '" & strDN & "' order by seq desc")
strLastReelCnt = Get_OracleStr("select count(*) from packing_detailed where dn_num = '" & strDN & "' and outbox_num = '" & strMaxOP & "' and inbox_num = '" & strMaxIP & "'   ")

If strMPN <> strLastMPN And strLastMPN <> "" Then
    strMaxOP = strMaxOP + 1
    strMaxIP = 1
Else

    If (strJob <> strLastJob And strLastJob <> "") Or (strLastReelCnt = 9) Then
        strMaxIP = strMaxIP + 1

        If strMaxIP = 13 Then
            strMaxIP = 1
            strMaxOP = strMaxOP + 1

        End If

    End If

End If

tD.TRAYID = strReelID
tD.INBOX_NUM = strMaxIP
tD.OUTBOX_NUM = strMaxOP
tD.DN_NUM = strDN
tD.JOB_ID = strJob
tD.QTY = lReelQty
tD.Customer_Device = strMPN
tD.SEQ = strseq
tD.KID = "K" & strMaxOP
tD.DC = get37_DC(tD.JOB_ID, tD.TRAYID)
tD.PSN = get37_PSN(tD)

strSql = "insert into PACKING_DETAILED(TRAYID,INBOX_NUM,OUTBOX_NUM,DN_NUM,JOB_ID,QTY,CUSTOMER_DEVICE,CREATE_DATE,CREATE_BY,PRINT_FLAG,FLAG,KID,SEQ,DATECODE,REELID) " & _
" values('" & tD.TRAYID & "', '" & tD.INBOX_NUM & "','" & tD.OUTBOX_NUM & "', '" & tD.DN_NUM & "','" & tD.JOB_ID & "','" & tD.QTY & "','" & tD.Customer_Device & "', sysdate, '" & gUserName & "' ,'0','0','" & tD.KID & "','" & tD.SEQ & "', '" & tD.DC & "','" & tD.PSN & "')"

AddSql (strSql)

strSql = "select sum(qty) from packing_detailed where dn_num = '" & strDN & "'"
lCurQty = Get_OracleNo(strSql)

strSql = "select sum(quantity) from customershippinguptbl where delivery = '" & strDN & "'"
lDNMaxQty = Get_OracleNo(strSql)

If lCurQty > lDNMaxQty Then
    Exit Function

End If

Cnn.CommitTrans

End Function

Public Function get37_DC(strJob As String, strTrayID As String) As String
Dim strWaferID  As String
Dim strDateCode As String
Dim strSql      As String
Dim strJobNew   As String
Dim strContent  As String
Dim str1 As String
Dim strBartenName As String

'str1 = "37_FIRST_FINISH_YYWW_MON"
str1 = "37_DATE_CODE"
strBartenName = "37TRAY.btw"

strSql = "select top 1 Content from erpdata..tblME_PrintInfo aa ," & _
"erpdata..tblErpInStockDetailInfo bb where bb.KEY_VALUE = '" & strTrayID & "' +  '|' +  '" & strJob & "' " & _
"and bb.keyid = aa.EVENT_ID and bb.KEY_NAME = 'CONTAINER_NAME'  and bb.KEY_TYPE = 'T' " & _
"and aa.BartenderName = '" & strBartenName & "' " & _
"order by ID desc"

strContent = Get_SqlStr(strSql)

If strContent = "" Then
    strSql = "select top 1 Content from erpdata..tblME_PrintInfo_BACK190603 aa ," & "erpdata..tblErpInStockDetailInfo bb where bb.KEY_VALUE = '" & strTrayID & "' +  '|' +  '" & strJob & "' " & "and bb.keyid = aa.EVENT_ID and bb.KEY_NAME = 'CONTAINER_NAME'  and bb.KEY_TYPE = 'T' " & "and aa.BartenderName = '" & strBartenName & "' " & "order by ID desc"
    strContent = Get_SqlStr(strSql)
    If strContent = "" Then
        strJobNew = Replace$(strJob, "M", "")
        strSql = "select distinct case when create_date >= to_date(to_char(create_date, 'yyyy') || '-12-31', 'yyyy-mm-dd') - mod(to_char(create_date, 'YYYY'), 7) - 5  then to_char(create_date, 'yyww') " & "else to_char(create_date + mod(mod(to_char(create_date, 'YYYY'), 7) + 5, 7),'yyww') end as PODATECODE " & "from customeroitbl_test a ,mappingdatatest b ,weight37 c where a.test_mtrl_desc = '" & strJobNew & "' and b.filename = to_char(a.id) and b.lotid = a.source_batch_id " & "and c.waferid = replace(b.substrateid,'+','') "
        get37_DC = Get_OracleStr(strSql)
        Exit Function
    
        'MsgBox "DATECODE无法获取,请联系3831 IT", vbInformation, "提示"
        'GetDateCode = ""
        'Exit Function

    End If
End If

strDateCode = Mid$(strContent, InStr(strContent, str1) + Len(str1) + 3, 4)
get37_DC = strDateCode

End Function

Public Sub InsertPkgID(strDN As String)

    Dim strSql     As String
      
    Dim strLastMPN As String
    
    Dim tD         As tSTData

    Dim i          As Integer, j As Integer, K As Integer

    Dim rs         As New ADODB.Recordset

    i = 0
    j = 1
    K = 1
    
    strLastMPN = ""

    strSql = "select dn, job,dev, qty,lotid, reelid,seq from ST_TR_SEQ where dn = '" & strDN & "' order by seq "
    Set rs = Get_OracleRs(strSql)

    If rs.RecordCount = 0 Then
        MsgBox "DN扫描记录被删除,请联系IT确认", vbInformation, "提示"
        rs.Close
        Set rs = Nothing
        Exit Sub

    End If

    rs.MoveFirst

    Do While Not rs.EOF
    
        If strLastMPN <> "" And rs!DEV <> strLastMPN Then
            K = K + 1
            i = 1
            j = 1
            
            strLastMPN = rs!DEV
        Else

            If strLastMPN = "" Then
                strLastMPN = rs!DEV

            End If
            
            i = i + 1
           
            If i = 10 Then
                i = 1
                j = j + 1
              
            End If
           
            If j = 13 Then
                j = 1
                K = K + 1

            End If

        End If
       
        tD.TRAYID = rs!REELID
        tD.INBOX_NUM = j
        tD.OUTBOX_NUM = K
        tD.DN_NUM = strDN
        tD.JOB_ID = rs!JOB
        tD.QTY = rs!QTY
        tD.Customer_Device = rs!DEV
        tD.SEQ = rs!SEQ
        tD.KID = "K" & K
        tD.DC = get37_DC(tD.JOB_ID, tD.TRAYID)
        tD.PSN = get37_PSN(tD)

        strSql = "insert into PACKING_DETAILED(TRAYID,INBOX_NUM,OUTBOX_NUM,DN_NUM,JOB_ID,QTY,CUSTOMER_DEVICE,CREATE_DATE,CREATE_BY,PRINT_FLAG,FLAG,KID,SEQ,DATECODE,REELID) " & " values('" & tD.TRAYID & "', '" & tD.INBOX_NUM & "','" & tD.OUTBOX_NUM & "', '" & tD.DN_NUM & "','" & tD.JOB_ID & "','" & tD.QTY & "','" & tD.Customer_Device & "', sysdate, '" & gUserName & "' ,'0','0','" & tD.KID & "','" & tD.SEQ & "', '" & tD.DC & "','" & tD.PSN & "')"
        
        AddSql (strSql)
       
        rs.MoveNext
    Loop

End Sub

Public Function get37_PSN(tD As tSTData) As String

Dim strPSN As String
Dim strSql As String
Dim strCPN As String
Dim strRand As String
Dim strMon As String
Dim lBase As Long, lCnt As Long

strSql = "select customerpartnumber from CUSTOMERSHIPPINGUPTBL where batchnumber = '" & tD.JOB_ID & "' and delivery = '" & tD.DN_NUM & "'"
strCPN = UCase(Get_OracleStr(strSql))

strMon = strCPN & Right(Year(Now), 2) & Hex(Month(Now))
lBase = 166576   ' 004LXR  - 004WB8 =  4* 10
strSql = "select nvl(count(*) + 1, 1) from REEL_REC_37 where mon = '" & strMon & "'"
lCnt = lBase + Get_OracleNo(strSql)

strRand = Right("000000" & Get10To33(lCnt), 6)

If Len(strCPN) = 8 Then
    strPSN = "P" & strCPN & "S" & Right(Year(Now), 2) & Hex(Month(Now)) & strRand
Else
    strPSN = "P" & strCPN & "/" & "S" & Right(Year(Now), 2) & Hex(Month(Now)) & strRand
End If

If Left(strRand, 4) = "004W" Then
    MsgBox "PSN流水段吃紧, 请及时联系IT", vbInformation, "提示"
End If

get37_PSN = strPSN

strSql = "insert into REEL_REC_37(REELID,MON,CREATE_DATE) values('" & tD.TRAYID & "','" & strMon & "', sysdate)"
AddSql (strSql)

End Function

Public Function Get10To33(lData As Long) As String
Dim strOut As String

strOut = ""

Do
    If (lData Mod 33) = 0 Then
        strOut = "0" & strOut
    
    Else
        
        strOut = get33Char(lData Mod 33) & strOut
    End If
    
    Get10To33 = strOut
    lData = lData \ 33
    
Loop Until (lData = 0)

End Function

Public Function get33Char(iCh As Long) As String

Dim str123 As String

str123 = "123456789ABCDEFGHJKLMNPQRSTUVWXY"
get33Char = Mid$(str123, iCh, 1)

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       updatePackingBYJOBID
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-16:05:32
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function updatePackingBYJOBID(strDN As String, strJobID As String)
Dim strSql  As String
Dim strSql2 As String
Dim rs      As New ADODB.Recordset
Dim Rs2     As New ADODB.Recordset
Dim i       As Integer    '外箱
Dim j       As Integer    '内箱
Dim strno   As String
Dim strBase As String

strSql = "select distinct outbox_num from PACKING_DETAILED where dn_num = '" & strDN & "' and job_id = '" & strJobID & "' order by outbox_num "
Set rs = Get_OracleRs(strSql)

If Not rs.EOF Then

    Do While Not rs.EOF
        i = rs!OUTBOX_NUM
        Call updateCID(strDN, i, strJobID)
        strSql2 = "select distinct inbox_num from PACKING_DETAILED where dn_num = '" & strDN & "' and job_id = '" & strJobID & "' and outbox_num = '" & i & "' order by inbox_num"
        Set Rs2 = Get_OracleRs(strSql2)

        If Not Rs2.EOF Then
            Do While Not Rs2.EOF
                j = Rs2!INBOX_NUM
                Call updateBID(strDN, i, j, strJobID)
                
                Rs2.MoveNext
            Loop
        
        End If

        Call updateQID(strDN, i)
        rs.MoveNext
    Loop

End If

Set rs = Nothing
Set Rs2 = Nothing

End Function

Public Sub updateCID(strDN As String, i As Integer, strJobID As String)
Dim strSql  As String
Dim strBase As String
Dim strseq  As String
Dim strKey  As String
Dim strCID  As String

strSql = "select distinct substr(trayid,1, InStr(trayid, '-') - 1) LOTID from PACKING_DETAILED where dn_num = '" & strDN & "' and job_id = '" & strJobID & "' "
strBase = Get_OracleStr(strSql) & "-C"
strSql = "select nvl(max(seq)+1, 1) from PKGIDSEQ_37 where val = '" & strBase & "' "
strseq = Get_OracleStr(strSql)
strCID = strBase & Right$("0" & strseq, 2)

strSql = "update PACKING_DETAILED set CARTONID = '" & strCID & "' where dn_num = '" & strDN & "' and outbox_num = '" & i & "' and job_id = '" & strJobID & "' "
AddSql (strSql)

strSql = "insert into PKGIDSEQ_37(VAL,SEQ,CREATE_DATE,DN) values('" & strBase & "', '" & strseq & "', sysdate, '" & strDN & "')"
AddSql (strSql)

End Sub

Public Sub updateBID(strDN As String, i As Integer, j As Integer, strJobID As String)
Dim strSql  As String
Dim strBase As String
Dim strseq  As String
Dim strKey  As String
Dim strBID  As String

strSql = "select distinct substr(trayid,1, InStr(trayid, '-') - 1) LOTID from PACKING_DETAILED where dn_num = '" & strDN & "' and job_id = '" & strJobID & "' "
strBase = Get_OracleStr(strSql) & "-B"
strSql = "select nvl(max(seq)+1, 1) from PKGIDSEQ_37 where val = '" & strBase & "' "
strseq = Get_OracleStr(strSql)
strBID = strBase & Right$("0" & strseq, 2)

strSql = "update PACKING_DETAILED set BOXID = '" & strBID & "' where dn_num = '" & strDN & "' and outbox_num = '" & i & "' and inbox_num = '" & j & "' and job_id = '" & strJobID & "' "
AddSql (strSql)

strSql = "insert into PKGIDSEQ_37(VAL,SEQ,CREATE_DATE,DN) values('" & strBase & "', '" & strseq & "', sysdate, '" & strDN & "')"
AddSql (strSql)

End Sub

Public Sub updateQID(strDN As String, i As Integer)
Dim strSql As String
Dim strQID As String
Dim strBID As String

strSql = "select BOXID from PACKING_DETAILED where dn_num = '" & strDN & "' and outbox_num = '" & i & "' and inbox_num = 1"
strBID = Get_OracleStr(strSql)

strSql = "select trglabelseq.QTSeq_NotMesQbox('" & strBID & "')  from dual"
strQID = Get_OracleStr(strSql)

strSql = "update PACKING_DETAILED set CARTON = '" & strQID & "' where dn_num = '" & strDN & "' and outbox_num = '" & i & "' "
AddSql (strSql)

End Sub

Public Sub updateDNStatus(strDN As String, strStatus As String)
Dim strSql As String

If strDN = "" Then Exit Sub

Select Case strStatus

    Case "new"
        strSql = "select * from PRINT_37FLAG where dn = '" & strDN & "'"
        If Get_OracleCnt(strSql) = 0 Then
            strSql = "insert into PRINT_37FLAG(dn, scaned,combined, printed) values('" & strDN & "','0','0','0')"
            AddSql (strSql)
        End If
        
    Case "scaned"
        strSql = "update PRINT_37FLAG set scaned = '1' where dn = '" & strDN & "'"
        AddSql (strSql)
    
    Case "combined"
        strSql = "update PRINT_37FLAG set combined = '1' where dn = '" & strDN & "'"
        AddSql (strSql)

    Case "printed"
        strSql = "update PRINT_37FLAG set printed = '1' where dn = '" & strDN & "'"
        AddSql (strSql)
    
End Select

End Sub

Public Function GetDevMark(sDev As String) As String

    Dim PartNOPre     As String

    Dim PRODUCTFAMILY As String

    Dim sCode         As String
            
    sCode = Left$(sDev, 2)
        
    Select Case sCode
        
        Case "RC"
            PartNOPre = "RCLAMP,{R}," & Mid$(sDev, 7)
            PRODUCTFAMILY = "RailClamp{R}"
        
        Case "UC"
            PartNOPre = "UCLAMP,{R}," & Mid$(sDev, 7)
            PRODUCTFAMILY = "MicroClamp{TM}"

        Case "EC"
            PartNOPre = "ECLAMP,{TM}," & Mid$(sDev, 7)
            PRODUCTFAMILY = "EMIClamp{TM}"

        Case "TC"
            PartNOPre = "TCLAMP,{TM}," & Mid$(sDev, 7)
            PRODUCTFAMILY = "TransClamp{TM}"

        Case "HC"
            PartNOPre = "HCLAMP,{TM}," & Mid$(sDev, 7)
            PRODUCTFAMILY = ""

        Case "PC"
            PartNOPre = "PCLAMP,{TM}," & Mid$(sDev, 7)
            PRODUCTFAMILY = ""

        Case "HS"
            PartNOPre = "HS"
            PRODUCTFAMILY = "HotSwitch{TM}"

    End Select
    
    GetDevMark = "," & PartNOPre & "," & PRODUCTFAMILY

End Function

Public Sub PrintINNERBOXFlag(strDN As String, iOp As Integer, iIp As Integer)

    Dim sContent  As String

    Dim sFileName As String
        
    sContent = "BOX_" & iOp & "_" & iIp
    sFileName = strDN & "-" & "FLAG_BOX_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")

    Call CreateTxt(sFileName, sContent, strFlagPath)
    
End Sub

Public Sub PrintJobFlag(strDN As String, strJob As String)

    Dim sContent  As String

    Dim sFileName As String
        
    sContent = "JOB_" & strJob
    sFileName = strDN & "-" & strJob & "-" & Format(Now(), "YYYYMMDDHHmmSS")

    Call CreateTxt(sFileName, sContent, strFlagPath)
    
End Sub

Public Sub PrintSTBoxLbl(sDN As String, iOp As Integer, iIp As Integer)

    Dim sOra          As String

    Dim sDatecode     As String

    Dim sTestDateCode As String

    Dim tSTBox        As STBox

    Dim sContent      As String

    Dim sFileName     As String

    Dim sPath         As String

    Dim sAdd          As String

    Dim rs            As New ADODB.Recordset

    sPath = ""
    sFileName = sDN & "-" & "STBoxLbl" & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
    sContent = ""

    sOra = "select job_id,CUSTOMER_DEVICE,boxid, sum(QTY) as qty from PACKING_DETAILED where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' and inbox_num =  '" & iIp & "'  group by job_id,CUSTOMER_DEVICE,boxid"
    Set rs = Get_OracleRs(sOra)

    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            tSTBox.JOB = Trim(rs!JOB_ID)
            tSTBox.DEV = Trim(rs!Customer_Device)
            tSTBox.lot = Trim(rs!BOXID)
            tSTBox.QTY = Trim(rs!QTY)
    
            sDatecode = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTBox.JOB & "'")
            tSTBox.DATECODE = sDatecode
        
            sTestDateCode = Get_SqlStr("select distinct HTDATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTBox.JOB & "'")
            tSTBox.testdateCode = sTestDateCode
        
            If tSTBox.DATECODE = "" Or tSTBox.testdateCode = "" Then
            
                tSTBox.DATECODE = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTBox.JOB & "'")
                tSTBox.testdateCode = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTBox.JOB & "'")

            End If
        
            sContent = sContent + tSTBox.DEV + "," + tSTBox.JOB + ",1T" + tSTBox.JOB + "," + tSTBox.DEV + "," + "1P" + tSTBox.DEV + "," + tSTBox.DATECODE + "," + tSTBox.DATECODE + "," + Mid(tSTBox.lot, 2) + "," + tSTBox.lot + ","
            sContent = sContent + tSTBox.QTY + ",Q" + tSTBox.QTY + "," + tSTBox.testdateCode + "," + tSTBox.testdateCode
        
            sAdd = ""
        
            sAdd = GetDevMark(tSTBox.DEV)
        
            sContent = sContent + sAdd + vbCrLf
        
            rs.MoveNext
        Loop

    End If

    Call CreateTxt(sFileName, sContent, str37BCIDPath)

End Sub

' 补打37内盒BID
Public Sub PrintSTBoxLbl2(strBID As String)
Dim sOra          As String
Dim sDatecode     As String
Dim sTestDateCode As String
Dim tSTBox        As STBox
Dim sContent      As String
Dim sFileName     As String
Dim sPath         As String
Dim sAdd          As String
Dim rs            As New ADODB.Recordset

sPath = ""
sFileName = "STBoxLbl-" & strBID & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select distinct job_id,CUSTOMER_DEVICE,boxid, sum(QTY) as qty from PACKING_DETAILED where BOXID = '" & strBID & "'  group by job_id,CUSTOMER_DEVICE,boxid"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tSTBox.JOB = Trim(rs!JOB_ID)
        tSTBox.DEV = Trim(rs!Customer_Device)
        tSTBox.lot = Trim(rs!BOXID)
        tSTBox.QTY = Trim(rs!QTY)
        sDatecode = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTBox.JOB & "'")
        tSTBox.DATECODE = sDatecode
        sTestDateCode = Get_SqlStr("select distinct HTDATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTBox.JOB & "'")
        tSTBox.testdateCode = sTestDateCode
        If tSTBox.DATECODE = "" Or tSTBox.testdateCode = "" Then
            tSTBox.DATECODE = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTBox.JOB & "'")
            tSTBox.testdateCode = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTBox.JOB & "'")

        End If

        sContent = sContent + tSTBox.DEV + "," + tSTBox.JOB + ",1T" + tSTBox.JOB + "," + tSTBox.DEV + "," + "1P" + tSTBox.DEV + "," + tSTBox.DATECODE + "," + tSTBox.DATECODE + "," + Mid(tSTBox.lot, 2) + "," + tSTBox.lot + ","
        sContent = sContent + tSTBox.QTY + ",Q" + tSTBox.QTY + "," + tSTBox.testdateCode + "," + tSTBox.testdateCode
        sAdd = ""
        sAdd = GetDevMark(tSTBox.DEV)
        sContent = sContent + sAdd + vbCrLf
        rs.MoveNext
    Loop
Else
    MsgBox "没有查询到该37内盒-B标签, 请扫描正确的37内盒-B标签", vbInformation, "提示"
    Exit Sub

End If

Call CreateTxt(sFileName, sContent, str37BCIDPath)
MsgBox "补打成功", vbInformation, "提示"

End Sub

Public Sub PrintREELFlag(sDN As String, iOp As Integer, iIp As Integer)

    Dim sContent  As String

    Dim sPath     As String

    Dim sFileName As String

    sPath = ""
    sContent = "REEL_" & iOp & "_" & iIp
    sFileName = sDN & "-" & "FLAG_REEL_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")

    Call CreateTxt(sFileName, sContent, strFlagPath)

End Sub

Public Sub PrintHWBoxLbl(sDN As String, iOp As Integer, iIp As Integer)

   Dim tHWBox    As HWBox

    Dim sContent  As String, strBarcode As String, strQrCode As String

    Dim sFileName As String

    Dim sPath     As String

    Dim sOra      As String

    Dim rs        As New ADODB.Recordset

    sPath = ""
    sFileName = sDN & "-" & "HWBoxLbl" & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
    sContent = ""

    sOra = "select job_id,mpn,cpn,datecode,sum(QTY) qty from LPSTBL where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' and inbox_num =  '" & iIp & "' group by job_id,mpn,cpn,datecode"
    Set rs = Get_OracleRs(sOra)

    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            tHWBox.CPN = UCase$(Trim$("" & rs!CPN))
            tHWBox.MPN = UCase$(Trim$("" & rs!MPN))
            tHWBox.lot = "" & rs!JOB_ID
            tHWBox.QTY = "" & rs!QTY
            tHWBox.PODATE = "" & rs!DATECODE
           
            strBarcode = tHWBox.CPN & "," & "" & "," & "" & "," & tHWBox.MPN & "," & tHWBox.PODATE & "," & tHWBox.lot & "," & tHWBox.QTY & ","
            strQrCode = "[)>" & Chr(30) & "06" & Chr(29) & "F01001P" & Chr(29) & "18VLEHWT" & Chr(29) & "F02010I" & Chr(29) & "1P" & tHWBox.CPN
            strQrCode = strQrCode & Chr(29) & "1V601024" & Chr(29) & "10D" & tHWBox.PODATE & Chr(29) & "1T" & tHWBox.lot & Chr(29) & "Q" & tHWBox.QTY & Chr(30) & Chr(4)
     
            sContent = sContent & strBarcode & strQrCode & vbCrLf
        
            rs.MoveNext
        Loop

    End If

    Call CreateTxt(sFileName, sContent, strHWBoxPath)

End Sub

' 补打华为内盒标签
Public Sub PrintHWBoxLbl2(strBID As String)
Dim tHWBox    As HWBox
Dim sContent  As String, strBarcode As String, strQrCode As String
Dim sFileName As String
Dim sPath     As String
Dim sOra      As String
Dim rs        As New ADODB.Recordset

sPath = ""
sFileName = "HWBoxLbl2" & "_" & strBID & "-" & Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select distinct job_id,mpn,cpn,datecode,sum(QTY) qty from LPSTBL where  boxid = '" & strBID & "' group by job_id,mpn,cpn,datecode"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tHWBox.CPN = UCase$(Trim$("" & rs!CPN))
        tHWBox.MPN = UCase$(Trim$("" & rs!MPN))
        tHWBox.lot = "" & rs!JOB_ID
        tHWBox.QTY = "" & rs!QTY
        tHWBox.PODATE = "" & rs!DATECODE
        strBarcode = tHWBox.CPN & "," & "" & "," & "" & "," & tHWBox.MPN & "," & tHWBox.PODATE & "," & tHWBox.lot & "," & tHWBox.QTY & ","
        strQrCode = "[)>" & Chr(30) & "06" & Chr(29) & "F01001P" & Chr(29) & "18VLEHWT" & Chr(29) & "F02010I" & Chr(29) & "1P" & tHWBox.CPN
        strQrCode = strQrCode & Chr(29) & "1V601024" & Chr(29) & "10D" & tHWBox.PODATE & Chr(29) & "1T" & tHWBox.lot & Chr(29) & "Q" & tHWBox.QTY & Chr(30) & Chr(4)
        sContent = sContent & strBarcode & strQrCode & vbCrLf
        rs.MoveNext
    Loop
Else
    MsgBox "查询不到该37内盒-B条码, 请扫描正确的37内盒-B条码", vbInformation, "提示"
    Exit Sub

End If

Call CreateTxt(sFileName, sContent, strHWBoxPath)
MsgBox "补打成功", vbInformation, "提示"

End Sub

Public Sub PrintHWReelLbl(sDN As String, iOp As Integer, iIp As Integer)

    Dim tHWBox    As HWBox

    Dim sContent  As String, strBarcode As String, strQrCode As String

    Dim sFileName As String

    Dim sPath     As String

    Dim sOra      As String

    Dim rs        As New ADODB.Recordset

    sPath = ""
    sFileName = sDN & "-" & "HWReelLbl" & "_" & iOp & "_" & iIp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
    sContent = ""

    sOra = "select job_id,mpn,cpn, QTY,datecode, reelid from LPSTBL where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' and inbox_num =  '" & iIp & "' order by seq"
    Set rs = Get_OracleRs(sOra)

    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            tHWBox.CPN = UCase$(Trim$("" & rs!CPN))
            tHWBox.MPN = UCase$(Trim$("" & rs!MPN))
            tHWBox.lot = "" & rs!JOB_ID
            tHWBox.QTY = "" & rs!QTY
            tHWBox.PODATE = "" & rs!DATECODE
            tHWBox.PSN = UCase$(Trim$("" & rs!REELID))
            
            strBarcode = tHWBox.CPN & "," & "" & "," & "" & "," & tHWBox.MPN & "," & tHWBox.PODATE & "," & tHWBox.lot & "," & tHWBox.QTY & "," & tHWBox.PSN & ","
            strQrCode = "[)>" & Chr(30) & "06" & Chr(29) & "F01001P" & Chr(29) & "52S" & tHWBox.PSN & Chr(29) & "18VLEHWT" & Chr(29) & "F02010I" & Chr(29) & "1P" & tHWBox.CPN
            strQrCode = strQrCode & Chr(29) & "1V601024" & Chr(29) & "10D" & tHWBox.PODATE & Chr(29) & "1T" & tHWBox.lot & Chr(29) & "Q" & tHWBox.QTY & Chr(30) & Chr(4)

            sContent = sContent & strBarcode & strQrCode & vbCrLf
        
            rs.MoveNext
        Loop

    End If

    Call CreateTxt(sFileName, sContent, strHWReelPath)

End Sub

' 补打华为卷盘
Public Sub PrintHWReelLbl2(strPSN As String)
Dim tHWBox    As HWBox
Dim sContent  As String, strBarcode As String, strQrCode As String
Dim sFileName As String
Dim sPath     As String
Dim sOra      As String
Dim rs        As New ADODB.Recordset
Dim tD        As tSTData

sPath = ""
sFileName = "HWReelLbl2" & "_" & strPSN & "-" & Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select distinct trayid,job_id,mpn,cpn, QTY,datecode,dn_num from LPSTBL where reelid = '" & strPSN & "'"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tHWBox.CPN = UCase$(Trim$("" & rs!CPN))
        tHWBox.MPN = UCase$(Trim$("" & rs!MPN))
        tHWBox.lot = "" & rs!JOB_ID
        tHWBox.QTY = "" & rs!QTY
        tHWBox.PODATE = "" & rs!DATECODE
        tD.JOB_ID = "" & rs!JOB_ID
        tD.DN_NUM = "" & rs!DN_NUM
        tD.TRAYID = "" & rs!TRAYID
       ' tHWBox.PSN = get37_PSN(tD)
        
        tHWBox.PSN = strPSN
        strBarcode = tHWBox.CPN & "," & "" & "," & "" & "," & tHWBox.MPN & "," & tHWBox.PODATE & "," & tHWBox.lot & "," & tHWBox.QTY & "," & tHWBox.PSN & ","
        strQrCode = "[)>" & Chr(30) & "06" & Chr(29) & "F01001P" & Chr(29) & "52S" & tHWBox.PSN & Chr(29) & "18VLEHWT" & Chr(29) & "F02010I" & Chr(29) & "1P" & tHWBox.CPN
        strQrCode = strQrCode & Chr(29) & "1V601024" & Chr(29) & "10D" & tHWBox.PODATE & Chr(29) & "1T" & tHWBox.lot & Chr(29) & "Q" & tHWBox.QTY & Chr(30) & Chr(4)
        sContent = sContent & strBarcode & strQrCode & vbCrLf
        ' 更新新的PSN
        Dim strSql As String

        strSql = "update packing_detailed set reelid = '" & tHWBox.PSN & "' where reelid = '" & strPSN & "' "
        AddSql (strSql)
        rs.MoveNext
    Loop
Else
    MsgBox "查询不到该华为PSN, 请确认是否扫描正确", vbInformation, "提示"
    Exit Sub

End If

Call CreateTxt(sFileName, sContent, strHWReelPath)
MsgBox "补打成功", vbInformation, "提示"

End Sub

Public Sub PrintOuterCFlag(sDN As String, iOp As Integer)

    Dim sContent  As String

    Dim sPath     As String

    Dim sFileName As String

    sPath = ""
    sContent = "CARTON_" & iOp
    sFileName = sDN & "-" & "FLAG_CARTON_" & iOp & "-" & Format(Now(), "YYYYMMDDHHmmSS")

    Call CreateTxt(sFileName, sContent, strFlagPath)

End Sub

Public Sub PrintSTCartonLbl(sDN As String, iOp As Integer)

    Dim sOra          As String

    Dim tSTCarton     As STCarton

    Dim sContent      As String

    Dim sFileName     As String

    Dim sDatecode     As String

    Dim sTestDateCode As String

    Dim rs            As New ADODB.Recordset

    Dim sPath         As String

    Dim sAdd          As String

    sPath = ""
    sFileName = sDN & "-" & "STCARTONLBL" & "_" & iOp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
    sContent = ""

    sOra = " select job_id,CUSTOMER_DEVICE,cartonid, sum(qty) as qty from PACKING_DETAILED where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "' group by job_id,CUSTOMER_DEVICE,cartonid"
    Set rs = Get_OracleRs(sOra)

    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            tSTCarton.JOB = Trim("" & rs!JOB_ID)
            tSTCarton.DEV = Trim$("" & rs!Customer_Device)
            tSTCarton.lot = Trim("" & rs!CARTONID)
            tSTCarton.QTY = Trim("" & rs!QTY)
        
            sDatecode = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTCarton.JOB & "'")
            tSTCarton.DATECODE = sDatecode
        
            sTestDateCode = Get_SqlStr("select distinct HTDATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTCarton.JOB & "'")
            tSTCarton.testdateCode = sTestDateCode

            If tSTCarton.DATECODE = "" Or tSTCarton.testdateCode = "" Then
            
                tSTCarton.DATECODE = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTCarton.JOB & "'")
                tSTCarton.testdateCode = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTCarton.JOB & "'")

            End If

            sContent = sContent + tSTCarton.DEV + "," + tSTCarton.JOB + ",1T" + tSTCarton.JOB + "," + tSTCarton.DEV + "," + "1P" + tSTCarton.DEV + "," + tSTCarton.DATECODE + "," + tSTCarton.DATECODE + "," + Mid(tSTCarton.lot, 2) + "," + tSTCarton.lot + ","
            sContent = sContent + tSTCarton.QTY + ",Q" + tSTCarton.QTY + "," + tSTCarton.testdateCode + "," + tSTCarton.testdateCode
        
            sAdd = GetDevMark(tSTCarton.DEV)
            sContent = sContent + sAdd + vbCrLf
        
            rs.MoveNext
        Loop

    End If

    Call CreateTxt(sFileName, sContent, str37BCIDPath)

End Sub

' 补打外箱C标签
Public Sub PrintSTCartonLbl2(strCID As String)
Dim sOra          As String
Dim tSTCarton     As STCarton
Dim sContent      As String
Dim sFileName     As String
Dim sDatecode     As String
Dim sTestDateCode As String
Dim rs            As New ADODB.Recordset
Dim sPath         As String
Dim sAdd          As String

sPath = ""
sFileName = "STCARTONLBL-" & strCID & "_" & iOp & "-" & Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = " select distinct job_id,CUSTOMER_DEVICE,cartonid, sum(qty) as qty from PACKING_DETAILED where CARTONID = '" & strCID & "' group by job_id,CUSTOMER_DEVICE,cartonid"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tSTCarton.JOB = Trim(rs!JOB_ID)
        tSTCarton.DEV = Trim$(rs!Customer_Device)
        tSTCarton.lot = Trim(rs!CARTONID)
        tSTCarton.QTY = Trim(rs!QTY)
        sDatecode = Get_SqlStr("select distinct PODATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTCarton.JOB & "'")
        tSTCarton.DATECODE = sDatecode
        sTestDateCode = Get_SqlStr("select distinct HTDATECODE from [erpdata].[dbo].[TblTSV_Tray_details] where  CUSTOMERLOTID = '" & tSTCarton.JOB & "'")
        tSTCarton.testdateCode = sTestDateCode
        If tSTCarton.DATECODE = "" Or tSTCarton.testdateCode = "" Then
            tSTCarton.DATECODE = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTCarton.JOB & "'")
            tSTCarton.testdateCode = Get_OracleStr("select datecode from PACKING_DETAILED where JOB_ID = '" & tSTCarton.JOB & "'")

        End If

        sContent = sContent + tSTCarton.DEV + "," + tSTCarton.JOB + ",1T" + tSTCarton.JOB + "," + tSTCarton.DEV + "," + "1P" + tSTCarton.DEV + "," + tSTCarton.DATECODE + "," + tSTCarton.DATECODE + "," + Mid(tSTCarton.lot, 2) + "," + tSTCarton.lot + ","
        sContent = sContent + tSTCarton.QTY + ",Q" + tSTCarton.QTY + "," + tSTCarton.testdateCode + "," + tSTCarton.testdateCode
        sAdd = GetDevMark(tSTCarton.DEV)
        sContent = sContent + sAdd + vbCrLf
        rs.MoveNext
    Loop
Else
    MsgBox "没有找打该37外箱-C标签, 请扫描正确的37外箱-C标签", vbInformation, "提示"
    Exit Sub

End If

Call CreateTxt(sFileName, sContent, str37BCIDPath)
MsgBox "补打成功", vbInformation, "提示"

End Sub

Public Sub PrintHTCartonLbl(sDN As String, iOp As Integer)

    Dim sCartonNo As String

    Dim sOra      As String

    Dim sFileName As String

    Dim sContent  As String

    Dim sPath     As String

    sPath = ""
    sFileName = sDN & "-" & "STCARTONLBL" + Format(Now(), "YYYYMMDDHHmmSS")
    sContent = ""

    sOra = "select distinct carton from PACKING_DETAILED where dn_num = '" & sDN & "' and outbox_num = '" & iOp & "'"
    sCartonNo = Get_OracleStr(sOra)

    sFileName = "HTCARTONLBL" & Format(Now(), "YYYYMMDDHHmmSS")
    sContent = sCartonNo

    Call CreateTxt(sFileName, sContent, strHTQCartonPath)

End Sub

Public Sub UpdatePrintStatus(strDN As String, iOp As Integer)

    Dim strSql As String

    strSql = "update PACKING_DETAILED set print_flag = '1' where dn_num = '" & strDN & "' and outbox_num = '" & iOp & "'"
    AddSql (strSql)

End Sub

Public Sub CreateTxt(filename As String, msgTxt As String, dirtemp As String)

    Dim fileNameTemp As String

    Dim dirNameTemp  As String

    Dim fileTemp     As String

    dirNameTemp = dirtemp
    fileNameTemp = Replace(filename, "'", "") & ".txt"
    fileTemp = dirNameTemp & fileNameTemp
    
    Open fileTemp For Output As #1
    Print #1, msgTxt
    Close #1

'    Sleep (1000)

End Sub

Public Function CopyFileToFtp(strSource As String, strDestination As String)
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
fs.CopyFile strSource, strDestination

End Function

Public Sub PrintCusCartonLbl(sDN As String, iOp As Integer)

    Dim sOra       As String

    Dim tCusCARTON As CUSCARTON

    Dim sFileName  As String

    Dim sContent   As String

    Dim rs         As New ADODB.Recordset

    Dim sPath      As String

    Dim KID        As String
    
    Dim sMaxOP As String
    
    sMaxOP = Get_OracleStr("select max(outbox_num) from PACKING_DETAILED where dn_num = '" & sDN & "'")
    
    sPath = ""
    sFileName = sDN & "-" & "CUSCARTONLBL" + Format(Now(), "YYYYMMDDHHmmSS")
    sContent = ""

    sOra = "select a.kid,a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & txtDN.text & "'" & "and b.delivery = '" & txtDN.text & "' and a.job_id = b.batchnumber and a.outbox_num = '" & iOp & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno, a.kid"
    
    Set rs = Get_OracleRs(sOra)

    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            tCusCARTON.dn = txtDN.text
            tCusCARTON.PO = rs!PO
            tCusCARTON.CPN = rs!CustomerPartnumber
            tCusCARTON.MPN = rs!Customer_Device
            tCusCARTON.QTY = rs!QTY
            KID = rs!KID
            sContent = sContent & tCusCARTON.dn & ",I" & tCusCARTON.dn & "," & tCusCARTON.PO & ",K" & tCusCARTON.PO & ",E2," & tCusCARTON.CPN & ",P" & tCusCARTON.CPN & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & ","
            sContent = sContent & tCusCARTON.QTY & ",Q" & tCusCARTON.QTY & ","
            sContent = sContent & Get_OracleStr("select distinct trim(a.freightforwarder)|| ',CHINA,' || substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3)||','||trim(a.city) || ' ' || trim(a.state) || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ','|| 'Attn:;Tel:' || trim(a.phone) || ','  from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & txtDN.text & "'")
            sContent = sContent & "N/A,PN/A,N/A,9DN/A," & iOp & "," & KID & "," & sMaxOP
        
            rs.MoveNext
        Loop

    End If

    Call CreateTxt(sFileName, sContent, strSSCartonPath)
End Sub

Public Sub PrintSTCartonStanderLbl(sDN As String, iOp As Integer)

    Dim sOra       As String

    Dim tCusCARTON As CUSCARTON

    Dim sFileName  As String

    Dim sContent   As String

    Dim rs         As New ADODB.Recordset

    Dim sPath      As String

    Dim sAdd       As String

    Dim sKid       As String
    
    Dim sMaxOP As String
    
    sMaxOP = Get_OracleStr("select max(outbox_num) from PACKING_DETAILED where dn_num = '" & sDN & "'")

    sPath = ""
    
    sFileName = sDN & "-" & "SemTechStanderCarton" + Format(Now(), "YYYYMMDDHHmmSS")
    sContent = ""

    sOra = "select a.CUSTOMER_DEVICE,a.kid, b.customerpartnumber,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & sDN & "'" & "and b.delivery = '" & sDN & "' and a.job_id = b.batchnumber and a.outbox_num = '" & iOp & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno,a.kid"
    
    Set rs = Get_OracleRs(sOra)

    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            tCusCARTON.dn = sDN
            tCusCARTON.PO = UCase(IIf(IsNull(rs!PO), "N/A", rs!PO))
            tCusCARTON.CPN = UCase(IIf(IsNull(rs!CustomerPartnumber), "N/A", rs!CustomerPartnumber))
            tCusCARTON.MPN = UCase(IIf(IsNull(rs!Customer_Device), "N/A", rs!Customer_Device))
            tCusCARTON.QTY = rs!QTY
            sKid = rs!KID
        
            sContent = sContent & Get_OracleStr("select distinct substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3) || ','||trim(a.city) || ' ' || trim(a.state)  || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ',' || trim(a.contactname) || ',' || trim(a.phone) from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & sDN & "' ") & ","
            sContent = sContent & tCusCARTON.dn & ",I" & tCusCARTON.dn & "," & tCusCARTON.PO & ",K" & tCusCARTON.PO & "," & tCusCARTON.CPN & ",P" & tCusCARTON.CPN & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & "," & tCusCARTON.QTY & ",Q" & tCusCARTON.QTY & "," & Get_OracleStr("select distinct freightforwarder from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & sDN & "'") & "," & "" & "," & "" & "," & "" & "," & "COO:CHINA" & "," & "CHINA"

            sAdd = "," & iOp & "," & sKid

            sContent = sContent & sAdd & "," & sMaxOP
            rs.MoveNext
        Loop

    End If

    Call CreateTxt(sFileName, sContent, str37CartonPath)
End Sub

Public Sub PrintSTCartonStanderLbl_ONELOT(sDN As String, iOp As Integer, iOpMax As Integer)

    Dim sOra       As String

    Dim tCusCARTON As CUSCARTON

    Dim sFileName  As String

    Dim sContent   As String

    Dim rs         As New ADODB.Recordset

    Dim sPath      As String

    Dim sAdd       As String

    Dim sKid       As String
    
    Dim sMaxOP As String
    
    sMaxOP = iOpMax

    sPath = ""
    
    sFileName = sDN & "-" & "SemTechStanderCarton" + Format(Now(), "YYYYMMDDHHmmSS")
    sContent = ""

    sOra = "select a.CUSTOMER_DEVICE,a.kid, b.customerpartnumber,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & sDN & "'" & "and b.delivery = '" & sDN & "' and a.job_id = b.batchnumber and a.outbox_num = '" & iOp & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno,a.kid"
    
    Set rs = Get_OracleRs(sOra)

    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            tCusCARTON.dn = sDN
            tCusCARTON.PO = UCase(IIf(IsNull(rs!PO), "N/A", rs!PO))
            tCusCARTON.CPN = UCase(IIf(IsNull(rs!CustomerPartnumber), "N/A", rs!CustomerPartnumber))
            tCusCARTON.MPN = UCase(IIf(IsNull(rs!Customer_Device), "N/A", rs!Customer_Device))
            tCusCARTON.QTY = rs!QTY
            sKid = rs!KID
        
            sContent = sContent & Get_OracleStr("select distinct substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3) || ','||trim(a.city) || ' ' || trim(a.state)  || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ',' || trim(a.contactname) || ',' || trim(a.phone) from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & sDN & "' ") & ","
            sContent = sContent & tCusCARTON.dn & ",I" & tCusCARTON.dn & "," & tCusCARTON.PO & ",K" & tCusCARTON.PO & "," & tCusCARTON.CPN & ",P" & tCusCARTON.CPN & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & "," & tCusCARTON.QTY & ",Q" & tCusCARTON.QTY & "," & Get_OracleStr("select distinct freightforwarder from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & sDN & "'") & "," & "" & "," & "" & "," & "" & "," & "COO:CHINA" & "," & "CHINA"

            sAdd = "," & iOp & "," & sKid

            sContent = sContent & sAdd & "," & sMaxOP
            rs.MoveNext
        Loop

    End If

    Call CreateTxt(sFileName, sContent, str37CartonPath)
End Sub

' 补打SEMTECH大标签
Public Sub PrintSTCartonStanderLbl2(strKey As String)
Dim sOra       As String
Dim tCusCARTON As CUSCARTON
Dim sFileName  As String
Dim sContent   As String
Dim rs         As New ADODB.Recordset
Dim sPath      As String
Dim sAdd       As String
Dim sKid       As String
Dim sMaxOP     As String
Dim sDN        As String
Dim iOp As Integer

sDN = Get_OracleStr("select distinct dn_num from PACKING_DETAILED where carton = '" & strKey & "'")
iOp = Get_OracleStr("select distinct outbox_num from PACKING_DETAILED where carton = '" & strKey & "'")
sMaxOP = Get_OracleStr("select max(outbox_num) from PACKING_DETAILED where dn_num = '" & sDN & "' ")
sPath = ""
sFileName = sDN & "-" & strKey & "-SemTechStanderCarton" + Format(Now(), "YYYYMMDDHHmmSS")
sContent = ""
sOra = "select a.CUSTOMER_DEVICE,a.kid, b.customerpartnumber,b.purchasingdocno as po, sum(a.qty) as qty from PACKING_DETAILED a, CUSTOMERSHIPPINGUPTBL b where a.dn_num = '" & sDN & "'" & "and b.delivery = '" & sDN & "' and a.job_id = b.batchnumber and a.carton = '" & strKey & "' group by a.CUSTOMER_DEVICE, b.customerpartnumber,b.purchasingdocno,a.kid"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tCusCARTON.dn = sDN
        tCusCARTON.PO = UCase(IIf(IsNull(rs!PO), "N/A", rs!PO))
        tCusCARTON.CPN = UCase(IIf(IsNull(rs!CustomerPartnumber), "N/A", rs!CustomerPartnumber))
        tCusCARTON.MPN = UCase(IIf(IsNull(rs!Customer_Device), "N/A", rs!Customer_Device))
        tCusCARTON.QTY = rs!QTY
        sKid = rs!KID
        sContent = sContent & Get_OracleStr("select distinct substr(trim(a.shiptoname), 0, 33) || ',' || trim(a.shiptostreet1) || ',' || trim(a.shiptostreet2) || ',' || trim(a.shiptostreet3) || ','||trim(a.city) || ' ' || trim(a.state)  || ' ' || trim(a.postalcode) || ',' || trim(a.countrykey) || ',' || trim(a.contactname) || ',' || trim(a.phone) from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & sDN & "' ") & ","
        sContent = sContent & tCusCARTON.dn & ",I" & tCusCARTON.dn & "," & tCusCARTON.PO & ",K" & tCusCARTON.PO & "," & tCusCARTON.CPN & ",P" & tCusCARTON.CPN & "," & tCusCARTON.MPN & ",Z" & tCusCARTON.MPN & "," & tCusCARTON.QTY & ",Q" & tCusCARTON.QTY & "," & Get_OracleStr("select distinct freightforwarder from CUSTOMERSHIPPINGUPTBL a where a.delivery = '" & sDN & "'") & "," & "" & "," & "" & "," & "" & "," & "COO:CHINA" & "," & "CHINA"
        sAdd = "," & iOp & "," & sKid
        sContent = sContent & sAdd & "," & sMaxOP
        rs.MoveNext
    Loop
Else
    MsgBox "找不到该大箱号条码" & strKey & " 请确认是否正确扫码", vbInformation, "提示"
    Exit Sub

End If

Call CreateTxt(sFileName, sContent, str37CartonPath)
MsgBox "华为外箱大标签补打成功", vbInformation, "提示"

End Sub

Public Sub setPrintPath()

' 标志
strFlagPath = "\\10.160.1.84\public\BarCode\37\37Flag\"

' 37 BID, CID小标签
str37BCIDPath = "\\10.160.1.84\public\BarCode\37\37内箱\"        ' 37B,C,R小标签
str37CartonPath = "\\10.160.1.84\public\BarCode\37\37外箱\"      ' 37自家外箱大标签

' 出三星标签
strSSBoxPath = "\\10.160.1.84\public\BarCode\37\37BoxNH\"      ' 三星内盒小标签
strSSReelPath = "\\10.160.1.84\public\BarCode\37\37BoxJP\"     ' 三星卷盘小标签
strSSCartonPath = "\\10.160.1.84\public\BarCode\37\37BoxOut\"  ' 三星外箱大标签

' 华天Q标签
strHTQCartonPath = "\\10.160.1.84\public\BarCode\37\37Box\"    ' 华天Q箱号小标签

' 出华为
strHWBoxPath = "\\10.160.1.84\public\BarCode\37\37HW\HW内盒\"
strHWReelPath = "\\10.160.1.84\public\BarCode\37\37HW\HW卷盘\"

End Sub

Public Sub setTestPrintPath()

' 标志
strFlagPath = "C:\test\"

' 37 BID, CID小标签
str37BCIDPath = "C:\test\"      ' 37B,C,R小标签
str37CartonPath = "C:\test\"     ' 37自家外箱大标签

' 出三星标签
strSSBoxPath = "C:\test\"      ' 三星内盒小标签
strSSReelPath = "C:\test\"    ' 三星卷盘小标签
strSSCartonPath = "C:\test\" ' 三星外箱大标签

' 华天Q标签
strHTQCartonPath = "C:\test\"   ' 华天Q箱号小标签

' 出华为
strHWBoxPath = "C:\test\"
strHWReelPath = "C:\test\"

End Sub

Public Sub DelToErp(strDN As String)
Dim strSql As String
Dim rs     As ADODB.Recordset
Dim strCartonID As String

On Error GoTo ERRON

INIadoCon.BeginTrans

strSql = "select distinct CARTON from PACKING_DETAILED where dn_num = '" & strDN & "' "
Set rs = Get_OracleRs(strSql)

If rs.EOF Then
    MsgBox "PACKING_DETAILED查询不到该DN", vbInformation, "提示"
    INIadoCon.RollbackTrans
    Exit Sub
End If

rs.MoveFirst

Do While Not rs.EOF
    strCartonID = Trim$("" & rs(0))
    
    strSql = "delete from [erpdata].[dbo].[tblPackTreeInf] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strSql)
    
    strSql = "delete from [erpdata].[dbo].[tblPackMainInf] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strSql)
 
    strSql = "update [erpdata].[dbo].[tblPackTreeInf] set 上级序号 = '', Memo = '' where 箱号 in (select trayid from erpbase..PACKING_DETAILED where carton = '" & strCartonID & "')  "
    AddSql2 (strSql)
 
    strSql = "delete from [erpdata].[dbo].[tblStockNumTree] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strSql)
 
    strSql = "Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='',Memo='', dn='' where 箱号 in (select trayid from erpbase..PACKING_DETAILED where carton = '" & strCartonID & "') "
    AddSql2 (strSql)

    rs.MoveNext
Loop

INIadoCon.CommitTrans

MsgBox "ERP箱号已删除,对应关系已解除", vbInformation, "提示"

Exit Sub

ERRON:
    INIadoCon.RollbackTrans
    MsgBox "错误:" & Err.DESCRIPTION, vbCritical, "警告"
    
End Sub

Public Sub TransToErp(strDN As String)

Dim strSql As String
Dim rs     As ADODB.Recordset
Dim strCartonID As String, strCartonQty As String
Dim ID As String
On Error GoTo ERRON
INIadoCon.BeginTrans

strSql = "select CARTON, SUM(QTY) from PACKING_DETAILED where dn_num = '" & strDN & "' group by CARTON"
Set rs = Get_OracleRs(strSql)

If rs.EOF Then
    MsgBox "PACKING_DETAILED查询不到该DN", vbInformation, "提示"
    INIadoCon.RollbackTrans
    Exit Sub
End If

rs.MoveFirst

Do While Not rs.EOF
    strCartonID = Trim$("" & rs(0))
    strCartonQty = Trim$("" & rs(1))
    ' ---------------------------------------------------删除
    '0
    strSql = "delete from [erpdata].[dbo].[tblPackTreeInf] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strSql)
    
    strSql = "delete from [erpdata].[dbo].[tblPackMainInf] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strSql)
 
    strSql = "update [erpdata].[dbo].[tblPackTreeInf] set 上级序号 = '', Memo = '' where 箱号 in (select trayid from erpbase..PACKING_DETAILED where carton = '" & strCartonID & "')  "
    AddSql2 (strSql)
 
    strSql = "delete from [erpdata].[dbo].[tblStockNumTree] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strSql)
 
    strSql = "Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='',Memo='', dn='' where 箱号 in (select trayid from erpbase..PACKING_DETAILED where carton = '" & strCartonID & "') "
    AddSql2 (strSql)
 
    ' --------------------------------------------------更新
    '1 insert [erpdata].[dbo].[tblPackMainInf]
    strSql = "insert into [erpdata].[dbo].[tblPackMainInf](箱号,客户代码,数量,合格标记,装箱标记,产线标记) values('" & strCartonID & "','37'," & strCartonQty & ",'0','1','1')"
    If AddSql2(strSql) = 0 Then
        MsgBox "1 insert [erpdata].[dbo].[tblPackMainInf]:failed!!! ", vbCritical, "提示"
        Exit Sub
    End If
    
    '2 insert - update [erpdata].[dbo].[tblPackTreeInf]
    strSql = "insert into [erpdata].[dbo].[tblPackTreeInf](箱号,上级序号,基层标记 ,Memo) values('" & strCartonID & "',0,1,'37')"
    If AddSql2(strSql) = 0 Then
        MsgBox "2 insert [erpdata].[dbo].[tblPackTreeInf]:failed!!!", vbCritical, "提示"
        Exit Sub
    End If
    
    ID = Get_SqlserverNo("select 序号 as ID from [erpdata].[dbo].[tblPackTreeInf] a where a.箱号='" & strCartonID & "' and Memo='37' ")
    
    strSql = "Update [erpdata].[dbo].[tblPackTreeInf] set 上级序号='" & ID & "',Memo='37' " & " where 箱号 in ( select trayid from  OPENQUERY(ORACLEDB, 'SELECT * from packing_detailed where carton = ''" & strCartonID & "'' ')) "
    If AddSql2(strSql) = 0 Then
        MsgBox "2 update [erpdata].[dbo].[tblPackTreeInf]:failed!!!", vbCritical, "提示"
        Exit Sub
    End If
    
    '3 insert - update [erpdata].[dbo].[tblStockNumTree]
    strSql = "insert into [erpdata].[dbo].[tblStockNumTree](序号,箱号,上级序号,基层标记 ,尺寸,重量,Memo,DN) values( " & ID & ",'" & strCartonID & "',0,1,'','','37','" & strDN & "')"
    If AddSql2(strSql) = 0 Then
        MsgBox "3 insert [erpdata].[dbo].[tblStockNumTree]", vbCritical, "提示"
        Exit Sub
    End If
    
    strSql = "Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='" & ID & "',Memo='37', dn='" & strDN & "' " & " where 箱号 in ( select trayid from  OPENQUERY(ORACLEDB, 'SELECT * from packing_detailed where carton = ''" & strCartonID & "'' ')) "
    If AddSql2(strSql) = 0 Then
        MsgBox "3 update [erpdata].[dbo].[tblStockNumTree]", vbCritical, "提示"
        Exit Sub
    End If
    
    rs.MoveNext
Loop

INIadoCon.CommitTrans

MsgBox "DN:" & strDN & "  :箱号已更新", vbInformation, "提示"

Exit Sub

ERRON:
    INIadoCon.RollbackTrans
    MsgBox "错误:" & Err.DESCRIPTION, vbCritical, "警告"
    
End Sub

Public Sub TransToErp_ONELOT(strDN As String, sCartonID As String, lQty As Long)

    Dim sqlTemp As String

    Dim idTemp  As String

    Dim boxdn   As String

    sqlTemp = "insert into [erpdata].[dbo].[tblPackMainInf](箱号,客户代码,数量,合格标记,装箱标记,产线标记) values('" & sCartonID & "','37'," & lQty & ",'0','1','1') "

    AddSql2 (sqlTemp)

    '插入Sqlserver   tblPackTreeInf
    sqlTemp = "insert into [erpdata].[dbo].[tblPackTreeInf](箱号,上级序号,基层标记 ,Memo) values('" & sCartonID & "',0,1,'37')"
    AddSql2 (sqlTemp)

    '再更新小箱的上级序号

    '把序号先查出来，再整体更新

    idTemp = Get37BigQboxIDV1(sCartonID)
    boxdn = strDN

    sqlTemp = "insert into [erpdata].[dbo].[tblStockNumTree](序号,箱号,上级序号,基层标记 ,尺寸,重量,Memo,DN) values( " & idTemp & ",'" & sCartonID & "',0,1,'','','37','" & boxdn & "')"
    AddSql2 (sqlTemp)

    sqlTemp = "Update [erpdata].[dbo].[tblPackTreeInf] set 上级序号='" & idTemp & "',Memo='37' " & " where 箱号 in ( select trayid from (select * from OPENQUERY(ORACLEDB, 'SELECT * from lpstbl' )) X where X.carton = '" & sCartonID & "') "
    AddSql2 (sqlTemp)

    sqlTemp = "Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='" & idTemp & "',Memo='37', dn='" & boxdn & "' " & " where 箱号 in (select trayid from (select * from OPENQUERY(ORACLEDB, 'SELECT * from lpstbl' )) X where X.carton = '" & sCartonID & "') "
    AddSql2 (sqlTemp)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       Get37TestDC
' Description:       获取37TESTDC
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/9/3-14:02:39
'
' Parameters :       strJobID (String)
'--------------------------------------------------------------------------------
Public Function Get37TestDC(strDN As String, strJobID As String) As String
Dim strSql As String
Dim strDC As String

strSql = "select distinct date_code from customershippinguptbl where delivery = '" & strDN & "' and batchnumber ='" & strJobID & "'"
strDC = Get_OracleStr(strSql)

If strDC = "" Then
    
    MsgBox "DC获取异常,请联系IT", vbInformation, "提示"

Else
    Get37TestDC = strDC
End If

'
'strJobID = Replace$(strJobID, "M", "")
'
'Do While Right$(strJobID, 1) = "R"
'    strJobID = Left(strJobID, Len(strJobID) - 1)
'Loop
'Get37TestDC = Get_OracleStr("select DC from TBL37TESTDC where JOBID = '" & strJobID & "'")
'If Get37TestDC = "" Then
'    MsgBox "DC获取异常,请联系IT", vbInformation, "提示"
'
'End If

End Function

Public Function chkPOPrice(strWaferID As String) As Boolean
Dim strSql As String
Dim strPONO As String
Dim strCusDev As String
Dim rs As New ADODB.Recordset

chkPOPrice = False
strSql = "select distinct b.po_num,b.mpn_desc from mappingdatatest a inner join customeroitbl_test b on to_char(b.id) = a.filename " & _
" and a.lotid = b.source_batch_id where  a.substrateid = '" & strWaferID & "' "

Set rs = Get_OracleRs(strSql)

strPONO = Trim("" & rs(0))
strCusDev = Trim("" & rs(1))

If strPONO = "" Then
    MsgBox "WAFERID: " & strWaferID & vbCrLf & "没有维护客户PO号" & vbCrLf & "请联系市场维护PO号,否则无法开立工单", vbCritical, "警告"
    Exit Function
End If

'strSql = "select * from TSV_MD_POPrice where PO_NUM = '" & strPONO & "' and PT = '" & strCusDev & "'"
'If Get_OracleCnt(strSql) = 0 Then
'    MsgBox "WAFERID: " & strWaferID & vbCrLf & "没有维护客户PO价格" & vbCrLf & "请联系市场维护PO价格,否则无法开立工单", vbCritical, "警告"
'    Exit Function
'End If

chkPOPrice = True
End Function

Public Function GetIni(appName As String, keyName As String) As String
Dim strDefault As String
Dim lngBuffLen As Long
Dim strResu    As String
Dim x          As Long
Dim strIniFile As String

'If Right(App.Path, 1) = "\" Then
'    strIniFile = App.Path & mc_strIniFileName
'Else
'    strIniFile = App.Path & "\" & mc_strIniFileName
'
'End If

strIniFile = "\\10.160.1.84\open\FileServer\" & mc_strIniFileName

strResu = String(1025, vbNullChar): lngBuffLen = 1025
strDefault = ""
x = GetPrivateProfileString(appName, keyName, strDefault, strResu, lngBuffLen, strIniFile)

GetIni = Left(strResu, x)

End Function

Public Sub WriteIni(appName As String, keyName As String, valueNew As String)
Dim x          As Long
Dim strIniFile As String

If Right(App.Path, 1) = "\" Then
    strIniFile = App.Path & mc_strIniFileName
Else
    strIniFile = App.Path & "\" & mc_strIniFileName

End If

x = WritePrivateProfileString(appName, keyName, valueNew, strIniFile)
Debug.Print x

End Sub

